# Daemon PR — TODO

## Correctness

### 3. `statSync` blocks event loop in cache `get()`

Every cache lookup calls `statSync()` for mtime validation. Under queued load this blocks the Node event loop synchronously.

**Fix**: Make `get()`, `getMtime()`, and `updateMtime()` async using `fs.promises.stat()`. Callers in `context.ts` are already async — just add `await`. For `put()`, accept the mtime as a parameter instead of re-stat'ing (the caller just opened/saved the file and already has it).

**Files**: `src/cli/file-cache.ts`, `src/cli/context.ts`

## Design

### 4. No memory visibility or limits

With `maxSize = 10`, large workbooks can consume ~1GB with no feedback. Users have no way to see memory usage or tune the cache size.

**Fix**: Expose `heapUsedMB` and `rssMB` (from `process.memoryUsage()`) in the `daemon status` response. Support `HSX_CACHE_SIZE` env var in `startDaemon()` to configure `maxSize`. Expose `maxSize` as a public getter on `FileCache`.

**Files**: `src/cli/daemon.ts`, `src/cli/file-cache.ts`

### 5. LRU eviction is O(n) scan

`put()` iterates all entries to find the oldest `lastAccess`. Fine for `maxSize=10`, but unnecessarily complex.

**Fix**: Exploit `Map` insertion-order. On cache hit in `get()`, delete and re-insert the entry to move it to the end. On eviction in `put()`, just delete `this.cache.keys().next().value` (the first/oldest key). This removes the `lastAccess` field entirely and makes eviction O(1).

**Files**: `src/cli/file-cache.ts`

### 6. Document `HSX_SOCKET_PATH` for multi-daemon use

`~/.hsx-daemon.sock` means only one daemon instance globally. Different projects or license keys can't run separate daemons. The `HSX_SOCKET_PATH` env var already exists to override the socket path.

**Fix**: Document `HSX_SOCKET_PATH` in the README as the way to run multiple daemons.

**Files**: `README.md`

### 7. Newline-delimited JSON protocol

The protocol assumes `JSON.stringify()` never produces raw `\n`. True today, but an implicit contract.

**Fix (optional)**: Switch to length-prefix framing (`4-byte big-endian length + payload`). ~20 lines of framing code on each side. Only worth doing if binary payloads or very large outputs become a concern. Low priority.

**Files**: `src/cli/daemon.ts`, `src/cli/client.ts`

## Minor

### 8. `enqueue` swallows errors silently

`queue.then(fn, fn)` drops the previous rejection reason. `socket.write` failures are caught with `catch {}`. Makes debugging impossible.

**Fix**: Add a `daemonLog(msg)` helper that writes to `process.stderr` with a timestamp. Use it in the `enqueue` catch and the `socket.write` catch. Add a `.catch()` at the end of the queue chain for defense in depth.

**Files**: `src/cli/daemon.ts`

### 9. `daemon stop` uses `setTimeout(shutdown, 100)` hack

Races against slow socket writes — the response may not flush before shutdown.

**Fix**: Remove `setTimeout(shutdown, 100)` from `handleRequest`. Instead, return a `shutdown: true` flag on the response object. In the `enqueue` callback, use the `socket.write(data, callback)` callback to call `shutdown()` only after the bytes are on the wire.

**Files**: `src/cli/daemon.ts`

### 10. `spawnDaemon` parses stdout for readiness

If SpreadJS prints warnings to stdout during `init()`, the startup handshake breaks. The daemon already has `process.send()` IPC code that goes unused.

**Fix**: Change `spawn()` stdio to `["ignore", "ignore", "ignore", "ipc"]`. Listen for `child.on("message", ...)` instead of parsing stdout. Call `child.disconnect()` + `child.unref()` after receiving the ready message. Keep the `process.stdout.write()` in `daemon.ts` as a fallback for interactive `daemon start`.

**Files**: `src/cli/client.ts`, `src/cli/daemon.ts`

## Documentation

### 11. README concurrency section is overstated

The current text says "you cannot safely run multiple workbook operations concurrently within the same process." This is too broad — the real constraint is don't overlap `init()`/`dispose()` lifecycles. Within a single lifecycle, multiple `ExcelFile` instances work concurrently (the daemon itself does this). The shims are stable once installed.

**Fix**: Rewrite to clarify that multiple concurrent workbook operations are safe within a single `init()`/`dispose()` lifecycle. The constraint is only against overlapping lifecycles or calling `dispose()` while operations are in-flight. Recommend child processes (not worker threads) for fully isolated runtimes, since `canvas` native addon thread-safety varies by version.

**Files**: `README.md`

## Priority

| Priority | Item | Effort | Impact |
|----------|------|--------|--------|
| 1 | #9 Write callback | Trivial | 5-line fix, removes a race |
| 2 | #10 IPC channel | Small | Removes fragile stdout parsing |
| 3 | #3 Async stat | Small | Unblocks event loop |
| 4 | #5 Map-based LRU | Small | Simpler code, fewer fields |
| 5 | #8 Error logging | Trivial | Debuggability |
| 6 | #4 Memory stats | Trivial | Observability |
| 7 | #11 README rewrite | Trivial | Accuracy |
| 8 | #6 Document HSX_SOCKET_PATH | Trivial | Document existing feature |
| 9 | #7 Length-prefix | Medium | Only if protocol issues arise |

## Done

### 1. Global mutable state in `output.ts` / `context.ts`

Replaced module-level globals (`capturedStdout`, `capturedStderr`, `pendingStdin`, `daemonRuntime`) with `AsyncLocalStorage` scoped per-request. Deleted `startCapture`, `stopCapture`, `setStdin`, `setDaemonRuntime`. Each request is wrapped in `runWithIo` + `runWithDaemonRuntime`.

### 2. Race condition on daemon startup

Eliminated by switching from TCP + port file to Unix domain sockets. `server.listen(socketPath)` is atomic — a second daemon probes the socket, sees the first is alive, and exits. No orphaned daemons. The port file and PID-based liveness check were deleted entirely.

### Transport: TCP → Unix domain sockets

Replaced `127.0.0.1` TCP + `~/.hsx-daemon.json` port file with Unix domain sockets (`~/.hsx-daemon.sock`) / named pipes (`\\.\pipe\hsx-daemon` on Windows). Socket path is the discovery mechanism — no port file needed. Added `HSX_SOCKET_PATH` env var for override. Stale socket detection via `connect()` probe.
