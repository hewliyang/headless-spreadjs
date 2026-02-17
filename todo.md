# Daemon PR — TODO

## Low priority (optional)

### 7. Newline-delimited JSON protocol

The protocol assumes `JSON.stringify()` never produces raw `\n`. True today, but an implicit contract.

**Fix (optional)**: Switch to length-prefix framing (`4-byte big-endian length + payload`). ~20 lines of framing code on each side. Only worth doing if binary payloads or very large outputs become a concern.

**Files**: `src/cli/daemon.ts`, `src/cli/client.ts`

## Done

### 1. Global mutable state in `output.ts` / `context.ts`

Replaced module-level globals (`capturedStdout`, `capturedStderr`, `pendingStdin`, `daemonRuntime`) with `AsyncLocalStorage` scoped per-request. Deleted `startCapture`, `stopCapture`, `setStdin`, `setDaemonRuntime`. Each request is wrapped in `runWithIo` + `runWithDaemonRuntime`.

### 2. Race condition on daemon startup

Eliminated by switching from TCP + port file to Unix domain sockets. `server.listen(socketPath)` is atomic — a second daemon probes the socket, sees the first is alive, and exits. No orphaned daemons.

### 3. `statSync` blocks event loop in cache `get()`

`get()`, `put()`, `updateMtime()` are now async using `fs.promises.stat()`. Callers in `context.ts` awaited.

### 4. No memory visibility or limits

`daemon status` now reports `heapUsedMB`, `rssMB`, and `maxCacheSize`. Cache size is configurable via `HSX_CACHE_SIZE` env var.

### 5. LRU eviction is O(n) scan

Replaced with Map insertion-order LRU. `get()` does delete + re-insert. `put()` evicts first key. Deleted `lastAccess` field.

### 6. Document `HSX_SOCKET_PATH` for multi-daemon use

Documented `HSX_SOCKET_PATH` and `HSX_CACHE_SIZE` env vars in README daemon section.

### 7 (transport). TCP → Unix domain sockets

Replaced `127.0.0.1` TCP + `~/.hsx-daemon.json` port file with Unix domain sockets (`~/.hsx-daemon.sock`) / named pipes (`\\.\pipe\hsx-daemon` on Windows). Added `HSX_SOCKET_PATH` env var. Stale socket detection via `connect()` probe.

### 8. `enqueue` swallows errors silently

Added `daemonLog()` helper. Queue catches prior rejections and logs them. `socket.write` errors are logged with context.

### 9. `daemon stop` uses `setTimeout(shutdown, 100)` hack

`handleRequest` returns a `shutdown` flag. `enqueue` uses `socket.write(data, callback)` to call `shutdown()` only after bytes flush.

### 10. `spawnDaemon` parses stdout for readiness

Switched to IPC channel (`stdio: ["ignore", "ignore", "ignore", "ipc"]`). Client listens for `child.on("message")` with `{ ready: true }`. `process.stdout.write()` kept as fallback for interactive `daemon start`.

### 11. README concurrency section is overstated

Rewritten to clarify that multiple concurrent workbook operations are safe within a single `init()`/`dispose()` lifecycle. Recommends child processes over worker threads.
