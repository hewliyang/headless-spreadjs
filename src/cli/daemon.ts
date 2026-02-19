import { existsSync, unlinkSync } from "node:fs";
import { connect, createServer, type Socket } from "node:net";
import { homedir } from "node:os";
import { join } from "node:path";
import { dispose as disposeRuntime, init } from "../index.js";
import { runWithDaemonRuntime } from "./context.js";
import { dispatch } from "./dispatch.js";
import { FileCache } from "./file-cache.js";
import { createIoContext, runWithIo } from "./output.js";

const IDLE_TIMEOUT_MS = 5 * 60 * 1000;
const DEFAULT_CACHE_SIZE = 10;
const PROBE_TIMEOUT_MS = 3_000;

function envEnabled(value: string | undefined): boolean {
  if (!value) return false;
  return /^(1|true|yes|on)$/i.test(value.trim());
}

function envPositiveInt(value: string | undefined, fallback: number): number {
  if (!value) return fallback;
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

function errorMessage(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

export function getSocketPath(): string {
  if (process.env.HSX_SOCKET_PATH) return process.env.HSX_SOCKET_PATH;
  return process.platform === "win32"
    ? "\\\\.\\pipe\\hsx-daemon"
    : join(homedir(), ".hsx-daemon.sock");
}

export function getLogPath(): string {
  const socketPath = getSocketPath();
  if (socketPath.startsWith("\\\\.\\pipe\\")) {
    return join(homedir(), ".hsx-daemon.log");
  }
  return `${socketPath.replace(/\.sock$/, "")}.log`;
}

function isDaemonListening(socketPath: string): Promise<boolean> {
  return new Promise((resolve) => {
    const probe = connect({ path: socketPath });
    const timeout = setTimeout(() => {
      probe.destroy();
      resolve(false);
    }, PROBE_TIMEOUT_MS);

    probe.on("connect", () => {
      clearTimeout(timeout);
      probe.end();
      resolve(true);
    });

    probe.on("error", () => {
      clearTimeout(timeout);
      resolve(false);
    });
  });
}

type DaemonRequest = { argv: string[]; cwd: string; stdin?: string };
type DaemonResponse = { stdout: string; stderr: string; exitCode: number };

export async function startDaemon(): Promise<void> {
  const socketPath = getSocketPath();

  if (process.platform !== "win32" && existsSync(socketPath)) {
    if (await isDaemonListening(socketPath)) {
      process.stderr.write(
        `${JSON.stringify({ error: "Daemon already running" })}\n`,
      );
      process.exit(1);
    }
    try {
      unlinkSync(socketPath);
    } catch {}
  }

  const { GC, ExcelFile } = await init();
  const cacheSize = envPositiveInt(
    process.env.HSX_CACHE_SIZE,
    DEFAULT_CACHE_SIZE,
  );
  const writeThrough = envEnabled(process.env.HSX_WRITE_THROUGH);
  const fileCache = new FileCache(cacheSize);

  let idleTimer: ReturnType<typeof setTimeout> | null = null;
  let activeConnections = 0;
  const server = createServer(handleConnection);

  function daemonLog(msg: string): void {
    const ts = new Date().toISOString();
    process.stderr.write(`[hsx-daemon ${ts}] ${msg}\n`);
  }

  function resetIdleTimer(): void {
    if (idleTimer) clearTimeout(idleTimer);
    idleTimer = setTimeout(() => {
      if (activeConnections === 0) {
        void shutdown().catch((err) => {
          daemonLog(`idle shutdown failed: ${errorMessage(err)}`);
        });
      }
    }, IDLE_TIMEOUT_MS);
  }

  let shuttingDown = false;

  async function shutdown(options?: { skipFlush?: boolean }): Promise<void> {
    if (shuttingDown) return;
    shuttingDown = true;

    if (idleTimer) clearTimeout(idleTimer);

    if (!options?.skipFlush) {
      try {
        const flushed = await fileCache.flushDirty();
        if (flushed > 0) {
          daemonLog(`flushed ${flushed} dirty file(s) to disk`);
        }
      } catch (err) {
        daemonLog(
          `shutdown aborted: failed to flush dirty files: ${errorMessage(err)}`,
        );
        shuttingDown = false;
        resetIdleTimer();
        throw err;
      }
    }

    fileCache.clear();

    server.close();
    try {
      unlinkSync(socketPath);
    } catch {}
    disposeRuntime();
    process.exit(0);
  }

  let queue: Promise<void> = Promise.resolve();
  function enqueue(fn: () => Promise<void>): Promise<void> {
    const task = queue
      .catch((err) => daemonLog(`queue error: ${errorMessage(err)}`))
      .then(fn);
    queue = task;
    return task;
  }

  async function handleRequest(
    request: DaemonRequest,
  ): Promise<DaemonResponse & { shutdown?: boolean }> {
    try {
      if (request.argv[0] === "daemon" && request.argv[1] === "stop") {
        const flushed = await fileCache.flushDirty();
        if (flushed > 0) {
          daemonLog(`flushed ${flushed} dirty file(s) before stop`);
        }
        return {
          stdout: `${JSON.stringify({ stopped: true })}\n`,
          stderr: "",
          exitCode: 0,
          shutdown: true,
        };
      }

      if (request.argv[0] === "daemon" && request.argv[1] === "flush") {
        const flushed = await fileCache.flushDirty();
        if (flushed > 0) {
          daemonLog(`flushed ${flushed} dirty file(s) via explicit request`);
        }
        return {
          stdout: `${JSON.stringify({
            flushed,
            dirtyFiles: fileCache.dirtyCount,
          })}\n`,
          stderr: "",
          exitCode: 0,
        };
      }

      if (request.argv[0] === "daemon" && request.argv[1] === "status") {
        const mem = process.memoryUsage();
        return {
          stdout: `${JSON.stringify({
            pid: process.pid,
            cachedFiles: fileCache.size,
            dirtyFiles: fileCache.dirtyCount,
            maxCacheSize: fileCache.maxCacheSize,
            writeThrough,
            uptime: Math.floor(process.uptime()),
            heapUsedMB: Math.round(mem.heapUsed / 1024 / 1024),
            rssMB: Math.round(mem.rss / 1024 / 1024),
            logFile: getLogPath(),
          })}\n`,
          stderr: "",
          exitCode: 0,
        };
      }

      const runtime = {
        GC,
        ExcelFile,
        fileCache,
        cwd: request.cwd,
        writeThrough,
      };
      const io = createIoContext(request.stdin);

      try {
        await runWithDaemonRuntime(runtime, () =>
          runWithIo(io, () => dispatch(request.argv)),
        );
        return { stdout: io.stdout, stderr: io.stderr, exitCode: 0 };
      } catch (err) {
        return {
          stdout: io.stdout,
          stderr: `${io.stderr}${JSON.stringify({ error: errorMessage(err) })}\n`,
          exitCode: 1,
        };
      }
    } catch (err) {
      return {
        stdout: "",
        stderr: `${JSON.stringify({ error: errorMessage(err) })}\n`,
        exitCode: 1,
      };
    }
  }

  function handleConnection(socket: Socket): void {
    activeConnections++;
    resetIdleTimer();

    let buffer = "";

    socket.on("data", (data) => {
      buffer += data.toString();

      let newlineIdx = buffer.indexOf("\n");
      while (newlineIdx !== -1) {
        const line = buffer.slice(0, newlineIdx);
        buffer = buffer.slice(newlineIdx + 1);

        if (line.trim()) {
          let request: DaemonRequest;
          try {
            request = JSON.parse(line) as DaemonRequest;
          } catch {
            socket.write(
              `${JSON.stringify({ stdout: "", stderr: "Invalid JSON\n", exitCode: 1 })}\n`,
            );
            newlineIdx = buffer.indexOf("\n");
            continue;
          }

          enqueue(async () => {
            const response = await handleRequest(request);
            const { shutdown: shouldShutdown, ...wire } = response;
            try {
              socket.write(`${JSON.stringify(wire)}\n`, (err) => {
                if (err) daemonLog(`socket.write error: ${err.message}`);
                if (shouldShutdown) {
                  void shutdown({ skipFlush: true }).catch((shutdownErr) => {
                    daemonLog(
                      `shutdown failed after stop ack: ${errorMessage(shutdownErr)}`,
                    );
                  });
                }
              });
            } catch (err) {
              daemonLog(`socket.write threw: ${errorMessage(err)}`);
            }
          }).catch((err) => {
            daemonLog(`request handling failed: ${errorMessage(err)}`);
          });
        }

        newlineIdx = buffer.indexOf("\n");
      }
    });

    let disconnected = false;
    const onDisconnect = () => {
      if (disconnected) return;
      disconnected = true;
      activeConnections--;
      resetIdleTimer();
    };

    socket.on("close", onDisconnect);
    socket.on("error", onDisconnect);
  }

  server.on("error", (err) => {
    daemonLog(`server error: ${err.message}`);
    try {
      unlinkSync(socketPath);
    } catch {}
    disposeRuntime();
    process.exit(1);
  });

  server.listen(socketPath, () => {
    if (process.send) process.send({ ready: true });
    process.stdout.write(
      `${JSON.stringify({ daemon: "started", pid: process.pid })}\n`,
    );
    resetIdleTimer();
  });

  process.on("SIGINT", () => {
    void shutdown().catch((err) => {
      daemonLog(`SIGINT shutdown failed: ${errorMessage(err)}`);
      process.exit(1);
    });
  });

  process.on("SIGTERM", () => {
    void shutdown().catch((err) => {
      daemonLog(`SIGTERM shutdown failed: ${errorMessage(err)}`);
      process.exit(1);
    });
  });
}
