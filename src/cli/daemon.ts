import { spawn } from "node:child_process";
import { existsSync, unlinkSync } from "node:fs";
import { createRequire } from "node:module";
import { connect, createServer, type Socket } from "node:net";
import { homedir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { setHookWriters } from "../hooks.js";
import { dispose as disposeRuntime, init } from "../index.js";
import { registerSignalTimeout, throwIfAborted } from "./abort.js";
import { DaemonProvider } from "./commands/watch.js";
import { runWithDaemonRuntime } from "./context.js";
import { dispatch } from "./dispatch.js";
import { FileCache } from "./file-cache.js";
import {
  createIoContext,
  runWithIo,
  writeStderr,
  writeStdout,
} from "./output.js";
import { WatchServer } from "./watch-server.js";

function envEnabled(value: string | undefined): boolean {
  if (!value) return false;
  return /^(1|true|yes|on)$/i.test(value.trim());
}

function envPositiveInt(value: string | undefined, fallback: number): number {
  if (!value) return fallback;
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

const IDLE_TIMEOUT_MS = 60 * 60 * 1000;
const AUTO_FLUSH_INTERVAL_MS = envPositiveInt(
  process.env.HSX_AUTO_FLUSH_MS,
  5_000,
);
const DEFAULT_CACHE_SIZE = 10;
const DEFAULT_REQUEST_TIMEOUT_MS = 30_000;
const PROBE_TIMEOUT_MS = 3_000;

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

type DaemonRequest = {
  argv: string[];
  cwd: string;
  stdin?: string;
  timeoutMs?: number;
};
type DaemonResponse = { stdout: string; stderr: string; exitCode: number };

function getCliSpawnSpec(): { command: string; args: string[] } {
  const thisFile = fileURLToPath(import.meta.url);
  const dir = dirname(thisFile);

  if (thisFile.endsWith(".ts")) {
    const require = createRequire(import.meta.url);
    const tsxCli = require.resolve("tsx/cli");
    return {
      command: process.execPath,
      args: [tsxCli, join(dir, "index.ts")],
    };
  }

  return {
    command: process.execPath,
    args: [join(dir, "index.js")],
  };
}

async function runIsolatedCli(
  request: DaemonRequest,
  signal: AbortSignal,
): Promise<DaemonResponse> {
  return new Promise((resolve, reject) => {
    const spawnSpec = getCliSpawnSpec();
    const child = spawn(
      spawnSpec.command,
      [...spawnSpec.args, "--no-daemon", ...request.argv],
      {
        cwd: request.cwd,
        env: { ...process.env },
        stdio: ["pipe", "pipe", "pipe"],
      },
    );

    let stdout = "";
    let stderr = "";
    let settled = false;
    let killTimer: ReturnType<typeof setTimeout> | null = null;
    let aborted = false;

    const cleanup = () => {
      if (killTimer) clearTimeout(killTimer);
      signal.removeEventListener("abort", onAbort);
    };

    const finish = (fn: () => void) => {
      if (settled) return;
      settled = true;
      cleanup();
      fn();
    };

    const onAbort = () => {
      aborted = true;
      child.kill("SIGTERM");
      killTimer = setTimeout(() => {
        child.kill("SIGKILL");
      }, 250);
    };

    if (signal.aborted) {
      onAbort();
    } else {
      signal.addEventListener("abort", onAbort, { once: true });
    }

    child.stdout.on("data", (chunk) => {
      stdout += chunk.toString();
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk.toString();
    });

    child.on("error", (err) => {
      finish(() => reject(err));
    });

    child.on("exit", (code, sig) => {
      finish(() => {
        if (aborted) {
          reject(
            signal.reason instanceof Error
              ? signal.reason
              : new Error(
                  `Eval aborted${sig ? ` (${sig})` : code !== null ? ` (exit ${code})` : ""}`,
                ),
          );
          return;
        }
        resolve({ stdout, stderr, exitCode: code ?? 1 });
      });
    });

    if (request.stdin !== undefined) {
      child.stdin.end(request.stdin);
    } else {
      child.stdin.end();
    }
  });
}

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

  setHookWriters(
    (data) => writeStdout(data),
    (data) => writeStderr(data),
  );
  const { GC, ExcelFile } = await init();
  const cacheSize = envPositiveInt(
    process.env.HSX_CACHE_SIZE,
    DEFAULT_CACHE_SIZE,
  );
  const writeThrough = envEnabled(process.env.HSX_WRITE_THROUGH);
  const fileCache = new FileCache(cacheSize);

  let idleTimer: ReturnType<typeof setTimeout> | null = null;
  let activeConnections = 0;
  let watchServer: WatchServer | null = null;
  let watchServerPort = 0;
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
    if (autoFlushTimer) clearInterval(autoFlushTimer);
    if (watchServer) {
      watchServer.stop();
      watchServer = null;
    }

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

  let autoFlushTimer: ReturnType<typeof setInterval> | null = null;
  if (!writeThrough) {
    autoFlushTimer = setInterval(() => {
      if (fileCache.dirtyCount > 0) {
        enqueue(async () => {
          try {
            const flushed = await fileCache.flushDirty();
            if (flushed > 0) {
              daemonLog(`auto-flush: saved ${flushed} dirty file(s) to disk`);
            }
          } catch (err) {
            daemonLog(`auto-flush failed: ${errorMessage(err)}`);
          }
        }).catch(() => {});
      }
    }, AUTO_FLUSH_INTERVAL_MS);
    autoFlushTimer.unref();
  }

  async function handleRequest(
    request: DaemonRequest,
    signal: AbortSignal,
  ): Promise<DaemonResponse & { shutdown?: boolean }> {
    try {
      throwIfAborted(signal);

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

      if (request.argv[0] === "daemon" && request.argv[1] === "files") {
        return {
          stdout: `${JSON.stringify({ files: fileCache.files() })}\n`,
          stderr: "",
          exitCode: 0,
        };
      }

      if (request.argv[0] === "daemon" && request.argv[1] === "watch") {
        const portIdx = request.argv.indexOf("--port");
        const port =
          portIdx !== -1
            ? Number.parseInt(request.argv[portIdx + 1], 10) || 8080
            : 8080;

        if (watchServer) {
          return {
            stdout: `${JSON.stringify({ watching: true, url: `http://127.0.0.1:${watchServerPort}`, alreadyRunning: true })}\n`,
            stderr: "",
            exitCode: 0,
          };
        }

        try {
          const provider = new DaemonProvider(fileCache);
          watchServer = new WatchServer(provider);
          watchServerPort = await watchServer.start(port);
          daemonLog(`watch server started on port ${watchServerPort}`);
          return {
            stdout: `${JSON.stringify({ watching: true, url: `http://127.0.0.1:${watchServerPort}` })}\n`,
            stderr: "",
            exitCode: 0,
          };
        } catch (err) {
          watchServer?.stop();
          watchServer = null;
          watchServerPort = 0;
          return {
            stdout: "",
            stderr: `${JSON.stringify({ error: `Failed to start watch server: ${errorMessage(err)}` })}\n`,
            exitCode: 1,
          };
        }
      }

      if (request.argv[0] === "daemon" && request.argv[1] === "unwatch") {
        if (watchServer) {
          watchServer.stop();
          watchServer = null;
          watchServerPort = 0;
          daemonLog("watch server stopped");
        }
        return {
          stdout: `${JSON.stringify({ stopped: true })}\n`,
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

      if (request.argv[0] === "eval") {
        if (fileCache.dirtyCount > 0) {
          await fileCache.flushDirty();
        }
        const response = await runIsolatedCli(request, signal);
        const evalFile = request.argv[1];
        if (evalFile) {
          fileCache.invalidate(evalFile, request.cwd);
        }
        return response;
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
          runWithIo(io, () => dispatch(request.argv, { signal })),
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

    const activeControllers = new Set<AbortController>();
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

          const controller = new AbortController();
          activeControllers.add(controller);

          const requestTimeoutMs = envPositiveInt(
            request.timeoutMs ? String(request.timeoutMs) : undefined,
            DEFAULT_REQUEST_TIMEOUT_MS,
          );
          registerSignalTimeout(
            controller.signal,
            requestTimeoutMs,
            `Command timed out after ${Math.ceil(requestTimeoutMs / 1000)}s`,
          );
          const timeout = setTimeout(() => {
            controller.abort(
              new Error(
                `Command timed out after ${Math.ceil(requestTimeoutMs / 1000)}s`,
              ),
            );
          }, requestTimeoutMs);

          enqueue(async () => {
            try {
              const response = await handleRequest(request, controller.signal);
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
            } finally {
              clearTimeout(timeout);
              activeControllers.delete(controller);
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
      for (const controller of activeControllers) {
        controller.abort(new Error("Client disconnected"));
      }
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
