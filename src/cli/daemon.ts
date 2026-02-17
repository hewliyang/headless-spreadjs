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

export function getSocketPath(): string {
  if (process.env.HSX_SOCKET_PATH) return process.env.HSX_SOCKET_PATH;
  return process.platform === "win32"
    ? "\\\\.\\pipe\\hsx-daemon"
    : join(homedir(), ".hsx-daemon.sock");
}

/**
 * Probe the socket to check if a daemon is already listening.
 * Returns true if a connection succeeds.
 */
function isDaemonListening(socketPath: string): Promise<boolean> {
  return new Promise((resolve) => {
    const probe = connect({ path: socketPath });
    probe.on("connect", () => {
      probe.end();
      resolve(true);
    });
    probe.on("error", () => resolve(false));
  });
}

type DaemonRequest = { argv: string[]; cwd: string; stdin?: string };
type DaemonResponse = { stdout: string; stderr: string; exitCode: number };

export async function startDaemon(): Promise<void> {
  const socketPath = getSocketPath();

  // Clean up stale socket from a previous crash (Unix only — named pipes
  // on Windows don't leave files behind).
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
  const fileCache = new FileCache(10);

  let idleTimer: ReturnType<typeof setTimeout> | null = null;
  let activeConnections = 0;

  const server = createServer(handleConnection);

  function resetIdleTimer() {
    if (idleTimer) clearTimeout(idleTimer);
    idleTimer = setTimeout(() => {
      if (activeConnections === 0) shutdown();
    }, IDLE_TIMEOUT_MS);
  }

  function shutdown() {
    if (idleTimer) clearTimeout(idleTimer);
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
    const task = queue.then(fn, fn);
    queue = task;
    return task;
  }

  async function handleRequest(
    request: DaemonRequest,
  ): Promise<DaemonResponse> {
    if (request.argv[0] === "daemon" && request.argv[1] === "stop") {
      setTimeout(shutdown, 100);
      return {
        stdout: `${JSON.stringify({ stopped: true })}\n`,
        stderr: "",
        exitCode: 0,
      };
    }

    if (request.argv[0] === "daemon" && request.argv[1] === "status") {
      return {
        stdout: `${JSON.stringify({
          pid: process.pid,
          cachedFiles: fileCache.size,
          uptime: Math.floor(process.uptime()),
        })}\n`,
        stderr: "",
        exitCode: 0,
      };
    }

    const runtime = { GC, ExcelFile, fileCache, cwd: request.cwd };
    const io = createIoContext(request.stdin);

    try {
      await runWithDaemonRuntime(runtime, () =>
        runWithIo(io, () => dispatch(request.argv)),
      );
      return { stdout: io.stdout, stderr: io.stderr, exitCode: 0 };
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      return {
        stdout: io.stdout,
        stderr: `${io.stderr}${JSON.stringify({ error: message })}\n`,
        exitCode: 1,
      };
    }
  }

  function handleConnection(socket: Socket) {
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
            try {
              socket.write(`${JSON.stringify(response)}\n`);
            } catch {}
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

  server.listen(socketPath, () => {
    if (process.send) process.send({ ready: true });
    process.stdout.write(
      `${JSON.stringify({ daemon: "started", pid: process.pid })}\n`,
    );
    resetIdleTimer();
  });

  process.on("SIGINT", shutdown);
  process.on("SIGTERM", shutdown);
}
