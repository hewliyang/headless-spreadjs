import { existsSync, readFileSync, unlinkSync, writeFileSync } from "node:fs";
import { createServer, type Socket } from "node:net";
import { homedir, tmpdir } from "node:os";
import { join } from "node:path";
import { dispose as disposeRuntime, init } from "../index.js";
import { setDaemonRuntime } from "./context.js";
import { dispatch } from "./dispatch.js";
import { FileCache } from "./file-cache.js";
import { setStdin, startCapture, stopCapture } from "./output.js";

const IDLE_TIMEOUT_MS = 5 * 60 * 1000;

export function getPortFilePath(): string {
  const dir = process.platform === "win32" ? tmpdir() : homedir();
  return join(dir, ".hsx-daemon.json");
}

export interface DaemonInfo {
  pid: number;
  port: number;
}

export function readDaemonInfo(): DaemonInfo | null {
  const portFile = getPortFilePath();
  try {
    if (!existsSync(portFile)) return null;
    const data = JSON.parse(readFileSync(portFile, "utf-8")) as DaemonInfo;
    try {
      process.kill(data.pid, 0);
      return data;
    } catch {
      try {
        unlinkSync(portFile);
      } catch {}
      return null;
    }
  } catch {
    return null;
  }
}

type DaemonRequest = { argv: string[]; cwd: string; stdin?: string };
type DaemonResponse = { stdout: string; stderr: string; exitCode: number };

export async function startDaemon(): Promise<void> {
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
    try {
      unlinkSync(getPortFilePath());
    } catch {}
    disposeRuntime();
    server.close();
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

    setDaemonRuntime({ GC, ExcelFile, fileCache, cwd: request.cwd });
    setStdin(request.stdin);
    startCapture();

    try {
      await dispatch(request.argv);
      const captured = stopCapture();
      return { stdout: captured.stdout, stderr: captured.stderr, exitCode: 0 };
    } catch (err) {
      const captured = stopCapture();
      const message = err instanceof Error ? err.message : String(err);
      return {
        stdout: captured.stdout,
        stderr: `${captured.stderr}${JSON.stringify({ error: message })}\n`,
        exitCode: 1,
      };
    } finally {
      setDaemonRuntime(null);
      setStdin(undefined);
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

  server.listen(0, "127.0.0.1", () => {
    const addr = server.address();
    if (!addr || typeof addr === "string") process.exit(1);

    const info: DaemonInfo = { pid: process.pid, port: addr.port };
    writeFileSync(getPortFilePath(), JSON.stringify(info));

    if (process.send) process.send({ ready: true, port: addr.port });
    process.stdout.write(`${JSON.stringify({ daemon: "started", ...info })}\n`);

    resetIdleTimer();
  });

  process.on("SIGINT", shutdown);
  process.on("SIGTERM", shutdown);
}
