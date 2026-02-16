/**
 * hsx daemon — long-running TCP server that keeps SpreadJS init'd
 * and files cached in memory for fast CLI operations.
 *
 * Protocol: newline-delimited JSON over TCP on localhost.
 * Request:  { argv: string[], cwd: string, stdin?: string }
 * Response: { stdout: string, stderr: string, exitCode: number }
 */

import { createServer, type Socket } from "node:net";
import { writeFileSync, unlinkSync, readFileSync, existsSync } from "node:fs";
import { join } from "node:path";
import { homedir, tmpdir } from "node:os";
import { init, dispose as disposeRuntime } from "../index.js";
import { FileCache } from "./file-cache.js";
import { setDaemonRuntime } from "./context.js";
import { startCapture, stopCapture, setStdin } from "./output.js";
import { dispatch } from "./main.js";

const IDLE_TIMEOUT_MS = 5 * 60 * 1000; // 5 minutes

export function getPortFilePath(): string {
  // Use tmpdir for cross-platform compat; include uid-like component
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
    // Check if process is still alive
    try {
      process.kill(data.pid, 0);
      return data;
    } catch {
      // Process is dead, clean up stale file
      try { unlinkSync(portFile); } catch {}
      return null;
    }
  } catch {
    return null;
  }
}

export async function startDaemon(): Promise<void> {
  const { GC, ExcelFile, dispose } = await init();
  const fileCache = new FileCache(10);

  let idleTimer: ReturnType<typeof setTimeout> | null = null;
  let activeConnections = 0;

  function resetIdleTimer() {
    if (idleTimer) clearTimeout(idleTimer);
    idleTimer = setTimeout(() => {
      if (activeConnections === 0) {
        shutdown();
      }
    }, IDLE_TIMEOUT_MS);
  }

  function shutdown() {
    if (idleTimer) clearTimeout(idleTimer);
    fileCache.clear();
    try { unlinkSync(getPortFilePath()); } catch {}
    disposeRuntime();
    server.close();
    process.exit(0);
  }

  // Sequential command queue
  let queue: Promise<void> = Promise.resolve();

  function enqueue(fn: () => Promise<void>): Promise<void> {
    const p = queue.then(fn, fn);
    queue = p;
    return p;
  }

  async function handleRequest(
    request: { argv: string[]; cwd: string; stdin?: string },
  ): Promise<{ stdout: string; stderr: string; exitCode: number }> {
    // Check for daemon control commands
    if (request.argv[0] === "daemon" && request.argv[1] === "stop") {
      // Respond first, then shutdown
      setTimeout(shutdown, 100);
      return { stdout: JSON.stringify({ stopped: true }) + "\n", stderr: "", exitCode: 0 };
    }

    if (request.argv[0] === "daemon" && request.argv[1] === "status") {
      return {
        stdout: JSON.stringify({
          pid: process.pid,
          cachedFiles: fileCache.size,
          uptime: Math.floor(process.uptime()),
        }) + "\n",
        stderr: "",
        exitCode: 0,
      };
    }

    // Set daemon runtime context for this command
    setDaemonRuntime({
      GC,
      ExcelFile,
      fileCache,
      cwd: request.cwd,
    });

    // Set stdin for readInput()
    setStdin(request.stdin);

    // Capture output
    startCapture();

    let exitCode = 0;
    try {
      await dispatch(request.argv);
    } catch (err) {
      const { stopCapture: sc } = await import("./output.js");
      // Write error to captured stderr
      const message = err instanceof Error ? err.message : String(err);
      const captured = stopCapture();
      return {
        stdout: captured.stdout,
        stderr: captured.stderr + JSON.stringify({ error: message }) + "\n",
        exitCode: 1,
      };
    }

    const captured = stopCapture();
    setDaemonRuntime(null);
    return { stdout: captured.stdout, stderr: captured.stderr, exitCode };
  }

  function handleConnection(socket: Socket) {
    activeConnections++;
    resetIdleTimer();

    let buffer = "";

    socket.on("data", (data) => {
      buffer += data.toString();

      // Process complete lines
      let newlineIdx: number;
      while ((newlineIdx = buffer.indexOf("\n")) !== -1) {
        const line = buffer.slice(0, newlineIdx);
        buffer = buffer.slice(newlineIdx + 1);

        if (!line.trim()) continue;

        let request: { argv: string[]; cwd: string; stdin?: string };
        try {
          request = JSON.parse(line);
        } catch {
          socket.write(JSON.stringify({ stdout: "", stderr: "Invalid JSON\n", exitCode: 1 }) + "\n");
          continue;
        }

        // Enqueue for sequential execution
        enqueue(async () => {
          const response = await handleRequest(request);
          try {
            socket.write(JSON.stringify(response) + "\n");
          } catch {
            // Socket may have closed
          }
        });
      }
    });

    socket.on("close", () => {
      activeConnections--;
      resetIdleTimer();
    });

    socket.on("error", () => {
      activeConnections--;
      resetIdleTimer();
    });
  }

  const server = createServer(handleConnection);

  // Listen on random port
  server.listen(0, "127.0.0.1", () => {
    const addr = server.address();
    if (!addr || typeof addr === "string") {
      process.exit(1);
    }

    const info: DaemonInfo = { pid: process.pid, port: addr.port };
    writeFileSync(getPortFilePath(), JSON.stringify(info));

    // Signal to parent that we're ready (if spawned by client)
    if (process.send) {
      process.send({ ready: true, port: addr.port });
    }

    // Also write to stdout for direct invocation
    process.stdout.write(JSON.stringify({ daemon: "started", ...info }) + "\n");

    resetIdleTimer();
  });

  // Clean up on signals
  for (const sig of ["SIGINT", "SIGTERM"] as const) {
    process.on(sig, shutdown);
  }
}
