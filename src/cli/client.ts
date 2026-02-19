import { spawn } from "node:child_process";
import { closeSync, openSync } from "node:fs";
import { connect } from "node:net";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { getLogPath, getSocketPath } from "./daemon.js";

const CLIENT_TIMEOUT_MS = 30_000;
const SPAWN_TIMEOUT_MS = 15_000;

function isDaemonUnavailableError(err: unknown): boolean {
  if (!(err instanceof Error)) return false;
  const code = (err as NodeJS.ErrnoException).code;
  return code === "ENOENT" || code === "ECONNREFUSED";
}

export type DaemonCommandResult = {
  stdout: string;
  stderr: string;
  exitCode: number;
};

async function sendCommand(
  socketPath: string,
  argv: string[],
  cwd: string,
  stdin?: string,
  timeoutMs = CLIENT_TIMEOUT_MS,
): Promise<DaemonCommandResult> {
  return new Promise((resolve, reject) => {
    const socket = connect({ path: socketPath }, () => {
      const request = JSON.stringify({ argv, cwd, stdin, timeoutMs });
      socket.write(`${request}\n`);
    });

    let buffer = "";
    const timeout = setTimeout(() => {
      socket.destroy();
      reject(new Error("Daemon request timed out"));
    }, timeoutMs);

    socket.on("data", (data) => {
      buffer += data.toString();
      const newlineIdx = buffer.indexOf("\n");
      if (newlineIdx !== -1) {
        clearTimeout(timeout);
        const line = buffer.slice(0, newlineIdx);
        socket.end();
        try {
          resolve(JSON.parse(line));
        } catch {
          reject(new Error("Invalid response from daemon"));
        }
      }
    });

    socket.on("error", (err) => {
      clearTimeout(timeout);
      reject(err);
    });

    socket.on("close", () => {
      clearTimeout(timeout);
      if (!buffer.includes("\n")) {
        reject(new Error("Connection closed before response"));
      }
    });
  });
}

export async function spawnDaemon(timeoutMs = SPAWN_TIMEOUT_MS): Promise<void> {
  const thisFile = fileURLToPath(import.meta.url);
  const cliDir = dirname(thisFile);
  const daemonEntry = join(cliDir, "daemon-entry.js");

  const logFd = openSync(getLogPath(), "a");

  return new Promise<void>((resolve, reject) => {
    const child = spawn(process.execPath, [daemonEntry], {
      detached: true,
      stdio: ["ignore", "ignore", logFd, "ipc"],
      env: { ...process.env },
    });

    closeSync(logFd);

    const timeout = setTimeout(() => {
      child.kill();
      reject(new Error("Daemon failed to start within timeout"));
    }, timeoutMs);

    child.on("message", (msg: unknown) => {
      const m = msg as Record<string, unknown>;
      if (m?.ready) {
        clearTimeout(timeout);
        child.disconnect();
        child.unref();
        resolve();
      }
    });

    child.on("error", (err) => {
      clearTimeout(timeout);
      reject(err);
    });

    child.on("exit", (code) => {
      clearTimeout(timeout);
      reject(new Error(`Daemon exited with code ${code}`));
    });
  });
}

export async function tryExistingDaemon(
  argv: string[],
  cwd: string,
  stdin?: string,
  timeoutMs = CLIENT_TIMEOUT_MS,
): Promise<DaemonCommandResult | null> {
  const socketPath = getSocketPath();
  try {
    return await sendCommand(socketPath, argv, cwd, stdin, timeoutMs);
  } catch (err) {
    if (isDaemonUnavailableError(err)) {
      return null;
    }
    throw err;
  }
}

export async function tryDaemon(
  argv: string[],
  cwd: string,
  stdin?: string,
  timeoutMs = CLIENT_TIMEOUT_MS,
): Promise<DaemonCommandResult | null> {
  const existing = await tryExistingDaemon(argv, cwd, stdin, timeoutMs);
  if (existing) {
    return existing;
  }

  const socketPath = getSocketPath();

  try {
    await spawnDaemon(timeoutMs);
  } catch {
    return null;
  }

  try {
    return await sendCommand(socketPath, argv, cwd, stdin, timeoutMs);
  } catch (err) {
    if (isDaemonUnavailableError(err)) {
      return null;
    }
    throw err;
  }
}
