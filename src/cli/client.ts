/**
 * Thin daemon client — sends command over TCP, returns response.
 * Auto-spawns daemon if not running.
 */

import { connect } from "node:net";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { readDaemonInfo, getPortFilePath, type DaemonInfo } from "./daemon.js";

const CLIENT_TIMEOUT_MS = 30_000;

async function sendCommand(
  port: number,
  argv: string[],
  cwd: string,
  stdin?: string,
): Promise<{ stdout: string; stderr: string; exitCode: number }> {
  return new Promise((resolve, reject) => {
    const socket = connect({ host: "127.0.0.1", port }, () => {
      const request = JSON.stringify({ argv, cwd, stdin });
      socket.write(request + "\n");
    });

    let buffer = "";
    const timeout = setTimeout(() => {
      socket.destroy();
      reject(new Error("Daemon request timed out"));
    }, CLIENT_TIMEOUT_MS);

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

async function spawnDaemon(): Promise<DaemonInfo> {
  const thisFile = fileURLToPath(import.meta.url);
  const cliDir = dirname(thisFile);
  const daemonEntry = join(cliDir, "daemon-entry.js");

  return new Promise<DaemonInfo>((resolve, reject) => {
    const child = spawn(process.execPath, [daemonEntry], {
      detached: true,
      stdio: ["ignore", "pipe", "ignore"],
      env: { ...process.env },
    });

    let stdout = "";
    const timeout = setTimeout(() => {
      child.kill();
      reject(new Error("Daemon failed to start within 15s"));
    }, 15_000);

    child.stdout!.on("data", (data: Buffer) => {
      stdout += data.toString();
      if (stdout.includes("\n")) {
        clearTimeout(timeout);
        try {
          const info = JSON.parse(stdout.trim());
          child.unref();
          resolve({ pid: info.pid, port: info.port });
        } catch {
          reject(new Error("Daemon returned invalid startup message"));
        }
      }
    });

    child.on("error", (err) => {
      clearTimeout(timeout);
      reject(err);
    });

    child.on("exit", (code) => {
      clearTimeout(timeout);
      if (!stdout.includes("\n")) {
        reject(new Error(`Daemon exited with code ${code}`));
      }
    });
  });
}

/**
 * Try to run a command via the daemon. Returns null if daemon is
 * unavailable and couldn't be started.
 */
export async function tryDaemon(
  argv: string[],
  cwd: string,
  stdin?: string,
): Promise<{ stdout: string; stderr: string; exitCode: number } | null> {
  let info = readDaemonInfo();

  if (!info) {
    try {
      info = await spawnDaemon();
    } catch {
      return null; // Fall back to direct mode
    }
  }

  try {
    return await sendCommand(info.port, argv, cwd, stdin);
  } catch {
    // Daemon might be stale, try respawning once
    try {
      info = await spawnDaemon();
      return await sendCommand(info.port, argv, cwd, stdin);
    } catch {
      return null;
    }
  }
}
