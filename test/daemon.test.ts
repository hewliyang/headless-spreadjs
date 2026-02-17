import assert from "node:assert/strict";
import { type ChildProcess, spawn } from "node:child_process";
import fs from "node:fs/promises";
import { connect } from "node:net";
import os from "node:os";
import path from "node:path";
import { afterAll, beforeAll, describe, it } from "vitest";

const DAEMON_ENTRY = path.resolve("src/cli/daemon-entry.ts");

type DaemonResponse = { stdout: string; stderr: string; exitCode: number };

function sendCommand(
  socketPath: string,
  argv: string[],
  cwd: string,
  stdin?: string,
): Promise<DaemonResponse> {
  return new Promise((resolve, reject) => {
    const socket = connect({ path: socketPath }, () => {
      socket.write(`${JSON.stringify({ argv, cwd, stdin })}\n`);
    });

    let buffer = "";
    const timeout = setTimeout(() => {
      socket.destroy();
      reject(new Error("Daemon request timed out"));
    }, 30_000);

    socket.on("data", (data) => {
      buffer += data.toString();
      const idx = buffer.indexOf("\n");
      if (idx !== -1) {
        clearTimeout(timeout);
        socket.end();
        try {
          resolve(JSON.parse(buffer.slice(0, idx)));
        } catch {
          reject(new Error("Invalid response from daemon"));
        }
      }
    });

    socket.on("error", (err) => {
      clearTimeout(timeout);
      reject(err);
    });
  });
}

/** Send multiple requests on a single connection (pipelining). */
function sendPipelined(
  socketPath: string,
  requests: { argv: string[]; cwd: string; stdin?: string }[],
): Promise<DaemonResponse[]> {
  return new Promise((resolve, reject) => {
    const socket = connect({ path: socketPath }, () => {
      for (const req of requests) {
        socket.write(`${JSON.stringify(req)}\n`);
      }
    });

    const responses: DaemonResponse[] = [];
    let buffer = "";
    const timeout = setTimeout(() => {
      socket.destroy();
      reject(new Error("Pipelined request timed out"));
    }, 30_000);

    socket.on("data", (data) => {
      buffer += data.toString();
      let idx = buffer.indexOf("\n");
      while (idx !== -1) {
        const line = buffer.slice(0, idx);
        buffer = buffer.slice(idx + 1);
        if (line.trim()) {
          try {
            responses.push(JSON.parse(line));
          } catch {
            clearTimeout(timeout);
            reject(new Error("Invalid response from daemon"));
            return;
          }
        }
        if (responses.length === requests.length) {
          clearTimeout(timeout);
          socket.end();
          resolve(responses);
          return;
        }
        idx = buffer.indexOf("\n");
      }
    });

    socket.on("error", (err) => {
      clearTimeout(timeout);
      reject(err);
    });
  });
}

let tmpDir: string;
let socketPath: string;
let daemonProc: ChildProcess;

beforeAll(async () => {
  tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "hsx-daemon-test-"));
  socketPath = path.join(tmpDir, "test-daemon.sock");

  daemonProc = spawn("tsx", [DAEMON_ENTRY], {
    stdio: ["ignore", "pipe", "pipe"],
    env: { ...process.env, HSX_SOCKET_PATH: socketPath },
  });

  await new Promise<void>((resolve, reject) => {
    let stdout = "";
    const timeout = setTimeout(() => {
      daemonProc.kill();
      reject(new Error("Daemon failed to start within 30s"));
    }, 30_000);

    daemonProc.stdout!.on("data", (d: Buffer) => {
      stdout += d.toString();
      if (stdout.includes("\n")) {
        clearTimeout(timeout);
        JSON.parse(stdout.trim()); // validate
        resolve();
      }
    });

    daemonProc.on("error", (err) => {
      clearTimeout(timeout);
      reject(err);
    });
    daemonProc.on("exit", (code) => {
      clearTimeout(timeout);
      if (!stdout.includes("\n"))
        reject(new Error(`Daemon exited early with code ${code}`));
    });
  });
}, 60_000);

afterAll(async () => {
  try {
    await sendCommand(socketPath, ["daemon", "stop"], tmpDir);
  } catch {}
  await new Promise((r) => setTimeout(r, 200));
  try {
    daemonProc?.kill();
  } catch {}
  await fs.rm(tmpDir, { recursive: true, force: true });
});

describe("daemon", () => {
  it("status", async () => {
    const res = await sendCommand(socketPath, ["daemon", "status"], tmpDir);
    assert.equal(res.exitCode, 0);
    const status = JSON.parse(res.stdout);
    assert.ok(typeof status.pid === "number");
    assert.ok(status.uptime >= 0);
    assert.ok(typeof status.cachedFiles === "number");
  });

  it("create + info", async () => {
    const file = path.join(tmpDir, "test.xlsx");
    const create = await sendCommand(
      socketPath,
      ["create", file],
      tmpDir,
    );
    assert.equal(create.exitCode, 0);
    assert.deepStrictEqual(JSON.parse(create.stdout), { created: file });

    const inf = await sendCommand(socketPath, ["info", file], tmpDir);
    assert.equal(inf.exitCode, 0);
    const data = JSON.parse(inf.stdout);
    assert.equal(data.sheets.length, 1);
    assert.equal(data.sheets[0].name, "Sheet1");
  });

  it("set + get round-trip", async () => {
    const file = path.join(tmpDir, "test.xlsx");
    const cells = [
      [{ value: "Name" }, { value: "Qty" }],
      [{ value: "Alice" }, { value: 4 }],
      [{ value: "Bob" }, { formula: "=B2+1" }],
    ];

    const setRes = await sendCommand(
      socketPath,
      ["set", file, "A1:B3", JSON.stringify(cells)],
      tmpDir,
    );
    assert.equal(setRes.exitCode, 0);
    assert.deepStrictEqual(JSON.parse(setRes.stdout), {
      written: 6,
      range: "A1:B3",
    });

    const getRes = await sendCommand(
      socketPath,
      ["get", file, "A1:B3", "--no-styles"],
      tmpDir,
    );
    assert.equal(getRes.exitCode, 0);
    const data = JSON.parse(getRes.stdout);
    assert.equal(data.cells.A1.value, "Name");
    assert.equal(data.cells.B2.value, 4);
    assert.equal(data.cells.B3.value, 5);
    assert.equal(data.cells.B3.formula, "B2+1");
  });

  it("set via stdin field", async () => {
    const file = path.join(tmpDir, "test.xlsx");
    const cells = [[{ value: "from-stdin" }]];

    const setRes = await sendCommand(
      socketPath,
      ["set", file, "D1"],
      tmpDir,
      JSON.stringify(cells),
    );
    assert.equal(setRes.exitCode, 0);

    const getRes = await sendCommand(
      socketPath,
      ["get", file, "D1", "--no-styles"],
      tmpDir,
    );
    assert.equal(getRes.exitCode, 0);
    assert.equal(JSON.parse(getRes.stdout).cells.D1.value, "from-stdin");
  });

  it("eval", async () => {
    const file = path.join(tmpDir, "test.xlsx");
    const res = await sendCommand(
      socketPath,
      ["eval", file, 'console.log("hi"); return sheet.getValue(0, 0);'],
      tmpDir,
    );
    assert.equal(res.exitCode, 0);
    const data = JSON.parse(res.stdout);
    assert.equal(data.result, "Name");
    assert.deepStrictEqual(data.logs, ["hi"]);
  });

  it("error gives exitCode 1", async () => {
    const res = await sendCommand(
      socketPath,
      ["info", "/tmp/nonexistent-hsx-daemon-test.xlsx"],
      tmpDir,
    );
    assert.equal(res.exitCode, 1);
    assert.ok(res.stderr.includes("error"));
  });

  it("each response contains only its own output", async () => {
    const file = path.join(tmpDir, "test.xlsx");

    const res1 = await sendCommand(
      socketPath,
      ["get", file, "A1", "--no-styles"],
      tmpDir,
    );
    const res2 = await sendCommand(
      socketPath,
      ["get", file, "B2", "--no-styles"],
      tmpDir,
    );

    // Each response is exactly one JSON line — no leakage
    assert.equal(res1.stdout.trim().split("\n").length, 1);
    assert.equal(res2.stdout.trim().split("\n").length, 1);
    assert.equal(res1.stderr, "");
    assert.equal(res2.stderr, "");

    assert.equal(JSON.parse(res1.stdout).cells.A1.value, "Name");
    assert.equal(JSON.parse(res2.stdout).cells.B2.value, 4);
  });

  it("pipelined requests on single connection stay isolated", async () => {
    const file = path.join(tmpDir, "test.xlsx");

    const responses = await sendPipelined(socketPath, [
      { argv: ["get", file, "A1", "--no-styles"], cwd: tmpDir },
      { argv: ["get", file, "B2", "--no-styles"], cwd: tmpDir },
      { argv: ["get", file, "A2", "--no-styles"], cwd: tmpDir },
    ]);

    assert.equal(responses.length, 3);
    for (const r of responses) {
      assert.equal(r.exitCode, 0);
      assert.equal(r.stderr, "");
      assert.equal(r.stdout.trim().split("\n").length, 1);
    }

    assert.equal(JSON.parse(responses[0].stdout).cells.A1.value, "Name");
    assert.equal(JSON.parse(responses[1].stdout).cells.B2.value, 4);
    assert.equal(JSON.parse(responses[2].stdout).cells.A2.value, "Alice");
  });

  it("file cache serves subsequent reads", async () => {
    const file = path.join(tmpDir, "test.xlsx");

    const statusBefore = JSON.parse(
      (await sendCommand(socketPath, ["daemon", "status"], tmpDir)).stdout,
    );

    // This file was already opened above — should be cached
    await sendCommand(
      socketPath,
      ["get", file, "A1", "--no-styles"],
      tmpDir,
    );

    const statusAfter = JSON.parse(
      (await sendCommand(socketPath, ["daemon", "status"], tmpDir)).stdout,
    );

    // Cache size should not have grown (file was already cached)
    assert.equal(statusBefore.cachedFiles, statusAfter.cachedFiles);
  });
});
