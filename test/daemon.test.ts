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
  timeoutMs?: number,
): Promise<DaemonResponse> {
  return new Promise((resolve, reject) => {
    const socket = connect({ path: socketPath }, () => {
      socket.write(`${JSON.stringify({ argv, cwd, stdin, timeoutMs })}\n`);
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
    stdio: ["ignore", "ignore", "pipe", "ipc"],
    env: { ...process.env, HSX_SOCKET_PATH: socketPath },
  });

  await new Promise<void>((resolve, reject) => {
    const timeout = setTimeout(() => {
      daemonProc.kill();
      reject(new Error("Daemon failed to start within 30s"));
    }, 30_000);

    daemonProc.on("message", (msg: unknown) => {
      const m = msg as Record<string, unknown>;
      if (m?.ready) {
        clearTimeout(timeout);
        resolve();
      }
    });

    daemonProc.on("error", (err) => {
      clearTimeout(timeout);
      reject(err);
    });
    daemonProc.on("exit", (code) => {
      clearTimeout(timeout);
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
    assert.equal(status.writeThrough, false);
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
      success: true,
      written: 6,
      range: "A1:B3",
      messages: [],
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

  it("request timeout aborts long-running command", async () => {
    const file = path.join(tmpDir, "timeout-abort.xlsx");
    await sendCommand(socketPath, ["create", file], tmpDir);

    const startedAt = Date.now();
    const res = await sendCommand(
      socketPath,
      [
        "eval",
        file,
        "await new Promise((r) => setTimeout(r, 1500)); return 1;",
      ],
      tmpDir,
      undefined,
      200,
    );
    const elapsedMs = Date.now() - startedAt;

    assert.equal(res.exitCode, 1);
    assert.ok(res.stderr.includes("timed out"));
    assert.ok(elapsedMs < 5_000);

    const status = await sendCommand(socketPath, ["daemon", "status"], tmpDir);
    assert.equal(status.exitCode, 0);
  });

  it("timed-out request does not block subsequent requests", async () => {
    const file = path.join(tmpDir, "timeout-unblock.xlsx");
    await sendCommand(socketPath, ["create", file], tmpDir);

    const longRequest = sendCommand(
      socketPath,
      [
        "eval",
        file,
        "await new Promise((r) => setTimeout(r, 1500)); return 1;",
      ],
      tmpDir,
      undefined,
      200,
    );

    await new Promise((r) => setTimeout(r, 50));
    const startedAt = Date.now();
    const statusRequest = sendCommand(socketPath, ["daemon", "status"], tmpDir);

    const [longRes, statusRes] = await Promise.all([longRequest, statusRequest]);
    const elapsedMs = Date.now() - startedAt;

    assert.equal(longRes.exitCode, 1);
    assert.ok(longRes.stderr.includes("timed out"));
    assert.equal(statusRes.exitCode, 0);
    assert.ok(elapsedMs < 5_000);
  });

  it("client disconnect aborts in-flight request", async () => {
    const file = path.join(tmpDir, "disconnect-abort.xlsx");
    await sendCommand(socketPath, ["create", file], tmpDir);

    await new Promise<void>((resolve, reject) => {
      const socket = connect({ path: socketPath }, () => {
        socket.write(
          `${JSON.stringify({
            argv: [
              "eval",
              file,
              "await new Promise((r) => setTimeout(r, 4000)); return 1;",
            ],
            cwd: tmpDir,
            timeoutMs: 5_000,
          })}\n`,
        );
        setTimeout(() => {
          socket.destroy();
          resolve();
        }, 50);
      });

      socket.on("error", (err) => {
        if ((err as NodeJS.ErrnoException).code === "ECONNRESET") {
          resolve();
          return;
        }
        reject(err);
      });
    });

    await new Promise((r) => setTimeout(r, 250));

    const status = await sendCommand(socketPath, ["daemon", "status"], tmpDir);
    assert.equal(status.exitCode, 0);
  });

  it("flush writes dirty files", async () => {
    const file = path.join(tmpDir, "flush.xlsx");
    await sendCommand(socketPath, ["create", file], tmpDir);
    await sendCommand(
      socketPath,
      ["set", file, "A1", '[[{"value":"flush-me"}]]'],
      tmpDir,
    );

    const before = JSON.parse(
      (await sendCommand(socketPath, ["daemon", "status"], tmpDir)).stdout,
    );
    assert.ok(before.dirtyFiles >= 1);

    const flush = await sendCommand(socketPath, ["daemon", "flush"], tmpDir);
    assert.equal(flush.exitCode, 0);
    const flushData = JSON.parse(flush.stdout);
    assert.ok(flushData.flushed >= 1);
    assert.equal(flushData.dirtyFiles, 0);

    const after = JSON.parse(
      (await sendCommand(socketPath, ["daemon", "status"], tmpDir)).stdout,
    );
    assert.equal(after.dirtyFiles, 0);
  });

  it("keeps previously dirty state after command error", async () => {
    const file = path.join(tmpDir, "dirty-survives-error.xlsx");
    await sendCommand(socketPath, ["create", file], tmpDir);
    await sendCommand(
      socketPath,
      ["set", file, "A1", '[[{"value":"persist"}]]'],
      tmpDir,
    );

    const evalFail = await sendCommand(
      socketPath,
      ["eval", file, 'throw new Error("boom")'],
      tmpDir,
    );
    assert.equal(evalFail.exitCode, 1);

    const flush = await sendCommand(socketPath, ["daemon", "flush"], tmpDir);
    assert.equal(flush.exitCode, 0);

    const get = await sendCommand(
      socketPath,
      ["get", file, "A1", "--no-styles"],
      tmpDir,
    );
    assert.equal(get.exitCode, 0);
    assert.equal(JSON.parse(get.stdout).cells.A1.value, "persist");
  });

  it("file cache serves subsequent reads", async () => {
    const file = path.join(tmpDir, "test.xlsx");

    const statusBefore = JSON.parse(
      (await sendCommand(socketPath, ["daemon", "status"], tmpDir)).stdout,
    );

    await sendCommand(
      socketPath,
      ["get", file, "A1", "--no-styles"],
      tmpDir,
    );

    const statusAfter = JSON.parse(
      (await sendCommand(socketPath, ["daemon", "status"], tmpDir)).stdout,
    );

    assert.equal(statusBefore.cachedFiles, statusAfter.cachedFiles);
  });

  it("flush failure returns structured error and daemon stays alive", async () => {
    const file = path.join(tmpDir, "flush-fail.xlsx");
    await sendCommand(socketPath, ["create", file], tmpDir);
    await sendCommand(
      socketPath,
      ["set", file, "A1", '[[{"value":"x"}]]'],
      tmpDir,
    );

    await fs.chmod(file, 0o444);

    const flush = await sendCommand(socketPath, ["daemon", "flush"], tmpDir);
    assert.equal(flush.exitCode, 1);
    assert.ok(flush.stderr.includes("error"));

    const status = await sendCommand(socketPath, ["daemon", "status"], tmpDir);
    assert.equal(status.exitCode, 0);
  });

  it("stop failure returns error and does not ack success", async () => {
    const file = path.join(tmpDir, "stop-fail.xlsx");
    await sendCommand(socketPath, ["create", file], tmpDir);
    await sendCommand(
      socketPath,
      ["set", file, "A1", '[[{"value":"x"}]]'],
      tmpDir,
    );

    await fs.chmod(file, 0o444);

    const stop = await sendCommand(socketPath, ["daemon", "stop"], tmpDir);
    assert.equal(stop.exitCode, 1);
    assert.equal(stop.stdout, "");
    assert.ok(stop.stderr.includes("error"));

    const status = await sendCommand(socketPath, ["daemon", "status"], tmpDir);
    assert.equal(status.exitCode, 0);
  });

  it("SIGTERM exits with code 1 when dirty flush fails", async () => {
    const file = path.join(tmpDir, "sigterm-fail.xlsx");
    await sendCommand(socketPath, ["create", file], tmpDir);
    await sendCommand(
      socketPath,
      ["set", file, "A1", '[[{"value":"x"}]]'],
      tmpDir,
    );

    await fs.chmod(file, 0o444);

    const status = JSON.parse(
      (await sendCommand(socketPath, ["daemon", "status"], tmpDir)).stdout,
    );
    process.kill(status.pid, "SIGTERM");

    const exitCode = await new Promise<number | null>((resolve, reject) => {
      const timeout = setTimeout(() => {
        reject(new Error("daemon did not exit after SIGTERM"));
      }, 5_000);

      daemonProc.once("exit", (code) => {
        clearTimeout(timeout);
        resolve(code);
      });
    });

    assert.equal(exitCode, 1);
  });
});
