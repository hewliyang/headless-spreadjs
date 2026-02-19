import assert from "node:assert/strict";
import { type ChildProcess, execFile, spawn } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { promisify } from "node:util";
import { afterAll, beforeAll, describe, it } from "vitest";

const exec = promisify(execFile);
const CLI = path.resolve("src/cli/index.ts");
const DAEMON_ENTRY = path.resolve("src/cli/daemon-entry.ts");

let tmpDir: string;
let testFile: string;
let socketPath: string;
let daemonProc: ChildProcess;
let testEnv: NodeJS.ProcessEnv;

beforeAll(async () => {
  tmpDir = await fs.mkdtemp(path.join(os.tmpdir(), "hsx-cli-test-"));
  testFile = path.join(tmpDir, "test.xlsx");
  socketPath = path.join(tmpDir, "cli-test-daemon.sock");
  testEnv = { ...process.env, HSX_SOCKET_PATH: socketPath };

  daemonProc = spawn("tsx", [DAEMON_ENTRY], {
    stdio: ["ignore", "ignore", "pipe", "ipc"],
    env: testEnv,
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
    await exec("tsx", [CLI, "daemon", "stop"], {
      env: testEnv,
      timeout: 10_000,
    });
  } catch {}
  await new Promise((r) => setTimeout(r, 200));
  try {
    daemonProc?.kill();
  } catch {}
  await fs.rm(tmpDir, { recursive: true, force: true });
});

function hsx(
  args: string[],
  input?: string,
): Promise<{ stdout: string; stderr: string }> {
  if (input !== undefined) {
    return new Promise((resolve, reject) => {
      const proc = spawn("tsx", [CLI, ...args], {
        stdio: ["pipe", "pipe", "pipe"],
        env: testEnv,
      });
      let stdout = "";
      let stderr = "";
      proc.stdout.on("data", (d: Buffer) => (stdout += d.toString()));
      proc.stderr.on("data", (d: Buffer) => (stderr += d.toString()));
      proc.on("close", (code) => {
        if (code === 0) resolve({ stdout, stderr });
        else
          reject(Object.assign(new Error(`exit ${code}`), { stdout, stderr }));
      });
      proc.stdin.end(input);
    });
  }
  return exec("tsx", [CLI, ...args], { timeout: 30_000, env: testEnv });
}

function hsxRaw(
  args: string[],
  env: NodeJS.ProcessEnv,
  input?: string,
): Promise<{ code: number; stdout: string; stderr: string }> {
  return new Promise((resolve) => {
    const proc = spawn("tsx", [CLI, ...args], {
      stdio: ["pipe", "pipe", "pipe"],
      env,
    });
    let stdout = "";
    let stderr = "";
    proc.stdout.on("data", (d: Buffer) => (stdout += d.toString()));
    proc.stderr.on("data", (d: Buffer) => (stderr += d.toString()));
    proc.on("close", (code) => {
      resolve({ code: code ?? 1, stdout, stderr });
    });
    if (input !== undefined) {
      proc.stdin.end(input);
    } else {
      proc.stdin.end();
    }
  });
}

describe("cli", () => {
  it("create + info", async () => {
    const { stdout } = await hsx(["create", testFile]);
    assert.deepStrictEqual(JSON.parse(stdout), { created: testFile });

    const { stdout: infoOut } = await hsx(["info", testFile]);
    const info = JSON.parse(infoOut);
    assert.equal(info.sheets.length, 1);
    assert.equal(info.sheets[0].name, "Sheet1");
    assert.equal(info.sheets[0].active, true);
  });

  it("set + get round-trip", async () => {
    const cells = [
      [{ value: "Name" }, { value: "Qty" }],
      [{ value: "Alice" }, { value: 4 }],
      [{ value: "Bob" }, { formula: "=B2+1" }],
    ];

    const { stdout: setOut } = await hsx([
      "set",
      testFile,
      "A1:B3",
      JSON.stringify(cells),
    ]);
    assert.deepStrictEqual(JSON.parse(setOut), {
      written: 6,
      range: "A1:B3",
    });

    const { stdout: getOut } = await hsx([
      "get",
      testFile,
      "A1:B3",
      "--no-styles",
    ]);
    const data = JSON.parse(getOut);
    assert.equal(data.cellCount, 6);
    assert.equal(data.cells.A1.value, "Name");
    assert.equal(data.cells.B2.value, 4);
    assert.equal(data.cells.B3.value, 5);
    assert.equal(data.cells.B3.formula, "B2+1");
  });

  it("get emits complete JSON for large ranges", async () => {
    const bigFile = path.join(tmpDir, "large-get.xlsx");
    await hsx(["create", bigFile]);
    await hsx([
      "eval",
      bigFile,
      "for (let r = 0; r < 120; r++) { for (let c = 0; c < 12; c++) { sheet.setValue(r, c, 'x'.repeat(120)); } }",
    ]);

    const res = await hsxRaw(["get", bigFile, "A1:L120", "--no-styles"], testEnv);
    assert.equal(res.code, 0);
    assert.ok(res.stdout.length > 65_536);

    const data = JSON.parse(res.stdout);
    assert.equal(data.cellCount, 120 * 12);
    assert.equal(data.cells.A1.value.length, 120);
  });

  it("csv output", async () => {
    const { stdout } = await hsx(["csv", testFile, "A1:B3"]);
    const lines = stdout.trim().split("\n");
    assert.equal(lines.length, 3);
    assert.equal(lines[0], "Name,Qty");
    assert.equal(lines[1], "Alice,4");
    assert.equal(lines[2], "Bob,5");
  });

  it("clear", async () => {
    await hsx(["clear", testFile, "B2"]);
    const { stdout } = await hsx(["get", testFile, "B2"]);
    const data = JSON.parse(stdout);
    assert.equal(data.cellCount, 0);

    const { stdout: b3Out } = await hsx(["get", testFile, "B3"]);
    const b3 = JSON.parse(b3Out);
    assert.equal(b3.cells.B3.value, 1);
  });

  it("sheet operations", async () => {
    await hsx(["sheet", testFile, "create", "Revenue"]);
    const { stdout: listOut } = await hsx(["sheet", testFile, "list"]);
    const sheets = JSON.parse(listOut).sheets;
    assert.equal(sheets.length, 2);
    assert.equal(sheets[1].name, "Revenue");

    await hsx(["sheet", testFile, "rename", "Revenue", "Sales"]);
    const { stdout: listOut2 } = await hsx(["sheet", testFile, "list"]);
    assert.equal(JSON.parse(listOut2).sheets[1].name, "Sales");

    await hsx(["sheet", testFile, "delete", "Sales"]);
    const { stdout: listOut3 } = await hsx(["sheet", testFile, "list"]);
    assert.equal(JSON.parse(listOut3).sheets.length, 1);
  });

  it("eval with return value and console.log", async () => {
    const { stdout } = await hsx([
      "eval",
      testFile,
      'console.log("hello"); return sheet.getValue(0, 0);',
    ]);
    const result = JSON.parse(stdout);
    assert.equal(result.result, "Name");
    assert.deepStrictEqual(result.logs, ["hello"]);
  });

  it("eval can create charts", async () => {
    const cells = [
      [{ value: "Month" }, { value: "Rev" }],
      [{ value: "Jan" }, { value: 10 }],
      [{ value: "Feb" }, { value: 20 }],
    ];
    await hsx(["set", testFile, "A1:B3", JSON.stringify(cells)]);

    await hsx([
      "eval",
      testFile,
      "sheet.charts.add('C1', GC.Spread.Sheets.Charts.ChartType.line, 0, 80, 400, 240, 'A1:B3');",
    ]);

    const { stdout } = await hsx(["objects", testFile]);
    const objs = JSON.parse(stdout).objects;
    assert.equal(objs.length, 1);
    assert.equal(objs[0].type, "chart");
    assert.equal(objs[0].name, "C1");
  });

  it("set with styles round-trips", async () => {
    const cells = [
      [
        {
          value: "Bold",
          cellStyles: {
            fontWeight: "bold",
            backgroundColor: "#FF0000",
          },
        },
      ],
    ];
    await hsx(["set", testFile, "D1", JSON.stringify(cells)]);

    const { stdout } = await hsx(["get", testFile, "D1"]);
    const data = JSON.parse(stdout);
    assert.equal(data.cells.D1.styles.bold, true);
    assert.equal(data.cells.D1.styles.backgroundColor.toLowerCase(), "#ff0000");
  });

  it("set via stdin", async () => {
    const cells = [[{ value: "stdin-test" }]];
    const { stdout } = await hsx(
      ["set", testFile, "E1"],
      JSON.stringify(cells),
    );
    assert.deepStrictEqual(JSON.parse(stdout), {
      written: 1,
      range: "E1",
    });

    const { stdout: getOut } = await hsx(["get", testFile, "E1"]);
    assert.equal(JSON.parse(getOut).cells.E1.value, "stdin-test");
  });

  it("set via stdin with '-' sentinel", async () => {
    const cells = [[{ value: "stdin-dash" }]];
    const { stdout } = await hsx(
      ["set", testFile, "E2", "-"],
      JSON.stringify(cells),
    );
    assert.deepStrictEqual(JSON.parse(stdout), {
      written: 1,
      range: "E2",
    });

    const { stdout: getOut } = await hsx(["get", testFile, "E2"]);
    assert.equal(JSON.parse(getOut).cells.E2.value, "stdin-dash");
  });

  it("search", async () => {
    const { stdout } = await hsx(["search", testFile, "Month"]);
    const data = JSON.parse(stdout);
    assert.ok(data.totalFound >= 1);
    assert.equal(data.matches[0].value, "Month");
  });

  it("search with regex", async () => {
    const { stdout } = await hsx(["search", testFile, "^(Jan|Feb)", "--regex"]);
    const data = JSON.parse(stdout);
    assert.equal(data.totalFound, 2);
  });

  it("copy range", async () => {
    await hsx(["copy", testFile, "A1:B1", "A10:B10"]);
    const { stdout } = await hsx(["get", testFile, "A10:B10", "--no-styles"]);
    const data = JSON.parse(stdout);
    assert.equal(data.cells.A10.value, "Month");
    assert.equal(data.cells.B10.value, "Rev");
  });

  it("insert and delete rows", async () => {
    await hsx(["rc", testFile, "insert", "rows", "--ref", "2"]);
    const { stdout } = await hsx(["get", testFile, "A2", "--no-styles"]);
    const data = JSON.parse(stdout);
    assert.equal(data.cellCount, 0);

    await hsx(["rc", testFile, "delete", "rows", "--ref", "2"]);
    const { stdout: after } = await hsx(["get", testFile, "A2", "--no-styles"]);
    const afterData = JSON.parse(after);
    assert.ok(afterData.cells.A2);
  });

  it("resize columns", async () => {
    await hsx(["resize", testFile, "--columns", "A:B", "--width", "120"]);
  });

  it("daemon flush persists buffered writes to disk", async () => {
    await hsx(["set", testFile, "F1", '[[{"value":"buffered"}]]']);

    const { stdout: before } = await hsx([
      "--no-daemon",
      "get",
      testFile,
      "F1",
      "--no-styles",
    ]);
    assert.equal(JSON.parse(before).cellCount, 0);

    const { stdout: flushOut } = await hsx(["daemon", "flush"]);
    const flush = JSON.parse(flushOut);
    assert.ok(flush.flushed >= 1);
    assert.equal(flush.dirtyFiles, 0);

    const { stdout: after } = await hsx([
      "--no-daemon",
      "get",
      testFile,
      "F1",
      "--no-styles",
    ]);
    assert.equal(JSON.parse(after).cells.F1.value, "buffered");
  });

  it("supports global --timeout option", async () => {
    const { stdout } = await hsx(["--timeout", "30", "info", testFile]);
    const info = JSON.parse(stdout);
    assert.equal(info.sheets.length, 1);
  });

  it("rejects invalid --timeout value", async () => {
    const res1 = await hsxRaw(["--timeout", "abc", "info", testFile], testEnv);
    assert.equal(res1.code, 1);
    assert.ok(res1.stderr.includes("error"));

    const res2 = await hsxRaw(["--timeout", "2s", "info", testFile], testEnv);
    assert.equal(res2.code, 1);
    assert.ok(res2.stderr.includes("error"));
  });

  it("times out long-running direct commands", async () => {
    const res = await hsxRaw(
      [
        "--no-daemon",
        "--timeout",
        "1",
        "eval",
        testFile,
        "await new Promise((r) => setTimeout(r, 1500)); return 1;",
      ],
      testEnv,
    );
    assert.equal(res.code, 1);
    assert.ok(res.stderr.includes("timed out"));
  });

  it("daemon status on missing socket does not auto-start", async () => {
    const missingEnv = {
      ...testEnv,
      HSX_SOCKET_PATH: path.join(tmpDir, "missing-status.sock"),
    };

    const res = await hsxRaw(["daemon", "status"], missingEnv);
    assert.equal(res.code, 0);
    assert.deepStrictEqual(JSON.parse(res.stdout), { running: false });
  });

  it("daemon stop on missing socket does not auto-start", async () => {
    const missingEnv = {
      ...testEnv,
      HSX_SOCKET_PATH: path.join(tmpDir, "missing-stop.sock"),
    };

    const res = await hsxRaw(["daemon", "stop"], missingEnv);
    assert.equal(res.code, 1);
    assert.deepStrictEqual(JSON.parse(res.stdout), {
      error: "No daemon running",
    });
  });

  it("write-through mode saves immediately", async () => {
    const wtDir = await fs.mkdtemp(path.join(tmpDir, "wt-"));
    const wtSocket = path.join(wtDir, "daemon.sock");
    const wtFile = path.join(wtDir, "wt.xlsx");
    const wtEnv = {
      ...process.env,
      HSX_SOCKET_PATH: wtSocket,
      HSX_WRITE_THROUGH: "1",
    };

    let wtDaemon: ChildProcess | null = null;

    try {
      wtDaemon = spawn("tsx", [DAEMON_ENTRY], {
        stdio: ["ignore", "ignore", "pipe", "ipc"],
        env: wtEnv,
      });

      await new Promise<void>((resolve, reject) => {
        const timeout = setTimeout(() => {
          wtDaemon?.kill();
          reject(new Error("write-through daemon failed to start"));
        }, 30_000);

        wtDaemon?.on("message", (msg: unknown) => {
          const m = msg as Record<string, unknown>;
          if (m?.ready) {
            clearTimeout(timeout);
            resolve();
          }
        });

        wtDaemon?.on("error", (err) => {
          clearTimeout(timeout);
          reject(err);
        });
        wtDaemon?.on("exit", (code) => {
          clearTimeout(timeout);
          reject(new Error(`write-through daemon exited early with code ${code}`));
        });
      });

      await exec("tsx", [CLI, "create", wtFile], { env: wtEnv, timeout: 30_000 });
      await exec("tsx", [CLI, "set", wtFile, "A1", '[[{"value":"sync"}]]'], {
        env: wtEnv,
        timeout: 30_000,
      });

      const { stdout } = await exec(
        "tsx",
        [CLI, "--no-daemon", "get", wtFile, "A1", "--no-styles"],
        { env: wtEnv, timeout: 30_000 },
      );
      assert.equal(JSON.parse(stdout).cells.A1.value, "sync");

      const { stdout: statusOut } = await exec("tsx", [CLI, "daemon", "status"], {
        env: wtEnv,
        timeout: 30_000,
      });
      assert.equal(JSON.parse(statusOut).writeThrough, true);
    } finally {
      await hsxRaw(["daemon", "stop"], wtEnv);
      try {
        wtDaemon?.kill();
      } catch {}
      await fs.rm(wtDir, { recursive: true, force: true });
    }
  });

  it("errors on nonexistent file", async () => {
    try {
      await hsx(["info", "/tmp/nonexistent-hsx-test.xlsx"]);
      assert.fail("should have thrown");
    } catch (err) {
      const e = err as Error & { stderr: string };
      assert.ok(e.stderr.includes("error"));
    }
  });
});
