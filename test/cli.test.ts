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
      success: true,
      written: 6,
      range: "A1:B3",
      messages: [],
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

  it("set value over existing formula clears formula", async () => {
    const file = path.join(tmpDir, "ghost-formula.xlsx");
    await hsx(["create", file]);

    await hsx(["set", file, "A1", '[[{"formula":"=1+1"}]]']);
    await hsx(["set", file, "A1", '[[{"value":99}]]']);

    const { stdout } = await hsx(["get", file, "A1", "--no-styles"]);
    const data = JSON.parse(stdout);
    assert.equal(data.cells.A1.value, 99);
    assert.equal(data.cells.A1.formula, undefined);
  });

  it("get emits complete JSON for large ranges", async () => {
    const bigFile = path.join(tmpDir, "large-get.xlsx");
    await hsx(["create", bigFile]);
    await hsx([
      "eval",
      bigFile,
      "for (let r = 0; r < 120; r++) { for (let c = 0; c < 12; c++) { sheet.setValue(r, c, 'x'.repeat(120)); } }",
    ]);

    const res = await hsxRaw(
      ["get", bigFile, "A1:L120", "--no-styles"],
      testEnv,
    );
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

  it("csv --mode formula", async () => {
    const { stdout } = await hsx([
      "csv",
      testFile,
      "A1:B3",
      "--mode",
      "formula",
    ]);
    const lines = stdout.trim().split("\n");
    assert.equal(lines[0], "Name,Qty");
    assert.equal(lines[1], "Alice,4");
    assert.equal(lines[2], "Bob,=B2+1");
  });

  it("csv --formulas shorthand", async () => {
    const { stdout } = await hsx(["csv", testFile, "A1:B3", "--formulas"]);
    const lines = stdout.trim().split("\n");
    assert.equal(lines[2], "Bob,=B2+1");
  });

  it("csv --mode both", async () => {
    const { stdout } = await hsx(["csv", testFile, "A1:B3", "--mode", "both"]);
    const lines = stdout.trim().split("\n");
    assert.equal(lines[2], "Bob,5 | =B2+1");
  });

  it("csv accepts flags between positional args", async () => {
    const { stdout } = await hsx(["csv", testFile, "--formulas", "A1:B3"]);
    const lines = stdout.trim().split("\n");
    assert.equal(lines[2], "Bob,=B2+1");
  });

  it("csv supports --mode=both", async () => {
    const { stdout } = await hsx(["csv", testFile, "A1:B3", "--mode=both"]);
    const lines = stdout.trim().split("\n");
    assert.equal(lines[2], "Bob,5 | =B2+1");
  });

  it("csv errors on invalid mode", async () => {
    try {
      await hsx(["csv", testFile, "A1:B3", "--mode", "wat"]);
      assert.fail("should have thrown");
    } catch (err) {
      const e = err as Error & { stderr: string };
      assert.ok(e.stderr.includes("Invalid --mode value"));
    }
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

  it("eval range(ref) resolves A1 on active + named sheets", async () => {
    const rangeFile = path.join(tmpDir, "eval-range-helper.xlsx");
    await hsx(["create", rangeFile]);
    await hsx(["sheet", rangeFile, "create", "Data Sheet"]);

    const { stdout } = await hsx([
      "eval",
      rangeFile,
      `
      workbook.getSheetFromName("Data Sheet").setValue(0, 0, "seed");
      range("B2").value(42);
      range("'Data Sheet'!A1").value("ok");
      return {
        activeB2: range("B2").value(),
        dataA1: range("'Data Sheet'!A1").value()
      };
      `,
    ]);

    const result = JSON.parse(stdout);
    assert.equal(result.result.activeB2, 42);
    assert.equal(result.result.dataA1, "ok");
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

  it("set with style.fontStyle object round-trips", async () => {
    const cells = [
      [
        {
          value: "Bold",
          style: {
            fontStyle: { bold: true, italic: true, underline: true },
            backColor: "#FF0000",
          },
        },
      ],
    ];
    await hsx(["set", testFile, "D1", JSON.stringify(cells)]);

    const { stdout } = await hsx(["get", testFile, "D1"]);
    const data = JSON.parse(stdout);
    assert.equal(data.cells.D1.style.fontStyle.bold, true);
    assert.equal(data.cells.D1.style.fontStyle.italic, true);
    assert.equal(data.cells.D1.style.fontStyle.underline, true);
    assert.equal(data.cells.D1.style.backColor.toLowerCase(), "#ff0000");
  });

  it("set auto-expands single-cell refs to input shape", async () => {
    const cells = [[{ value: "r1" }], [{ value: "r2" }]];
    const { stdout } = await hsx([
      "set",
      testFile,
      "A20",
      JSON.stringify(cells),
    ]);

    const result = JSON.parse(stdout);
    assert.equal(result.success, true);
    assert.equal(result.range, "A20:A21");
    assert.equal(result.written, 2);
    assert.equal(result.messages.length, 1);
    assert.match(result.messages[0], /Adjusted range from A20 to A20:A21/);

    const { stdout: getOut } = await hsx(["get", testFile, "A20:A21"]);
    const data = JSON.parse(getOut);
    assert.equal(data.cells.A20.value, "r1");
    assert.equal(data.cells.A21.value, "r2");
  });

  it("set supports style.fontStyle.bold", async () => {
    const cells = [[{ value: "Bold alias", style: { fontStyle: { bold: true } } }]];
    await hsx(["set", testFile, "G1", JSON.stringify(cells)]);

    const { stdout } = await hsx(["get", testFile, "G1"]);
    const data = JSON.parse(stdout);
    assert.equal(data.cells.G1.style.fontStyle.bold, true);
  });

  it("set via stdin", async () => {
    const cells = [[{ value: "stdin-test" }]];
    const { stdout } = await hsx(
      ["set", testFile, "E1"],
      JSON.stringify(cells),
    );
    assert.deepStrictEqual(JSON.parse(stdout), {
      success: true,
      written: 1,
      range: "E1",
      messages: [],
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
      success: true,
      written: 1,
      range: "E2",
      messages: [],
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

  it("search respects sparse used-range bounds", async () => {
    const sparseFile = path.join(tmpDir, "search-sparse.xlsx");
    await hsx(["create", sparseFile]);
    await hsx(["eval", sparseFile, 'sheet.setValue(999999, 0, "tail")']);

    const res = await hsxRaw(["--timeout", "2", "search", sparseFile, "nomatch"], testEnv);
    assert.equal(res.code, 0, res.stderr);

    const data = JSON.parse(res.stdout) as { totalFound: number };
    assert.equal(data.totalFound, 0);
  });

  it("copy range", async () => {
    await hsx(["copy", testFile, "A1:B1", "A10:B10"]);
    const { stdout } = await hsx(["get", testFile, "A10:B10", "--no-styles"]);
    const data = JSON.parse(stdout);
    assert.equal(data.cells.A10.value, "Month");
    assert.equal(data.cells.B10.value, "Rev");
  });

  it("diff reports changed values and formulas", async () => {
    const left = path.join(tmpDir, "diff-left.xlsx");
    const right = path.join(tmpDir, "diff-right.xlsx");

    await hsx(["create", left]);
    await hsx(["create", right]);

    await hsx([
      "set",
      left,
      "A1:B2",
      '[[{"value":"Name"},{"value":1}],[{"value":"Alice"},{"formula":"=B1+1"}]]',
    ]);
    await hsx([
      "set",
      right,
      "A1:B2",
      '[[{"value":"Name"},{"value":2}],[{"value":"Alicia"},{"formula":"=B1+2"}]]',
    ]);

    const { stdout } = await hsx(["diff", left, right]);
    const result = JSON.parse(stdout) as {
      changedCells: number;
      outputMode: string;
      summary: string;
      diffs: Array<{
        cell: string;
        left: { value: unknown; formula: string | null };
        right: { value: unknown; formula: string | null };
      }>;
    };

    assert.equal(result.changedCells, 3);
    assert.equal(result.outputMode, "inline");
    assert.ok(result.summary.includes("changed cells"));

    const byCell = new Map<string, (typeof result.diffs)[number]>(
      result.diffs.map((d) => [d.cell, d] as const),
    );

    assert.equal(byCell.get("A2")?.left.value, "Alice");
    assert.equal(byCell.get("A2")?.right.value, "Alicia");
    assert.equal(byCell.get("B1")?.left.value, 1);
    assert.equal(byCell.get("B1")?.right.value, 2);
    assert.equal(byCell.get("B2")?.left.formula, "B1+1");
    assert.equal(byCell.get("B2")?.right.formula, "B1+2");
  });

  it("diff spills large output to tmp file", async () => {
    const left = path.join(tmpDir, "diff-big-left.xlsx");
    const right = path.join(tmpDir, "diff-big-right.xlsx");

    await hsx(["create", left]);
    await hsx(["create", right]);

    await hsx([
      "eval",
      left,
      "for (let r = 0; r < 6; r++) { for (let c = 0; c < 6; c++) { sheet.setValue(r, c, `L-${r}-${c}`); } }",
    ]);
    await hsx([
      "eval",
      right,
      "for (let r = 0; r < 6; r++) { for (let c = 0; c < 6; c++) { sheet.setValue(r, c, `R-${r}-${c}`); } }",
    ]);

    const { stdout } = await hsx([
      "diff",
      left,
      right,
      "--inline-limit",
      "5",
      "--preview-limit",
      "3",
    ]);
    const result = JSON.parse(stdout);

    assert.equal(result.changedCells, 36);
    assert.equal(result.outputMode, "tmpfile");
    assert.equal(result.diffs.length, 3);
    assert.ok(typeof result.diffFile === "string");

    const diffFile = await fs.readFile(result.diffFile, "utf8");
    const lines = diffFile.trim().split("\n");
    assert.equal(lines.length, 36);
    assert.ok(lines[0].includes('"sheet":"Sheet1"'));
  });

  it("deps supports recursive tracing with hop counts", async () => {
    const traceFile = path.join(tmpDir, "trace-deps.xlsx");
    await hsx(["create", traceFile]);
    await hsx(["sheet", traceFile, "create", "Sheet2"]);
    await hsx(["sheet", traceFile, "create", "Sheet3"]);

    await hsx([
      "eval",
      traceFile,
      'workbook.getSheetFromName("Sheet1").setValue(0, 0, 2); workbook.getSheetFromName("Sheet2").setFormula(0, 0, "Sheet1!A1*5"); workbook.getSheetFromName("Sheet3").setFormula(0, 0, "Sheet2!A1+1");',
    ]);

    const { stdout } = await hsx([
      "deps",
      traceFile,
      "Sheet3!A1",
      "--recursive",
    ]);
    const result = JSON.parse(stdout) as {
      dependencies: Array<{ sheet: string; ref: string; hop: number }>;
    };

    const byHop = new Map(
      result.dependencies.map((d) => [`${d.sheet}!${d.ref}`, d.hop]),
    );
    assert.equal(byHop.get("Sheet2!A1"), 1);
    assert.equal(byHop.get("Sheet1!A1"), 2);
  });

  it("refs defaults to one-hop", async () => {
    const traceFile = path.join(tmpDir, "trace-deps.xlsx");
    const { stdout } = await hsx(["refs", traceFile, "Sheet1!A1"]);
    const result = JSON.parse(stdout) as {
      references: Array<{ sheet: string; cell: string; hop: number }>;
    };

    const refsSet = new Set(
      result.references.map((r) => `${r.sheet}!${r.cell}`),
    );
    assert.ok(refsSet.has("Sheet2!A1"));
    assert.ok(!refsSet.has("Sheet3!A1"));
  });

  it("refs supports recursive multi-hop tracing", async () => {
    const traceFile = path.join(tmpDir, "trace-deps.xlsx");
    const { stdout } = await hsx([
      "refs",
      traceFile,
      "Sheet1!A1",
      "--recursive",
    ]);
    const result = JSON.parse(stdout) as {
      references: Array<{ sheet: string; cell: string; hop: number }>;
    };

    const byHop = new Map(
      result.references.map((r) => [`${r.sheet}!${r.cell}`, r.hop]),
    );
    assert.equal(byHop.get("Sheet2!A1"), 1);
    assert.equal(byHop.get("Sheet3!A1"), 2);
  });

  it("refs avoids sparse data-only used-range scans", async () => {
    const sparseFile = path.join(tmpDir, "trace-refs-sparse.xlsx");
    await hsx(["create", sparseFile]);
    await hsx([
      "eval",
      sparseFile,
      'sheet.setValue(999999, 0, "tail"); sheet.setValue(0, 0, 123);',
    ]);

    const res = await hsxRaw(
      ["--timeout", "2", "refs", sparseFile, "Sheet1!A1"],
      testEnv,
    );
    assert.equal(res.code, 0, res.stderr);

    const result = JSON.parse(res.stdout) as {
      references: Array<{ sheet: string; cell: string; hop: number }>;
      stats: { scannedFormulaCells: number; indexedFormulaCells: number };
    };

    assert.deepEqual(result.references, []);
    assert.equal(result.stats.scannedFormulaCells, 0);
    assert.equal(result.stats.indexedFormulaCells, 0);
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

  it("rc validates dimension names", async () => {
    const res = await hsxRaw(
      ["rc", testFile, "freeze", "banana", "--ref", "2"],
      testEnv,
    );
    assert.equal(res.code, 1);
    assert.ok(res.stderr.includes("Invalid rc dimension"));
  });

  it("rc insert row alias shifts cell data", async () => {
    const shiftFile = path.join(tmpDir, "rc-insert-row-alias.xlsx");
    await hsx(["create", shiftFile]);
    await hsx(["set", shiftFile, "A20", '[[{"value":"mark"}]]']);

    await hsx(["rc", shiftFile, "insert", "row", "--ref", "20", "--count", "2"]);

    const { stdout } = await hsx(["get", shiftFile, "A22", "--no-styles"]);
    const data = JSON.parse(stdout);
    assert.equal(data.cells.A22.value, "mark");
  });

  it("rc accepts sheet-qualified row refs", async () => {
    const shiftFile = path.join(tmpDir, "rc-insert-sheet-qualified.xlsx");
    await hsx(["create", shiftFile]);
    await hsx(["sheet", shiftFile, "create", "Data"]);
    await hsx(["set", shiftFile, "Data!A20", '[[{"value":"mark"}]]']);

    await hsx([
      "rc",
      shiftFile,
      "insert",
      "rows",
      "--ref",
      "Data!20",
      "--count",
      "2",
    ]);

    const { stdout } = await hsx(["get", shiftFile, "Data!A22", "--no-styles"]);
    const data = JSON.parse(stdout);
    assert.equal(data.cells.A22.value, "mark");
  });

  it("resize without explicit --rows limits to used range", async () => {
    const resizeFile = path.join(tmpDir, "resize-used-range.xlsx");
    await hsx(["create", resizeFile]);
    await hsx(["set", resizeFile, "A1", '[[{"value":"x"}]]']);

    const res = await hsxRaw(
      ["--timeout", "1", "resize", resizeFile, "--height", "20"],
      testEnv,
    );
    assert.equal(res.code, 0, res.stderr);
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
    // Auto-flush may have already persisted writes before this explicit flush.
    assert.ok(flush.flushed >= 0);
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

  it("parses --timeout even after command args", async () => {
    const res = await hsxRaw(
      ["--no-daemon", "eval", testFile, "--timeout"],
      testEnv,
    );
    assert.equal(res.code, 1);
    assert.ok(res.stderr.includes("Usage: --timeout <seconds>"));
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

  it("times out long-running daemon commands", async () => {
    const res = await hsxRaw(
      [
        "--timeout",
        "1",
        "eval",
        testFile,
        "await new Promise((r) => setTimeout(r, 1500)); return 1;",
      ],
      testEnv,
    );
    assert.equal(res.code, 1);
    assert.ok(
      res.stderr.includes("timed out") || res.stderr.includes("aborted"),
    );
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
          reject(
            new Error(`write-through daemon exited early with code ${code}`),
          );
        });
      });

      await exec("tsx", [CLI, "create", wtFile], {
        env: wtEnv,
        timeout: 30_000,
      });
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

      const { stdout: statusOut } = await exec(
        "tsx",
        [CLI, "daemon", "status"],
        {
          env: wtEnv,
          timeout: 30_000,
        },
      );
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
