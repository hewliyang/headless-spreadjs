import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { afterEach, beforeAll, describe, expect, it } from "vitest";
import { FileCache } from "../src/cli/file-cache.js";
import { runWithDaemonRuntime, withFile, withNewFile } from "../src/cli/context.js";
import { ExcelFile, init } from "../src/index.js";
import {
  clearHooks,
  createHookAPI,
  getDiscoveryErrors,
  runWithHooksDisabled,
  setCurrentCommand,
} from "../src/hooks.js";
import type { HookContext, PostCommandContext, PreCommandContext } from "../src/hooks.js";

beforeAll(async () => {
  await init();
});

afterEach(() => {
  clearHooks();
});

function makeCtx(overrides?: Partial<HookContext>): HookContext {
  const file = new ExcelFile();
  return {
    command: "set",
    args: ["model.xlsx", "A1:B2", "[[1,2]]"],
    filePath: "/tmp/model.xlsx",
    file,
    workbook: file.workbook,
    GC: {} as any,
    mutatedRanges: [],
    ...overrides,
  };
}

async function withTempDir<T>(fn: (dir: string) => Promise<T>): Promise<T> {
  const dir = await fs.mkdtemp(path.join(os.tmpdir(), "hsx-hooks-test-"));
  try {
    return await fn(dir);
  } finally {
    await fs.rm(dir, { recursive: true, force: true });
  }
}

describe("hook context has command + workbook info", () => {
  it("onOpen receives command, args, filePath, and workbook", async () => {
    const { runOnOpenHooks } = await import("../src/hooks.js");
    const hsx = createHookAPI();

    let seen: HookContext | null = null;
    hsx.on("onOpen", (ctx) => {
      seen = ctx;
    });

    const ctx = makeCtx();
    await runOnOpenHooks(ctx);

    expect(seen).not.toBeNull();
    expect(seen!.command).toBe("set");
    expect(seen!.args).toEqual(["model.xlsx", "A1:B2", "[[1,2]]"]);
    expect(seen!.filePath).toBe("/tmp/model.xlsx");
    expect(seen!.workbook).toBeDefined();
    expect(seen!.workbook.getSheetCount()).toBeGreaterThan(0);
  });

  it("preSave receives full context", async () => {
    const { runPreSaveHooks } = await import("../src/hooks.js");
    const hsx = createHookAPI();

    let seen: HookContext | null = null;
    hsx.on("preSave", (ctx) => {
      seen = ctx;
    });

    const ctx = makeCtx({ command: "eval", args: ["model.xlsx", "code"] });
    await runPreSaveHooks(ctx);

    expect(seen!.command).toBe("eval");
    expect(seen!.file).toBeDefined();
    expect(seen!.GC).toBeDefined();
  });
});

describe("workbook manipulation in hooks", () => {
  it("onOpen can modify workbook (e.g., hide gridlines)", async () => {
    const { runOnOpenHooks } = await import("../src/hooks.js");
    const hsx = createHookAPI();

    hsx.on("onOpen", (ctx) => {
      const sheet = ctx.workbook.getActiveSheet();
      sheet.options.gridline = {
        showVerticalGridline: false,
        showHorizontalGridline: false,
      };
    });

    const ctx = makeCtx();
    await runOnOpenHooks(ctx);

    const gridline = ctx.workbook.getActiveSheet().options.gridline;
    expect(gridline.showVerticalGridline).toBe(false);
    expect(gridline.showHorizontalGridline).toBe(false);
  });

  it("preSave can apply financial color coding (negative = red)", async () => {
    const { runPreSaveHooks } = await import("../src/hooks.js");
    const hsx = createHookAPI();

    hsx.on("preSave", (ctx) => {
      const sheet = ctx.workbook.getActiveSheet();
      for (let r = 0; r < 3; r++) {
        const val = sheet.getValue(r, 0);
        if (typeof val === "number" && val < 0) {
          sheet.getCell(r, 0).foreColor("red");
        }
      }
    });

    const ctx = makeCtx();
    const sheet = ctx.workbook.getActiveSheet();
    sheet.setValue(0, 0, 100);
    sheet.setValue(1, 0, -50);
    sheet.setValue(2, 0, 200);

    await runPreSaveHooks(ctx);

    expect(sheet.getCell(0, 0).foreColor()).not.toBe("red");
    expect(sheet.getCell(1, 0).foreColor()).toBe("red");
    expect(sheet.getCell(2, 0).foreColor()).not.toBe("red");
  });

  it("hook can conditionally act based on command", async () => {
    const { runOnOpenHooks } = await import("../src/hooks.js");
    const hsx = createHookAPI();

    const actions: string[] = [];
    hsx.on("onOpen", (ctx) => {
      if (ctx.command === "set" || ctx.command === "eval") {
        actions.push(`formatting for ${ctx.command}`);
      } else {
        actions.push(`skip for ${ctx.command}`);
      }
    });

    await runOnOpenHooks(makeCtx({ command: "set" }));
    await runOnOpenHooks(makeCtx({ command: "get" }));
    await runOnOpenHooks(makeCtx({ command: "eval" }));

    expect(actions).toEqual([
      "formatting for set",
      "skip for get",
      "formatting for eval",
    ]);
  });
});

describe("command context threading", () => {
  it("setCurrentCommand/getCurrentCommand threads through to hooks", async () => {
    const { getCurrentCommand, runOnOpenHooks } = await import(
      "../src/hooks.js"
    );
    const hsx = createHookAPI();

    setCurrentCommand({ command: "set", args: ["model.xlsx", "A1", "42"] });

    const cmd = getCurrentCommand();
    expect(cmd.command).toBe("set");
    expect(cmd.args).toEqual(["model.xlsx", "A1", "42"]);

    // When withFile builds HookContext, it spreads getCurrentCommand() in
    let seen: HookContext | null = null;
    hsx.on("onOpen", (ctx) => {
      seen = ctx;
    });

    await runOnOpenHooks(makeCtx({ ...cmd }));
    expect(seen!.command).toBe("set");
    expect(seen!.args).toEqual(["model.xlsx", "A1", "42"]);

    setCurrentCommand(null);
    const empty = getCurrentCommand();
    expect(empty.command).toBe("");
  });
});

describe("hook ordering and errors", () => {
  it("multiple hooks run in registration order", async () => {
    const { runOnOpenHooks } = await import("../src/hooks.js");
    const hsx = createHookAPI();

    const calls: string[] = [];
    hsx.on("onOpen", () => { calls.push("first"); });
    hsx.on("onOpen", () => { calls.push("second"); });
    hsx.on("onOpen", () => { calls.push("third"); });

    await runOnOpenHooks(makeCtx());
    expect(calls).toEqual(["first", "second", "third"]);
  });

  it("hook errors propagate", async () => {
    const { runOnOpenHooks } = await import("../src/hooks.js");
    const hsx = createHookAPI();

    hsx.on("onOpen", () => {
      throw new Error("hook boom");
    });

    await expect(runOnOpenHooks(makeCtx())).rejects.toThrow("hook boom");
  });

  it("async hooks are awaited", async () => {
    const { runPreSaveHooks } = await import("../src/hooks.js");
    const hsx = createHookAPI();

    const calls: string[] = [];
    hsx.on("preSave", async () => {
      await new Promise((r) => setTimeout(r, 10));
      calls.push("async-done");
    });

    await runPreSaveHooks(makeCtx());
    expect(calls).toEqual(["async-done"]);
  });

  it("runWithHooksDisabled suppresses registered hooks", async () => {
    const { runPreCommandHooks } = await import("../src/hooks.js");
    const hsx = createHookAPI();

    const calls: string[] = [];
    hsx.on("preCommand", () => {
      calls.push("ran");
    });

    await runWithHooksDisabled(true, () =>
      runPreCommandHooks({ command: "info", args: ["model.xlsx"] }),
    );

    expect(calls).toEqual([]);
  });
});

describe("output capture", () => {
  it("captures console output with [hook-type:name] prefix to stderr", async () => {
    const { runOnOpenHooks } = await import("../src/hooks.js");
    const hsx = createHookAPI();

    const chunks: string[] = [];
    const orig = process.stderr.write;
    process.stderr.write = ((chunk: string | Uint8Array) => {
      chunks.push(typeof chunk === "string" ? chunk : chunk.toString());
      return true;
    }) as typeof process.stderr.write;

    hsx.on("onOpen", function hideGridlines() {
      console.log("gridlines hidden");
    });

    try {
      await runOnOpenHooks(makeCtx());
    } finally {
      process.stderr.write = orig;
    }

    const output = chunks.join("");
    expect(output).toContain("[on-open:hideGridlines]");
    expect(output).toContain("gridlines hidden");
  });

  it("output=stdout sends to stdout", async () => {
    const { runPreSaveHooks } = await import("../src/hooks.js");
    const hsx = createHookAPI();

    const chunks: string[] = [];
    const orig = process.stdout.write;
    process.stdout.write = ((chunk: string | Uint8Array) => {
      chunks.push(typeof chunk === "string" ? chunk : chunk.toString());
      return true;
    }) as typeof process.stdout.write;

    hsx.on("preSave", { output: "stdout" }, function formatCells() {
      console.log("formatted");
    });

    try {
      await runPreSaveHooks(makeCtx());
    } finally {
      process.stdout.write = orig;
    }

    expect(chunks.join("")).toContain("formatted");
  });

  it("output=none suppresses all output", async () => {
    const { runPostSaveHooks } = await import("../src/hooks.js");
    const hsx = createHookAPI();

    const stderr: string[] = [];
    const stdout: string[] = [];
    const origErr = process.stderr.write;
    const origOut = process.stdout.write;
    process.stderr.write = ((c: string | Uint8Array) => {
      stderr.push(typeof c === "string" ? c : c.toString());
      return true;
    }) as typeof process.stderr.write;
    process.stdout.write = ((c: string | Uint8Array) => {
      stdout.push(typeof c === "string" ? c : c.toString());
      return true;
    }) as typeof process.stdout.write;

    hsx.on("postSave", { output: "none" }, function quiet() {
      console.log("should vanish");
    });

    try {
      await runPostSaveHooks(makeCtx());
    } finally {
      process.stderr.write = origErr;
      process.stdout.write = origOut;
    }

    expect(stderr.join("")).not.toContain("should vanish");
    expect(stdout.join("")).not.toContain("should vanish");
  });
});

describe("hooks.on registers hooks with shared state", () => {
  it("registers hooks via hooks.on() with shared state", async () => {
    const { runOnOpenHooks, runPreSaveHooks, runPostSaveHooks } = await import(
      "../src/hooks.js"
    );
    const hsx = createHookAPI();

    let saveCount = 0;
    hsx.on("onOpen", (ctx) => {
      ctx.workbook.getActiveSheet().setValue(0, 0, "DEFAULT");
    });
    hsx.on("preSave", () => {
      saveCount++;
    });
    hsx.on("postSave", () => {
      saveCount++;
    });

    const ctx = makeCtx();
    await runOnOpenHooks(ctx);
    expect(ctx.workbook.getActiveSheet().getValue(0, 0)).toBe("DEFAULT");

    await runPreSaveHooks(ctx);
    await runPostSaveHooks(ctx);
    expect(saveCount).toBe(2);
  });
});

describe("hook lifecycle through file context helpers", () => {
  it("withNewFile runs workbook hooks for create flow", async () => {
    const hsx = createHookAPI();
    const calls: string[] = [];

    hsx.on("onOpen", (ctx) => {
      calls.push(`open:${ctx.command}`);
      ctx.workbook.getActiveSheet().setValue(0, 0, "DEFAULT");
    });
    hsx.on("preSave", (ctx) => {
      calls.push(`pre:${ctx.command}`);
    });
    hsx.on("postSave", (ctx) => {
      calls.push(`post:${ctx.command}`);
    });

    await withTempDir(async (tmpDir) => {
      const filePath = path.join(tmpDir, "created.xlsx");
      setCurrentCommand({ command: "create", args: [filePath] });
      try {
        await withNewFile(filePath);
      } finally {
        setCurrentCommand(null);
      }

      await init();
      const reopened = await ExcelFile.open(filePath);
      expect(reopened.workbook.getActiveSheet().getValue(0, 0)).toBe(
        "DEFAULT",
      );
    });

    expect(calls).toEqual(["open:create", "pre:create", "post:create"]);
  });

  it("invalidates daemon cache when onOpen hooks fail before save", async () => {
    await withTempDir(async (tmpDir) => {
      const filePath = path.join(tmpDir, "cache.xlsx");
      await withNewFile(filePath);

      clearHooks();
      const hsx = createHookAPI();
      hsx.on("onOpen", (ctx) => {
        ctx.workbook.getActiveSheet().setValue(0, 0, "mutated");
      });
      hsx.on("onOpen", () => {
        throw new Error("hook boom");
      });

      const { GC } = await init();
      const fileCache = new FileCache(5);

      await expect(
        runWithDaemonRuntime(
          {
            GC,
            ExcelFile,
            fileCache,
            cwd: tmpDir,
            writeThrough: false,
          },
          () => withFile(filePath, () => undefined),
        ),
      ).rejects.toThrow("hook boom");

      const cached = await fileCache.get(filePath, tmpDir);
      expect(cached).toBeNull();

      const reopened = await ExcelFile.open(filePath);
      expect(reopened.workbook.getActiveSheet().getValue(0, 0)).toBeNull();
    });
  });
});

describe("clearHooks", () => {
  it("resets all hooks and discovery state", () => {
    const hsx = createHookAPI();
    hsx.on("onOpen", () => {});
    hsx.on("preSave", () => {});
    hsx.on("postSave", () => {});
    hsx.on("preCommand", () => {});
    hsx.on("postCommand", () => {});

    clearHooks();
    expect(getDiscoveryErrors()).toEqual([]);
  });
});

describe("preCommand / postCommand hooks", () => {
  it("preCommand receives command and args", async () => {
    const { runPreCommandHooks } = await import("../src/hooks.js");
    const hsx = createHookAPI();

    let seen: PreCommandContext | null = null;
    hsx.on("preCommand", (ctx) => {
      seen = ctx;
    });

    await runPreCommandHooks({ command: "set", args: ["model.xlsx", "A1"] });

    expect(seen).not.toBeNull();
    expect(seen!.command).toBe("set");
    expect(seen!.args).toEqual(["model.xlsx", "A1"]);
  });

  it("postCommand receives command, args, and error", async () => {
    const { runPostCommandHooks } = await import("../src/hooks.js");
    const hsx = createHookAPI();

    let seen: PostCommandContext | null = null;
    hsx.on("postCommand", (ctx) => {
      seen = ctx;
    });

    const error = new Error("test error");
    await runPostCommandHooks({
      command: "get",
      args: ["file.xlsx", "B2"],
      error,
    });

    expect(seen).not.toBeNull();
    expect(seen!.command).toBe("get");
    expect(seen!.args).toEqual(["file.xlsx", "B2"]);
    expect(seen!.error).toBe(error);
  });

  it("postCommand receives no error on success", async () => {
    const { runPostCommandHooks } = await import("../src/hooks.js");
    const hsx = createHookAPI();

    let seen: PostCommandContext | null = null;
    hsx.on("postCommand", (ctx) => {
      seen = ctx;
    });

    await runPostCommandHooks({ command: "info", args: ["file.xlsx"] });

    expect(seen!.error).toBeUndefined();
  });

  it("multiple preCommand hooks run in registration order", async () => {
    const { runPreCommandHooks } = await import("../src/hooks.js");
    const hsx = createHookAPI();

    const calls: string[] = [];
    hsx.on("preCommand", () => { calls.push("first"); });
    hsx.on("preCommand", () => { calls.push("second"); });

    await runPreCommandHooks({ command: "get", args: [] });
    expect(calls).toEqual(["first", "second"]);
  });

  it("hooks.on registers preCommand/postCommand", async () => {
    const { runPreCommandHooks, runPostCommandHooks } = await import(
      "../src/hooks.js"
    );
    const hsx = createHookAPI();

    const log: string[] = [];
    hsx.on("preCommand", (ctx) => { log.push(`pre:${ctx.command}`); });
    hsx.on("postCommand", (ctx) => { log.push(`post:${ctx.command}`); });

    await runPreCommandHooks({ command: "set", args: [] });
    await runPostCommandHooks({ command: "set", args: [] });

    expect(log).toEqual(["pre:set", "post:set"]);
  });

  it("preCommand hook errors propagate", async () => {
    const { runPreCommandHooks } = await import("../src/hooks.js");
    const hsx = createHookAPI();

    hsx.on("preCommand", () => {
      throw new Error("pre-command boom");
    });

    await expect(
      runPreCommandHooks({ command: "set", args: [] }),
    ).rejects.toThrow("pre-command boom");
  });

  it("async preCommand hooks are awaited", async () => {
    const { runPreCommandHooks } = await import("../src/hooks.js");
    const hsx = createHookAPI();

    const calls: string[] = [];
    hsx.on("preCommand", async () => {
      await new Promise((r) => setTimeout(r, 10));
      calls.push("async-done");
    });

    await runPreCommandHooks({ command: "get", args: [] });
    expect(calls).toEqual(["async-done"]);
  });
});
