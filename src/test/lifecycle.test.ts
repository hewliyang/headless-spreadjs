import assert from "node:assert/strict";
import { describe, it } from "vitest";
import { dispose, ExcelFile, init } from "../index.js";
import { disposeShims, installShims, isShimInstalled } from "../shims.js";
import { readWithFileReader } from "./helpers.js";

describe("lifecycle", () => {
  it("ExcelFile requires init()", () => {
    dispose();
    assert.throws(
      () => new ExcelFile(),
      /headless-spreadjs not initialized\. Call init\(\) first\./,
    );
  });

  it("init() is idempotent and loads all extensions", async () => {
    const first = await init();

    try {
      const second = await init();

      assert.strictEqual(first.GC, second.GC);
      assert.strictEqual(first.ExcelFile, second.ExcelFile);

      assert.ok(first.GC.Spread.Sheets.Charts?.ChartType);
      assert.ok(first.GC.Spread.Pivot?.PivotTableFieldType);
      assert.ok(first.GC.Spread.Sheets.Shapes?.AutoShapeType);
      assert.ok(first.GC.Spread.Sheets.Slicers?.SlicerStyles);
      assert.ok(first.GC.Spread.Sheets.Sparklines?.SparklineType);
    } finally {
      first.dispose();
    }
  });

  it("shims + FileReader work headlessly", async () => {
    disposeShims();
    assert.equal(isShimInstalled(), false);

    installShims();
    assert.equal(isShimInstalled(), true);

    const g = globalThis as Record<string, unknown>;
    const currentWindow = g.window;
    installShims();
    assert.strictEqual(g.window, currentWindow);

    const blob = new Blob(["hello world"], { type: "text/plain" });

    const text = await readWithFileReader("readAsText", blob);
    assert.equal(text, "hello world");

    const arrayBuffer = (await readWithFileReader(
      "readAsArrayBuffer",
      blob,
    )) as ArrayBuffer;
    assert.equal(Buffer.from(arrayBuffer).toString("utf8"), "hello world");

    const dataUrl = (await readWithFileReader("readAsDataURL", blob)) as string;
    assert.match(dataUrl, /^data:text\/plain;base64,/);

    disposeShims();
    assert.equal(isShimInstalled(), false);
  });
});
