import assert from "node:assert/strict";
import { describe, it } from "vitest";
import { withRuntime } from "../helpers.js";

describe("sparklines extension", () => {
  it("roundtrips sparkline through XLSX", async () => {
    await withRuntime(async ({ Workbook, GC }) => {
      const wb = new Workbook();
      const sheet = wb.getActiveSheet();
      assert.ok(sheet);

      sheet.setValue(0, 0, 4);
      sheet.setValue(0, 1, 6);
      sheet.setValue(0, 2, 5);
      sheet.setValue(0, 3, 7);

      const settings = new GC.Spread.Sheets.Sparklines.SparklineSetting();
      sheet.setSparkline(
        1,
        0,
        new GC.Spread.Sheets.Range(0, 0, 1, 4),
        GC.Spread.Sheets.Sparklines.DataOrientation.horizontal,
        GC.Spread.Sheets.Sparklines.SparklineType.line,
        settings,
      );

      const reopened = await Workbook.openFromBuffer(await wb.saveToBuffer());
      const restoredSheet = reopened.getActiveSheet();
      assert.ok(restoredSheet);

      const sparkline = restoredSheet.getSparkline(1, 0);
      assert.ok(sparkline);
      assert.equal(
        sparkline.sparklineType(),
        GC.Spread.Sheets.Sparklines.SparklineType.line,
      );
    });
  });
});
