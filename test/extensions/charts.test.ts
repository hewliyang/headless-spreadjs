import assert from "node:assert/strict";
import { describe, it } from "vitest";
import { withRuntime } from "../helpers.js";

describe("charts extension", () => {
  it("roundtrips chart through XLSX", async () => {
    await withRuntime(async ({ ExcelFile, GC }) => {
      const wb = new ExcelFile();
      const sheet = wb.workbook.getActiveSheet();
      assert.ok(sheet);

      sheet.setValue(0, 0, "Month");
      sheet.setValue(0, 1, "Revenue");
      sheet.setValue(1, 0, "Jan");
      sheet.setValue(1, 1, 10);
      sheet.setValue(2, 0, "Feb");
      sheet.setValue(2, 1, 20);
      sheet.setValue(3, 0, "Mar");
      sheet.setValue(3, 1, 15);

      sheet.charts.add(
        "RevenueChart",
        GC.Spread.Sheets.Charts.ChartType.line,
        0,
        120,
        400,
        240,
        "A1:B4",
      );

      const reopened = await ExcelFile.openFromBuffer(await wb.saveToBuffer());
      const restoredSheet = reopened.workbook.getActiveSheet();
      assert.ok(restoredSheet);

      assert.equal(restoredSheet.charts.all().length, 1);
      const restoredChart = restoredSheet.charts.get("RevenueChart");
      assert.ok(restoredChart);
      assert.equal(
        restoredChart.chartType(),
        GC.Spread.Sheets.Charts.ChartType.line,
      );
    });
  });
});
