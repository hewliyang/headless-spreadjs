import assert from "node:assert/strict";
import { describe, it } from "vitest";
import { withRuntime } from "../helpers.js";

describe("slicers extension", () => {
  it("roundtrips slicer and table through XLSX", async () => {
    await withRuntime(async ({ Workbook, GC }) => {
      const wb = new Workbook();
      const sheet = wb.getActiveSheet();
      assert.ok(sheet);

      sheet.setValue(0, 0, "Region");
      sheet.setValue(0, 1, "Sales");
      sheet.setValue(1, 0, "East");
      sheet.setValue(1, 1, 10);
      sheet.setValue(2, 0, "West");
      sheet.setValue(2, 1, 20);

      sheet.tables.add(
        "SalesTable",
        0,
        0,
        3,
        2,
        GC.Spread.Sheets.Tables.TableThemes.medium2,
      );
      sheet.slicers.add("RegionSlicer", "SalesTable", "Region");

      const reopened = await Workbook.openFromBuffer(await wb.saveToBuffer());
      const restoredSheet = reopened.getActiveSheet();
      assert.ok(restoredSheet);

      assert.equal(restoredSheet.tables.all().length, 1);
      assert.equal(restoredSheet.slicers.all().length, 1);

      const slicer = restoredSheet.slicers.all()[0];
      assert.equal(slicer.name(), "RegionSlicer");
      assert.equal(slicer.captionName(), "Region");
    });
  });
});
