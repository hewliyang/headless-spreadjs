import assert from "node:assert/strict";
import { describe, it } from "vitest";
import { withRuntime } from "../helpers.js";

describe("pivot extension", () => {
  it("roundtrips pivot through XLSX and JSON", async () => {
    await withRuntime(async ({ Workbook, GC }) => {
      const wb = new Workbook();
      const dataSheet = wb.getActiveSheet();
      assert.ok(dataSheet);
      dataSheet.name("Sales Data");

      dataSheet.setValue(0, 0, "Region");
      dataSheet.setValue(0, 1, "Quarter");
      dataSheet.setValue(0, 2, "Sales");
      dataSheet.setValue(1, 0, "East");
      dataSheet.setValue(1, 1, "Q1");
      dataSheet.setValue(1, 2, 10);
      dataSheet.setValue(2, 0, "East");
      dataSheet.setValue(2, 1, "Q2");
      dataSheet.setValue(2, 2, 12);
      dataSheet.setValue(3, 0, "West");
      dataSheet.setValue(3, 1, "Q1");
      dataSheet.setValue(3, 2, 14);
      dataSheet.setValue(4, 0, "West");
      dataSheet.setValue(4, 1, "Q2");
      dataSheet.setValue(4, 2, 16);

      const pivotSheet = wb.addSheet("Pivot");
      const pivot = pivotSheet.pivotTables.add(
        "SalesPivot",
        "'Sales Data'!A1:C5",
        1,
        0,
        GC.Spread.Pivot.PivotTableLayoutType.outline,
      );

      pivot.add(
        "Region",
        "Region",
        GC.Spread.Pivot.PivotTableFieldType.rowField,
      );
      pivot.add(
        "Quarter",
        "Quarter",
        GC.Spread.Pivot.PivotTableFieldType.columnField,
      );
      pivot.add(
        "TotalSales",
        "Sales",
        GC.Spread.Pivot.PivotTableFieldType.valueField,
      );

      const reopened = await Workbook.openFromBuffer(await wb.saveToBuffer());
      const restoredPivotSheet = reopened.getSheet(1);
      assert.equal(restoredPivotSheet.pivotTables.all().length, 1);

      const restoredPivot = restoredPivotSheet.pivotTables.all()[0];
      assert.equal(restoredPivot.name(), "SalesPivot");
      assert.match(restoredPivot.getSource(), /Sales Data/);

      const json = wb.toJSON();
      const fromJson = new Workbook();
      fromJson.fromJSON(json);
      assert.equal(fromJson.getSheet(1).pivotTables.all().length, 1);
    });
  });
});
