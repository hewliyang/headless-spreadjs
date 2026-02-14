import assert from "node:assert/strict";
import { describe, it } from "vitest";
import { withRuntime } from "../helpers.js";

describe("shapes extension", () => {
  it("roundtrips shape through XLSX", async () => {
    await withRuntime(async ({ ExcelFile, GC }) => {
      const wb = new ExcelFile();
      const sheet = wb.workbook.getActiveSheet();
      assert.ok(sheet);

      const shape = sheet.shapes.add(
        "StatusShape",
        GC.Spread.Sheets.Shapes.AutoShapeType.roundedRectangle,
        120,
        50,
        220,
        70,
      );
      shape.text("Headless shape");

      const reopened = await ExcelFile.openFromBuffer(await wb.saveToBuffer());
      const restoredShape = reopened.workbook
        .getActiveSheet()
        .shapes.get("StatusShape");

      assert.ok(restoredShape);
      assert.equal(restoredShape.text(), "Headless shape");
    });
  });
});
