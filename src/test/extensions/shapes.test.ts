import assert from "node:assert/strict";
import { describe, it } from "vitest";
import { withRuntime } from "../helpers.js";

describe("shapes extension", () => {
  it("roundtrips shape through XLSX", async () => {
    await withRuntime(async ({ Workbook, GC }) => {
      const wb = new Workbook();
      const sheet = wb.getActiveSheet();
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

      const reopened = await Workbook.openFromBuffer(await wb.saveToBuffer());
      const restoredShape = reopened.getActiveSheet().shapes.get("StatusShape");

      assert.ok(restoredShape);
      assert.equal(restoredShape.text(), "Headless shape");
    });
  });
});
