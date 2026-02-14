import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { describe, it } from "vitest";
import { withRuntime, withTempDir } from "./helpers.js";

describe("workbook core", () => {
  it("new workbook has at least one sheet", async () => {
    await withRuntime(async ({ ExcelFile }) => {
      const wb = new ExcelFile();
      assert.ok(wb.workbook.getSheetCount() >= 1);
      assert.ok(wb.workbook.getActiveSheet());
    });
  });

  it("addSheet appends by default and inserts at index", async () => {
    await withRuntime(async ({ ExcelFile, GC }) => {
      const wb = new ExcelFile();
      const originalName = wb.workbook.getSheet(0).name();
      assert.equal(wb.workbook.getSheetCount(), 1);

      const appended = new GC.Spread.Sheets.Worksheet("Appended");
      wb.workbook.addSheet(wb.workbook.getSheetCount(), appended);
      assert.equal(appended.name(), "Appended");
      assert.equal(wb.workbook.getSheetCount(), 2);
      assert.equal(wb.workbook.getSheet(1).name(), "Appended");

      const inserted = new GC.Spread.Sheets.Worksheet("Inserted");
      wb.workbook.addSheet(0, inserted);
      assert.equal(inserted.name(), "Inserted");
      assert.equal(wb.workbook.getSheetCount(), 3);
      assert.equal(wb.workbook.getSheet(0).name(), "Inserted");
      assert.equal(wb.workbook.getSheet(1).name(), originalName);
      assert.equal(wb.workbook.getSheet(2).name(), "Appended");
    });
  });

  it("roundtrip via file, buffer, and JSON", async () => {
    await withRuntime(async ({ ExcelFile }) => {
      await withTempDir(async (tempDir) => {
        const wb = new ExcelFile();
        const sheet = wb.workbook.getActiveSheet();
        sheet.name("Core");
        sheet.setValue(0, 0, "Name");
        sheet.setValue(0, 1, "Qty");
        sheet.setValue(1, 0, "Alice");
        sheet.setValue(1, 1, 4);
        sheet.setValue(2, 0, "Bob");
        sheet.setValue(2, 1, 6);
        sheet.setFormula(3, 1, "SUM(B2:B3)");

        // save creates nested directories
        const filePath = path.join(tempDir, "nested", "core.xlsx");
        await wb.save(filePath);
        assert.ok((await fs.stat(filePath)).isFile());

        // open from file
        const fromFile = await ExcelFile.open(filePath);
        const fileSheet = fromFile.workbook.getActiveSheet();
        assert.equal(fileSheet.getValue(1, 0), "Alice");
        assert.equal(fileSheet.getValue(3, 1), 10);
        assert.equal(fileSheet.getFormula(3, 1), "SUM(B2:B3)");

        // open from buffer
        const fromBuffer = await ExcelFile.openFromBuffer(
          await fromFile.saveToBuffer(),
        );
        assert.equal(
          fromBuffer.workbook.getActiveSheet().getValue(2, 0),
          "Bob",
        );
        assert.equal(fromBuffer.workbook.getActiveSheet().getValue(3, 1), 10);

        // JSON roundtrip
        const fromJson = new ExcelFile();
        fromJson.fromJSON(fromBuffer.toJSON());
        assert.equal(
          fromJson.workbook.getActiveSheet().getValue(1, 0),
          "Alice",
        );
        assert.equal(fromJson.workbook.getActiveSheet().getValue(3, 1), 10);
      });
    });
  });

  it("roundtrips diverse cell types through XLSX", async () => {
    await withRuntime(async ({ ExcelFile }) => {
      const wb = new ExcelFile();
      const sheet = wb.workbook.getActiveSheet();

      sheet.setValue(0, 0, "text");
      sheet.setValue(1, 0, 42);
      sheet.setValue(2, 0, 3.14);
      sheet.setValue(3, 0, true);
      sheet.setValue(4, 0, false);
      sheet.setValue(5, 0, "");

      const reopened = await ExcelFile.openFromBuffer(await wb.saveToBuffer());
      const s = reopened.workbook.getActiveSheet();

      assert.equal(s.getValue(0, 0), "text");
      assert.equal(s.getValue(1, 0), 42);
      assert.equal(s.getValue(2, 0), 3.14);
      assert.equal(s.getValue(3, 0), true);
      assert.equal(s.getValue(4, 0), false);
      assert.equal(s.getValue(5, 0), "");
    });
  });

  it("cross-sheet formulas calculate and roundtrip", async () => {
    await withRuntime(async ({ ExcelFile, GC }) => {
      const wb = new ExcelFile();
      const data = wb.workbook.getActiveSheet();
      data.name("Data");
      data.setValue(0, 0, 10);
      data.setValue(1, 0, 20);

      const summary = new GC.Spread.Sheets.Worksheet("Summary");
      wb.workbook.addSheet(wb.workbook.getSheetCount(), summary);
      summary.setFormula(0, 0, "SUM(Data!A1:A2)");

      assert.equal(summary.getValue(0, 0), 30);

      const reopened = await ExcelFile.openFromBuffer(await wb.saveToBuffer());
      assert.equal(reopened.workbook.getSheet(1).getValue(0, 0), 30);
      assert.match(
        reopened.workbook.getSheet(1).getFormula(0, 0),
        /SUM\(Data!A1:A2\)/,
      );
    });
  });

  it("open, modify, re-save, reopen preserves changes", async () => {
    await withRuntime(async ({ ExcelFile }) => {
      await withTempDir(async (tempDir) => {
        const filePath = path.join(tempDir, "edit.xlsx");

        const wb = new ExcelFile();
        wb.workbook.getActiveSheet().setValue(0, 0, "original");
        await wb.save(filePath);

        const wb2 = await ExcelFile.open(filePath);
        wb2.workbook.getActiveSheet().setValue(0, 0, "modified");
        wb2.workbook.getActiveSheet().setValue(1, 0, "added");
        await wb2.save(filePath);

        const wb3 = await ExcelFile.open(filePath);
        assert.equal(wb3.workbook.getActiveSheet().getValue(0, 0), "modified");
        assert.equal(wb3.workbook.getActiveSheet().getValue(1, 0), "added");
      });
    });
  });

  it("batch() computes formulas correctly after resuming", async () => {
    await withRuntime(async ({ ExcelFile }) => {
      const wb = new ExcelFile();
      const sheet = wb.workbook.getActiveSheet();

      wb.batch(() => {
        sheet.setValue(0, 0, 10);
        sheet.setValue(1, 0, 20);
        sheet.setValue(2, 0, 30);
        sheet.setFormula(3, 0, "SUM(A1:A3)");
      });

      assert.equal(sheet.getValue(3, 0), 60);
    });
  });

  it("batch() recovers from sync and async errors", async () => {
    await withRuntime(async ({ ExcelFile }) => {
      const wb = new ExcelFile();
      const sheet = wb.workbook.getActiveSheet();
      sheet.setValue(0, 0, 5);
      sheet.setFormula(1, 0, "A1*2");
      assert.equal(sheet.getValue(1, 0), 10);

      // sync error — formulas still work afterward
      assert.throws(() => {
        wb.batch(() => {
          throw new Error("sync-fail");
        });
      }, /sync-fail/);

      sheet.setValue(0, 0, 7);
      assert.equal(sheet.getValue(1, 0), 14);

      // async error — formulas still work afterward
      await assert.rejects(
        wb.batch(async () => {
          throw new Error("async-fail");
        }),
        /async-fail/,
      );

      sheet.setValue(0, 0, 3);
      assert.equal(sheet.getValue(1, 0), 6);
    });
  });

  it("open() rejects for non-existent file", async () => {
    await withRuntime(async ({ ExcelFile }) => {
      await assert.rejects(ExcelFile.open("/no/such/file.xlsx"), /ENOENT/);
    });
  });

  it("openFromBuffer() rejects for invalid XLSX", async () => {
    await withRuntime(async ({ ExcelFile }) => {
      await assert.rejects(
        ExcelFile.openFromBuffer(Buffer.from("not-an-xlsx")),
      );
    });
  });

  it("strips 'Evaluation Version' sheet on save/open", async () => {
    await withRuntime(async ({ ExcelFile, GC }) => {
      const wb = new ExcelFile();
      wb.workbook.getActiveSheet().name("Main");
      const evalSheet = new GC.Spread.Sheets.Worksheet("Evaluation Version");
      wb.workbook.addSheet(wb.workbook.getSheetCount(), evalSheet);
      const tailSheet = new GC.Spread.Sheets.Worksheet("Tail");
      wb.workbook.addSheet(wb.workbook.getSheetCount(), tailSheet);

      const reopened = await ExcelFile.openFromBuffer(await wb.saveToBuffer());
      const names = Array.from(
        { length: reopened.workbook.getSheetCount() },
        (_, i) => reopened.workbook.getSheet(i).name(),
      );

      assert.equal(names.includes("Evaluation Version"), false);
      assert.ok(names.includes("Main"));
      assert.ok(names.includes("Tail"));
    });
  });
});
