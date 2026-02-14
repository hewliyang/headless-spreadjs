import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { describe, it } from "vitest";
import { Workbook } from "../index.js";
import { withRuntime, withTempDir } from "./helpers.js";

describe("workbook core", () => {
  it("new workbook has at least one sheet", async () => {
    await withRuntime(async ({ Workbook }) => {
      const wb = new Workbook();
      assert.ok(wb.getSheetCount() >= 1);
      assert.ok(wb.getActiveSheet());
    });
  });

  it("addSheet() appends by default and inserts at index", async () => {
    await withRuntime(async ({ Workbook }) => {
      const wb = new Workbook();
      const originalName = wb.getSheet(0).name();
      assert.equal(wb.getSheetCount(), 1);

      const appended = wb.addSheet("Appended");
      assert.equal(appended.name(), "Appended");
      assert.equal(wb.getSheetCount(), 2);
      assert.equal(wb.getSheet(1).name(), "Appended");

      const inserted = wb.addSheet("Inserted", 0);
      assert.equal(inserted.name(), "Inserted");
      assert.equal(wb.getSheetCount(), 3);
      assert.equal(wb.getSheet(0).name(), "Inserted");
      assert.equal(wb.getSheet(1).name(), originalName);
      assert.equal(wb.getSheet(2).name(), "Appended");
    });
  });

  it("roundtrip via file, buffer, and JSON", async () => {
    await withRuntime(async ({ Workbook }) => {
      await withTempDir(async (tempDir) => {
        const wb = new Workbook();
        const sheet = wb.getActiveSheet();
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
        const fromFile = await Workbook.open(filePath);
        const fileSheet = fromFile.getActiveSheet();
        assert.equal(fileSheet.getValue(1, 0), "Alice");
        assert.equal(fileSheet.getValue(3, 1), 10);
        assert.equal(fileSheet.getFormula(3, 1), "SUM(B2:B3)");

        // open from buffer
        const fromBuffer = await Workbook.openFromBuffer(
          await fromFile.saveToBuffer(),
        );
        assert.equal(fromBuffer.getActiveSheet().getValue(2, 0), "Bob");
        assert.equal(fromBuffer.getActiveSheet().getValue(3, 1), 10);

        // JSON roundtrip
        const fromJson = new Workbook();
        fromJson.fromJSON(fromBuffer.toJSON());
        assert.equal(fromJson.getActiveSheet().getValue(1, 0), "Alice");
        assert.equal(fromJson.getActiveSheet().getValue(3, 1), 10);
      });
    });
  });

  it("roundtrips diverse cell types through XLSX", async () => {
    await withRuntime(async ({ Workbook }) => {
      const wb = new Workbook();
      const sheet = wb.getActiveSheet();

      sheet.setValue(0, 0, "text");
      sheet.setValue(1, 0, 42);
      sheet.setValue(2, 0, 3.14);
      sheet.setValue(3, 0, true);
      sheet.setValue(4, 0, false);
      sheet.setValue(5, 0, "");

      const reopened = await Workbook.openFromBuffer(await wb.saveToBuffer());
      const s = reopened.getActiveSheet();

      assert.equal(s.getValue(0, 0), "text");
      assert.equal(s.getValue(1, 0), 42);
      assert.equal(s.getValue(2, 0), 3.14);
      assert.equal(s.getValue(3, 0), true);
      assert.equal(s.getValue(4, 0), false);
      assert.equal(s.getValue(5, 0), "");
    });
  });

  it("cross-sheet formulas calculate and roundtrip", async () => {
    await withRuntime(async ({ Workbook }) => {
      const wb = new Workbook();
      const data = wb.getActiveSheet();
      data.name("Data");
      data.setValue(0, 0, 10);
      data.setValue(1, 0, 20);

      const summary = wb.addSheet("Summary");
      summary.setFormula(0, 0, "SUM(Data!A1:A2)");

      assert.equal(summary.getValue(0, 0), 30);

      const reopened = await Workbook.openFromBuffer(await wb.saveToBuffer());
      assert.equal(reopened.getSheet(1).getValue(0, 0), 30);
      assert.match(reopened.getSheet(1).getFormula(0, 0), /SUM\(Data!A1:A2\)/);
    });
  });

  it("open, modify, re-save, reopen preserves changes", async () => {
    await withRuntime(async ({ Workbook }) => {
      await withTempDir(async (tempDir) => {
        const filePath = path.join(tempDir, "edit.xlsx");

        const wb = new Workbook();
        wb.getActiveSheet().setValue(0, 0, "original");
        await wb.save(filePath);

        const wb2 = await Workbook.open(filePath);
        wb2.getActiveSheet().setValue(0, 0, "modified");
        wb2.getActiveSheet().setValue(1, 0, "added");
        await wb2.save(filePath);

        const wb3 = await Workbook.open(filePath);
        assert.equal(wb3.getActiveSheet().getValue(0, 0), "modified");
        assert.equal(wb3.getActiveSheet().getValue(1, 0), "added");
      });
    });
  });

  it("batch() computes formulas correctly after resuming", async () => {
    await withRuntime(async ({ Workbook }) => {
      const wb = new Workbook();
      const sheet = wb.getActiveSheet();

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
    await withRuntime(async ({ Workbook }) => {
      const wb = new Workbook();
      const sheet = wb.getActiveSheet();
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
    await withRuntime(async ({ Workbook }) => {
      await assert.rejects(Workbook.open("/no/such/file.xlsx"), /ENOENT/);
    });
  });

  it("openFromBuffer() rejects for invalid XLSX", async () => {
    await withRuntime(async ({ Workbook }) => {
      await assert.rejects(Workbook.openFromBuffer(Buffer.from("not-an-xlsx")));
    });
  });

  it("strips 'Evaluation Version' sheet on save/open", async () => {
    await withRuntime(async ({ Workbook }) => {
      const wb = new Workbook();
      wb.getActiveSheet().name("Main");
      wb.addSheet("Evaluation Version");
      wb.addSheet("Tail");

      const reopened = await Workbook.openFromBuffer(await wb.saveToBuffer());
      const names = Array.from({ length: reopened.getSheetCount() }, (_, i) =>
        reopened.getSheet(i).name(),
      );

      assert.equal(names.includes("Evaluation Version"), false);
      assert.ok(names.includes("Main"));
      assert.ok(names.includes("Tail"));
    });
  });
});
