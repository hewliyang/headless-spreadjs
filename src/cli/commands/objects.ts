import { withFile } from "../context.js";
import { ok } from "../output.js";

interface ObjectInfo {
  sheet: string;
  type: string;
  name: string;
}

export async function objects(
  filePath: string,
  sheetFilter?: string,
): Promise<void> {
  await withFile(filePath, ({ workbook }) => {
    const result: ObjectInfo[] = [];

    for (let i = 0; i < workbook.getSheetCount(); i++) {
      const sheet = workbook.getSheet(i);
      const sheetName = sheet.name();

      if (sheetFilter && sheetName !== sheetFilter) continue;

      // Charts
      try {
        const charts = sheet.charts?.all?.() ?? [];
        for (const chart of charts) {
          result.push({
            sheet: sheetName,
            type: "chart",
            name: chart.name(),
          });
        }
      } catch {
        // charts extension may not be loaded
      }

      // Tables
      try {
        const tables = sheet.tables?.all?.() ?? [];
        for (const table of tables) {
          result.push({
            sheet: sheetName,
            type: "table",
            name: String(table.name()),
          });
        }
      } catch {
        // tables may not be available
      }

      // Pivot tables
      try {
        const pivots = sheet.pivotTables?.all?.() ?? [];
        for (const pivot of pivots) {
          result.push({
            sheet: sheetName,
            type: "pivotTable",
            name: String(pivot.name()),
          });
        }
      } catch {
        // pivot extension may not be loaded
      }

      // Slicers
      try {
        const slicers = sheet.slicers?.all?.() ?? [];
        for (const slicer of slicers) {
          result.push({
            sheet: sheetName,
            type: "slicer",
            name: String(slicer.name()),
          });
        }
      } catch {
        // slicers extension may not be loaded
      }
    }

    ok({ objects: result });
  });
}
