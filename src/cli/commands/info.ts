import { cellToA1 } from "../a1.js";
import { withFile } from "../context.js";
import { ok } from "../output.js";

export async function info(filePath: string): Promise<void> {
  await withFile(filePath, ({ workbook, GC }) => {
    const sheetCount = workbook.getSheetCount();
    const activeIndex = workbook.getActiveSheetIndex();
    const sheets: {
      index: number;
      name: string;
      usedRange: string | null;
      active: boolean;
    }[] = [];

    for (let i = 0; i < sheetCount; i++) {
      const sheet = workbook.getSheet(i);
      let usedRange: string | null = null;

      try {
        const range = sheet.getUsedRange(
          GC.Spread.Sheets.UsedRangeType.data |
            GC.Spread.Sheets.UsedRangeType.formula,
        );
        if (range && range.rowCount > 0 && range.colCount > 0) {
          const startA1 = cellToA1(range.row, range.col);
          const endA1 = cellToA1(
            range.row + range.rowCount - 1,
            range.col + range.colCount - 1,
          );
          usedRange = `${startA1}:${endA1}`;
        }
      } catch {
        // getUsedRange may fail on empty sheets
      }

      sheets.push({
        index: i,
        name: sheet.name(),
        usedRange,
        active: i === activeIndex,
      });
    }

    ok({ sheets });
  });
}
