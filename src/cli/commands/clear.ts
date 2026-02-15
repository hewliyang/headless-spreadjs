import { parseRef, rangeDimensions } from "../a1.js";
import { withFile } from "../context.js";
import { fail, ok } from "../output.js";

export async function clear(
  filePath: string,
  ref: string,
  clearType: "values" | "styles" | "all" = "values",
): Promise<void> {
  const parsed = parseRef(ref);

  await withFile(
    filePath,
    ({ workbook, GC }) => {
      const sheet = parsed.sheet
        ? workbook.getSheetFromName(parsed.sheet)
        : workbook.getActiveSheet();

      if (!sheet) {
        fail(`Sheet not found: ${parsed.sheet ?? "(active)"}`);
      }

      const { rows, cols } = rangeDimensions(parsed);
      const range = new GC.Spread.Sheets.Range(
        parsed.start.row,
        parsed.start.col,
        rows,
        cols,
      );

      if (clearType === "values" || clearType === "all") {
        sheet.clear(
          range.row,
          range.col,
          range.rowCount,
          range.colCount,
          GC.Spread.Sheets.SheetArea.viewport,
          GC.Spread.Sheets.StorageType.data,
        );
      }
      if (clearType === "styles" || clearType === "all") {
        sheet.clear(
          range.row,
          range.col,
          range.rowCount,
          range.colCount,
          GC.Spread.Sheets.SheetArea.viewport,
          GC.Spread.Sheets.StorageType.style,
        );
      }

      ok({ cleared: ref, type: clearType });
    },
    { save: true },
  );
}
