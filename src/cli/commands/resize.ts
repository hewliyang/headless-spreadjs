import { colToIndex, parseRef } from "../a1.js";
import { withFile } from "../context.js";
import { fail, ok } from "../output.js";

export async function resize(
  filePath: string,
  sheetName: string | undefined,
  options: {
    columns?: string;
    rows?: string;
    width?: number;
    height?: number;
  },
): Promise<void> {
  if (!options.width && !options.height) {
    fail("Specify --width and/or --height.");
  }

  await withFile(
    filePath,
    ({ workbook }) => {
      const sheet = sheetName
        ? workbook.getSheetFromName(sheetName)
        : workbook.getActiveSheet();

      if (!sheet) fail(`Sheet not found: ${sheetName ?? "(active)"}`);

      if (options.width !== undefined) {
        if (options.columns) {
          const [startStr, endStr] = options.columns.split(":");
          const startCol = colToIndex(startStr);
          const endCol = endStr ? colToIndex(endStr) : startCol;
          for (let c = startCol; c <= endCol; c++) {
            sheet.setColumnWidth(c, options.width);
          }
        } else {
          // All columns
          const colCount = sheet.getColumnCount();
          for (let c = 0; c < colCount; c++) {
            sheet.setColumnWidth(c, options.width);
          }
        }
      }

      if (options.height !== undefined) {
        if (options.rows) {
          const [startStr, endStr] = options.rows.split(":");
          const startRow = parseInt(startStr, 10) - 1;
          const endRow = endStr ? parseInt(endStr, 10) - 1 : startRow;
          for (let r = startRow; r <= endRow; r++) {
            sheet.setRowHeight(r, options.height);
          }
        } else {
          // All rows
          const rowCount = sheet.getRowCount();
          for (let r = 0; r < rowCount; r++) {
            sheet.setRowHeight(r, options.height);
          }
        }
      }

      ok({ resized: true });
    },
    { save: true },
  );
}
