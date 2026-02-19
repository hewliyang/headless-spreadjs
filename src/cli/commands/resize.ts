import { colToIndex } from "../a1.js";
import { throwIfAborted } from "../abort.js";
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
    signal?: AbortSignal | null;
  },
): Promise<void> {
  if (!options.width && !options.height) {
    fail("Specify --width and/or --height.");
  }

  const signal = options.signal;

  await withFile(
    filePath,
    ({ file, workbook }) => {
      const sheet = sheetName
        ? workbook.getSheetFromName(sheetName)
        : workbook.getActiveSheet();

      if (!sheet) fail(`Sheet not found: ${sheetName ?? "(active)"}`);

      file.batch(() => {
        if (options.width !== undefined) {
          if (options.columns) {
            const [startStr, endStr] = options.columns.split(":");
            const startCol = colToIndex(startStr);
            const endCol = endStr ? colToIndex(endStr) : startCol;
            for (let c = startCol; c <= endCol; c++) {
              throwIfAborted(signal);
              sheet.setColumnWidth(c, options.width);
            }
          } else {
            const colCount = sheet.getColumnCount();
            for (let c = 0; c < colCount; c++) {
              throwIfAborted(signal);
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
              throwIfAborted(signal);
              sheet.setRowHeight(r, options.height);
            }
          } else {
            const rowCount = sheet.getRowCount();
            for (let r = 0; r < rowCount; r++) {
              throwIfAborted(signal);
              sheet.setRowHeight(r, options.height);
            }
          }
        }
      });

      ok({ resized: true });
    },
    { save: true, signal },
  );
}
