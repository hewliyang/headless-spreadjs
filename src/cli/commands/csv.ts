import { parseRef, rangeDimensions } from "../a1.js";
import { throwIfAborted } from "../abort.js";
import { withFile } from "../context.js";
import { fail, writeStdout } from "../output.js";

export async function csv(
  filePath: string,
  ref: string,
  options?: { signal?: AbortSignal | null },
): Promise<void> {
  const parsed = parseRef(ref);
  const signal = options?.signal;

  await withFile(
    filePath,
    ({ workbook }) => {
      const sheet = parsed.sheet
        ? workbook.getSheetFromName(parsed.sheet)
        : workbook.getActiveSheet();

      if (!sheet) {
        fail(`Sheet not found: ${parsed.sheet ?? "(active)"}`);
      }

      const { rows, cols } = rangeDimensions(parsed);
      const lines: string[] = [];

      for (let r = 0; r < rows; r++) {
        throwIfAborted(signal);
        const row: string[] = [];
        for (let c = 0; c < cols; c++) {
          const value = sheet.getValue(
            parsed.start.row + r,
            parsed.start.col + c,
          );
          const str =
            value === null || value === undefined ? "" : String(value);
          if (str.includes(",") || str.includes('"') || str.includes("\n")) {
            row.push(`"${str.replace(/"/g, '""')}"`);
          } else {
            row.push(str);
          }
        }
        lines.push(row.join(","));
      }

      writeStdout(`${lines.join("\n")}\n`);
    },
    { signal },
  );
}
