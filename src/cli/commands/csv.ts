import { parseRef, rangeDimensions } from "../a1.js";
import { throwIfAborted } from "../abort.js";
import { withFile } from "../context.js";
import { fail, writeStdout } from "../output.js";

export type CsvMode = "value" | "formula" | "both";

export async function csv(
  filePath: string,
  ref: string,
  options: { signal?: AbortSignal | null; mode?: CsvMode } = {},
): Promise<void> {
  const parsed = parseRef(ref);
  const signal = options?.signal;
  const mode = options.mode ?? "value";

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
          const rowIndex = parsed.start.row + r;
          const colIndex = parsed.start.col + c;
          row.push(toCsvCell(sheet, rowIndex, colIndex, mode));
        }
        lines.push(row.join(","));
      }

      writeStdout(`${lines.join("\n")}\n`);
    },
    { signal },
  );
}

function toCsvCell(
  sheet: {
    getValue(row: number, col: number): unknown;
    getFormula(row: number, col: number): string | null;
  },
  row: number,
  col: number,
  mode: CsvMode,
): string {
  const value = sheet.getValue(row, col);
  const formula = sheet.getFormula(row, col);

  const valueText = value === null || value === undefined ? "" : String(value);
  const formulaText = formula ? `=${formula}` : "";

  let text = "";
  if (mode === "formula") {
    text = formulaText || valueText;
  } else if (mode === "both") {
    text = formulaText
      ? valueText
        ? `${valueText} | ${formulaText}`
        : formulaText
      : valueText;
  } else {
    text = valueText;
  }

  return escapeCsv(text);
}

function escapeCsv(text: string): string {
  if (text.includes(",") || text.includes('"') || text.includes("\n")) {
    return `"${text.replace(/"/g, '""')}"`;
  }
  return text;
}
