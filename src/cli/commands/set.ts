import { cellToA1, parseRef, type RangeRef, rangeDimensions } from "../a1.js";
import { throwIfAborted } from "../abort.js";
import { withFile } from "../context.js";
import { ok, readInput } from "../output.js";
import { ensureSheetSize } from "../sheet-size.js";
import { applyStyles, type CellStyle } from "../styles.js";

interface CellInput {
  value?: unknown;
  formula?: string;
  style?: CellStyle;
}

function formatRange(ref: RangeRef): string {
  const start = cellToA1(ref.start.row, ref.start.col);
  const end = cellToA1(ref.end.row, ref.end.col);
  const base = start === end ? start : `${start}:${end}`;

  if (!ref.sheet) return base;

  const safeSheet = /^[A-Za-z0-9_]+$/.test(ref.sheet)
    ? ref.sheet
    : `'${ref.sheet.replace(/'/g, "''")}'`;
  return `${safeSheet}!${base}`;
}

export async function set(
  filePath: string,
  ref: string,
  jsonArg: string | undefined,
  options?: { signal?: AbortSignal | null },
): Promise<void> {
  const signal = options?.signal;
  const input = await readInput(jsonArg);
  let cells: CellInput[][];

  try {
    cells = JSON.parse(input);
  } catch {
    throw new Error("Invalid JSON input for cells.");
  }

  if (!Array.isArray(cells) || !Array.isArray(cells[0])) {
    throw new Error("Cells must be a 2D array: [[{value: ...}, ...], ...]");
  }

  const parsed = parseRef(ref);
  const target: RangeRef = {
    sheet: parsed.sheet,
    start: { ...parsed.start },
    end: { ...parsed.end },
  };

  const { rows: originalRows, cols: originalCols } = rangeDimensions(parsed);
  const isSingleCellRef = originalRows === 1 && originalCols === 1;

  const dataRows = cells.length;
  const dataCols = cells[0].length;

  for (let r = 0; r < dataRows; r++) {
    throwIfAborted(signal);
    if (!Array.isArray(cells[r])) {
      throw new Error(`Row ${r} is not an array.`);
    }
    if (cells[r].length !== dataCols) {
      throw new Error(
        `Column count mismatch in row ${r}: expected ${dataCols} cols but got ${cells[r].length} cols.`,
      );
    }
  }

  const messages: string[] = [];

  if (isSingleCellRef) {
    target.end.row = target.start.row + dataRows - 1;
    target.end.col = target.start.col + dataCols - 1;

    const rowDiff = dataRows - originalRows;
    const colDiff = dataCols - originalCols;
    if (rowDiff !== 0 || colDiff !== 0) {
      messages.push(
        `Adjusted range from ${formatRange(parsed)} to ${formatRange(target)} (row diff: ${rowDiff}, col diff: ${colDiff})`,
      );
    }
  } else {
    if (dataRows !== originalRows) {
      throw new Error(
        `Row count mismatch: range has ${originalRows} rows but got ${dataRows} rows.`,
      );
    }
    for (let r = 0; r < dataRows; r++) {
      throwIfAborted(signal);
      if (cells[r].length !== originalCols) {
        throw new Error(
          `Column count mismatch in row ${r}: range has ${originalCols} cols but got ${cells[r].length} cols.`,
        );
      }
    }
  }

  const { rows, cols } = rangeDimensions(target);

  await withFile(
    filePath,
    ({ file, workbook, GC, markMutated }) => {
      const sheet = target.sheet
        ? workbook.getSheetFromName(target.sheet)
        : workbook.getActiveSheet();

      if (!sheet) {
        throw new Error(`Sheet not found: ${target.sheet ?? "(active)"}`);
      }

      let written = 0;

      ensureSheetSize(sheet, target.end.row + 1, target.end.col + 1);

      file.batch(() => {
        for (let r = 0; r < rows; r++) {
          throwIfAborted(signal);
          for (let c = 0; c < cols; c++) {
            const cell = cells[r]?.[c];
            if (!cell) continue;

            const row = target.start.row + r;
            const col = target.start.col + c;

            if (cell.formula) {
              const f = cell.formula.startsWith("=")
                ? cell.formula.slice(1)
                : cell.formula;
              sheet.setFormula(row, col, f);
              written++;
            } else if (cell.value !== undefined) {
              sheet.setFormula(row, col, "");
              sheet.setValue(row, col, cell.value);
              written++;
            }

            if (cell.style) {
              applyStyles(sheet, row, col, cell.style, GC);
            }
          }
        }
      });

      markMutated({
        sheet: target.sheet ?? sheet.name(),
        startRow: target.start.row,
        startCol: target.start.col,
        endRow: target.end.row,
        endCol: target.end.col,
      });

      ok({
        success: true,
        written,
        range: formatRange(target),
        messages,
      });
    },
    { save: true, signal },
  );
}
