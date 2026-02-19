import { parseRef, rangeDimensions } from "../a1.js";
import { throwIfAborted } from "../abort.js";
import { withFile } from "../context.js";
import { fail, ok, readInput } from "../output.js";
import { applyStyles, type CellStyles } from "../styles.js";

interface CellInput {
  value?: unknown;
  formula?: string;
  cellStyles?: CellStyles;
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
    fail("Invalid JSON input for cells.");
    return;
  }

  if (!Array.isArray(cells) || !Array.isArray(cells[0])) {
    fail("Cells must be a 2D array: [[{value: ...}, ...], ...]");
    return;
  }

  const parsed = parseRef(ref);
  const { rows, cols } = rangeDimensions(parsed);

  if (cells.length !== rows) {
    fail(
      `Row count mismatch: range has ${rows} rows but got ${cells.length} rows.`,
    );
    return;
  }
  for (let r = 0; r < cells.length; r++) {
    throwIfAborted(signal);
    if (cells[r].length !== cols) {
      fail(
        `Column count mismatch in row ${r}: range has ${cols} cols but got ${cells[r].length} cols.`,
      );
      return;
    }
  }

  await withFile(
    filePath,
    ({ file, workbook, GC }) => {
      const sheet = parsed.sheet
        ? workbook.getSheetFromName(parsed.sheet)
        : workbook.getActiveSheet();

      if (!sheet) {
        fail(`Sheet not found: ${parsed.sheet ?? "(active)"}`);
      }

      let written = 0;

      file.batch(() => {
        for (let r = 0; r < rows; r++) {
          throwIfAborted(signal);
          for (let c = 0; c < cols; c++) {
            const cell = cells[r][c];
            if (!cell) continue;

            const row = parsed.start.row + r;
            const col = parsed.start.col + c;

            if (cell.formula) {
              const f = cell.formula.startsWith("=")
                ? cell.formula.slice(1)
                : cell.formula;
              sheet.setFormula(row, col, f);
              written++;
            } else if (cell.value !== undefined) {
              sheet.setValue(row, col, cell.value);
              written++;
            }

            if (cell.cellStyles) {
              applyStyles(sheet, row, col, cell.cellStyles, GC);
            }
          }
        }
      });

      ok({ written, range: ref });
    },
    { save: true, signal },
  );
}
