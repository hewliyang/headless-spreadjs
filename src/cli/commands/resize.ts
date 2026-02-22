import { colToIndex } from "../a1.js";
import { throwIfAborted } from "../abort.js";
import { withFile } from "../context.js";
import { fail, ok } from "../output.js";

const USED_RANGE_TYPE_DATA_FORMULA = 16 | 32;

function parseColumnSpan(ref: string): { start: number; end: number } {
  const trimmed = ref.trim();
  const match = trimmed.match(/^([A-Za-z]+)(?::([A-Za-z]+))?$/);
  if (!match) {
    fail(`Invalid --columns reference: ${ref}`);
  }

  const start = colToIndex(match[1]);
  const end = match[2] ? colToIndex(match[2]) : start;
  if (start < 0 || end < 0 || end < start) {
    fail(`Invalid --columns reference: ${ref}`);
  }

  return { start, end };
}

function parseRowSpan(ref: string): { start: number; end: number } {
  const trimmed = ref.trim();
  const match = trimmed.match(/^(\d+)(?::(\d+))?$/);
  if (!match) {
    fail(`Invalid --rows reference: ${ref}`);
  }

  const start = Number.parseInt(match[1], 10) - 1;
  const end = match[2] ? Number.parseInt(match[2], 10) - 1 : start;
  if (start < 0 || end < 0 || end < start) {
    fail(`Invalid --rows reference: ${ref}`);
  }

  return { start, end };
}

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
        let usedRange: ReturnType<typeof sheet.getUsedRange> | null = null;
        if (!options.columns || !options.rows) {
          try {
            usedRange = sheet.getUsedRange(USED_RANGE_TYPE_DATA_FORMULA);
          } catch {
            usedRange = null;
          }
        }

        if (options.width !== undefined) {
          if (options.columns) {
            const span = parseColumnSpan(options.columns);
            for (let c = span.start; c <= span.end; c++) {
              throwIfAborted(signal);
              sheet.setColumnWidth(c, options.width);
            }
          } else if (usedRange && usedRange.colCount > 0) {
            const endCol = usedRange.col + usedRange.colCount;
            for (let c = usedRange.col; c < endCol; c++) {
              throwIfAborted(signal);
              sheet.setColumnWidth(c, options.width);
            }
          }
        }

        if (options.height !== undefined) {
          if (options.rows) {
            const span = parseRowSpan(options.rows);
            for (let r = span.start; r <= span.end; r++) {
              throwIfAborted(signal);
              sheet.setRowHeight(r, options.height);
            }
          } else if (usedRange && usedRange.rowCount > 0) {
            const endRow = usedRange.row + usedRange.rowCount;
            for (let r = usedRange.row; r < endRow; r++) {
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
