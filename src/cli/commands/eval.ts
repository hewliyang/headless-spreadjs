import { parseRef } from "../a1.js";
import { withFile } from "../context.js";
import { fail, ok, readInput } from "../output.js";
import { ensureSheetSize } from "../sheet-size.js";

type MarkMutatedFn = (range: {
  sheet?: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}) => void;

function toIndex(value: unknown): number | null {
  return typeof value === "number" && Number.isInteger(value) && value >= 0
    ? value
    : null;
}

function markCell(
  markMutated: MarkMutatedFn,
  sheetName: string,
  row: number,
  col: number,
): void {
  markMutated({
    sheet: sheetName,
    startRow: row,
    startCol: col,
    endRow: row,
    endCol: col,
  });
}

function markRect(
  markMutated: MarkMutatedFn,
  sheetName: string,
  row: number,
  col: number,
  rowCount: number,
  colCount: number,
): void {
  markMutated({
    sheet: sheetName,
    startRow: row,
    startCol: col,
    endRow: row + rowCount - 1,
    endCol: col + colCount - 1,
  });
}

function wrapSheet(sheet: any, markMutated: MarkMutatedFn): any {
  const proxy = new Proxy(sheet, {
    get(target, prop, receiver) {
      const value = Reflect.get(target, prop, receiver);
      if (typeof value !== "function") return value;

      return (...args: unknown[]) => {
        const sheetName = target.name();

        switch (prop) {
          case "setValue":
          case "setFormula":
          case "setStyle":
          case "setText":
          case "setTag": {
            const row = toIndex(args[0]);
            const col = toIndex(args[1]);
            if (row !== null && col !== null) {
              ensureSheetSize(target, row + 1, col + 1);
              markCell(markMutated, sheetName, row, col);
            }
            break;
          }

          case "getValue":
          case "getFormula":
          case "getStyle":
          case "getText":
          case "getTag":
          case "getCell": {
            const row = toIndex(args[0]);
            const col = toIndex(args[1]);
            if (row !== null && col !== null) {
              ensureSheetSize(target, row + 1, col + 1);
            }
            break;
          }

          case "getRange":
          case "clear": {
            const row = toIndex(args[0]);
            const col = toIndex(args[1]);
            const rowCount = toIndex(args[2]);
            const colCount = toIndex(args[3]);
            if (
              row !== null &&
              col !== null &&
              rowCount !== null &&
              colCount !== null &&
              rowCount > 0 &&
              colCount > 0
            ) {
              ensureSheetSize(target, row + rowCount, col + colCount);
              if (prop === "clear") {
                markRect(markMutated, sheetName, row, col, rowCount, colCount);
              }
            }
            break;
          }

          case "copyTo": {
            const dstRow = toIndex(args[2]);
            const dstCol = toIndex(args[3]);
            const rowCount = toIndex(args[4]);
            const colCount = toIndex(args[5]);
            if (
              dstRow !== null &&
              dstCol !== null &&
              rowCount !== null &&
              colCount !== null &&
              rowCount > 0 &&
              colCount > 0
            ) {
              ensureSheetSize(target, dstRow + rowCount, dstCol + colCount);
              markRect(
                markMutated,
                sheetName,
                dstRow,
                dstCol,
                rowCount,
                colCount,
              );
            }
            break;
          }

          case "setArray": {
            const row = toIndex(args[0]);
            const col = toIndex(args[1]);
            const values = args[2];
            if (row !== null && col !== null && Array.isArray(values)) {
              const rowCount = values.length;
              const colCount = values.reduce(
                (max, item) =>
                  Array.isArray(item) ? Math.max(max, item.length) : max,
                0,
              );
              if (rowCount > 0 && colCount > 0) {
                ensureSheetSize(target, row + rowCount, col + colCount);
                markRect(markMutated, sheetName, row, col, rowCount, colCount);
              }
            }
            break;
          }
        }

        return value.apply(target, args);
      };
    },
  });

  return proxy;
}

function wrapWorkbook(workbook: any, markMutated: MarkMutatedFn): any {
  const sheetCache = new WeakMap<object, any>();

  const getWrappedSheet = (sheet: any) => {
    if (!sheet || typeof sheet !== "object") return sheet;
    const cached = sheetCache.get(sheet);
    if (cached) return cached;
    const wrapped = wrapSheet(sheet, markMutated);
    sheetCache.set(sheet, wrapped);
    return wrapped;
  };

  return new Proxy(workbook, {
    get(target, prop, receiver) {
      const value = Reflect.get(target, prop, receiver);
      if (typeof value !== "function") return value;

      if (
        prop === "getActiveSheet" ||
        prop === "getSheet" ||
        prop === "getSheetFromName"
      ) {
        return (...args: unknown[]) =>
          getWrappedSheet(value.apply(target, args));
      }

      return (...args: unknown[]) => value.apply(target, args);
    },
  });
}

export async function evalCode(
  filePath: string,
  codeArg: string | undefined,
  options?: { signal?: AbortSignal | null },
): Promise<void> {
  const code = await readInput(codeArg);

  await withFile(
    filePath,
    async ({ file, workbook, GC, markMutated }) => {
      const wrappedWorkbook = wrapWorkbook(workbook, markMutated);
      const sheet = wrappedWorkbook.getActiveSheet();

      const range = (ref: string) => {
        if (typeof ref !== "string" || ref.trim().length === 0) {
          throw new Error(
            "range(ref) expects a non-empty A1 reference string.",
          );
        }

        const parsed = parseRef(ref.trim());
        const targetSheet = parsed.sheet
          ? wrappedWorkbook.getSheetFromName(parsed.sheet)
          : wrappedWorkbook.getActiveSheet();

        if (!targetSheet) {
          throw new Error(`Sheet not found: ${parsed.sheet ?? "(active)"}`);
        }

        const rowCount = parsed.end.row - parsed.start.row + 1;
        const colCount = parsed.end.col - parsed.start.col + 1;

        if (rowCount < 1 || colCount < 1) {
          throw new Error(`Invalid range reference: ${ref}`);
        }

        ensureSheetSize(targetSheet, parsed.end.row + 1, parsed.end.col + 1);

        return targetSheet.getRange(
          parsed.start.row,
          parsed.start.col,
          rowCount,
          colCount,
        );
      };

      const logs: string[] = [];
      const origLog = console.log;
      const origWarn = console.warn;
      const origError = console.error;
      console.log = (...args: unknown[]) =>
        logs.push(args.map(String).join(" "));
      console.warn = (...args: unknown[]) =>
        logs.push(`[warn] ${args.map(String).join(" ")}`);
      console.error = (...args: unknown[]) =>
        logs.push(`[error] ${args.map(String).join(" ")}`);

      try {
        const AsyncFunction = Object.getPrototypeOf(async () => {}).constructor;
        const fn = new AsyncFunction(
          "workbook",
          "sheet",
          "GC",
          "file",
          "range",
          code,
        );
        const result = await fn(wrappedWorkbook, sheet, GC, file, range);

        const output: Record<string, unknown> = {};
        if (result !== undefined) {
          output.result = result;
        }
        if (logs.length > 0) {
          output.logs = logs;
        }
        if (Object.keys(output).length === 0) {
          output.result = null;
        }

        ok(output);
      } catch (err) {
        fail(err instanceof Error ? err.message : String(err));
      } finally {
        console.log = origLog;
        console.warn = origWarn;
        console.error = origError;
      }
    },
    { save: true, signal: options?.signal },
  );
}
