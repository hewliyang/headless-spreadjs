import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { cellToA1 } from "../a1.js";
import { throwIfAborted } from "../abort.js";
import { withFile } from "../context.js";
import { ok } from "../output.js";

const DEFAULT_INLINE_LIMIT = 2000;
const DEFAULT_PREVIEW_LIMIT = 200;
const USED_RANGE_TYPE_DATA_FORMULA = 16 | 32;

interface CellSnapshot {
  row: number;
  col: number;
  value: unknown;
  formula: string | null;
}

interface DiffEntry {
  sheet: string;
  cell: string;
  left: { value: unknown; formula: string | null };
  right: { value: unknown; formula: string | null };
}

function toCellKey(row: number, col: number): string {
  return `${row}:${col}`;
}

function fromCellKey(key: string): { row: number; col: number } {
  const [row, col] = key.split(":");
  return {
    row: Number.parseInt(row, 10),
    col: Number.parseInt(col, 10),
  };
}

function collectUsedCells(
  sheet: {
    getUsedRange(type: number):
      | { row: number; col: number; rowCount: number; colCount: number }
      | null;
    getValue(row: number, col: number): unknown;
    getFormula(row: number, col: number): string | null;
  },
  signal?: AbortSignal | null,
): Map<string, CellSnapshot> {
  const cells = new Map<string, CellSnapshot>();
  let usedRange: { row: number; col: number; rowCount: number; colCount: number } | null =
    null;

  try {
    usedRange = sheet.getUsedRange(USED_RANGE_TYPE_DATA_FORMULA);
  } catch {
    return cells;
  }

  if (!usedRange || usedRange.rowCount <= 0 || usedRange.colCount <= 0) {
    return cells;
  }

  const rowEnd = usedRange.row + usedRange.rowCount;
  const colEnd = usedRange.col + usedRange.colCount;

  for (let row = usedRange.row; row < rowEnd; row++) {
    throwIfAborted(signal);
    for (let col = usedRange.col; col < colEnd; col++) {
      const value = sheet.getValue(row, col);
      const formula = sheet.getFormula(row, col);
      const hasFormula = !!formula;
      const hasValue = value !== null && value !== undefined && value !== "";

      if (!hasFormula && !hasValue) {
        continue;
      }

      cells.set(toCellKey(row, col), {
        row,
        col,
        value: value ?? null,
        formula: formula || null,
      });
    }
  }

  return cells;
}

function valuesEqual(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) {
    return true;
  }

  if (left instanceof Date && right instanceof Date) {
    return left.getTime() === right.getTime();
  }

  if (
    left &&
    right &&
    typeof left === "object" &&
    typeof right === "object"
  ) {
    try {
      return JSON.stringify(left) === JSON.stringify(right);
    } catch {
      return false;
    }
  }

  return false;
}

function summarizeDivergence(
  changedCells: number,
  comparedCells: number,
  sheetsWithDifferences: number,
  sheetsOnlyLeft: number,
  sheetsOnlyRight: number,
): string {
  if (changedCells === 0) {
    return "No differences found.";
  }

  const ratio = comparedCells > 0 ? changedCells / comparedCells : 1;
  const level = ratio < 0.01 ? "Low divergence" : ratio < 0.2 ? "Moderate divergence" : "High divergence";

  return `${level}: ${changedCells} changed cells across ${sheetsWithDifferences} sheet(s). Left-only sheets: ${sheetsOnlyLeft}, right-only sheets: ${sheetsOnlyRight}.`;
}

export async function diff(
  leftFilePath: string,
  rightFilePath: string,
  options?: {
    signal?: AbortSignal | null;
    inlineLimit?: number;
    previewLimit?: number;
  },
): Promise<void> {
  const signal = options?.signal;
  const inlineLimit =
    options?.inlineLimit && options.inlineLimit > 0
      ? options.inlineLimit
      : DEFAULT_INLINE_LIMIT;
  const previewLimit =
    options?.previewLimit && options.previewLimit > 0
      ? options.previewLimit
      : DEFAULT_PREVIEW_LIMIT;

  await withFile(leftFilePath, ({ workbook: leftWorkbook }) =>
    withFile(rightFilePath, ({ workbook: rightWorkbook }) => {
      const leftSheetNames = new Set<string>();
      const rightSheetNames = new Set<string>();

      for (let i = 0; i < leftWorkbook.getSheetCount(); i++) {
        throwIfAborted(signal);
        leftSheetNames.add(leftWorkbook.getSheet(i).name());
      }

      for (let i = 0; i < rightWorkbook.getSheetCount(); i++) {
        throwIfAborted(signal);
        rightSheetNames.add(rightWorkbook.getSheet(i).name());
      }

      const allSheetNames = Array.from(
        new Set([...leftSheetNames, ...rightSheetNames]),
      ).sort((a, b) => a.localeCompare(b));

      const previews: DiffEntry[] = [];
      let bufferedDiffs: DiffEntry[] = [];
      let outputFilePath: string | null = null;
      let outputDirPath: string | null = null;

      const sheetStats: Record<string, { changedCells: number }> = {};

      let changedCells = 0;
      let comparedCells = 0;

      const ensureOutputFile = (): string => {
        if (outputFilePath) {
          return outputFilePath;
        }

        outputDirPath = fs.mkdtempSync(path.join(os.tmpdir(), "hsx-diff-"));
        outputFilePath = path.join(outputDirPath, "diff.ndjson");

        if (bufferedDiffs.length > 0) {
          const bootstrap = bufferedDiffs
            .map((entry) => `${JSON.stringify(entry)}\n`)
            .join("");
          fs.writeFileSync(outputFilePath, bootstrap, "utf8");
          bufferedDiffs = [];
        }

        return outputFilePath;
      };

      const recordDiff = (entry: DiffEntry): void => {
        changedCells++;
        sheetStats[entry.sheet] ??= { changedCells: 0 };
        sheetStats[entry.sheet].changedCells++;

        if (previews.length < previewLimit) {
          previews.push(entry);
        }

        if (outputFilePath) {
          fs.appendFileSync(outputFilePath, `${JSON.stringify(entry)}\n`, "utf8");
          return;
        }

        bufferedDiffs.push(entry);
        if (bufferedDiffs.length > inlineLimit) {
          ensureOutputFile();
        }
      };

      for (const sheetName of allSheetNames) {
        throwIfAborted(signal);
        const leftSheet = leftWorkbook.getSheetFromName(sheetName);
        const rightSheet = rightWorkbook.getSheetFromName(sheetName);

        const leftCells = leftSheet ? collectUsedCells(leftSheet, signal) : new Map();
        const rightCells = rightSheet
          ? collectUsedCells(rightSheet, signal)
          : new Map();

        const cellKeys = new Set<string>([
          ...leftCells.keys(),
          ...rightCells.keys(),
        ]);

        comparedCells += cellKeys.size;

        for (const key of cellKeys) {
          throwIfAborted(signal);
          const coords = fromCellKey(key);
          const left = leftCells.get(key) ?? {
            row: coords.row,
            col: coords.col,
            value: null,
            formula: null,
          };
          const right = rightCells.get(key) ?? {
            row: coords.row,
            col: coords.col,
            value: null,
            formula: null,
          };

          const sameFormula = (left.formula || null) === (right.formula || null);
          const sameValue = valuesEqual(left.value, right.value);
          if (sameFormula && sameValue) {
            continue;
          }

          recordDiff({
            sheet: sheetName,
            cell: cellToA1(left.row, left.col),
            left: {
              value: left.value,
              formula: left.formula,
            },
            right: {
              value: right.value,
              formula: right.formula,
            },
          });
        }
      }

      const sheetsWithDifferences = Object.keys(sheetStats).length;
      const sheetsOnlyLeft = Array.from(leftSheetNames).filter(
        (name) => !rightSheetNames.has(name),
      );
      const sheetsOnlyRight = Array.from(rightSheetNames).filter(
        (name) => !leftSheetNames.has(name),
      );

      const summary = summarizeDivergence(
        changedCells,
        comparedCells,
        sheetsWithDifferences,
        sheetsOnlyLeft.length,
        sheetsOnlyRight.length,
      );

      const inlineDiffs = outputFilePath ? previews : bufferedDiffs;

      ok({
        summary,
        leftFile: leftFilePath,
        rightFile: rightFilePath,
        changedCells,
        comparedCells,
        sheetsCompared: allSheetNames.length,
        sheetsWithDifferences,
        sheetsOnlyLeft,
        sheetsOnlyRight,
        inlineLimit,
        previewLimit,
        outputMode: outputFilePath ? "tmpfile" : "inline",
        ...(outputDirPath && outputFilePath
          ? {
              diffFile: outputFilePath,
              tempDir: outputDirPath,
              note:
                "Diff set is large; full diff was written to diffFile as NDJSON for grep-friendly inspection.",
            }
          : {}),
        diffs: inlineDiffs,
        sheetStats,
      });
    }),
  );
}
