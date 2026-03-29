/**
 * Financial modeling color conventions hook.
 *
 * Automatically applies and lints financial color coding on every save:
 *   - Blue:   Hardcoded numeric inputs
 *   - Black:  Formulas referencing same-sheet cells
 *   - Green:  Formulas referencing other sheets (cross-sheet links)
 *
 * When the command reports which cells it mutated (e.g. `hsx set`),
 * only those cells are colored. For opaque commands (e.g. `hsx eval`),
 * falls back to scanning every used cell.
 *
 * Install:
 *   mkdir -p .headless-spreadjs/hooks
 *   cp examples/hooks/financial-colors.ts .headless-spreadjs/hooks/
 *
 * Now every `hsx set`, `hsx eval`, etc. will auto-color before saving
 * and lint color violations in a postSave pass.
 */

import type { HookAPI, HookContext } from "@hewliyang/headless-spreadjs/hooks";

const COLORS = {
  HARDCODE: "Blue",
  FORMULA: "Black",
  EXTERNAL_LINK: "Green",
} as const;

type FinancialColor = (typeof COLORS)[keyof typeof COLORS];

function removeStringLiterals(formula: string): string {
  let result = "";
  let inString = false;
  for (let i = 0; i < formula.length; i++) {
    const ch = formula[i];
    if (ch === '"') {
      if (inString && i + 1 < formula.length && formula[i + 1] === '"') {
        i++;
        continue;
      }
      inString = !inString;
      result += ch;
    } else if (!inString) {
      result += ch;
    }
  }
  return result;
}

function inferColor(
  value: unknown,
  formula: string | null,
): FinancialColor | null {
  if (formula) {
    const stripped = removeStringLiterals(formula);
    if (stripped.includes("!")) return COLORS.EXTERNAL_LINK;
    return COLORS.FORMULA;
  }
  if (typeof value === "number") return COLORS.HARDCODE;
  return null;
}

function colorName(color: string | undefined): string {
  if (!color) return "none";
  const lower = color.toLowerCase();
  if (lower === "blue" || lower === "#0000ff") return "HARDCODE (blue)";
  if (lower === "black" || lower === "#000000") return "FORMULA (black)";
  if (lower === "green" || lower === "#008000") return "EXTERNAL_LINK (green)";
  return color;
}

type CellVisitor = (info: {
  ws: any;
  sheetName: string;
  row: number;
  col: number;
  value: unknown;
  formula: string | null;
}) => void;

function forEachMutatedCell(ctx: HookContext, callback: CellVisitor) {
  for (const range of ctx.mutatedRanges) {
    const ws = range.sheet
      ? ctx.workbook.getSheetFromName(range.sheet)
      : ctx.workbook.getActiveSheet();
    if (!ws) continue;
    const sheetName = ws.name();

    for (let r = range.startRow; r <= range.endRow; r++) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        const value = ws.getValue(r, c);
        const formula = ws.getFormula(r, c) || null;
        if (value == null && !formula) continue;
        callback({ ws, sheetName, row: r, col: c, value, formula });
      }
    }
  }
}

function forEachUsedCell(workbook: any, callback: CellVisitor) {
  const sheetCount = workbook.getSheetCount();
  for (let si = 0; si < sheetCount; si++) {
    const ws = workbook.getSheet(si);
    const sheetName = ws.name();

    let range: any;
    try {
      range = ws.getUsedRange(16 | 32);
    } catch {
      continue;
    }
    if (!range || range.rowCount <= 0 || range.colCount <= 0) continue;

    const endRow = range.row + range.rowCount;
    const endCol = range.col + range.colCount;

    for (let r = range.row; r < endRow; r++) {
      for (let c = range.col; c < endCol; c++) {
        const value = ws.getValue(r, c);
        const formula = ws.getFormula(r, c) || null;
        if (value == null && !formula) continue;
        callback({ ws, sheetName, row: r, col: c, value, formula });
      }
    }
  }
}

function visitCells(ctx: HookContext, callback: CellVisitor) {
  if (ctx.mutatedRanges.length > 0) {
    forEachMutatedCell(ctx, callback);
  } else {
    forEachUsedCell(ctx.workbook, callback);
  }
}

// ---------------------------------------------------------------------------
// Hook entry point
// ---------------------------------------------------------------------------
export default function (hooks: HookAPI) {
  // Apply: color cells according to their type before each save
  hooks.on("preSave", function applyFinancialColors(ctx: HookContext) {
    let colored = 0;

    visitCells(ctx, ({ ws, row, col, value, formula }) => {
      const target = inferColor(value, formula);
      if (!target) return;

      const cell = ws.getCell(row, col);
      const current = cell.foreColor();
      if (current !== target) {
        cell.foreColor(target);
        colored++;
      }
    });

    if (colored > 0) {
      const scope =
        ctx.mutatedRanges.length > 0
          ? `${ctx.mutatedRanges.length} range(s)`
          : "full workbook";
      console.log(`Applied financial colors to ${colored} cell(s) [${scope}]`);
    }
  });

  // Lint: report any remaining violations after save
  hooks.on("postSave", function lintFinancialColors(ctx: HookContext) {
    const violations: Array<{
      ref: string;
      value: unknown;
      expected: string;
      current: string | undefined;
    }> = [];

    visitCells(ctx, ({ sheetName, row, col, value, formula }) => {
      const expected = inferColor(value, formula);
      if (!expected) return;

      const cell = ctx.workbook.getSheetFromName(sheetName).getCell(row, col);
      const current: string | undefined = cell.foreColor();

      if (current !== expected) {
        const colLetter = String.fromCharCode(65 + (col % 26));
        const ref = `${sheetName}!${colLetter}${row + 1}`;
        violations.push({ ref, value, expected, current });
      }
    });

    if (violations.length > 0) {
      console.log(`Color violations (${violations.length}):`);
      const limit = 10;
      for (const v of violations.slice(0, limit)) {
        console.log(
          `  ${v.ref}: need ${colorName(v.expected)}, have ${colorName(v.current)}`,
        );
      }
      if (violations.length > limit) {
        console.log(`  ... +${violations.length - limit} more`);
      }
    }
  });
}
