/**
 * Formula hardcode lint hook.
 *
 * Flags formulas that appear to embed numeric assumptions directly instead of
 * referencing driver / assumption cells.
 *
 * Heuristics:
 *   - Flags decimals (e.g. 0.08, 1.05) and large thresholds (e.g. 1000000)
 *   - Flags arithmetic constants in formulas (e.g. +5000, *0.25)
 *   - Ignores common low-noise cases like ROUND(..., 2), INDEX(..., 1, 2),
 *     LEFT/RIGHT/MID lengths, OFFSET row/col offsets, DATE/TIME builders,
 *     and simple 0 / 1 / -1 constants
 *
 * Install:
 *   mkdir -p .headless-spreadjs/hooks
 *   cp examples/hooks/hardcode-lint.ts .headless-spreadjs/hooks/
 *
 * Now every save will print suspicious formula constants for changed cells.
 */

import type { HookAPI, HookContext } from "@hewliyang/headless-spreadjs/hooks";

const MAX_FINDINGS = 20;

/** Flags for getUsedRange: formulas (16) + values (32) */
const USED_RANGE_FORMULAS_AND_VALUES = 16 | 32;

interface FormulaCell {
  sheetName: string;
  row: number;
  col: number;
  formula: string;
}

interface NumberToken {
  raw: string;
  value: number;
  start: number;
  end: number;
}

interface FunctionContext {
  name: string;
  argIndex: number;
}

interface Finding {
  sheetName: string;
  row: number;
  col: number;
  formula: string;
  literals: string[];
}

const CONDITIONAL_FUNCTIONS = new Set([
  "IF",
  "IFS",
  "SWITCH",
  "CHOOSE",
  "MIN",
  "MAX",
  "MEDIAN",
]);

const SAFE_FUNCTION_ARGS = new Map<string, Set<number>>([
  ["ROUND", new Set([1])],
  ["ROUNDUP", new Set([1])],
  ["ROUNDDOWN", new Set([1])],
  ["MROUND", new Set([1])],
  ["CEILING", new Set([1])],
  ["FLOOR", new Set([1])],
  ["LEFT", new Set([1])],
  ["RIGHT", new Set([1])],
  ["MID", new Set([1, 2])],
  ["INDEX", new Set([1, 2])],
  ["OFFSET", new Set([1, 2, 3, 4])],
  ["DATE", new Set([0, 1, 2])],
  ["TIME", new Set([0, 1, 2])],
  ["EDATE", new Set([1])],
  ["EOMONTH", new Set([1])],
  ["WORKDAY", new Set([1])],
  ["WORKDAY.INTL", new Set([1, 2])],
  ["NETWORKDAYS", new Set([2])],
  ["NETWORKDAYS.INTL", new Set([2, 3])],
  ["YEAR", new Set([0])],
  ["MONTH", new Set([0])],
  ["DAY", new Set([0])],
  ["WEEKDAY", new Set([1])],
  ["VLOOKUP", new Set([2, 3])],
  ["HLOOKUP", new Set([2, 3])],
  ["MATCH", new Set([2])],
  ["XLOOKUP", new Set([4, 5])],
  ["XMATCH", new Set([2, 3])],
  ["LARGE", new Set([1])],
  ["SMALL", new Set([1])],
  ["RANK", new Set([2])],
  ["RANK.EQ", new Set([2])],
  ["RANK.AVG", new Set([2])],
  ["CHOOSE", new Set([0])],
  ["SUBSTITUTE", new Set([3])],
  ["FIND", new Set([2])],
  ["SEARCH", new Set([2])],
  ["TRUNC", new Set([1])],
  ["FIXED", new Set([1])],
  ["SUBTOTAL", new Set([0])],
  ["AGGREGATE", new Set([0, 1])],
  ["REPLACE", new Set([1, 2])],
  ["REPT", new Set([1])],
  ["ADDRESS", new Set([2, 3])],
  ["INDIRECT", new Set([1])],
  ["YEARFRAC", new Set([2])],
  ["DAYS360", new Set([2])],
  ["DB", new Set([4])],
  ["SORT", new Set([1, 2, 3])],
  ["SORTBY", new Set([2, 4, 6, 8])],
  ["UNIQUE", new Set([1, 2])],
  ["SEQUENCE", new Set([0, 1, 2, 3])],
  ["TAKE", new Set([1, 2])],
  ["DROP", new Set([1, 2])],
  ["WRAPROWS", new Set([1])],
  ["WRAPCOLS", new Set([1])],
  ["TOCOL", new Set([1, 2])],
  ["TOROW", new Set([1, 2])],
  ["EXPAND", new Set([1, 2, 3])],
  ["BASE", new Set([1, 2])],
  ["DECIMAL", new Set([1])],
  ["ROMAN", new Set([1])],
]);

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
      result += " ";
    } else if (inString) {
      result += " ";
    } else {
      result += ch;
    }
  }
  return result;
}

function maskSheetQualifiers(formula: string): string {
  return formula
    .replace(/'[^']+'!/g, "")
    .replace(/\b[A-Za-z_][A-Za-z0-9_.]*!/g, "");
}

function maskReferences(formula: string): string {
  return formula
    .replace(/\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?/gi, "__REF__")
    .replace(/\$?[A-Z]{1,3}:\$?[A-Z]{1,3}/gi, "__COLRANGE__")
    .replace(/\b\d+:\d+\b/g, "__ROWRANGE__");
}

function isIdentifierChar(ch: string | undefined): boolean {
  return !!ch && /[A-Za-z0-9_.]/.test(ch);
}

function previousNonSpace(text: string, index: number): string | undefined {
  for (let i = index - 1; i >= 0; i--) {
    if (!/\s/.test(text[i])) return text[i];
  }
  return undefined;
}

function nextNonSpace(text: string, index: number): string | undefined {
  for (let i = index; i < text.length; i++) {
    if (!/\s/.test(text[i])) return text[i];
  }
  return undefined;
}

function extractNumericTokens(formula: string): NumberToken[] {
  const regex = /(?:\d+\.\d*|\.\d+|\d+)(?:[Ee][-+]?\d+)?/g;
  const out: NumberToken[] = [];

  for (const match of formula.matchAll(regex)) {
    const raw = match[0];
    const start = match.index ?? 0;
    const end = start + raw.length;
    const prev = formula[start - 1];
    const next = formula[end];

    if (isIdentifierChar(prev) || isIdentifierChar(next)) continue;

    const value = Number(raw);
    if (Number.isNaN(value)) continue;

    out.push({ raw, value, start, end });
  }

  return out;
}

function getFunctionContextAt(
  formula: string,
  index: number,
): FunctionContext | null {
  const stack: Array<{ name: string | null; argIndex: number }> = [];

  for (let i = 0; i < index; i++) {
    const ch = formula[i];

    if (/[A-Za-z_]/.test(ch)) {
      let j = i + 1;
      while (j < index && /[A-Za-z0-9_.]/.test(formula[j])) j++;
      let k = j;
      while (k < index && /\s/.test(formula[k])) k++;
      if (k < index && formula[k] === "(") {
        stack.push({ name: formula.slice(i, j).toUpperCase(), argIndex: 0 });
        i = k;
        continue;
      }
      i = j - 1;
      continue;
    }

    if (ch === "(") {
      stack.push({ name: null, argIndex: 0 });
      continue;
    }

    if (ch === "," || ch === ";") {
      for (let s = stack.length - 1; s >= 0; s--) {
        if (stack[s].name) {
          stack[s].argIndex++;
          break;
        }
      }
      continue;
    }

    if (ch === ")") {
      stack.pop();
    }
  }

  for (let i = stack.length - 1; i >= 0; i--) {
    const entry = stack[i];
    if (entry.name) return { name: entry.name, argIndex: entry.argIndex };
  }

  return null;
}

function isSafeFunctionArg(ctx: FunctionContext | null): boolean {
  if (!ctx) return false;
  const safeArgs = SAFE_FUNCTION_ARGS.get(ctx.name);
  return safeArgs?.has(ctx.argIndex) ?? false;
}

function isSuspiciousNumber(token: NumberToken, formula: string): boolean {
  if ([0, 1, -1].includes(token.value)) return false;

  const ctx = getFunctionContextAt(formula, token.start);
  if (isSafeFunctionArg(ctx)) return false;

  const prev = previousNonSpace(formula, token.start);
  const next = nextNonSpace(formula, token.end);
  const abs = Math.abs(token.value);
  const hasDecimal = token.raw.includes(".") || /e/i.test(token.raw);

  if (ctx && CONDITIONAL_FUNCTIONS.has(ctx.name)) return true;

  if (hasDecimal && abs !== 0 && abs !== 1) return true;
  if (abs >= 1000) return true;

  if (
    prev &&
    /[+\-*/^&><=]/.test(prev) &&
    (!next || /[+\-*/^&><=),]/.test(next))
  ) {
    return true;
  }

  if ((!prev || /[,(]/.test(prev)) && next && /[+\-*/^&><=)]/.test(next)) {
    return true;
  }

  return false;
}

function colToLetters(col: number): string {
  let s = "";
  let n = col;
  while (n >= 0) {
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

function cellRef(sheetName: string, row: number, col: number): string {
  return `${sheetName}!${colToLetters(col)}${row + 1}`;
}

function forEachMutatedFormulaCell(
  ctx: HookContext,
  callback: (cell: FormulaCell) => void,
) {
  for (const range of ctx.mutatedRanges) {
    const ws = range.sheet
      ? ctx.workbook.getSheetFromName(range.sheet)
      : ctx.workbook.getActiveSheet();
    if (!ws) continue;
    const sheetName = ws.name();

    for (let r = range.startRow; r <= range.endRow; r++) {
      for (let c = range.startCol; c <= range.endCol; c++) {
        const formula = ws.getFormula(r, c);
        if (!formula) continue;
        callback({ sheetName, row: r, col: c, formula });
      }
    }
  }
}

function forEachUsedFormulaCell(
  ctx: HookContext,
  callback: (cell: FormulaCell) => void,
) {
  const sheetCount = ctx.workbook.getSheetCount();
  for (let si = 0; si < sheetCount; si++) {
    const ws = ctx.workbook.getSheet(si);
    const sheetName = ws.name();

    let range: any;
    try {
      range = ws.getUsedRange(USED_RANGE_FORMULAS_AND_VALUES);
    } catch {
      continue;
    }
    if (!range || range.rowCount <= 0 || range.colCount <= 0) continue;

    for (let r = range.row; r < range.row + range.rowCount; r++) {
      for (let c = range.col; c < range.col + range.colCount; c++) {
        const formula = ws.getFormula(r, c);
        if (!formula) continue;
        callback({ sheetName, row: r, col: c, formula });
      }
    }
  }
}

function visitFormulaCells(
  ctx: HookContext,
  callback: (cell: FormulaCell) => void,
) {
  if (ctx.mutatedRanges.length > 0) {
    forEachMutatedFormulaCell(ctx, callback);
  } else {
    forEachUsedFormulaCell(ctx, callback);
  }
}

function findHardcodes(formula: string): string[] {
  const raw = formula.startsWith("=") ? formula.slice(1) : formula;
  const noSheets = maskSheetQualifiers(raw);
  const stripped = removeStringLiterals(noSheets);
  const masked = maskReferences(stripped);
  const tokens = extractNumericTokens(masked);

  const findings = new Set<string>();
  for (const token of tokens) {
    if (isSuspiciousNumber(token, masked)) {
      findings.add(token.raw);
    }
  }

  return [...findings];
}

export default function (hsx: HookAPI) {
  hsx.on("postSave", function lintFormulaHardcodes(ctx: HookContext) {
    const findings: Finding[] = [];

    visitFormulaCells(ctx, ({ sheetName, row, col, formula }) => {
      const literals = findHardcodes(formula);
      if (literals.length === 0) return;
      findings.push({ sheetName, row, col, formula, literals });
    });

    if (findings.length === 0) return;

    console.log(`Suspicious formula hardcodes (${findings.length}):`);
    for (const finding of findings.slice(0, MAX_FINDINGS)) {
      console.log(
        `  ${cellRef(finding.sheetName, finding.row, finding.col)}: ${finding.formula} -> ${finding.literals.join(", ")}`,
      );
    }
    if (findings.length > MAX_FINDINGS) {
      console.log(`  ... +${findings.length - MAX_FINDINGS} more`);
    }
  });
}
