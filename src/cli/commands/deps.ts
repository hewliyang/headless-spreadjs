import type { SpreadWorkbook, SpreadWorksheet } from "../../types.js";
import { cellToA1, parseRef } from "../a1.js";
import { throwIfAborted } from "../abort.js";
import { withFile } from "../context.js";
import { fail, ok } from "../output.js";

const USED_RANGE_TYPE_DATA_FORMULA = 16 | 32;
const DEFAULT_RANGE_EXPANSION_LIMIT = 128;
const DEFAULT_MAX_FORMULA_SCAN = 250_000;

interface PrecedentRange {
  sheet: string;
  row: number;
  col: number;
  rowCount: number;
  colCount: number;
}

interface TraceCell {
  sheet: string;
  row: number;
  col: number;
}

interface FormulaNode {
  sheet: string;
  row: number;
  col: number;
  precedents: PrecedentRange[];
}

type RawPrecedent = NonNullable<
  ReturnType<SpreadWorksheet["getPrecedents"]>
>[number];
type RawUsedRange = ReturnType<SpreadWorksheet["getUsedRange"]>;

function makeCellKey(sheet: string, row: number, col: number): string {
  return `${sheet}|${row}|${col}`;
}

function makeRangeKey(range: PrecedentRange): string {
  return `${range.sheet}|${range.row}|${range.col}|${range.rowCount}|${range.colCount}`;
}

function toA1Ref(range: PrecedentRange): string {
  const start = cellToA1(range.row, range.col);
  if (range.rowCount === 1 && range.colCount === 1) {
    return start;
  }
  const end = cellToA1(
    range.row + range.rowCount - 1,
    range.col + range.colCount - 1,
  );
  return `${start}:${end}`;
}

function normalizePrecedents(
  raw: ReturnType<SpreadWorksheet["getPrecedents"]>,
  defaultSheet: string,
): PrecedentRange[] {
  if (!raw || raw.length === 0) return [];

  return raw
    .map((item: RawPrecedent) => ({
      sheet: item.sheetName || defaultSheet,
      row: typeof item.row === "number" ? item.row : -1,
      col: typeof item.col === "number" ? item.col : -1,
      rowCount: typeof item.rowCount === "number" ? item.rowCount : 1,
      colCount: typeof item.colCount === "number" ? item.colCount : 1,
    }))
    .filter(
      (item) =>
        item.row >= 0 &&
        item.col >= 0 &&
        item.rowCount > 0 &&
        item.colCount > 0,
    );
}

function isSingleCell(range: PrecedentRange): boolean {
  return range.rowCount === 1 && range.colCount === 1;
}

function rangeContainsCell(range: PrecedentRange, cell: TraceCell): boolean {
  if (range.sheet !== cell.sheet) return false;
  return (
    cell.row >= range.row &&
    cell.row < range.row + range.rowCount &&
    cell.col >= range.col &&
    cell.col < range.col + range.colCount
  );
}

function parseTarget(ref: string): {
  sheet?: string;
  row: number;
  col: number;
} {
  const parsed = parseRef(ref);
  if (
    parsed.start.row !== parsed.end.row ||
    parsed.start.col !== parsed.end.col
  ) {
    fail(`Expected a single-cell reference, got range: ${ref}`);
  }

  return {
    sheet: parsed.sheet,
    row: parsed.start.row,
    col: parsed.start.col,
  };
}

export async function deps(
  filePath: string,
  ref: string,
  options?: {
    signal?: AbortSignal | null;
    depth?: number;
    rangeExpansionLimit?: number;
  },
): Promise<void> {
  const signal = options?.signal;
  const depth = options?.depth && options.depth > 0 ? options.depth : 1;
  const rangeExpansionLimit =
    options?.rangeExpansionLimit && options.rangeExpansionLimit > 0
      ? options.rangeExpansionLimit
      : DEFAULT_RANGE_EXPANSION_LIMIT;

  await withFile(
    filePath,
    ({ workbook }) => {
      const target = parseTarget(ref);
      const defaultSheet = workbook.getActiveSheet().name();
      const startSheetName = target.sheet || defaultSheet;
      const startSheet = workbook.getSheetFromName(startSheetName);

      if (!startSheet) {
        fail(`Sheet not found: ${startSheetName}`);
      }

      let frontier = new Map<string, TraceCell>([
        [
          makeCellKey(startSheetName, target.row, target.col),
          {
            sheet: startSheetName,
            row: target.row,
            col: target.col,
          },
        ],
      ]);

      const visited = new Set<string>(frontier.keys());
      const discovered = new Map<
        string,
        { range: PrecedentRange; hop: number }
      >();

      let hopsTraversed = 0;
      let cellsExpanded = 0;
      let rangesNotExpanded = 0;
      let sourceCellsVisited = 0;

      for (let hop = 1; hop <= depth && frontier.size > 0; hop++) {
        throwIfAborted(signal);
        hopsTraversed = hop;
        const next = new Map<string, TraceCell>();

        for (const cell of frontier.values()) {
          throwIfAborted(signal);
          sourceCellsVisited++;
          const sheet = workbook.getSheetFromName(cell.sheet);
          if (!sheet) continue;

          const precedents = normalizePrecedents(
            sheet.getPrecedents(cell.row, cell.col),
            cell.sheet,
          );

          for (const precedent of precedents) {
            const rangeKey = makeRangeKey(precedent);
            const existing = discovered.get(rangeKey);
            if (!existing || hop < existing.hop) {
              discovered.set(rangeKey, { range: precedent, hop });
            }

            if (hop >= depth) continue;

            const size = precedent.rowCount * precedent.colCount;
            if (size > rangeExpansionLimit) {
              rangesNotExpanded++;
              continue;
            }

            for (let r = 0; r < precedent.rowCount; r++) {
              for (let c = 0; c < precedent.colCount; c++) {
                const row = precedent.row + r;
                const col = precedent.col + c;
                const key = makeCellKey(precedent.sheet, row, col);
                if (visited.has(key)) continue;
                visited.add(key);
                cellsExpanded++;
                next.set(key, {
                  sheet: precedent.sheet,
                  row,
                  col,
                });
              }
            }
          }
        }

        frontier = next;
      }

      const dependencies = Array.from(discovered.values())
        .map(({ range, hop }) => ({
          sheet: range.sheet,
          ref: toA1Ref(range),
          hop,
          type: isSingleCell(range) ? "cell" : "range",
        }))
        .sort((a, b) => {
          if (a.hop !== b.hop) return a.hop - b.hop;
          if (a.sheet !== b.sheet) return a.sheet.localeCompare(b.sheet);
          return a.ref.localeCompare(b.ref);
        });

      ok({
        target: {
          sheet: startSheetName,
          cell: cellToA1(target.row, target.col),
        },
        requestedDepth: depth,
        hopsTraversed,
        oneHop: depth === 1,
        dependencies,
        stats: {
          sourceCellsVisited,
          cellsExpanded,
          rangesNotExpanded,
          rangeExpansionLimit,
        },
        note: "Best-effort trace. Dynamic formulas (e.g. INDIRECT/OFFSET) may be incomplete.",
      });
    },
    { signal },
  );
}

function collectFormulaNodes(
  workbook: SpreadWorkbook,
  maxFormulaCells: number,
  signal?: AbortSignal | null,
): { nodes: FormulaNode[]; scannedFormulaCells: number; truncated: boolean } {
  const nodes: FormulaNode[] = [];
  let scannedFormulaCells = 0;
  let truncated = false;

  for (let si = 0; si < workbook.getSheetCount(); si++) {
    throwIfAborted(signal);
    const sheet = workbook.getSheet(si);
    const sheetName = sheet.name();

    let usedRange: RawUsedRange | null = null;

    try {
      usedRange = sheet.getUsedRange(USED_RANGE_TYPE_DATA_FORMULA);
    } catch {
      continue;
    }

    if (!usedRange || usedRange.rowCount <= 0 || usedRange.colCount <= 0) {
      continue;
    }

    const rowEnd = usedRange.row + usedRange.rowCount;
    const colEnd = usedRange.col + usedRange.colCount;

    for (let row = usedRange.row; row < rowEnd; row++) {
      throwIfAborted(signal);
      for (let col = usedRange.col; col < colEnd; col++) {
        const formula = sheet.getFormula(row, col);
        if (!formula) continue;

        scannedFormulaCells++;
        if (scannedFormulaCells > maxFormulaCells) {
          truncated = true;
          return { nodes, scannedFormulaCells, truncated };
        }

        const precedents = normalizePrecedents(
          sheet.getPrecedents(row, col),
          sheetName,
        );

        nodes.push({
          sheet: sheetName,
          row,
          col,
          precedents,
        });
      }
    }
  }

  return { nodes, scannedFormulaCells, truncated };
}

function matchesAnyFrontierCell(
  precedents: PrecedentRange[],
  frontierBySheet: Map<string, TraceCell[]>,
  frontierExact: Set<string>,
): boolean {
  for (const precedent of precedents) {
    const frontierCells = frontierBySheet.get(precedent.sheet);
    if (!frontierCells || frontierCells.length === 0) continue;

    if (isSingleCell(precedent)) {
      if (
        frontierExact.has(
          makeCellKey(precedent.sheet, precedent.row, precedent.col),
        )
      ) {
        return true;
      }
      continue;
    }

    for (const cell of frontierCells) {
      if (rangeContainsCell(precedent, cell)) {
        return true;
      }
    }
  }

  return false;
}

function groupBySheet(cells: TraceCell[]): Map<string, TraceCell[]> {
  const bySheet = new Map<string, TraceCell[]>();
  for (const cell of cells) {
    const list = bySheet.get(cell.sheet);
    if (list) {
      list.push(cell);
    } else {
      bySheet.set(cell.sheet, [cell]);
    }
  }
  return bySheet;
}

export async function refs(
  filePath: string,
  ref: string,
  options?: {
    signal?: AbortSignal | null;
    depth?: number;
    maxFormulaCells?: number;
  },
): Promise<void> {
  const signal = options?.signal;
  const depth = options?.depth && options.depth > 0 ? options.depth : 1;
  const maxFormulaCells =
    options?.maxFormulaCells && options.maxFormulaCells > 0
      ? options.maxFormulaCells
      : DEFAULT_MAX_FORMULA_SCAN;

  await withFile(
    filePath,
    ({ workbook }) => {
      const target = parseTarget(ref);
      const defaultSheet = workbook.getActiveSheet().name();
      const startSheetName = target.sheet || defaultSheet;
      const startSheet = workbook.getSheetFromName(startSheetName);

      if (!startSheet) {
        fail(`Sheet not found: ${startSheetName}`);
      }

      const { nodes, scannedFormulaCells, truncated } = collectFormulaNodes(
        workbook,
        maxFormulaCells,
        signal,
      );

      let frontier: TraceCell[] = [
        { sheet: startSheetName, row: target.row, col: target.col },
      ];
      let frontierExact = new Set<string>([
        makeCellKey(startSheetName, target.row, target.col),
      ]);

      const seenDependents = new Set<string>();
      const discovered = new Map<
        string,
        { sheet: string; row: number; col: number; hop: number }
      >();

      let hopsTraversed = 0;

      for (let hop = 1; hop <= depth && frontier.length > 0; hop++) {
        throwIfAborted(signal);
        hopsTraversed = hop;

        const frontierBySheet = groupBySheet(frontier);
        const next: TraceCell[] = [];
        const nextExact = new Set<string>();

        for (const node of nodes) {
          throwIfAborted(signal);
          const nodeKey = makeCellKey(node.sheet, node.row, node.col);
          if (seenDependents.has(nodeKey)) continue;

          if (
            !matchesAnyFrontierCell(
              node.precedents,
              frontierBySheet,
              frontierExact,
            )
          ) {
            continue;
          }

          if (nodeKey === makeCellKey(startSheetName, target.row, target.col)) {
            continue;
          }

          seenDependents.add(nodeKey);
          discovered.set(nodeKey, {
            sheet: node.sheet,
            row: node.row,
            col: node.col,
            hop,
          });

          next.push({ sheet: node.sheet, row: node.row, col: node.col });
          nextExact.add(nodeKey);
        }

        frontier = next;
        frontierExact = nextExact;
      }

      const references = Array.from(discovered.values())
        .map((item) => ({
          sheet: item.sheet,
          cell: cellToA1(item.row, item.col),
          hop: item.hop,
        }))
        .sort((a, b) => {
          if (a.hop !== b.hop) return a.hop - b.hop;
          if (a.sheet !== b.sheet) return a.sheet.localeCompare(b.sheet);
          return a.cell.localeCompare(b.cell);
        });

      ok({
        target: {
          sheet: startSheetName,
          cell: cellToA1(target.row, target.col),
        },
        requestedDepth: depth,
        hopsTraversed,
        oneHop: depth === 1,
        references,
        stats: {
          scannedFormulaCells,
          indexedFormulaCells: nodes.length,
          maxFormulaCells,
          truncated,
        },
        note: "Best-effort trace. Dynamic formulas (e.g. INDIRECT/OFFSET) may be incomplete.",
      });
    },
    { signal },
  );
}
