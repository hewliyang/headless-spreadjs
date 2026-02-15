/**
 * A1-notation ↔ 0-indexed (row, col) conversion utilities.
 */

export interface CellRef {
  row: number;
  col: number;
}

export interface RangeRef {
  sheet?: string;
  start: CellRef;
  end: CellRef;
}

export function colToIndex(col: string): number {
  let index = 0;
  for (const c of col.toUpperCase()) {
    index = index * 26 + (c.charCodeAt(0) - 64);
  }
  return index - 1;
}

export function indexToCol(index: number): string {
  let col = "";
  let n = index;
  while (n >= 0) {
    col = String.fromCharCode((n % 26) + 65) + col;
    n = Math.floor(n / 26) - 1;
  }
  return col;
}

export function cellToA1(row: number, col: number): string {
  return `${indexToCol(col)}${row + 1}`;
}

export function parseCell(ref: string): CellRef {
  const match = ref.match(/^\$?([A-Z]+)\$?(\d+)$/i);
  if (!match) throw new Error(`Invalid cell reference: ${ref}`);
  return {
    row: parseInt(match[2], 10) - 1,
    col: colToIndex(match[1]),
  };
}

/**
 * Parse a reference like "Sheet1!A1:C10", "'My Sheet'!A1:C10", or "A1:C10".
 */
export function parseRef(ref: string): RangeRef {
  let sheet: string | undefined;
  let rangeStr = ref;

  // Handle quoted sheet names: 'My Sheet'!A1:C10
  const quotedMatch = ref.match(/^'([^']+)'!(.+)$/);
  if (quotedMatch) {
    sheet = quotedMatch[1];
    rangeStr = quotedMatch[2];
  } else {
    // Handle unquoted: Sheet1!A1:C10
    const bangIdx = ref.indexOf("!");
    if (bangIdx !== -1) {
      sheet = ref.slice(0, bangIdx);
      rangeStr = ref.slice(bangIdx + 1);
    }
  }

  const parts = rangeStr.split(":");
  const start = parseCell(parts[0]);
  const end = parts.length > 1 ? parseCell(parts[1]) : { ...start };

  return { sheet, start, end };
}

/**
 * Get the dimensions of a range: { rows, cols }.
 */
export function rangeDimensions(ref: RangeRef): { rows: number; cols: number } {
  return {
    rows: ref.end.row - ref.start.row + 1,
    cols: ref.end.col - ref.start.col + 1,
  };
}
