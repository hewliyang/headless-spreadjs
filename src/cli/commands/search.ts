import { cellToA1 } from "../a1.js";
import { throwIfAborted } from "../abort.js";
import { withFile } from "../context.js";
import { ok } from "../output.js";

interface SearchMatch {
  sheet: string;
  cell: string;
  value: unknown;
  formula: string | null;
}

export async function search(
  filePath: string,
  term: string,
  options: {
    sheet?: string;
    matchCase?: boolean;
    regex?: boolean;
    maxResults?: number;
    signal?: AbortSignal | null;
  },
): Promise<void> {
  const signal = options.signal;

  await withFile(
    filePath,
    ({ workbook }) => {
      const matches: SearchMatch[] = [];
      const max = options.maxResults ?? 500;
      const lowerTerm = options.matchCase ? term : term.toLowerCase();

      const pattern = options.regex
        ? new RegExp(term, options.matchCase ? "" : "i")
        : null;

      const sheetCount = workbook.getSheetCount();

      for (let si = 0; si < sheetCount && matches.length < max; si++) {
        throwIfAborted(signal);
        const ws = workbook.getSheet(si);
        const sheetName = ws.name();

        if (options.sheet && sheetName !== options.sheet) continue;

        let startRow = 0;
        let endRow = 0;
        let startCol = 0;
        let endCol = 0;
        try {
          const range = ws.getUsedRange(16 | 32); // data | formula
          if (range && range.rowCount > 0 && range.colCount > 0) {
            startRow = range.row;
            endRow = range.row + range.rowCount;
            startCol = range.col;
            endCol = range.col + range.colCount;
          }
        } catch {
          continue;
        }

        for (let r = startRow; r < endRow && matches.length < max; r++) {
          throwIfAborted(signal);
          for (let c = startCol; c < endCol && matches.length < max; c++) {
            const value = ws.getValue(r, c);
            let formula: string | null = null;

            if (value === null || value === undefined) {
              formula = ws.getFormula(r, c);
              if (!formula) continue;
            }

            const text = String(value ?? "");
            let isMatch = false;

            if (pattern) {
              isMatch = pattern.test(text);
            } else if (options.matchCase) {
              isMatch = text.includes(term);
            } else {
              isMatch = text.toLowerCase().includes(lowerTerm);
            }

            if (!isMatch) continue;

            formula ??= ws.getFormula(r, c);

            matches.push({
              sheet: sheetName,
              cell: cellToA1(r, c),
              value,
              formula: formula || null,
            });
          }
        }
      }

      ok({
        matches,
        totalFound: matches.length,
        hasMore: matches.length >= max,
        searchTerm: term,
      });
    },
    { signal },
  );
}
