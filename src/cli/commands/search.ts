import { cellToA1 } from "../a1.js";
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
  },
): Promise<void> {
  await withFile(filePath, ({ workbook }) => {
    const matches: SearchMatch[] = [];
    const max = options.maxResults ?? 500;

    const pattern = options.regex
      ? new RegExp(term, options.matchCase ? "" : "i")
      : null;

    const sheetCount = workbook.getSheetCount();

    for (let si = 0; si < sheetCount && matches.length < max; si++) {
      const ws = workbook.getSheet(si);
      const sheetName = ws.name();

      if (options.sheet && sheetName !== options.sheet) continue;

      let usedRows = 0;
      let usedCols = 0;
      try {
        const range = ws.getUsedRange(16 | 32); // data | formula
        if (range && range.rowCount > 0) {
          usedRows = range.row + range.rowCount;
          usedCols = range.col + range.colCount;
        }
      } catch {
        continue;
      }

      for (let r = 0; r < usedRows && matches.length < max; r++) {
        for (let c = 0; c < usedCols && matches.length < max; c++) {
          const value = ws.getValue(r, c);
          const formula = ws.getFormula(r, c);

          if (value === null || value === undefined) {
            if (!formula) continue;
          }

          const text = String(value ?? "");
          let isMatch = false;

          if (pattern) {
            isMatch = pattern.test(text);
          } else if (options.matchCase) {
            isMatch = text.includes(term);
          } else {
            isMatch = text.toLowerCase().includes(term.toLowerCase());
          }

          if (!isMatch) continue;

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
  });
}
