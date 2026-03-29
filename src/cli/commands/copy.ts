import { parseRef, rangeDimensions } from "../a1.js";
import { throwIfAborted } from "../abort.js";
import { withFile } from "../context.js";
import { fail, ok } from "../output.js";

export async function copy(
  filePath: string,
  srcRef: string,
  dstRef: string,
  options?: { signal?: AbortSignal | null },
): Promise<void> {
  const src = parseRef(srcRef);
  const dst = parseRef(dstRef);
  const signal = options?.signal;

  await withFile(
    filePath,
    ({ file, workbook, markMutated }) => {
      const srcSheet = src.sheet
        ? workbook.getSheetFromName(src.sheet)
        : workbook.getActiveSheet();
      const dstSheet = dst.sheet
        ? workbook.getSheetFromName(dst.sheet)
        : srcSheet;

      if (!srcSheet) fail(`Source sheet not found: ${src.sheet}`);
      if (!dstSheet) fail(`Destination sheet not found: ${dst.sheet}`);

      const { rows: srcRows, cols: srcCols } = rangeDimensions(src);
      const { rows: dstRows, cols: dstCols } = rangeDimensions(dst);

      file.batch(() => {
        for (let r = 0; r < dstRows; r++) {
          throwIfAborted(signal);
          for (let c = 0; c < dstCols; c++) {
            const sr = src.start.row + (r % srcRows);
            const sc = src.start.col + (c % srcCols);
            const dr = dst.start.row + r;
            const dc = dst.start.col + c;

            const formula = srcSheet.getFormula(sr, sc);
            if (formula) {
              dstSheet.setFormula(dr, dc, formula);
            } else {
              const value = srcSheet.getValue(sr, sc);
              if (value !== null && value !== undefined) {
                dstSheet.setValue(dr, dc, value);
              }
            }

            const style = srcSheet.getStyle(sr, sc);
            if (style) {
              dstSheet.setStyle(dr, dc, style);
            }
          }
        }
      });

      markMutated({
        sheet: dst.sheet ?? dstSheet.name(),
        startRow: dst.start.row,
        startCol: dst.start.col,
        endRow: dst.start.row + dstRows - 1,
        endCol: dst.start.col + dstCols - 1,
      });

      ok({
        source: srcRef,
        destination: dstRef,
        cellsCopied: dstRows * dstCols,
      });
    },
    { save: true, signal },
  );
}
