import { cellToA1, parseRef, rangeDimensions } from "../a1.js";
import { withFile } from "../context.js";
import { fail, ok } from "../output.js";
import { type SerializedStyle, serializeStyle } from "../styles.js";

export async function get(
  filePath: string,
  ref: string,
  options: { styles?: boolean },
): Promise<void> {
  const parsed = parseRef(ref);

  await withFile(filePath, ({ workbook }) => {
    const sheet = parsed.sheet
      ? workbook.getSheetFromName(parsed.sheet)
      : workbook.getActiveSheet();

    if (!sheet) {
      fail(`Sheet not found: ${parsed.sheet ?? "(active)"}`);
    }

    const { rows, cols } = rangeDimensions(parsed);
    const cells: Record<
      string,
      {
        value: unknown;
        formula?: string;
        styles?: SerializedStyle;
      }
    > = {};
    let cellCount = 0;

    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const row = parsed.start.row + r;
        const col = parsed.start.col + c;
        const value = sheet.getValue(row, col);
        const formula = sheet.getFormula(row, col);
        const hasValue = value !== null && value !== undefined && value !== "";
        const hasFormula = !!formula;

        if (!hasValue && !hasFormula) continue;

        const a1 = cellToA1(row, col);
        const cell: (typeof cells)[string] = {
          value: value ?? null,
        };

        if (hasFormula) {
          cell.formula = formula;
        }

        if (options.styles !== false) {
          const style = sheet.getStyle(row, col);
          if (style) {
            const serialized = serializeStyle(style);
            if (serialized) {
              cell.styles = serialized;
            }
          }
        }

        cells[a1] = cell;
        cellCount++;
      }
    }

    ok({
      sheet: sheet.name(),
      range: ref,
      cellCount,
      cells,
    });
  });
}
