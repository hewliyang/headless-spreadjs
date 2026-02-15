import { colToIndex } from "../a1.js";
import { withFile } from "../context.js";
import { fail, ok } from "../output.js";

type Op = "insert" | "delete" | "hide" | "unhide" | "freeze" | "unfreeze";
type Dim = "rows" | "columns";

export async function rowsCols(
  filePath: string,
  op: Op,
  dim: Dim,
  options: {
    sheet?: string;
    ref?: string;
    count?: number;
  },
): Promise<void> {
  await withFile(
    filePath,
    ({ workbook }) => {
      const sheet = options.sheet
        ? workbook.getSheetFromName(options.sheet)
        : workbook.getActiveSheet();

      if (!sheet) fail(`Sheet not found: ${options.sheet ?? "(active)"}`);

      if (op === "unfreeze") {
        sheet.frozenRowCount(0);
        sheet.frozenColumnCount(0);
        ok({ operation: "unfreeze" });
        return;
      }

      if (op === "freeze") {
        if (!options.ref)
          fail(
            "freeze requires --ref (e.g. --ref 2 for rows or --ref C for columns)",
          );
        if (dim === "rows") {
          sheet.frozenRowCount(parseInt(options.ref, 10));
        } else {
          sheet.frozenColumnCount(colToIndex(options.ref) + 1);
        }
        ok({ operation: "freeze", dimension: dim, ref: options.ref });
        return;
      }

      if (!options.ref) fail(`${op} requires --ref`);
      const count = options.count ?? 1;
      const isRow = dim === "rows";
      const index = isRow
        ? parseInt(options.ref, 10) - 1
        : colToIndex(options.ref);

      switch (op) {
        case "insert":
          if (isRow) {
            sheet.addRows(index, count);
          } else {
            sheet.addColumns(index, count);
          }
          break;

        case "delete":
          if (isRow) {
            sheet.deleteRows(index, count);
          } else {
            sheet.deleteColumns(index, count);
          }
          break;

        case "hide":
          for (let i = 0; i < count; i++) {
            if (isRow) {
              sheet.setRowVisible(index + i, false);
            } else {
              sheet.setColumnVisible(index + i, false);
            }
          }
          break;

        case "unhide":
          for (let i = 0; i < count; i++) {
            if (isRow) {
              sheet.setRowVisible(index + i, true);
            } else {
              sheet.setColumnVisible(index + i, true);
            }
          }
          break;
      }

      ok({ operation: op, dimension: dim, ref: options.ref, count });
    },
    { save: true },
  );
}
