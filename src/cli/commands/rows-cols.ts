import { colToIndex } from "../a1.js";
import { throwIfAborted } from "../abort.js";
import { withFile } from "../context.js";
import { fail, ok } from "../output.js";

function splitQualifiedRef(ref: string): { sheet?: string; localRef: string } {
  const quotedMatch = ref.match(/^'([^']+)'!(.+)$/);
  if (quotedMatch) {
    return { sheet: quotedMatch[1], localRef: quotedMatch[2] };
  }

  const bangIndex = ref.indexOf("!");
  if (bangIndex === -1) {
    return { localRef: ref };
  }

  return {
    sheet: ref.slice(0, bangIndex),
    localRef: ref.slice(bangIndex + 1),
  };
}

function parseRowRef(ref: string): number {
  if (!/^\d+$/.test(ref)) {
    fail(`Invalid row reference: ${ref}`);
  }
  const parsed = Number.parseInt(ref, 10);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    fail(`Invalid row reference: ${ref}`);
  }
  return parsed;
}

function parseColumnRef(ref: string): number {
  if (!/^[A-Za-z]+$/.test(ref)) {
    fail(`Invalid column reference: ${ref}`);
  }
  const index = colToIndex(ref);
  if (!Number.isFinite(index) || index < 0) {
    fail(`Invalid column reference: ${ref}`);
  }
  return index;
}

export type RcOp =
  | "insert"
  | "delete"
  | "hide"
  | "unhide"
  | "freeze"
  | "unfreeze";
export type RcDim = "rows" | "columns";

export async function rowsCols(
  filePath: string,
  op: RcOp,
  dim: RcDim,
  options: {
    sheet?: string;
    ref?: string;
    count?: number;
    signal?: AbortSignal | null;
  },
): Promise<void> {
  const signal = options.signal;

  await withFile(
    filePath,
    ({ file, workbook }) => {
      const qualified = options.ref ? splitQualifiedRef(options.ref) : null;
      if (
        options.sheet &&
        qualified?.sheet &&
        options.sheet !== qualified.sheet
      ) {
        fail(
          `Sheet mismatch: --sheet ${options.sheet} does not match ref sheet ${qualified.sheet}`,
        );
      }

      const effectiveSheetName = options.sheet ?? qualified?.sheet;
      const sheet = effectiveSheetName
        ? workbook.getSheetFromName(effectiveSheetName)
        : workbook.getActiveSheet();

      if (!sheet) fail(`Sheet not found: ${effectiveSheetName ?? "(active)"}`);

      if (op === "unfreeze") {
        file.batch(() => {
          sheet.frozenRowCount(0);
          sheet.frozenColumnCount(0);
        });
        ok({ operation: "unfreeze" });
        return;
      }

      if (op === "freeze") {
        if (!options.ref)
          fail(
            "freeze requires --ref (e.g. --ref 2 for rows or --ref C for columns)",
          );
        const ref = qualified?.localRef ?? options.ref;
        file.batch(() => {
          if (dim === "rows") {
            sheet.frozenRowCount(parseRowRef(ref));
          } else {
            sheet.frozenColumnCount(parseColumnRef(ref) + 1);
          }
        });
        ok({ operation: "freeze", dimension: dim, ref: options.ref });
        return;
      }

      if (!options.ref) fail(`${op} requires --ref`);
      const count = options.count ?? 1;
      if (!Number.isInteger(count) || count <= 0) {
        fail(`Invalid --count value: ${options.count}`);
      }
      const localRef = qualified?.localRef ?? options.ref;
      const isRow = dim === "rows";
      const index = isRow
        ? parseRowRef(localRef) - 1
        : parseColumnRef(localRef);

      file.batch(() => {
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
              throwIfAborted(signal);
              if (isRow) {
                sheet.setRowVisible(index + i, false);
              } else {
                sheet.setColumnVisible(index + i, false);
              }
            }
            break;

          case "unhide":
            for (let i = 0; i < count; i++) {
              throwIfAborted(signal);
              if (isRow) {
                sheet.setRowVisible(index + i, true);
              } else {
                sheet.setColumnVisible(index + i, true);
              }
            }
            break;
        }
      });

      ok({ operation: op, dimension: dim, ref: options.ref, count });
    },
    { save: true, signal },
  );
}
