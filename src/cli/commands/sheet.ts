import type { SpreadWorkbook } from "../../types.js";
import { throwIfAborted } from "../abort.js";
import { withFile } from "../context.js";
import { fail, ok } from "../output.js";

export type SheetOp = "list" | "create" | "delete" | "rename";

export async function sheet(
  filePath: string,
  op: SheetOp,
  args: string[],
  options?: { signal?: AbortSignal | null },
): Promise<void> {
  const needsSave = op !== "list";
  const signal = options?.signal;

  await withFile(
    filePath,
    ({ file, workbook, GC }) => {
      switch (op) {
        case "list": {
          const sheets: { index: number; name: string }[] = [];
          for (let i = 0; i < workbook.getSheetCount(); i++) {
            throwIfAborted(signal);
            sheets.push({ index: i, name: workbook.getSheet(i).name() });
          }
          ok({ sheets });
          break;
        }

        case "create": {
          const name = args[0];
          if (!name) fail("Usage: hsx sheet <file> create <name>");
          file.batch(() => {
            const ws = new GC.Spread.Sheets.Worksheet(name);
            workbook.addSheet(workbook.getSheetCount(), ws);
          });
          ok({ created: name, index: workbook.getSheetCount() - 1 });
          break;
        }

        case "delete": {
          const name = args[0];
          if (!name) fail("Usage: hsx sheet <file> delete <name>");
          const index = findSheet(workbook, name, signal);
          if (index === -1) fail(`Sheet not found: ${name}`);
          if (workbook.getSheetCount() <= 1) {
            fail("Cannot delete the only sheet.");
          }
          file.batch(() => {
            workbook.removeSheet(index);
          });
          ok({ deleted: name });
          break;
        }

        case "rename": {
          const [oldName, newName] = args;
          if (!oldName || !newName) {
            fail("Usage: hsx sheet <file> rename <old> <new>");
          }
          const index = findSheet(workbook, oldName, signal);
          if (index === -1) fail(`Sheet not found: ${oldName}`);
          file.batch(() => {
            workbook.getSheet(index).name(newName);
          });
          ok({ renamed: { from: oldName, to: newName } });
          break;
        }

        default:
          fail(`Unknown sheet operation: ${op}`);
      }
    },
    { save: needsSave, signal },
  );
}

function findSheet(
  workbook: SpreadWorkbook,
  name: string,
  signal?: AbortSignal | null,
): number {
  for (let i = 0; i < workbook.getSheetCount(); i++) {
    throwIfAborted(signal);
    if (workbook.getSheet(i).name() === name) return i;
  }
  return -1;
}
