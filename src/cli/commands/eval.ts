import { parseRef } from "../a1.js";
import { withFile } from "../context.js";
import { fail, ok, readInput } from "../output.js";

export async function evalCode(
  filePath: string,
  codeArg: string | undefined,
  options?: { signal?: AbortSignal | null },
): Promise<void> {
  const code = await readInput(codeArg);

  await withFile(
    filePath,
    async ({ file, workbook, GC }) => {
      const sheet = workbook.getActiveSheet();

      const range = (ref: string) => {
        if (typeof ref !== "string" || ref.trim().length === 0) {
          throw new Error(
            "range(ref) expects a non-empty A1 reference string.",
          );
        }

        const parsed = parseRef(ref.trim());
        const targetSheet = parsed.sheet
          ? workbook.getSheetFromName(parsed.sheet)
          : workbook.getActiveSheet();

        if (!targetSheet) {
          throw new Error(`Sheet not found: ${parsed.sheet ?? "(active)"}`);
        }

        const rowCount = parsed.end.row - parsed.start.row + 1;
        const colCount = parsed.end.col - parsed.start.col + 1;

        if (rowCount < 1 || colCount < 1) {
          throw new Error(`Invalid range reference: ${ref}`);
        }

        return targetSheet.getRange(
          parsed.start.row,
          parsed.start.col,
          rowCount,
          colCount,
        );
      };

      // Capture console.log output
      const logs: string[] = [];
      const origLog = console.log;
      const origWarn = console.warn;
      const origError = console.error;
      console.log = (...args: unknown[]) =>
        logs.push(args.map(String).join(" "));
      console.warn = (...args: unknown[]) =>
        logs.push(`[warn] ${args.map(String).join(" ")}`);
      console.error = (...args: unknown[]) =>
        logs.push(`[error] ${args.map(String).join(" ")}`);

      try {
        const AsyncFunction = Object.getPrototypeOf(async () => {}).constructor;
        const fn = new AsyncFunction(
          "workbook",
          "sheet",
          "GC",
          "file",
          "range",
          code,
        );
        const result = await fn(workbook, sheet, GC, file, range);

        const output: Record<string, unknown> = {};
        if (result !== undefined) {
          output.result = result;
        }
        if (logs.length > 0) {
          output.logs = logs;
        }
        if (Object.keys(output).length === 0) {
          output.result = null;
        }

        ok(output);
      } catch (err) {
        fail(err instanceof Error ? err.message : String(err));
      } finally {
        console.log = origLog;
        console.warn = origWarn;
        console.error = origError;
      }
    },
    { save: true, signal: options?.signal },
  );
}
