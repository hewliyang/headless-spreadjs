import { withFile } from "../context.js";
import { fail, ok, readInput } from "../output.js";

export async function evalCode(
  filePath: string,
  codeArg: string | undefined,
): Promise<void> {
  const code = await readInput(codeArg);

  await withFile(
    filePath,
    async ({ file, workbook, GC }) => {
      const sheet = workbook.getActiveSheet();

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
        const fn = new AsyncFunction("workbook", "sheet", "GC", "file", code);
        const result = await fn(workbook, sheet, GC, file);

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
    { save: true },
  );
}
