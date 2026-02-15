/**
 * Lifecycle wrapper: init → open/create → execute → save → dispose.
 */

import { ExcelFile, init } from "../index.js";
import type { GCNamespace, SpreadWorkbook } from "../types.js";

export interface FileContext {
  file: ExcelFile;
  workbook: SpreadWorkbook;
  GC: GCNamespace;
}

export async function withFile<T>(
  filePath: string,
  fn: (ctx: FileContext) => T | Promise<T>,
  options?: { save?: boolean },
): Promise<T> {
  const { GC, ExcelFile: EF, dispose } = await init();
  try {
    const file = await EF.open(filePath);
    const result = await fn({ file, workbook: file.workbook, GC });
    if (options?.save) {
      await file.save(filePath);
    }
    return result;
  } finally {
    dispose();
  }
}

export async function withNewFile<T>(
  filePath: string,
  fn?: (ctx: FileContext) => T | Promise<T>,
): Promise<T | void> {
  const { GC, ExcelFile: EF, dispose } = await init();
  try {
    const file = new EF();
    let result: T | undefined;
    if (fn) {
      result = await fn({ file, workbook: file.workbook, GC });
    }
    await file.save(filePath);
    return result as T;
  } finally {
    dispose();
  }
}
