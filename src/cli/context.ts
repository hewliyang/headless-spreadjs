/**
 * Lifecycle wrapper: init → open/create → execute → save → dispose.
 *
 * In daemon mode, uses shared runtime and file cache instead of
 * init/dispose per call.
 */

import { resolve } from "node:path";
import { type ExcelFile, init } from "../index.js";
import type { GCNamespace, SpreadWorkbook } from "../types.js";
import type { FileCache } from "./file-cache.js";

export interface FileContext {
  file: ExcelFile;
  workbook: SpreadWorkbook;
  GC: GCNamespace;
}

/** Daemon shared state — set by daemon, null in direct mode. */
interface DaemonRuntime {
  GC: GCNamespace;
  ExcelFile: typeof ExcelFile;
  fileCache: FileCache;
  cwd: string;
}

let daemonRuntime: DaemonRuntime | null = null;

export function setDaemonRuntime(runtime: DaemonRuntime | null): void {
  daemonRuntime = runtime;
}

export function getDaemonRuntime(): DaemonRuntime | null {
  return daemonRuntime;
}

export async function withFile<T>(
  filePath: string,
  fn: (ctx: FileContext) => T | Promise<T>,
  options?: { save?: boolean },
): Promise<T> {
  if (daemonRuntime) {
    return withFileDaemon(filePath, fn, options);
  }

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

async function withFileDaemon<T>(
  filePath: string,
  fn: (ctx: FileContext) => T | Promise<T>,
  options?: { save?: boolean },
): Promise<T> {
  const rt = daemonRuntime!;
  const absPath = resolve(rt.cwd, filePath);

  // Try cache first
  let cached = rt.fileCache.get(filePath, rt.cwd);
  if (!cached) {
    const file = await rt.ExcelFile.open(absPath);
    rt.fileCache.put(filePath, rt.cwd, file);
    cached = { file, absPath };
  }

  const result = await fn({
    file: cached.file,
    workbook: cached.file.workbook,
    GC: rt.GC,
  });

  if (options?.save) {
    await cached.file.save(absPath);
    rt.fileCache.updateMtime(filePath, rt.cwd);
  }

  return result;
}

export async function withNewFile<T>(
  filePath: string,
  fn?: (ctx: FileContext) => T | Promise<T>,
): Promise<T | undefined> {
  if (daemonRuntime) {
    return withNewFileDaemon(filePath, fn);
  }

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

async function withNewFileDaemon<T>(
  filePath: string,
  fn?: (ctx: FileContext) => T | Promise<T>,
): Promise<T | undefined> {
  const rt = daemonRuntime!;
  const absPath = resolve(rt.cwd, filePath);
  const file = new rt.ExcelFile();
  let result: T | undefined;
  if (fn) {
    result = await fn({ file, workbook: file.workbook, GC: rt.GC });
  }
  await file.save(absPath);
  // Cache the newly created file
  rt.fileCache.put(filePath, rt.cwd, file);
  rt.fileCache.updateMtime(filePath, rt.cwd);
  return result as T;
}
