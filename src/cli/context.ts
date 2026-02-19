import { AsyncLocalStorage } from "node:async_hooks";
import { resolve } from "node:path";
import { type ExcelFile, init } from "../index.js";
import type { GCNamespace, SpreadWorkbook } from "../types.js";
import type { FileCache } from "./file-cache.js";

export interface FileContext {
  file: ExcelFile;
  workbook: SpreadWorkbook;
  GC: GCNamespace;
}

interface DaemonRuntime {
  GC: GCNamespace;
  ExcelFile: typeof ExcelFile;
  fileCache: FileCache;
  cwd: string;
  writeThrough: boolean;
}

const runtimeStore = new AsyncLocalStorage<DaemonRuntime>();

export function runWithDaemonRuntime<T>(
  runtime: DaemonRuntime,
  fn: () => T | Promise<T>,
): T | Promise<T> {
  return runtimeStore.run(runtime, fn);
}

export function getDaemonRuntime(): DaemonRuntime | null {
  return runtimeStore.getStore() ?? null;
}

export async function withFile<T>(
  filePath: string,
  fn: (ctx: FileContext) => T | Promise<T>,
  options?: { save?: boolean },
): Promise<T> {
  const daemonRuntime = runtimeStore.getStore();
  if (daemonRuntime) {
    return withFileDaemon(daemonRuntime, filePath, fn, options);
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
  rt: DaemonRuntime,
  filePath: string,
  fn: (ctx: FileContext) => T | Promise<T>,
  options?: { save?: boolean },
): Promise<T> {
  const absPath = resolve(rt.cwd, filePath);

  let cached = await rt.fileCache.get(filePath, rt.cwd);
  if (!cached) {
    const file = await rt.ExcelFile.open(absPath);
    await rt.fileCache.put(filePath, rt.cwd, file);
    cached = { file, absPath };
  }

  const wasDirty = rt.fileCache.isDirty(filePath, rt.cwd);

  let result: T;
  try {
    result = await fn({
      file: cached.file,
      workbook: cached.file.workbook,
      GC: rt.GC,
    });
  } catch (err) {
    if (!wasDirty) {
      rt.fileCache.invalidate(filePath, rt.cwd);
    }
    throw err;
  }

  if (options?.save) {
    if (rt.writeThrough) {
      try {
        await cached.file.save(cached.absPath);
        await rt.fileCache.updateMtime(filePath, rt.cwd);
      } catch (err) {
        rt.fileCache.invalidate(filePath, rt.cwd);
        throw err;
      }
    } else {
      rt.fileCache.markDirty(filePath, rt.cwd);
    }
  }

  return result;
}

export async function withNewFile<T>(
  filePath: string,
  fn?: (ctx: FileContext) => T | Promise<T>,
): Promise<T | undefined> {
  const daemonRuntime = runtimeStore.getStore();
  if (daemonRuntime) {
    return withNewFileDaemon(daemonRuntime, filePath, fn);
  }

  const { GC, ExcelFile: EF, dispose } = await init();
  try {
    const file = new EF();
    const result = fn
      ? await fn({ file, workbook: file.workbook, GC })
      : undefined;
    await file.save(filePath);
    return result;
  } finally {
    dispose();
  }
}

async function withNewFileDaemon<T>(
  rt: DaemonRuntime,
  filePath: string,
  fn?: (ctx: FileContext) => T | Promise<T>,
): Promise<T | undefined> {
  const absPath = resolve(rt.cwd, filePath);
  const file = new rt.ExcelFile();
  const result = fn
    ? await fn({ file, workbook: file.workbook, GC: rt.GC })
    : undefined;
  await file.save(absPath);
  await rt.fileCache.put(filePath, rt.cwd, file);
  await rt.fileCache.updateMtime(filePath, rt.cwd);
  return result;
}
