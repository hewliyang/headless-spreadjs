import { AsyncLocalStorage } from "node:async_hooks";
import { resolve } from "node:path";
import {
  getCurrentCommand,
  type MutatedRange,
  runOnOpenHooks,
  runPostSaveHooks,
  runPreSaveHooks,
} from "../hooks.js";
import { type ExcelFile, init } from "../index.js";
import type { GCNamespace, SpreadWorkbook } from "../types.js";
import { throwIfAborted } from "./abort.js";
import type { FileCache } from "./file-cache.js";

export interface FileContext {
  file: ExcelFile;
  workbook: SpreadWorkbook;
  GC: GCNamespace;
  markMutated(range: MutatedRange): void;
}

interface DaemonRuntime {
  GC: GCNamespace;
  ExcelFile: typeof ExcelFile;
  fileCache: FileCache;
  cwd: string;
  writeThrough: boolean;
}

interface FileOptions {
  save?: boolean;
  signal?: AbortSignal | null;
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
  options?: FileOptions,
): Promise<T> {
  const signal = options?.signal;
  throwIfAborted(signal);

  const daemonRuntime = runtimeStore.getStore();
  if (daemonRuntime) {
    return withFileDaemon(daemonRuntime, filePath, fn, options);
  }

  const { GC, ExcelFile: EF, dispose } = await init();
  try {
    throwIfAborted(signal);
    const file = await EF.open(filePath);
    throwIfAborted(signal);

    const mutatedRanges: MutatedRange[] = [];
    const hookCtx = {
      ...getCurrentCommand(),
      filePath,
      file,
      workbook: file.workbook,
      GC,
      mutatedRanges,
    };
    await runOnOpenHooks(hookCtx);

    const result = await fn({
      file,
      workbook: file.workbook,
      GC,
      markMutated: (r) => mutatedRanges.push(r),
    });
    if (options?.save) {
      throwIfAborted(signal);
      await runPreSaveHooks(hookCtx);
      await file.save(filePath);
      await runPostSaveHooks(hookCtx);
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
  options?: FileOptions,
): Promise<T> {
  const signal = options?.signal;
  throwIfAborted(signal);

  const absPath = resolve(rt.cwd, filePath);

  let cached = await rt.fileCache.get(filePath, rt.cwd);
  if (!cached) {
    throwIfAborted(signal);
    const file = await rt.ExcelFile.open(absPath);
    throwIfAborted(signal);
    await rt.fileCache.put(filePath, rt.cwd, file);
    cached = { file, absPath };
  }

  const mutatedRanges: MutatedRange[] = [];
  const hookCtx = {
    ...getCurrentCommand(),
    filePath: cached.absPath,
    file: cached.file,
    workbook: cached.file.workbook,
    GC: rt.GC,
    mutatedRanges,
  };
  await runOnOpenHooks(hookCtx);

  const wasDirty = rt.fileCache.isDirty(filePath, rt.cwd);

  let result: T;
  try {
    throwIfAborted(signal);
    result = await fn({
      file: cached.file,
      workbook: cached.file.workbook,
      GC: rt.GC,
      markMutated: (r) => mutatedRanges.push(r),
    });
  } catch (err) {
    if (!wasDirty) {
      rt.fileCache.invalidate(filePath, rt.cwd);
    }
    throw err;
  }

  if (options?.save) {
    throwIfAborted(signal);
    await runPreSaveHooks(hookCtx);
    if (rt.writeThrough) {
      try {
        await cached.file.save(cached.absPath);
        throwIfAborted(signal);
        await rt.fileCache.updateMtime(filePath, rt.cwd);
        await runPostSaveHooks(hookCtx);
      } catch (err) {
        rt.fileCache.invalidate(filePath, rt.cwd);
        throw err;
      }
    } else {
      rt.fileCache.markDirty(filePath, rt.cwd);
      await runPostSaveHooks(hookCtx);
    }
  }

  return result;
}

export async function withNewFile<T>(
  filePath: string,
  fn?: (ctx: FileContext) => T | Promise<T>,
  options?: { signal?: AbortSignal | null },
): Promise<T | undefined> {
  const signal = options?.signal;
  throwIfAborted(signal);

  const daemonRuntime = runtimeStore.getStore();
  if (daemonRuntime) {
    return withNewFileDaemon(daemonRuntime, filePath, fn, options);
  }

  const { GC, ExcelFile: EF, dispose } = await init();
  try {
    throwIfAborted(signal);
    const file = new EF();
    const noop = () => {};
    const result = fn
      ? await fn({ file, workbook: file.workbook, GC, markMutated: noop })
      : undefined;
    throwIfAborted(signal);
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
  options?: { signal?: AbortSignal | null },
): Promise<T | undefined> {
  const signal = options?.signal;
  throwIfAborted(signal);

  const absPath = resolve(rt.cwd, filePath);
  const file = new rt.ExcelFile();
  const noop = () => {};
  const result = fn
    ? await fn({ file, workbook: file.workbook, GC: rt.GC, markMutated: noop })
    : undefined;
  throwIfAborted(signal);
  await file.save(absPath);
  throwIfAborted(signal);
  await rt.fileCache.put(filePath, rt.cwd, file);
  await rt.fileCache.updateMtime(filePath, rt.cwd);
  return result;
}
