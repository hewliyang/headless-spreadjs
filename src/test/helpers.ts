import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { init } from "../index.js";

export type Runtime = Awaited<ReturnType<typeof init>>;

export async function withRuntime<T>(
  fn: (runtime: Runtime) => Promise<T> | T,
): Promise<T> {
  const runtime = await init();
  try {
    return await fn(runtime);
  } finally {
    runtime.dispose();
  }
}

export async function withTempDir<T>(
  fn: (dir: string) => Promise<T>,
): Promise<T> {
  const tempDir = await fs.mkdtemp(
    path.join(os.tmpdir(), "headless-spreadjs-test-"),
  );

  try {
    return await fn(tempDir);
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
}

export function readWithFileReader(
  method: "readAsArrayBuffer" | "readAsDataURL" | "readAsText",
  input: unknown,
): Promise<unknown> {
  const ReaderCtor = (globalThis as Record<string, unknown>)
    .FileReader as unknown as new () => {
    result: unknown;
    error: unknown;
    onload: ((event: { target: { result: unknown } }) => void) | null;
    onerror: (() => void) | null;
    [key: string]: unknown;
  };

  const reader = new ReaderCtor();

  return new Promise((resolve, reject) => {
    reader.onload = (event) => resolve(event.target.result);
    reader.onerror = () => reject(reader.error);
    (reader[method] as (value: unknown) => void)(input);
  });
}
