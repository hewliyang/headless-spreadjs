/**
 * Structured JSON output helpers for CLI commands.
 *
 * In daemon mode, output is captured via AsyncLocalStorage per-request
 * instead of written to process streams.
 */

import { AsyncLocalStorage } from "node:async_hooks";

export interface IoContext {
  stdout: string;
  stderr: string;
  stdin: string | null;
}

const ioStore = new AsyncLocalStorage<IoContext>();

/**
 * Run `fn` with per-request IO capture.
 * After fn completes (or throws), the IoContext is accessible
 * via the returned/outer `io` object.
 */
export function runWithIo<T>(
  io: IoContext,
  fn: () => T | Promise<T>,
): T | Promise<T> {
  return ioStore.run(io, fn);
}

/** Create a fresh IoContext, optionally pre-loaded with stdin. */
export function createIoContext(stdin?: string): IoContext {
  return { stdout: "", stderr: "", stdin: stdin ?? null };
}

export function writeStdout(data: string): void {
  const io = ioStore.getStore();
  if (io) {
    io.stdout += data;
  } else {
    process.stdout.write(data);
  }
}

export function writeStderr(data: string): void {
  const io = ioStore.getStore();
  if (io) {
    io.stderr += data;
  } else {
    process.stderr.write(data);
  }
}

export function ok(data: unknown): void {
  writeStdout(`${JSON.stringify(data)}\n`);
}

export function fail(message: string): never {
  throw new Error(message);
}

/**
 * Read input from last positional arg or stdin.
 * In daemon mode, stdin comes from the IoContext.
 */
export async function readInput(argValue: string | undefined): Promise<string> {
  if (argValue && argValue !== "-") {
    return argValue;
  }

  // In IO-captured context, use provided stdin
  const io = ioStore.getStore();
  if (io?.stdin != null) {
    const text = io.stdin;
    io.stdin = null; // consume once
    if (!text) {
      fail("No input provided. Pass as argument or pipe via stdin.");
    }
    return text;
  }

  // Read from stdin
  const chunks: Buffer[] = [];
  for await (const chunk of process.stdin) {
    chunks.push(chunk as Buffer);
  }
  const text = Buffer.concat(chunks).toString("utf-8").trim();
  if (!text) {
    fail("No input provided. Pass as argument or pipe via stdin.");
  }
  return text;
}
