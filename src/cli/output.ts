/**
 * Structured JSON output helpers for CLI commands.
 *
 * In daemon mode, output is captured instead of written to process streams.
 */

let capturedStdout: string | null = null;
let capturedStderr: string | null = null;

/** Enable output capture (daemon mode). */
export function startCapture(): void {
  capturedStdout = "";
  capturedStderr = "";
}

/** Disable capture and return captured output. */
export function stopCapture(): { stdout: string; stderr: string } {
  const result = {
    stdout: capturedStdout ?? "",
    stderr: capturedStderr ?? "",
  };
  capturedStdout = null;
  capturedStderr = null;
  return result;
}

export function isCapturing(): boolean {
  return capturedStdout !== null;
}

export function writeStdout(data: string): void {
  if (capturedStdout !== null) {
    capturedStdout += data;
  } else {
    process.stdout.write(data);
  }
}

export function writeStderr(data: string): void {
  if (capturedStderr !== null) {
    capturedStderr += data;
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
 * In daemon mode, stdin comes from the request payload.
 */
let pendingStdin: string | null = null;

export function setStdin(data: string | undefined): void {
  pendingStdin = data ?? null;
}

export async function readInput(argValue: string | undefined): Promise<string> {
  if (argValue && argValue !== "-") {
    return argValue;
  }

  // In daemon mode, use provided stdin
  if (pendingStdin !== null) {
    const text = pendingStdin;
    pendingStdin = null;
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
