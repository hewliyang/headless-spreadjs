/**
 * Structured JSON output helpers for CLI commands.
 */

export function ok(data: unknown): void {
  process.stdout.write(JSON.stringify(data) + "\n");
}

export function fail(message: string): never {
  throw new Error(message);
}

/**
 * Read input from last positional arg or stdin.
 */
export async function readInput(
  argValue: string | undefined,
): Promise<string> {
  if (argValue && argValue !== "-") {
    return argValue;
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
