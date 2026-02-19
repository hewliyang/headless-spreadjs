import { withNewFile } from "../context.js";
import { ok } from "../output.js";

export async function create(
  filePath: string,
  options?: { signal?: AbortSignal | null },
): Promise<void> {
  await withNewFile(filePath, undefined, { signal: options?.signal });
  ok({ created: filePath });
}
