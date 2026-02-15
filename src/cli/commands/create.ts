import { withNewFile } from "../context.js";
import { ok } from "../output.js";

export async function create(filePath: string): Promise<void> {
  await withNewFile(filePath);
  ok({ created: filePath });
}
