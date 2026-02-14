import { disposeShims, installShims } from "./shims.js";
import type { GCNamespace } from "./types.js";
import { setGC, Workbook } from "./workbook.js";

export type {
  GCNamespace,
  SpreadStyle,
  SpreadWorkbook,
  SpreadWorksheet,
} from "./types.js";
export { Workbook } from "./workbook.js";

export interface InitOptions {
  licenseKey?: string;
}

export interface InitResult {
  GC: GCNamespace;
  Workbook: typeof Workbook;
  dispose: () => void;
}

const optionalAddons = [
  "@mescius/spread-sheets-charts",
  "@mescius/spread-sheets-pivot-addon",
] as const;

let initialized = false;
let cachedGC: GCNamespace | null = null;

async function importOptional(moduleName: (typeof optionalAddons)[number]) {
  try {
    await import(moduleName);
  } catch {
    // Optional addon is not installed.
  }
}

export async function init(options?: InitOptions): Promise<InitResult> {
  if (initialized && cachedGC) {
    return { GC: cachedGC, Workbook, dispose };
  }

  installShims();

  const spreadModule = await import("@mescius/spread-sheets");
  const gc = (spreadModule.default ?? spreadModule) as GCNamespace;

  await import("@mescius/spread-sheets-io");
  await Promise.all(optionalAddons.map(importOptional));

  if (options?.licenseKey) {
    gc.Spread.Sheets.LicenseKey = options.licenseKey;
  }

  setGC(gc);
  cachedGC = gc;
  initialized = true;

  return { GC: gc, Workbook, dispose };
}

export function dispose(): void {
  disposeShims();
  cachedGC = null;
  initialized = false;
  setGC(null);
}
