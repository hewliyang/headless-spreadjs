import { ExcelFile, setGC } from "./excel-file.js";
import { disposeShims, installShims } from "./shims.js";
import type { GCNamespace } from "./types.js";

export { ExcelFile } from "./excel-file.js";
export type {
  GCNamespace,
  SpreadStyle,
  SpreadWorkbook,
  SpreadWorksheet,
} from "./types.js";

export interface InitOptions {
  licenseKey?: string;
}

export interface InitResult {
  GC: GCNamespace;
  ExcelFile: typeof ExcelFile;
  dispose: () => void;
}

const optionalAddons = [
  "@mescius/spread-sheets-charts",
  "@mescius/spread-sheets-pivot-addon",
  "@mescius/spread-sheets-shapes",
  "@mescius/spread-sheets-slicers",
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
    return { GC: cachedGC, ExcelFile, dispose };
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

  return { GC: gc, ExcelFile, dispose };
}

export function dispose(): void {
  disposeShims();
  cachedGC = null;
  initialized = false;
  setGC(null);
}
