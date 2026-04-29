import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { throwIfAborted } from "./abort.js";

export const SJS_VERSION = "18.2.5";
export const SJS_CACHE_ENV = "HEADLESS_SPREADJS_SJS_DIR";
export const SJS_ASSET_PREFIX = "/_sjs/";

export type ServedAsset = { buffer: Buffer; contentType: string };

interface BrowserAsset {
  filename: string;
  url: string;
  sri: string;
  contentType: string;
}

const SJS_ASSETS = {
  css: {
    filename: "gc.spread.sheets.excel2013white.min.css",
    url: `https://cdn.jsdelivr.net/npm/@grapecity/spread-sheets@${SJS_VERSION}/styles/gc.spread.sheets.excel2013white.min.css`,
    sri: "sha384-8xszHlJ5bAtwSy2RsLvLQWrw5LZaEO3LsN6ltnFrvR1vutHpk2hVMw3hpTi0J5ve",
    contentType: "text/css; charset=utf-8",
  },
  sheets: {
    filename: "gc.spread.sheets.all.min.js",
    url: `https://cdn.jsdelivr.net/npm/@grapecity/spread-sheets@${SJS_VERSION}/dist/gc.spread.sheets.all.min.js`,
    sri: "sha384-/dzRVKyy13/kDppFpstvr423l6zSyRU1iqa1xk50JvBPhRNHVBcz/bZj6DlvjW+b",
    contentType: "application/javascript; charset=utf-8",
  },
  shapes: {
    filename: "gc.spread.sheets.shapes.min.js",
    url: `https://cdn.jsdelivr.net/npm/@grapecity/spread-sheets-shapes@${SJS_VERSION}/dist/gc.spread.sheets.shapes.min.js`,
    sri: "sha384-/3K6eZ4RcDKWKtKUPpayueXZ6RJGq1ot8+14i5ykB7bj+iIez8pdJ5XitfJ6fWVq",
    contentType: "application/javascript; charset=utf-8",
  },
  charts: {
    filename: "gc.spread.sheets.charts.min.js",
    url: `https://cdn.jsdelivr.net/npm/@grapecity/spread-sheets-charts@${SJS_VERSION}/dist/gc.spread.sheets.charts.min.js`,
    sri: "sha384-hS1RVsLVUx9BPZoFqgSqPcephSC1mXPiyJFQgn/ds6yvOtODN5LaQB/kVdRIiT07",
    contentType: "application/javascript; charset=utf-8",
  },
  pivot: {
    filename: "gc.spread.pivot.pivottables.min.js",
    url: `https://cdn.jsdelivr.net/npm/@grapecity/spread-sheets-pivot-addon@${SJS_VERSION}/dist/gc.spread.pivot.pivottables.min.js`,
    sri: "sha384-HzxnQ4cIZiOuXFElx8nPIHVepJZ/aTbWak0YBrgbc3gTCHEgtlO/o2+jsv+j7oW8",
    contentType: "application/javascript; charset=utf-8",
  },
  slicers: {
    filename: "gc.spread.sheets.slicers.min.js",
    url: `https://cdn.jsdelivr.net/npm/@grapecity/spread-sheets-slicers@${SJS_VERSION}/dist/gc.spread.sheets.slicers.min.js`,
    sri: "sha384-J9dpcXulhzecs/8iRsmCTAtWG61wAsgWU4RrdeFkMaTHIKlXMIFYEBzfyIB3PFMA",
    contentType: "application/javascript; charset=utf-8",
  },
  excelio: {
    filename: "gc.spread.excelio.min.js",
    url: `https://cdn.jsdelivr.net/npm/@grapecity/spread-excelio@${SJS_VERSION}/dist/gc.spread.excelio.min.js`,
    sri: "sha384-8YVQAx4YGZwqNtl2kAQUqg+g18pCTVtyuYPZwvBq0GzjVbAs37+hkqByGB7h1wqo",
    contentType: "application/javascript; charset=utf-8",
  },
} as const satisfies Record<string, BrowserAsset>;

export type BrowserAssetKey = keyof typeof SJS_ASSETS;
export type BrowserScriptKey = Exclude<BrowserAssetKey, "css">;

export function browserScriptKeys(loadPivot: boolean): BrowserScriptKey[] {
  const keys: BrowserScriptKey[] = ["sheets", "shapes", "charts"];
  if (loadPivot) keys.push("pivot");
  keys.push("slicers", "excelio");
  return keys;
}

export function assetPath(key: BrowserAssetKey): string {
  return `${SJS_ASSET_PREFIX}${SJS_ASSETS[key].filename}`;
}

export function assetSri(key: BrowserAssetKey): string {
  return SJS_ASSETS[key].sri;
}

function sriOf(buf: Buffer): string {
  return `sha384-${crypto.createHash("sha384").update(buf).digest("base64")}`;
}

function sjsCacheDir(): string {
  const override = process.env[SJS_CACHE_ENV];
  if (override) return path.resolve(override);
  return path.join(os.homedir(), ".headless-spreadjs", `sjs-${SJS_VERSION}`);
}

async function readVerifiedAsset(
  file: string,
  asset: BrowserAsset,
): Promise<Buffer | null> {
  let buf: Buffer;
  try {
    buf = await fs.readFile(file);
  } catch {
    return null;
  }

  return sriOf(buf) === asset.sri ? buf : null;
}

async function writeAssetAtomically(file: string, buf: Buffer): Promise<void> {
  await fs.mkdir(path.dirname(file), { recursive: true });
  const tmp = path.join(
    path.dirname(file),
    `${path.basename(file)}.tmp.${process.pid}.${crypto.randomUUID()}`,
  );

  try {
    await fs.writeFile(tmp, buf);
    await fs.rename(tmp, file);
  } catch (err) {
    await fs.unlink(tmp).catch(() => undefined);
    throw err;
  }
}

async function fetchAsset(
  asset: BrowserAsset,
  signal: AbortSignal | null | undefined,
): Promise<Buffer> {
  let res: Response;
  try {
    res = await fetch(asset.url, { signal: signal ?? undefined });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    throw new Error(
      `failed to fetch SpreadJS asset ${asset.filename} from ${asset.url}: ${message}. ` +
        `For offline use, set ${SJS_CACHE_ENV} to a directory containing the pinned files.`,
    );
  }

  if (!res.ok) {
    throw new Error(
      `failed to fetch SpreadJS asset ${asset.filename} from ${asset.url}: HTTP ${res.status} ${res.statusText}`,
    );
  }

  const buf = Buffer.from(await res.arrayBuffer());
  const got = sriOf(buf);
  if (got !== asset.sri) {
    throw new Error(
      `integrity check failed for ${asset.filename}: expected ${asset.sri}, got ${got}`,
    );
  }

  return buf;
}

async function loadAsset(
  key: BrowserAssetKey,
  cacheRoot: string,
  signal: AbortSignal | null | undefined,
): Promise<ServedAsset> {
  const asset = SJS_ASSETS[key];
  const file = path.join(cacheRoot, asset.filename);
  const cached = await readVerifiedAsset(file, asset);
  if (cached) return { buffer: cached, contentType: asset.contentType };

  throwIfAborted(signal);
  const fetched = await fetchAsset(asset, signal);
  await writeAssetAtomically(file, fetched);
  return { buffer: fetched, contentType: asset.contentType };
}

export async function loadBrowserAssets(
  keys: BrowserAssetKey[],
  signal: AbortSignal | null | undefined,
): Promise<Map<string, ServedAsset>> {
  const cacheRoot = sjsCacheDir();
  const entries = await Promise.all(
    keys.map(async (key) => {
      const asset = SJS_ASSETS[key];
      return [asset.filename, await loadAsset(key, cacheRoot, signal)] as const;
    }),
  );

  return new Map(entries);
}
