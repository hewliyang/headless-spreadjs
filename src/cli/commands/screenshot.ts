import http from "node:http";
import path from "node:path";
import type { SpreadWorkbook } from "../../types.js";
import { parseRef, rangeDimensions } from "../a1.js";
import { throwIfAborted } from "../abort.js";
import { withFile } from "../context.js";
import { fail, ok } from "../output.js";
import {
  assetPath,
  assetSri,
  type BrowserScriptKey,
  browserScriptKeys,
  loadBrowserAssets,
  type ServedAsset,
  SJS_ASSET_PREFIX,
} from "../screenshot-assets.js";

const DEFAULT_VIEWPORT = { width: 1280, height: 800 };
const RANGE_PADDING_PX = 50;
const RENDER_SETTLE_MS = 300;

interface ScreenshotOptions {
  ref?: string;
  out?: string;
  signal?: AbortSignal | null;
}

function hasPivot(workbook: SpreadWorkbook): boolean {
  for (let i = 0; i < workbook.getSheetCount(); i++) {
    if (workbook.getSheet(i).pivotTables.all().length > 0) return true;
  }
  return false;
}

function scriptTag(key: BrowserScriptKey): string {
  return `<script src="${assetPath(key)}" integrity="${assetSri(key)}" crossorigin="anonymous"></script>`;
}

function buildViewerHTML(scriptKeys: BrowserScriptKey[]): string {
  const scripts = scriptKeys.map(scriptTag).join("\n");

  return `<!DOCTYPE html>
<html><head>
<link rel="stylesheet" href="${assetPath("css")}" integrity="${assetSri("css")}" crossorigin="anonymous">
<style>
  body { margin: 0; padding: 0; background: white; }
  #ss { width: 100%; height: 100vh; }
</style>
<script>
window.__ready = false;
window.__error = null;
function __screenshotErrorText(err) {
  if (!err) return "unknown error";
  if (typeof err === "string") return err;
  if (err.message) return err.message;
  try { return JSON.stringify(err); } catch (_) { return String(err); }
}
function __markScreenshotError(err) {
  window.__error = __screenshotErrorText(err);
  window.__ready = true;
}
window.addEventListener("error", function(event) {
  __markScreenshotError(event.error || event.message);
});
window.addEventListener("unhandledrejection", function(event) {
  __markScreenshotError(event.reason);
});
</script>
</head><body>
<div id="ss"></div>
${scripts}
<script>
(function(){
  try {
    if (!window.GC || !GC.Spread || !GC.Spread.Sheets || !GC.Spread.Excel) {
      throw new Error("SpreadJS browser bundles did not load");
    }

    if (window.CanvasRenderingContext2D) {
      var orig = CanvasRenderingContext2D.prototype.fillText;
      CanvasRenderingContext2D.prototype.fillText = function(text) {
        if (text && typeof text === 'string' && (
          text.includes('GrapeCity') || text.includes('MESCIUS') ||
          text.includes('EVALUATION') || text.includes('Powered') ||
          text.includes('deployment') || text.includes('grapecity.com') ||
          text.includes('mescius.com') || text.includes('Email us') ||
          text === 'Evaluation Version'
        )) return;
        return orig.apply(this, arguments);
      };
    }

    var spread = new GC.Spread.Sheets.Workbook(document.getElementById('ss'));

    fetch('/file.xlsx?t=' + Date.now())
      .then(function(r) {
        if (!r.ok) throw new Error('failed to load workbook: HTTP ' + r.status);
        return r.blob();
      })
      .then(function(blob) {
        var io = new GC.Spread.Excel.IO();
        io.open(blob, function(json) {
          try {
            if (json.sheets) {
              Object.values(json.sheets).forEach(function(s) {
                s.rowCount = Math.max(s.rowCount || 0, 500);
                s.columnCount = Math.max(s.columnCount || 0, 100);
              });
            }
            spread.fromJSON(json);
            spread.refresh();
            window.__spread = spread;
            window.__ready = true;
          } catch (err) {
            __markScreenshotError(err);
          }
        }, __markScreenshotError);
      })
      .catch(__markScreenshotError);
  } catch (err) {
    __markScreenshotError(err);
  }
})();
</script>
</body></html>`;
}

function requestPath(req: http.IncomingMessage): string {
  try {
    return new URL(req.url ?? "/", "http://127.0.0.1").pathname;
  } catch {
    return "/";
  }
}

function assetNameFromPath(pathname: string): string | null {
  if (!pathname.startsWith(SJS_ASSET_PREFIX)) return null;

  let name: string;
  try {
    name = decodeURIComponent(pathname.slice(SJS_ASSET_PREFIX.length));
  } catch {
    return null;
  }

  return name && !name.includes("/") && !name.includes("\\") ? name : null;
}

function startServer(
  html: string,
  xlsxBuf: Buffer,
  assets: Map<string, ServedAsset>,
): Promise<{ server: http.Server; port: number }> {
  return new Promise((resolve, reject) => {
    const server = http.createServer((req, res) => {
      const pathname = requestPath(req);
      const assetName = assetNameFromPath(pathname);

      if (pathname === "/" || pathname === "/index.html") {
        res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
        res.end(html);
        return;
      }

      if (pathname === "/file.xlsx") {
        res.writeHead(200, {
          "Content-Type": "application/octet-stream",
          "Cache-Control": "no-cache",
        });
        res.end(xlsxBuf);
        return;
      }

      if (assetName) {
        const asset = assets.get(assetName);
        if (asset) {
          res.writeHead(200, {
            "Content-Type": asset.contentType,
            "Cache-Control": "no-cache",
            "Access-Control-Allow-Origin": "*",
          });
          res.end(asset.buffer);
          return;
        }
      }

      res.writeHead(404);
      res.end();
    });

    const onError = (err: Error) => reject(err);
    server.once("error", onError);
    server.listen(0, "127.0.0.1", () => {
      server.off("error", onError);
      const addr = server.address();
      if (!addr || typeof addr === "string") {
        reject(new Error("bind failed"));
        return;
      }
      resolve({ server, port: addr.port });
    });
  });
}

async function closeServer(server: http.Server): Promise<void> {
  if (!server.listening) return;
  await new Promise<void>((resolve) => server.close(() => resolve()));
}

const CALC_RANGE_SIZE_SCRIPT = `(opts) => {
  const spread = window.__spread;
  if (!spread) return null;
  if (opts.sheetName) {
    const idx = spread.getSheetIndex(opts.sheetName);
    if (idx === -1) return null;
    spread.setActiveSheetIndex(idx);
  }
  const sheet = spread.getActiveSheet();
  let w = 0, h = 0;
  for (let c = opts.startCol; c < opts.startCol + opts.cols; c++) {
    w += sheet.getColumnWidth(c);
  }
  for (let r = opts.startRow; r < opts.startRow + opts.rows; r++) {
    h += sheet.getRowHeight(r);
  }
  const rhw = sheet.options.rowHeaderVisible !== false
    ? sheet.getColumnWidth(0, GC.Spread.Sheets.SheetArea.rowHeader) : 0;
  const chh = sheet.options.colHeaderVisible !== false
    ? sheet.getRowHeight(0, GC.Spread.Sheets.SheetArea.colHeader) : 0;
  return { width: Math.ceil(w + rhw), height: Math.ceil(h + chh) };
}`;

const CLIP_RANGE_SCRIPT = `(opts) => {
  const spread = window.__spread;
  if (!spread) return null;

  if (opts.sheetName) {
    const idx = spread.getSheetIndex(opts.sheetName);
    if (idx === -1) return { error: "Sheet not found: " + opts.sheetName };
    spread.setActiveSheetIndex(idx);
  }

  const sheet = spread.getActiveSheet();
  sheet.showRow(opts.startRow, GC.Spread.Sheets.VerticalPosition.top);
  sheet.showColumn(opts.startCol, GC.Spread.Sheets.HorizontalPosition.left);
  spread.refresh();

  const hostRect = document.getElementById("ss").getBoundingClientRect();
  const first = sheet.getCellRect(opts.startRow, opts.startCol);
  const last  = sheet.getCellRect(
    opts.startRow + opts.rows - 1,
    opts.startCol + opts.cols - 1,
  );

  if (!first || first.x == null || !last || last.x == null) return null;

  const rhw = sheet.options.rowHeaderVisible !== false
    ? sheet.getColumnWidth(0, GC.Spread.Sheets.SheetArea.rowHeader) : 0;
  const chh = sheet.options.colHeaderVisible !== false
    ? sheet.getRowHeight(0, GC.Spread.Sheets.SheetArea.colHeader) : 0;

  const x = first.x + hostRect.left - rhw;
  const y = first.y + hostRect.top - chh;
  const right  = last.x + last.width  + hostRect.left;
  const bottom = last.y + last.height + hostRect.top;

  return {
    x:      Math.max(0, Math.floor(x)),
    y:      Math.max(0, Math.floor(y)),
    width:  Math.ceil(right - x),
    height: Math.ceil(bottom - y),
  };
}`;

export async function screenshot(
  filePath: string,
  options: ScreenshotOptions,
): Promise<void> {
  const signal = options.signal;
  const outPath = options.out
    ? path.resolve(options.out)
    : path.resolve(`${path.basename(filePath, path.extname(filePath))}.png`);

  const { xlsxBuf, loadPivot } = await withFile(
    filePath,
    async ({ file }) => {
      throwIfAborted(signal);
      return {
        xlsxBuf: await file.saveToBuffer(),
        loadPivot: hasPivot(file.workbook),
      };
    },
    { signal },
  );

  throwIfAborted(signal);

  let chromium: typeof import("playwright")["chromium"];
  try {
    const pw = await import("playwright");
    chromium = pw.chromium;
  } catch {
    fail(
      "playwright is required for screenshots. Install it:\n  npm i -D playwright && npx playwright install chromium",
    );
  }

  const rangeRef = options.ref ? parseRef(options.ref) : null;
  const scriptKeys = browserScriptKeys(loadPivot);
  const assetMap = await loadBrowserAssets(["css", ...scriptKeys], signal);
  throwIfAborted(signal);

  const html = buildViewerHTML(scriptKeys);
  const { server, port } = await startServer(html, xlsxBuf, assetMap);

  try {
    throwIfAborted(signal);
    const browser = await chromium.launch();
    try {
      const page = await browser.newPage({ viewport: DEFAULT_VIEWPORT });

      await page.goto(`http://127.0.0.1:${port}`, {
        waitUntil: "domcontentloaded",
      });
      await page.waitForFunction("window.__ready", { timeout: 30_000 });

      const loadError = await page.evaluate("window.__error");
      if (loadError) fail(`SpreadJS IO error: ${loadError}`);

      if (rangeRef) {
        const { rows, cols } = rangeDimensions(rangeRef);
        const rangeOptions = {
          sheetName: rangeRef.sheet,
          startRow: rangeRef.start.row,
          startCol: rangeRef.start.col,
          rows,
          cols,
        };

        const calcFn = new Function(`return (${CALC_RANGE_SIZE_SCRIPT})`)();
        const rangeSize = (await page.evaluate(
          calcFn as never,
          rangeOptions,
        )) as { width: number; height: number } | null;

        if (rangeSize) {
          const needW = rangeSize.width + RANGE_PADDING_PX;
          const needH = rangeSize.height + RANGE_PADDING_PX;
          if (
            needW > DEFAULT_VIEWPORT.width ||
            needH > DEFAULT_VIEWPORT.height
          ) {
            await page.setViewportSize({
              width: Math.max(DEFAULT_VIEWPORT.width, needW),
              height: Math.max(DEFAULT_VIEWPORT.height, needH),
            });
            await page.waitForTimeout(200);
          }
        }

        const clipFn = new Function(`return (${CLIP_RANGE_SCRIPT})`)();
        const clipRect = (await page.evaluate(
          clipFn as never,
          rangeOptions,
        )) as {
          x: number;
          y: number;
          width: number;
          height: number;
          error?: string;
        } | null;

        if (clipRect?.error) fail(clipRect.error);
        await page.waitForTimeout(RENDER_SETTLE_MS);

        if (clipRect && clipRect.width > 0 && clipRect.height > 0) {
          await page.screenshot({ path: outPath, clip: clipRect });
        } else {
          await page.locator("#ss").screenshot({ path: outPath });
        }
      } else {
        await page.waitForTimeout(RENDER_SETTLE_MS);
        await page.locator("#ss").screenshot({ path: outPath });
      }
    } finally {
      await browser.close();
    }
  } finally {
    await closeServer(server);
  }

  ok({ file: outPath });
}
