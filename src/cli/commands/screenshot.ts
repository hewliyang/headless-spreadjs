import http from "node:http";
import path from "node:path";
import { parseRef, rangeDimensions } from "../a1.js";
import { throwIfAborted } from "../abort.js";
import { withFile } from "../context.js";
import { fail, ok } from "../output.js";

/**
 * CDN URLs for SpreadJS (trial/eval, browser-only rendering).
 * Uses the older @grapecity packages which work in eval mode with fromJSON.
 */
const CDN = {
  css: "https://cdn.jsdelivr.net/npm/@grapecity/spread-sheets/styles/gc.spread.sheets.excel2013white.min.css",
  sheets:
    "https://cdn.jsdelivr.net/npm/@grapecity/spread-sheets/dist/gc.spread.sheets.all.min.js",
  excelio:
    "https://cdn.jsdelivr.net/npm/@grapecity/spread-excelio/dist/gc.spread.excelio.min.js",
};

interface ScreenshotOptions {
  ref?: string;
  out?: string;
  signal?: AbortSignal | null;
}

function buildViewerHTML(): string {
  return `<!DOCTYPE html>
<html><head>
<link rel="stylesheet" href="${CDN.css}">
<style>
  body { margin: 0; padding: 0; background: white; }
  #ss { width: 100%; height: 100vh; }
</style>
</head><body>
<div id="ss"></div>
<script src="${CDN.sheets}"></script>
<script src="${CDN.excelio}"></script>
<script>
(function(){
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
})();

var spread = new GC.Spread.Sheets.Workbook(document.getElementById('ss'));

fetch('/file.xlsx?t=' + Date.now()).then(function(r){ return r.blob(); }).then(function(blob){
  var io = new GC.Spread.Excel.IO();
  io.open(blob, function(json){
    if (json.sheets) {
      Object.values(json.sheets).forEach(function(s) {
        s.rowCount  = Math.max(s.rowCount  || 0, 500);
        s.columnCount = Math.max(s.columnCount || 0, 100);
      });
    }
    spread.fromJSON(json);
    spread.refresh();
    window.__spread = spread;
    window.__ready  = true;
  }, function(err){
    window.__error = String(err);
    window.__ready = true;
  });
});
</script>
</body></html>`;
}

function startServer(
  html: string,
  xlsxBuf: Buffer,
): Promise<{ server: http.Server; port: number }> {
  return new Promise((resolve, reject) => {
    const server = http.createServer((req, res) => {
      if (req.url === "/" || req.url === "/index.html") {
        res.writeHead(200, { "Content-Type": "text/html" });
        res.end(html);
      } else if (req.url?.startsWith("/file.xlsx")) {
        res.writeHead(200, {
          "Content-Type": "application/octet-stream",
          "Cache-Control": "no-cache",
        });
        res.end(xlsxBuf);
      } else {
        res.writeHead(404);
        res.end();
      }
    });
    server.on("error", reject);
    server.listen(0, "127.0.0.1", () => {
      const addr = server.address();
      if (!addr || typeof addr === "string")
        return reject(new Error("bind failed"));
      resolve({ server, port: addr.port });
    });
  });
}

/**
 * Browser-side script that navigates to a range and returns its pixel rect.
 * Injected via page.evaluate() as a string to avoid DOM type issues.
 */
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

  const x = first.x + hostRect.left;
  const y = first.y + hostRect.top;
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

  // Resolve output path
  const outPath = options.out
    ? path.resolve(options.out)
    : path.resolve(`${path.basename(filePath, path.extname(filePath))}.png`);

  // Get xlsx buffer from the headless engine
  const xlsxBuf = await withFile(
    filePath,
    async ({ file }) => {
      throwIfAborted(signal);
      return file.saveToBuffer();
    },
    { signal },
  );

  throwIfAborted(signal);

  // Import playwright lazily (optional dependency)
  let chromium: typeof import("playwright")["chromium"];
  try {
    const pw = await import("playwright");
    chromium = pw.chromium;
  } catch {
    fail(
      "playwright is required for screenshots. Install it:\n  npm i -D playwright && npx playwright install chromium",
    );
  }

  // Parse range ref (if given)
  const rangeRef = options.ref ? parseRef(options.ref) : null;

  // Start local HTTP server
  const html = buildViewerHTML();
  const { server, port } = await startServer(html, xlsxBuf);

  try {
    throwIfAborted(signal);
    const browser = await chromium.launch();
    try {
      // Start with a default viewport; we may resize later for large ranges
      const page = await browser.newPage({
        viewport: { width: 1280, height: 800 },
      });

      await page.goto(`http://127.0.0.1:${port}`, {
        waitUntil: "networkidle",
      });

      // Wait for SpreadJS to finish loading the xlsx
      await page.waitForFunction("window.__ready", { timeout: 15_000 });

      // Check for load errors
      const loadError = await page.evaluate("window.__error");
      if (loadError) fail(`SpreadJS IO error: ${loadError}`);

      if (rangeRef) {
        const { rows, cols } = rangeDimensions(rangeRef);

        // Compute the pixel size the range needs so we can resize the viewport
        const CALC_SIZE_SCRIPT = `(opts) => {
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
          // Add space for row/column headers if visible
          const rhw = sheet.options.rowHeaderVisible !== false ? sheet.getColumnWidth(0, GC.Spread.Sheets.SheetArea.rowHeader) : 0;
          const chh = sheet.options.colHeaderVisible !== false ? sheet.getRowHeight(0, GC.Spread.Sheets.SheetArea.colHeader) : 0;
          return { width: Math.ceil(w + rhw), height: Math.ceil(h + chh) };
        }`;

        const calcFn = new Function(`return (${CALC_SIZE_SCRIPT})`)();
        const rangeSize = (await page.evaluate(calcFn as never, {
          sheetName: rangeRef.sheet,
          startRow: rangeRef.start.row,
          startCol: rangeRef.start.col,
          rows,
          cols,
        })) as { width: number; height: number } | null;

        // Resize viewport if the range needs more space (with some padding)
        if (rangeSize) {
          const needW = rangeSize.width + 50;
          const needH = rangeSize.height + 50;
          if (needW > 1280 || needH > 800) {
            await page.setViewportSize({
              width: Math.max(1280, needW),
              height: Math.max(800, needH),
            });
            await page.waitForTimeout(200);
          }
        }

        // Navigate to range and get clip rect
        const clipFn = new Function(`return (${CLIP_RANGE_SCRIPT})`)();
        const clipRect = (await page.evaluate(clipFn as never, {
          sheetName: rangeRef.sheet,
          startRow: rangeRef.start.row,
          startCol: rangeRef.start.col,
          rows,
          cols,
        })) as {
          x: number;
          y: number;
          width: number;
          height: number;
          error?: string;
        } | null;

        if (clipRect && "error" in clipRect && clipRect.error) {
          fail(clipRect.error);
        }

        await page.waitForTimeout(300);

        if (clipRect && clipRect.width > 0 && clipRect.height > 0) {
          await page.screenshot({ path: outPath, clip: clipRect });
        } else {
          await page.locator("#ss").screenshot({ path: outPath });
        }
      } else {
        await page.waitForTimeout(300);
        await page.locator("#ss").screenshot({ path: outPath });
      }
    } finally {
      await browser.close();
    }
  } finally {
    server.close();
  }

  ok({ file: outPath });
}
