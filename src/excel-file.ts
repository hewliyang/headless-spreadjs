import fs from "node:fs";
import path from "node:path";
import { type Unzipped, unzipSync, zipSync } from "fflate";
import type { GCNamespace, SpreadWorkbook } from "./types.js";

let gcRef: GCNamespace | null = null;

function getGC(): GCNamespace {
  if (!gcRef) {
    throw new Error("headless-spreadjs not initialized. Call init() first.");
  }
  return gcRef;
}

function stripEvalSheet(spread: SpreadWorkbook): void {
  for (let i = spread.getSheetCount() - 1; i >= 0; i--) {
    if (spread.getSheet(i).name() === "Evaluation Version") {
      spread.removeSheet(i);
    }
  }
}

type UnknownMethod = (...args: unknown[]) => unknown;

function getMethod(obj: unknown, name: string): UnknownMethod | null {
  if (!obj || typeof obj !== "object") return null;
  const value = (obj as Record<string, unknown>)[name];
  return typeof value === "function" ? (value as UnknownMethod) : null;
}

function toArrayBuffer(buffer: Buffer): ArrayBuffer {
  return buffer.buffer.slice(
    buffer.byteOffset,
    buffer.byteOffset + buffer.byteLength,
  ) as ArrayBuffer;
}

function stripEvalSheetFromZip(buf: Buffer): Buffer {
  const zip = unzipSync(new Uint8Array(buf));

  const wbXmlBytes = zip["xl/workbook.xml"];
  if (!wbXmlBytes) return buf;

  const decoder = new TextDecoder();
  const wbXml = decoder.decode(wbXmlBytes);

  const sheetRe = /<sheet\s[^>]*name="Evaluation Version"[^>]*\/>/g;
  const match = sheetRe.exec(wbXml);
  if (!match) return buf;

  const ridMatch = /r:id="(rId\d+)"/.exec(match[0]);
  if (!ridMatch) return buf;
  const rId = ridMatch[1];

  const allSheets = [...wbXml.matchAll(/<sheet\s[^>]*\/>/g)];
  const evalIndex = allSheets.findIndex((m) => m[0] === match[0]);

  let newWbXml = wbXml.replace(match[0], "");

  newWbXml = newWbXml.replace(
    /activeTab="(\d+)"/g,
    (_full: string, tabStr: string) => {
      let tab = parseInt(tabStr, 10);
      if (evalIndex >= 0 && tab >= evalIndex) {
        tab = Math.max(0, tab - 1);
      }
      return `activeTab="${tab}"`;
    },
  );

  const relsBytes = zip["xl/_rels/workbook.xml.rels"];
  if (!relsBytes) return buf;
  const relsXml = decoder.decode(relsBytes);

  const relRe = new RegExp(`<Relationship[^>]*Id="${rId}"[^>]*/>`, "g");
  const relMatch = relRe.exec(relsXml);
  let sheetPath: string | null = null;
  if (relMatch) {
    const targetMatch = /Target="([^"]+)"/.exec(relMatch[0]);
    if (targetMatch) {
      sheetPath = `xl/${targetMatch[1]}`;
    }
  }

  const newRelsXml = relsXml.replace(relRe, "");

  const ctBytes = zip["[Content_Types].xml"];
  let newCtXml: string | null = null;
  if (ctBytes && sheetPath) {
    const ctXml = decoder.decode(ctBytes);
    const partName = `/${sheetPath}`;
    const overrideRe = new RegExp(
      `<Override[^>]*PartName="${partName.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}"[^>]*/>`,
      "g",
    );
    newCtXml = ctXml.replace(overrideRe, "");
  }

  const encoder = new TextEncoder();
  const newZip: Unzipped = {};
  for (const [name, data] of Object.entries(zip)) {
    if (sheetPath && name === sheetPath) continue;
    if (name === "xl/workbook.xml") {
      newZip[name] = encoder.encode(newWbXml);
    } else if (name === "xl/_rels/workbook.xml.rels") {
      newZip[name] = encoder.encode(newRelsXml);
    } else if (name === "[Content_Types].xml" && newCtXml) {
      newZip[name] = encoder.encode(newCtXml);
    } else {
      newZip[name] = data;
    }
  }

  return Buffer.from(zipSync(newZip));
}

function saveAsBuffer(
  spread: SpreadWorkbook,
  gc: GCNamespace,
): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    spread.save(
      (blob: Blob) => {
        blob
          .arrayBuffer()
          .then((bytes) => resolve(stripEvalSheetFromZip(Buffer.from(bytes))))
          .catch(reject);
      },
      (err: Error) => reject(err),
      { fileType: gc.Spread.Sheets.FileType.excel } as never,
    );
  });
}

function importWorkbook(
  bytes: ArrayBuffer,
  gc: GCNamespace,
): Promise<SpreadWorkbook> {
  return new Promise((resolve, reject) => {
    const spread = new gc.Spread.Sheets.Workbook();

    spread.import(
      bytes as unknown as File,
      () => {
        stripEvalSheet(spread);
        resolve(spread);
      },
      (err: Error) => reject(err),
      { fileType: gc.Spread.Sheets.FileType.excel } as never,
    );
  });
}

/**
 * SpreadJS defaults to 200 rows × 20 columns and **silently drops**
 * any writes outside that range.  Expand every sheet in the workbook
 * to the Excel-maximum so callers never hit this.
 *
 * Cost: ~16 MB heap + ~180 ms per workbook (one-time).
 * File size is unaffected — SpreadJS only serializes populated cells.
 */
const EXCEL_MAX_ROWS = 1_048_576;
const EXCEL_MAX_COLS = 16_384;

function expandSheetDefaults(spread: SpreadWorkbook): void {
  const n = spread.getSheetCount();
  for (let i = 0; i < n; i++) {
    const s = spread.getSheet(i);
    if (s.getRowCount() < EXCEL_MAX_ROWS) s.setRowCount(EXCEL_MAX_ROWS);
    if (s.getColumnCount() < EXCEL_MAX_COLS) s.setColumnCount(EXCEL_MAX_COLS);
  }
}

export function setGC(gc: GCNamespace | null): void {
  gcRef = gc;
}

export class ExcelFile {
  public readonly workbook: SpreadWorkbook;
  /** The file path this workbook was opened from (if any). */
  public sourcePath: string | undefined;

  constructor(workbook?: SpreadWorkbook) {
    const gc = getGC();
    this.workbook = workbook ?? new gc.Spread.Sheets.Workbook();
    expandSheetDefaults(this.workbook);

    this.workbook.options.allowDynamicArray = true;
  }

  batch<T>(fn: () => T): T;
  batch<T>(fn: () => Promise<T>): Promise<T>;
  batch<T>(fn: () => T | Promise<T>): T | Promise<T> {
    const suspendPaint = getMethod(this.workbook, "suspendPaint");
    const resumePaint = getMethod(this.workbook, "resumePaint");
    const suspendEvent = getMethod(this.workbook, "suspendEvent");
    const resumeEvent = getMethod(this.workbook, "resumeEvent");

    suspendPaint?.call(this.workbook);
    suspendEvent?.call(this.workbook);
    this.workbook.suspendCalcService(false);

    const finalize = () => {
      this.workbook.resumeCalcService(true);
      resumeEvent?.call(this.workbook);
      resumePaint?.call(this.workbook);
    };

    try {
      const result = fn();

      if (result instanceof Promise) {
        return result.finally(finalize);
      }

      finalize();
      return result;
    } catch (error) {
      finalize();
      throw error;
    }
  }

  async save(filePath?: string): Promise<void> {
    filePath ??= this.sourcePath;
    if (!filePath) {
      throw new Error(
        "No file path provided. Pass a path to save() or open the file with ExcelFile.open().",
      );
    }
    const bytes = await this.saveToBuffer();
    const dir = path.dirname(filePath);

    if (dir && dir !== ".") {
      await fs.promises.mkdir(dir, { recursive: true });
    }

    await fs.promises.writeFile(filePath, bytes);
  }

  saveToBuffer(): Promise<Buffer> {
    return saveAsBuffer(this.workbook, getGC());
  }

  toJSON(): object {
    return this.workbook.toJSON();
  }

  fromJSON(json: object): void {
    this.workbook.fromJSON(json);
  }

  static async open(filePath: string): Promise<ExcelFile> {
    const gc = getGC();
    const bytes = await fs.promises.readFile(filePath);
    const spread = await importWorkbook(toArrayBuffer(bytes), gc);
    const ef = new ExcelFile(spread);
    ef.sourcePath = path.resolve(filePath);
    return ef;
  }

  static async openFromBuffer(buffer: Buffer): Promise<ExcelFile> {
    const gc = getGC();
    const spread = await importWorkbook(toArrayBuffer(buffer), gc);
    return new ExcelFile(spread);
  }
}
