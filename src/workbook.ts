import fs from "node:fs";
import path from "node:path";
import { unzipSync, zipSync, type Unzipped } from "fflate";
import type { GCNamespace, SpreadWorkbook, SpreadWorksheet } from "./types.js";

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
      sheetPath = "xl/" + targetMatch[1];
    }
  }

  const newRelsXml = relsXml.replace(relRe, "");

  const ctBytes = zip["[Content_Types].xml"];
  let newCtXml: string | null = null;
  if (ctBytes && sheetPath) {
    const ctXml = decoder.decode(ctBytes);
    const partName = "/" + sheetPath;
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

export function setGC(gc: GCNamespace | null): void {
  gcRef = gc;
}

export class Workbook {
  public readonly spread: SpreadWorkbook;

  constructor(spread?: SpreadWorkbook) {
    const gc = getGC();
    this.spread = spread ?? new gc.Spread.Sheets.Workbook();
  }

  batch<T>(fn: () => T): T;
  batch<T>(fn: () => Promise<T>): Promise<T>;
  batch<T>(fn: () => T | Promise<T>): T | Promise<T> {
    this.spread.suspendCalcService(false);

    try {
      const result = fn();

      if (result instanceof Promise) {
        return result.finally(() => {
          this.spread.resumeCalcService(true);
        });
      }

      this.spread.resumeCalcService(true);
      return result;
    } catch (error) {
      this.spread.resumeCalcService(true);
      throw error;
    }
  }

  getActiveSheet(): SpreadWorksheet {
    return this.spread.getActiveSheet();
  }

  getSheet(index: number): SpreadWorksheet {
    return this.spread.getSheet(index);
  }

  getSheetCount(): number {
    return this.spread.getSheetCount();
  }

  addSheet(name: string, index?: number): SpreadWorksheet {
    const gc = getGC();
    const sheet = new gc.Spread.Sheets.Worksheet(name);
    this.spread.addSheet(index ?? this.spread.getSheetCount(), sheet);
    return sheet;
  }

  async save(filePath: string): Promise<void> {
    const bytes = await this.saveToBuffer();
    const dir = path.dirname(filePath);

    if (dir && dir !== ".") {
      await fs.promises.mkdir(dir, { recursive: true });
    }

    await fs.promises.writeFile(filePath, bytes);
  }

  saveToBuffer(): Promise<Buffer> {
    return saveAsBuffer(this.spread, getGC());
  }


  toJSON(): object {
    return this.spread.toJSON();
  }

  fromJSON(json: object): void {
    this.spread.fromJSON(json);
  }

  static async open(filePath: string): Promise<Workbook> {
    const gc = getGC();
    const bytes = await fs.promises.readFile(filePath);
    const spread = await importWorkbook(toArrayBuffer(bytes), gc);
    return new Workbook(spread);
  }

  static async openFromBuffer(buffer: Buffer): Promise<Workbook> {
    const gc = getGC();
    const spread = await importWorkbook(toArrayBuffer(buffer), gc);
    return new Workbook(spread);
  }
}
