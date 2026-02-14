import fs from "node:fs";
import path from "node:path";
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

function saveAsBuffer(
  spread: SpreadWorkbook,
  gc: GCNamespace,
): Promise<Buffer> {
  return new Promise((resolve, reject) => {
    spread.save(
      (blob: Blob) => {
        blob
          .arrayBuffer()
          .then((bytes) => resolve(Buffer.from(bytes)))
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
