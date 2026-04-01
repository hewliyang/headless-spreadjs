import { EventEmitter } from "node:events";
import { stat } from "node:fs/promises";
import { resolve } from "node:path";
import type { ExcelFile } from "../index.js";

interface CacheEntry {
  file: ExcelFile;
  absPath: string;
  mtime: number;
  dirty: boolean;
}

export interface FileCacheEvents {
  opened: [absPath: string, file: ExcelFile];
  changed: [absPath: string, file: ExcelFile];
}

export class FileCache {
  private cache = new Map<string, CacheEntry>();
  readonly events = new EventEmitter<FileCacheEvents>();

  constructor(private readonly maxSize = 10) {}

  get maxCacheSize(): number {
    return this.maxSize;
  }

  private key(filePath: string, cwd: string): string {
    return resolve(cwd, filePath);
  }

  private async getMtime(absPath: string): Promise<number> {
    try {
      return (await stat(absPath)).mtimeMs;
    } catch {
      return -1;
    }
  }

  async get(
    filePath: string,
    cwd: string,
  ): Promise<{ file: ExcelFile; absPath: string } | null> {
    const absPath = this.key(filePath, cwd);
    const entry = this.cache.get(absPath);
    if (!entry) return null;

    if (!entry.dirty) {
      const currentMtime = await this.getMtime(absPath);
      if (currentMtime !== entry.mtime) {
        this.cache.delete(absPath);
        return null;
      }
    }

    this.cache.delete(absPath);
    this.cache.set(absPath, entry);
    return { file: entry.file, absPath };
  }

  async put(filePath: string, cwd: string, file: ExcelFile): Promise<void> {
    const absPath = this.key(filePath, cwd);
    const mtime = await this.getMtime(absPath);

    if (this.cache.size >= this.maxSize && !this.cache.has(absPath)) {
      const oldest = this.cache.keys().next().value as string | undefined;
      if (oldest !== undefined) {
        const entry = this.cache.get(oldest);
        if (entry?.dirty) {
          await entry.file.save(entry.absPath);
        }
        this.cache.delete(oldest);
      }
    }

    const isNew = !this.cache.has(absPath);
    this.cache.set(absPath, { file, absPath, mtime, dirty: false });
    if (isNew) {
      this.events.emit("opened", absPath, file);
    }
  }

  markDirty(filePath: string, cwd: string): void {
    const absPath = this.key(filePath, cwd);
    const entry = this.cache.get(absPath);
    if (entry) {
      entry.dirty = true;
      this.events.emit("changed", absPath, entry.file);
    }
  }

  async updateMtime(filePath: string, cwd: string): Promise<void> {
    const absPath = this.key(filePath, cwd);
    const entry = this.cache.get(absPath);
    if (entry) {
      entry.mtime = await this.getMtime(absPath);
      const wasDirty = entry.dirty;
      entry.dirty = false;
      if (wasDirty) {
        this.events.emit("changed", absPath, entry.file);
      }
    }
  }

  isDirty(filePath: string, cwd: string): boolean {
    const entry = this.cache.get(this.key(filePath, cwd));
    return !!entry?.dirty;
  }

  invalidate(filePath: string, cwd: string): void {
    this.cache.delete(this.key(filePath, cwd));
  }

  async flushDirty(): Promise<number> {
    let flushed = 0;
    for (const entry of this.cache.values()) {
      if (entry.dirty) {
        await entry.file.save(entry.absPath);
        entry.mtime = await this.getMtime(entry.absPath);
        entry.dirty = false;
        flushed++;
      }
    }
    return flushed;
  }

  get dirtyCount(): number {
    let count = 0;
    for (const entry of this.cache.values()) {
      if (entry.dirty) count++;
    }
    return count;
  }

  clear(): void {
    this.cache.clear();
  }

  get size(): number {
    return this.cache.size;
  }

  files(): { absPath: string; dirty: boolean }[] {
    return [...this.cache.values()].map((e) => ({
      absPath: e.absPath,
      dirty: e.dirty,
    }));
  }

  getFile(absPath: string): ExcelFile | null {
    return this.cache.get(absPath)?.file ?? null;
  }
}
