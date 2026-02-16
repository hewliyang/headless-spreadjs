/**
 * LRU file cache with mtime invalidation for the daemon.
 */

import { statSync } from "node:fs";
import { resolve } from "node:path";
import type { ExcelFile } from "../index.js";

interface CacheEntry {
  file: ExcelFile;
  absPath: string;
  mtime: number;
  lastAccess: number;
}

export class FileCache {
  private cache = new Map<string, CacheEntry>();
  private maxSize: number;

  constructor(maxSize = 10) {
    this.maxSize = maxSize;
  }

  private key(filePath: string, cwd: string): string {
    return resolve(cwd, filePath);
  }

  private getMtime(absPath: string): number {
    try {
      return statSync(absPath).mtimeMs;
    } catch {
      return -1;
    }
  }

  get(
    filePath: string,
    cwd: string,
  ): { file: ExcelFile; absPath: string } | null {
    const absPath = this.key(filePath, cwd);
    const entry = this.cache.get(absPath);
    if (!entry) return null;

    // Check mtime — if file changed on disk, invalidate
    const currentMtime = this.getMtime(absPath);
    if (currentMtime !== entry.mtime) {
      this.cache.delete(absPath);
      return null;
    }

    entry.lastAccess = Date.now();
    return { file: entry.file, absPath };
  }

  put(
    filePath: string,
    cwd: string,
    file: ExcelFile,
  ): void {
    const absPath = this.key(filePath, cwd);
    const mtime = this.getMtime(absPath);

    // Evict LRU if at capacity
    if (this.cache.size >= this.maxSize && !this.cache.has(absPath)) {
      let oldest: string | null = null;
      let oldestTime = Infinity;
      for (const [key, entry] of this.cache) {
        if (entry.lastAccess < oldestTime) {
          oldestTime = entry.lastAccess;
          oldest = key;
        }
      }
      if (oldest) this.cache.delete(oldest);
    }

    this.cache.set(absPath, {
      file,
      absPath,
      mtime,
      lastAccess: Date.now(),
    });
  }

  /** Update mtime after saving a file */
  updateMtime(filePath: string, cwd: string): void {
    const absPath = this.key(filePath, cwd);
    const entry = this.cache.get(absPath);
    if (entry) {
      entry.mtime = this.getMtime(absPath);
    }
  }

  invalidate(filePath: string, cwd: string): void {
    this.cache.delete(this.key(filePath, cwd));
  }

  clear(): void {
    this.cache.clear();
  }

  get size(): number {
    return this.cache.size;
  }
}
