/**
 * LRU file cache with mtime invalidation for the daemon.
 *
 * Uses Map insertion-order for O(1) LRU eviction: on cache hit,
 * delete + re-insert moves the entry to the end. On eviction,
 * the first key is always the least-recently-used.
 */

import { stat } from "node:fs/promises";
import { resolve } from "node:path";
import type { ExcelFile } from "../index.js";

interface CacheEntry {
  file: ExcelFile;
  absPath: string;
  mtime: number;
}

export class FileCache {
  private cache = new Map<string, CacheEntry>();
  private maxSize: number;

  constructor(maxSize = 10) {
    this.maxSize = maxSize;
  }

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

    // Check mtime — if file changed on disk, invalidate
    const currentMtime = await this.getMtime(absPath);
    if (currentMtime !== entry.mtime) {
      this.cache.delete(absPath);
      return null;
    }

    // Move to end (most-recently-used)
    this.cache.delete(absPath);
    this.cache.set(absPath, entry);
    return { file: entry.file, absPath };
  }

  async put(filePath: string, cwd: string, file: ExcelFile): Promise<void> {
    const absPath = this.key(filePath, cwd);
    const mtime = await this.getMtime(absPath);

    // Evict LRU (first key) if at capacity
    if (this.cache.size >= this.maxSize && !this.cache.has(absPath)) {
      const oldest = this.cache.keys().next().value;
      if (oldest) this.cache.delete(oldest);
    }

    this.cache.set(absPath, { file, absPath, mtime });
  }

  /** Update mtime after saving a file */
  async updateMtime(filePath: string, cwd: string): Promise<void> {
    const absPath = this.key(filePath, cwd);
    const entry = this.cache.get(absPath);
    if (entry) {
      entry.mtime = await this.getMtime(absPath);
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
