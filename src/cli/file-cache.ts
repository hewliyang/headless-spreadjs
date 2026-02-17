/**
 * LRU file cache with mtime invalidation and deferred saves for the daemon.
 *
 * Uses Map insertion-order for O(1) LRU eviction: on cache hit,
 * delete + re-insert moves the entry to the end. On eviction,
 * the first key is always the least-recently-used.
 *
 * Write operations mark entries as dirty instead of saving immediately.
 * Dirty entries are flushed to disk on eviction, daemon stop, or
 * process exit to guarantee eventual consistency.
 */

import { stat } from "node:fs/promises";
import { resolve } from "node:path";
import type { ExcelFile } from "../index.js";

interface CacheEntry {
  file: ExcelFile;
  absPath: string;
  mtime: number;
  dirty: boolean;
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

    // Dirty entries always win — skip mtime check since our
    // in-memory state is ahead of disk.
    if (!entry.dirty) {
      // Check mtime — if file changed on disk, invalidate
      const currentMtime = await this.getMtime(absPath);
      if (currentMtime !== entry.mtime) {
        this.cache.delete(absPath);
        return null;
      }
    }

    // Move to end (most-recently-used)
    this.cache.delete(absPath);
    this.cache.set(absPath, entry);
    return { file: entry.file, absPath };
  }

  async put(filePath: string, cwd: string, file: ExcelFile): Promise<void> {
    const absPath = this.key(filePath, cwd);
    const mtime = await this.getMtime(absPath);

    // Evict LRU (first key) if at capacity — flush if dirty
    if (this.cache.size >= this.maxSize && !this.cache.has(absPath)) {
      const oldest = this.cache.keys().next().value;
      if (oldest) {
        const entry = this.cache.get(oldest);
        if (entry?.dirty) {
          await entry.file.save(entry.absPath);
        }
        this.cache.delete(oldest);
      }
    }

    this.cache.set(absPath, { file, absPath, mtime, dirty: false });
  }

  /** Mark a cached entry as dirty (has unsaved mutations). */
  markDirty(filePath: string, cwd: string): void {
    const absPath = this.key(filePath, cwd);
    const entry = this.cache.get(absPath);
    if (entry) entry.dirty = true;
  }

  /** Update mtime after saving a file and clear dirty flag. */
  async updateMtime(filePath: string, cwd: string): Promise<void> {
    const absPath = this.key(filePath, cwd);
    const entry = this.cache.get(absPath);
    if (entry) {
      entry.mtime = await this.getMtime(absPath);
      entry.dirty = false;
    }
  }

  invalidate(filePath: string, cwd: string): void {
    this.cache.delete(this.key(filePath, cwd));
  }

  /** Flush all dirty entries to disk. */
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

  /** Number of entries with unsaved changes. */
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
}
