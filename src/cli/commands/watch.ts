import { exec } from "node:child_process";
import { existsSync, watch as fsWatch, statSync } from "node:fs";
import { readFile } from "node:fs/promises";
import { basename, extname, resolve } from "node:path";
import { tryExistingDaemon } from "../client.js";
import type { FileCache } from "../file-cache.js";
import { ok, writeStderr } from "../output.js";
import {
  type WatchEventListener,
  type WatchFileProvider,
  WatchServer,
} from "../watch-server.js";

// ---------------------------------------------------------------------------
// Disk-backed provider (standalone / --no-daemon)
// ---------------------------------------------------------------------------

interface WatchedFile {
  absPath: string;
  lastSize: number;
  lastMtime: number;
}

function errorMessage(err: unknown): string {
  return err instanceof Error ? err.message : String(err);
}

function getFileStat(absPath: string): { size: number; mtime: number } | null {
  try {
    const s = statSync(absPath);
    return { size: s.size, mtime: s.mtimeMs };
  } catch {
    return null;
  }
}

class DiskProvider implements WatchFileProvider {
  private files: WatchedFile[] = [];
  private buffers = new Map<string, Buffer>();
  private listeners = new Set<WatchEventListener>();
  private watchers: ReturnType<typeof fsWatch>[] = [];
  private debounceTimers = new Map<string, ReturnType<typeof setTimeout>>();

  async addFiles(absPaths: string[]): Promise<void> {
    for (const absPath of absPaths) {
      const s = getFileStat(absPath);
      this.files.push({
        absPath,
        lastSize: s?.size ?? 0,
        lastMtime: s?.mtime ?? 0,
      });
      try {
        this.buffers.set(absPath, await readFile(absPath));
      } catch (err) {
        throw new Error(`Failed to read ${absPath}: ${errorMessage(err)}`);
      }
      this.startWatching(absPath);
    }
  }

  private startWatching(absPath: string): void {
    const f = this.files.find((x) => x.absPath === absPath);
    if (!f) return;

    const watcher = fsWatch(f.absPath, { persistent: true }, () => {
      const existing = this.debounceTimers.get(f.absPath);
      if (existing) clearTimeout(existing);

      this.debounceTimers.set(
        f.absPath,
        setTimeout(async () => {
          this.debounceTimers.delete(f.absPath);
          const s = getFileStat(f.absPath);
          if (!s) return;
          if (s.size === f.lastSize && s.mtime === f.lastMtime) return;
          f.lastSize = s.size;
          f.lastMtime = s.mtime;

          try {
            this.buffers.set(f.absPath, await readFile(f.absPath));
          } catch (err) {
            writeStderr(
              `Warning: failed to reload ${f.absPath}: ${errorMessage(err)}\n`,
            );
            return;
          }

          for (const l of this.listeners) {
            l({ type: "changed", absPath: f.absPath });
          }
        }, 200),
      );
    });
    this.watchers.push(watcher);
  }

  listFiles(): string[] {
    return this.files.map((f) => f.absPath);
  }

  async getBuffer(absPath: string): Promise<Buffer | null> {
    return this.buffers.get(absPath) ?? null;
  }

  subscribe(listener: WatchEventListener): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  stop(): void {
    for (const w of this.watchers) w.close();
    for (const t of this.debounceTimers.values()) clearTimeout(t);
    this.watchers = [];
    this.debounceTimers.clear();
  }
}

// ---------------------------------------------------------------------------
// Daemon-backed provider (in-process, used when watch runs inside daemon)
// ---------------------------------------------------------------------------

export class DaemonProvider implements WatchFileProvider {
  private listeners = new Set<WatchEventListener>();

  constructor(private fileCache: FileCache) {
    fileCache.events.on("opened", (absPath) => {
      for (const l of this.listeners) {
        l({ type: "opened", absPath });
      }
    });

    fileCache.events.on("changed", (absPath) => {
      for (const l of this.listeners) {
        l({ type: "changed", absPath });
      }
    });
  }

  listFiles(): string[] {
    return this.fileCache.files().map((f) => f.absPath);
  }

  async getBuffer(absPath: string): Promise<Buffer | null> {
    const file = this.fileCache.getFile(absPath);
    if (!file) return null;
    return file.saveToBuffer();
  }

  subscribe(listener: WatchEventListener): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }
}

// ---------------------------------------------------------------------------
// Daemon auto-discovery (when watch runs as separate process)
// ---------------------------------------------------------------------------

async function discoverFromDaemon(): Promise<string[] | null> {
  try {
    const result = await tryExistingDaemon(["daemon", "files"], process.cwd());
    if (!result || result.exitCode !== 0) return null;
    const parsed = JSON.parse(result.stdout.trim()) as {
      files: { absPath: string; dirty: boolean }[];
    };
    if (!parsed.files || parsed.files.length === 0) return null;
    return parsed.files.map((f) => f.absPath);
  } catch (err) {
    writeStderr(
      `Warning: failed to discover files from daemon: ${errorMessage(err)}\n`,
    );
    return null;
  }
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

export interface WatchOptions {
  port?: number;
  open?: boolean;
  signal?: AbortSignal | null;
}

function openUrl(url: string, onError?: (message: string) => void): void {
  const cmd =
    process.platform === "darwin"
      ? `open "${url}"`
      : process.platform === "win32"
        ? `start "" "${url}"`
        : `xdg-open "${url}"`;

  const child = exec(cmd, (err) => {
    if (err) {
      onError?.(err.message);
    }
  });
  child.on("error", (err) => onError?.(err.message));
}

/**
 * Standalone watch (runs as its own process, reads from disk).
 * Auto-discovers from daemon cache if no files specified.
 */
export async function watch(
  filePaths: string[],
  options: WatchOptions = {},
): Promise<void> {
  const httpPort = options.port ?? 8080;

  // Auto-discover from daemon when no files specified
  if (filePaths.length === 0) {
    const daemonFiles = await discoverFromDaemon();
    if (daemonFiles && daemonFiles.length > 0) {
      filePaths = daemonFiles;
      writeStderr(
        `Auto-discovered ${filePaths.length} file(s) from daemon cache\n`,
      );
    }
  }

  // Resolve and validate
  const absPaths: string[] = [];
  for (const fp of filePaths) {
    const absPath = resolve(fp);
    if (!existsSync(absPath)) {
      throw new Error(`File not found: ${fp}`);
    }
    const ext = extname(absPath).toLowerCase();
    if (ext !== ".xlsx" && ext !== ".xlsm") {
      throw new Error(`Expected .xlsx or .xlsm file, got: ${ext}`);
    }
    absPaths.push(absPath);
  }

  const provider = new DiskProvider();
  if (absPaths.length > 0) {
    await provider.addFiles(absPaths);
  }

  const server = new WatchServer(provider);
  const actualPort = await server.start(httpPort);

  const fileList =
    absPaths.length > 0
      ? absPaths.map((p) => basename(p)).join(", ")
      : "(waiting for files)";

  writeStderr(
    `Watching ${absPaths.length} file(s): ${fileList}\n` +
      `  http://127.0.0.1:${actualPort}\n` +
      `\nPress Ctrl+C to stop\n`,
  );

  ok({
    watching: absPaths.map((p) => basename(p)),
    url: `http://127.0.0.1:${actualPort}`,
  });

  if (options.open) {
    const url = `http://127.0.0.1:${actualPort}`;
    openUrl(
      url,
      (message) => writeStderr(`Warning: failed to open browser: ${message}\n`),
    );
  }

  return new Promise<void>((res) => {
    const cleanup = () => {
      provider.stop();
      server.stop();
      res();
    };

    if (options.signal) {
      options.signal.addEventListener("abort", cleanup, { once: true });
    }
    process.on("SIGINT", cleanup);
    process.on("SIGTERM", cleanup);
  });
}
