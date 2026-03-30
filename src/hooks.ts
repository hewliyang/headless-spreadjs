/**
 * Extension hook system for headless-spreadjs.
 *
 * Hooks allow users to run custom code at specific points in the CLI workflow.
 * Common use cases include:
 * - Applying formatting conventions (e.g., financial modeling colors)
 * - Setting workbook defaults (e.g., hiding gridlines)
 * - Validation and linting before/after save
 * - Logging and auditing tool invocations
 *
 * Discovery:
 *   Hooks are auto-discovered from these locations (TypeScript via jiti):
 *   1. ./.headless-spreadjs/hooks/*.{ts,js,mjs} (project-local, takes precedence)
 *   2. ~/.headless-spreadjs/hooks/*.{ts,js,mjs} (global fallback)
 *
 * Discovery can be disabled with HSX_NO_HOOKS=1.
 *
 * Hook Points:
 *   - onOpen:      Called after a workbook is opened/created, before the command runs
 *   - preSave:     Called before each save() operation
 *   - postSave:    Called after each save() operation
 *   - preCommand:  Called before each CLI command (no workbook context)
 *   - postCommand: Called after each CLI command (no workbook context)
 *
 * Every workbook hook receives a HookContext with both command info and workbook access:
 *   ctx.command   — "get", "set", "eval", etc.
 *   ctx.args      — remaining CLI args after the command
 *   ctx.filePath  — resolved path to the workbook file
 *   ctx.file      — ExcelFile instance
 *   ctx.workbook  — SpreadJS Workbook instance
 *   ctx.GC        — GC.Spread.Sheets namespace
 *
 * Hook files export a default function that receives a HookAPI instance:
 *
 *   // .headless-spreadjs/hooks/defaults.ts
 *   import type { HookAPI, HookContext } from "@hewliyang/headless-spreadjs/hooks";
 *
 *   export default function (hsx: HookAPI) {
 *     hsx.on("onOpen", (ctx: HookContext) => {
 *       for (let i = 0; i < ctx.workbook.getSheetCount(); i++) {
 *         const sheet = ctx.workbook.getSheet(i);
 *         sheet.options.gridline = {
 *           showVerticalGridline: false,
 *           showHorizontalGridline: false,
 *         };
 *       }
 *     });
 *
 *     hsx.on("preSave", (ctx: HookContext) => {
 *       const sheet = ctx.workbook.getActiveSheet();
 *       const rowCount = sheet.getUsedRange().rowCount;
 *       for (let r = 0; r < rowCount; r++) {
 *         const val = sheet.getValue(r, 0);
 *         if (typeof val === "number" && val < 0) {
 *           sheet.getCell(r, 0).foreColor("red");
 *         }
 *       }
 *       console.log(`Formatted ${ctx.filePath} (command: ${ctx.command})`);
 *     });
 *   }
 *
 * Output Configuration:
 *   Hook output is captured and prefixed. By default, output goes to stderr.
 *
 *   hsx.on("preSave", { output: "stdout" }, (ctx) => { ... });
 *   hsx.on("preSave", { output: "none" }, (ctx) => { ... });
 */

import { existsSync, readdirSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { createJiti } from "jiti";
import type { ExcelFile } from "./excel-file.js";
import type { GCNamespace, SpreadWorkbook } from "./types.js";

// ============================================================================
// Context
// ============================================================================

export interface MutatedRange {
  /** Sheet name (undefined = active sheet at time of mutation) */
  sheet?: string;
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

export interface HookContext {
  /** CLI command name (e.g. "get", "set", "eval", "create") */
  command: string;
  /** CLI args after the command */
  args: string[];
  /** Resolved file path */
  filePath: string;
  /** ExcelFile wrapper */
  file: ExcelFile;
  /** SpreadJS Workbook instance — full API for reading/writing cells, styles, etc. */
  workbook: SpreadWorkbook;
  /** GC.Spread.Sheets namespace */
  GC: GCNamespace;
  /** Ranges mutated by the current command (empty for read-only commands or eval) */
  mutatedRanges: MutatedRange[];
}

// ============================================================================
// Hook Function Types
// ============================================================================

export type HookOutput = "stderr" | "stdout" | "none";

export interface PreCommandContext {
  /** CLI command name (e.g. "get", "set", "eval", "create") */
  command: string;
  /** CLI args after the command */
  args: string[];
}

export interface PostCommandContext extends PreCommandContext {
  /** Error thrown during command execution, if any */
  error?: Error;
}

export type PreCommandFn = (ctx: PreCommandContext) => void | Promise<void>;
export type PostCommandFn = (ctx: PostCommandContext) => void | Promise<void>;
export type OnOpenFn = (ctx: HookContext) => void | Promise<void>;
export type PreSaveFn = (ctx: HookContext) => void | Promise<void>;
export type PostSaveFn = (ctx: HookContext) => void | Promise<void>;

export interface HookEntry<T> {
  fn: T;
  output: HookOutput;
}

interface HookOptions {
  output?: HookOutput;
}

// ============================================================================
// Hook event map (for type-safe hsx.on())
// ============================================================================

interface HookEventMap {
  preCommand: PreCommandFn;
  postCommand: PostCommandFn;
  onOpen: OnOpenFn;
  preSave: PreSaveFn;
  postSave: PostSaveFn;
}

type HookEvent = keyof HookEventMap;

// ============================================================================
// Registry
// ============================================================================

export interface HookRegistry {
  preCommand: HookEntry<PreCommandFn>[];
  postCommand: HookEntry<PostCommandFn>[];
  onOpen: HookEntry<OnOpenFn>[];
  preSave: HookEntry<PreSaveFn>[];
  postSave: HookEntry<PostSaveFn>[];
  _discovered: boolean;
  _discoveryErrors: string[];
}

const _registry: HookRegistry = {
  preCommand: [],
  postCommand: [],
  onOpen: [],
  preSave: [],
  postSave: [],
  _discovered: false,
  _discoveryErrors: [],
};

// ============================================================================
// HookAPI — passed to hook files
// ============================================================================

export interface HookAPI {
  on<E extends HookEvent>(event: E, fn: HookEventMap[E]): void;
  on<E extends HookEvent>(
    event: E,
    opts: HookOptions,
    fn: HookEventMap[E],
  ): void;
}

function _createHookAPI(registry: HookRegistry): HookAPI {
  return {
    on<E extends HookEvent>(
      event: E,
      fnOrOpts: HookEventMap[E] | HookOptions,
      maybeFn?: HookEventMap[E],
    ): void {
      const list = registry[event] as HookEntry<HookEventMap[E]>[];
      if (typeof fnOrOpts === "function") {
        list.push({ fn: fnOrOpts as HookEventMap[E], output: "stderr" });
      } else {
        const opts = fnOrOpts as HookOptions;
        if (!maybeFn || typeof maybeFn !== "function") {
          throw new Error("hooks.on() requires a function argument.");
        }
        list.push({ fn: maybeFn, output: opts.output ?? "stderr" });
      }
    },
  };
}

/** Create a HookAPI instance for programmatic hook registration. */
export function createHookAPI(): HookAPI {
  return _createHookAPI(_registry);
}

// ============================================================================
// Output writers (injectable for CLI/daemon piping)
// ============================================================================

type WriteFn = (data: string) => void;

let _writeStdout: WriteFn = (data) => process.stdout.write(data);
let _writeStderr: WriteFn = (data) => process.stderr.write(data);

export function setHookWriters(stdout: WriteFn, stderr: WriteFn): void {
  _writeStdout = stdout;
  _writeStderr = stderr;
}

// ============================================================================
// Command context threading
// ============================================================================

interface CommandInfo {
  command: string;
  args: string[];
}

let _currentCommand: CommandInfo | null = null;

export function setCurrentCommand(info: CommandInfo | null): void {
  _currentCommand = info;
}

export function getCurrentCommand(): CommandInfo {
  return _currentCommand ?? { command: "", args: [] };
}

// ============================================================================
// Discovery
// ============================================================================

const HOOK_EXTENSIONS = new Set([".ts", ".js", ".mjs"]);

function isHookFile(name: string): boolean {
  if (name.startsWith("_")) return false;
  for (const ext of HOOK_EXTENSIONS) {
    if (name.endsWith(ext)) return true;
  }
  return false;
}

function getAliases(): Record<string, string> {
  const __dirname = dirname(fileURLToPath(import.meta.url));
  const distDir = __dirname;
  return {
    "@hewliyang/headless-spreadjs/hooks": join(distDir, "hooks.js"),
    "@hewliyang/headless-spreadjs": join(distDir, "index.js"),
  };
}

async function loadHookFile(filePath: string): Promise<void> {
  try {
    const jiti = createJiti(import.meta.url, {
      moduleCache: false,
      alias: getAliases(),
    });
    const mod = await jiti.import(filePath, { default: true });

    // Hook files export a default function that receives a HookAPI instance
    const factory =
      typeof mod === "function"
        ? mod
        : mod && typeof (mod as any).default === "function"
          ? (mod as any).default
          : null;

    if (factory) {
      const api = _createHookAPI(_registry);
      factory(api);
    }
  } catch (err) {
    const msg = `${filePath}: ${err instanceof Error ? err.message : String(err)}`;
    _registry._discoveryErrors.push(msg);
  }
}

async function loadHooksFromDir(dir: string): Promise<void> {
  if (!existsSync(dir)) return;

  let entries: string[];
  try {
    entries = readdirSync(dir).sort();
  } catch {
    return;
  }

  for (const entry of entries) {
    if (isHookFile(entry)) {
      await loadHookFile(join(dir, entry));
    }
  }
}

export async function discoverHooks(): Promise<void> {
  if (_registry._discovered) return;
  _registry._discovered = true;

  if (process.env.HSX_NO_HOOKS === "1") {
    return;
  }

  const localDir = resolve(".headless-spreadjs/hooks");
  if (existsSync(localDir)) {
    await loadHooksFromDir(localDir);
    return;
  }

  const globalDir = join(homedir(), ".headless-spreadjs", "hooks");
  if (existsSync(globalDir)) {
    await loadHooksFromDir(globalDir);
  }
}

// ============================================================================
// Hook Runners
// ============================================================================

function getHookName(fn: unknown): string {
  if (typeof fn === "function" && fn.name) return fn.name;
  return "(anonymous)";
}

async function runHookWithPrefix(
  fn: unknown,
  runFn: () => void | Promise<void>,
  hookType: string,
  output: HookOutput,
): Promise<void> {
  const hookName = getHookName(fn);
  const prefix = `[${hookType}:${hookName}]`;

  const origLog = console.log;
  const origWarn = console.warn;
  const origError = console.error;
  const captured: string[] = [];

  console.log = (...args: unknown[]) =>
    captured.push(args.map(String).join(" "));
  console.warn = (...args: unknown[]) =>
    captured.push(`[warn] ${args.map(String).join(" ")}`);
  console.error = (...args: unknown[]) =>
    captured.push(`[error] ${args.map(String).join(" ")}`);

  try {
    await runFn();
  } finally {
    console.log = origLog;
    console.warn = origWarn;
    console.error = origError;
  }

  if (output === "none" || captured.length === 0) return;

  const write = output === "stdout" ? _writeStdout : _writeStderr;
  for (const line of captured) {
    write(`${prefix} ${line}\n`);
  }
}

export async function getRegistry(): Promise<HookRegistry> {
  await discoverHooks();
  return _registry;
}

export async function runPreCommandHooks(
  ctx: PreCommandContext,
): Promise<void> {
  const registry = await getRegistry();
  for (const entry of registry.preCommand) {
    await runHookWithPrefix(
      entry.fn,
      () => entry.fn(ctx),
      "pre-command",
      entry.output,
    );
  }
}

export async function runPostCommandHooks(
  ctx: PostCommandContext,
): Promise<void> {
  const registry = await getRegistry();
  for (const entry of registry.postCommand) {
    await runHookWithPrefix(
      entry.fn,
      () => entry.fn(ctx),
      "post-command",
      entry.output,
    );
  }
}

export async function runOnOpenHooks(ctx: HookContext): Promise<void> {
  const registry = await getRegistry();
  for (const entry of registry.onOpen) {
    await runHookWithPrefix(
      entry.fn,
      () => entry.fn(ctx),
      "on-open",
      entry.output,
    );
  }
}

export async function runPreSaveHooks(ctx: HookContext): Promise<void> {
  const registry = await getRegistry();
  for (const entry of registry.preSave) {
    await runHookWithPrefix(
      entry.fn,
      () => entry.fn(ctx),
      "pre-save",
      entry.output,
    );
  }
}

export async function runPostSaveHooks(ctx: HookContext): Promise<void> {
  const registry = await getRegistry();
  for (const entry of registry.postSave) {
    await runHookWithPrefix(
      entry.fn,
      () => entry.fn(ctx),
      "post-save",
      entry.output,
    );
  }
}

// ============================================================================
// Utilities
// ============================================================================

export function clearHooks(): void {
  _registry.preCommand.length = 0;
  _registry.postCommand.length = 0;
  _registry.onOpen.length = 0;
  _registry.preSave.length = 0;
  _registry.postSave.length = 0;
  _registry._discovered = false;
  _registry._discoveryErrors.length = 0;
  _currentCommand = null;
}

export function getDiscoveryErrors(): string[] {
  return [..._registry._discoveryErrors];
}
