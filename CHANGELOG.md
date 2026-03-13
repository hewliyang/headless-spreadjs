# Changelog

## [Unreleased]

### Docs

- Updated `skills/spreadjs/SKILL.md` with modern Excel / dynamic-array formula guidance for `hsx` workflows.

## [0.0.8] - 2026-03-13

### Fixed

- **CLI:** `set` now expands newly created sheets before writing, preventing silent data loss on non-active sheets beyond SpreadJS' default `200 × 20` grid.
- **CLI:** `eval` and `range(ref)` now expand newly created target sheets before access, fixing silent drops and truncated results on off-grid writes.
- **CLI:** `sheet create` now initializes new worksheets at full size so later commands behave consistently with existing sheets.
- Added regression tests covering non-active sheet writes and `eval` access beyond `U201` / `V202`.

## [0.0.7] - 2026-03-04

### Features

- Enable dynamic-array formula support (`allowDynamicArray`) by default — unlocks FILTER, SORT, UNIQUE, SEQUENCE, VSTACK, HSTACK, MAP, REDUCE, and other spill-based functions.

## [0.0.6] - 2026-02-22

### Features

- **Daemon mode** — `hsx` now runs a background daemon by default to avoid re-initializing SpreadJS on every command. Auto-starts on first use, caches open workbooks (LRU), auto-exits after idle timeout. Commands: `hsx daemon start|stop|status|flush`.
- **CLI:** `hsx screenshot` — render sheets to PNG via Playwright (optional peer dependency).
- **CLI:** `hsx diff` — compare two workbooks cell-by-cell, with `--epsilon` for floating-point tolerance.
- **CLI:** `hsx deps` — show formula dependency graph for a cell reference.
- **CLI:** `hsx csv` now supports `--mode value|formula|both`, plus `--formulas` shorthand for formula-focused exports.
- **CLI:** `hsx eval` gains a `range(ref)` helper for easy A1-notation cell access in scripts.

### Fixed

- **CLI:** `set` now properly clears existing formulas when overwriting with plain values.
- **CLI:** `set` style bugs and performance improvements.
- **CLI:** Range expansion for set operations.
- **CLI:** CSV argument parsing and daemon auto-flush on interval.
- **CLI:** Screenshots now include column and row headers.

### Chores

- Moved Playwright to an optional peer dependency.
- Renamed `skill/` to `skills/` for CLI discovery conventions.
- Added `AGENTS.md` with project overview, structure, dev commands, and release workflow.

## [0.0.3] - 2026-02-15

### Features

- **CLI:** Added `--version` / `-v` flag to print the package version.

### Fixed

- **CLI:** Extracted fatal error handling into a shared `reportFatalError` helper, removing duplicated `uncaughtException` / `unhandledRejection` logic. Errors now propagate from `main()` via `.catch()` instead of a redundant inner `try/catch`.
- **CLI:** Long error messages are now truncated (with `...`) instead of being silently replaced. The generic "unsupported features" fallback message is only used when the error looks like a known SpreadJS canvas/shape/chart/form-control failure.
- **Shims:** Canvas dimensions are now synced — the backing node-canvas is recreated when the element's `width` or `height` changes, preventing stale context bugs.
- **Shims:** `toBlob` callback is now dispatched via `queueMicrotask` to match the browser-spec async timing.
- **Shims:** Removed blanket `as never` type casts in canvas shim; all canvas method calls now use precise overload-matching types from the `canvas` package.

## [0.0.2] - 2026-02-15

### Features

- **CLI (`hsx`)** — New command-line interface for working with Excel files directly from the terminal. Supports `get`, `set`, `clear`, `copy`, `eval`, `search`, `info`, `csv`, `create`, `sheet`, `resize`, `rows-cols`, and `objects` commands with A1-style range addressing and formatted table output.

### Chores

- Added AI skill definition (`skill/spreadjs/`) with TypeScript declarations for LLM-assisted spreadsheet workflows.
- Moved tests from `src/test/` to top-level `test/` directory with dedicated `tsconfig.json`.
- Added typecheck step for tests in CI.
- Cleaned up types and formatting.

## [0.0.1] - 2026-02-15

Initial release — headless Excel workbook engine for Node.js powered by SpreadJS with DOM shims. Supports workbook creation, I/O (xlsx/ssjson/csv), formula evaluation, and extensions (charts, pivot tables, shapes, slicers, sparklines).
