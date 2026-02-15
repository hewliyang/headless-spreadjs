# Changelog

## [Unreleased]

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
