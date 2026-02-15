# Changelog

## [Unreleased]

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
