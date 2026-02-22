# AGENTS.md

## Project Overview

**headless-spreadjs** is a headless Excel workbook engine for Node.js powered by [SpreadJS](https://developer.mescius.com/spreadjs) with DOM shims. It provides a programmatic SDK and a CLI (`hsx`) for reading, writing, and manipulating `.xlsx`/`.xlsm` files without a browser or Excel. Published as `@hewliyang/headless-spreadjs` on npm.

## Tech Stack

- **Language**: TypeScript
- **Runtime**: Node.js ≥ 18
- **Build**: `tsc` (pure TypeScript compilation)
- **Test**: Vitest
- **Linter/Formatter**: Biome
- **Excel Engine**: SpreadJS (`@mescius/spread-sheets` + optional addons for charts, pivots, shapes, slicers)
- **DOM Shim**: `happy-dom` (provides browser globals for SpreadJS in Node)
- **Canvas**: `node-canvas` (Cairo-based, for rendering/screenshots)
- **Compression**: `fflate` (for xlsx I/O)
- **Optional Peer**: Playwright (for screenshot command)

## Project Structure

```
headless-spreadjs/
├── src/
│   ├── index.ts                  # SDK entry — init(), dispose(), exports
│   ├── excel-file.ts             # ExcelFile class (open, save, toJSON, batch)
│   ├── shims.ts                  # DOM shim install/dispose (happy-dom + canvas)
│   ├── types.ts                  # GCNamespace, SpreadWorkbook, SpreadWorksheet, etc.
│   ├── vendor.d.ts               # SpreadJS vendor type declarations
│   ├── spreadjs-patches.d.ts     # SpreadJS type patches
│   └── cli/
│       ├── index.ts              # CLI entry point (bin: hsx)
│       ├── main.ts               # Argument parsing, command dispatch
│       ├── dispatch.ts           # Command router
│       ├── context.ts            # CLI context (init, file open)
│       ├── a1.ts                 # A1-style range parser
│       ├── output.ts             # Table/output formatting
│       ├── styles.ts             # Cell style utilities
│       ├── daemon.ts             # Background daemon (avoid re-init)
│       ├── daemon-entry.ts       # Daemon process entry
│       ├── client.ts             # Daemon client
│       ├── file-cache.ts         # LRU workbook cache for daemon
│       ├── abort.ts              # Abort/signal handling
│       └── commands/
│           ├── get.ts            # Read cell ranges
│           ├── set.ts            # Write cell data
│           ├── clear.ts          # Clear cells
│           ├── copy.ts           # Copy ranges
│           ├── create.ts         # Create new workbooks
│           ├── csv.ts            # CSV export (value/formula/both)
│           ├── deps.ts           # Formula dependency graph
│           ├── diff.ts           # Workbook diff
│           ├── eval.ts           # Evaluate JS against workbook
│           ├── info.ts           # Workbook/sheet info
│           ├── objects.ts        # List charts/tables/pivots
│           ├── resize.ts         # Resize ranges
│           ├── rows-cols.ts      # Row/column operations
│           ├── screenshot.ts     # Sheet screenshot (via Playwright)
│           ├── search.ts         # Search sheet data
│           └── sheet.ts          # Sheet management
├── test/
│   ├── cli.test.ts               # CLI integration tests
│   ├── daemon.test.ts            # Daemon tests
│   ├── lifecycle.test.ts         # init/dispose lifecycle tests
│   ├── workbook-core.test.ts     # Core workbook operations
│   ├── extensions/               # Extension-specific tests
│   ├── helpers.ts                # Test utilities
│   └── tsconfig.json             # Test-specific TS config
├── examples/
│   ├── basics.ts                 # Basic SDK usage
│   ├── charts.ts                 # Chart creation
│   ├── pivot-table.ts            # Pivot table creation
│   ├── shapes.ts                 # Shapes
│   ├── slicers.ts                # Slicers
│   ├── sparklines.ts             # Sparklines
│   ├── open-modify-save.ts       # Open, modify, save workflow
│   └── tsconfig.json             # Examples TS config
├── skills/
│   └── spreadjs/                 # AI skill definitions for LLM agents
├── CHANGELOG.md                  # Release changelog
├── biome.json                    # Biome linter/formatter config
├── tsconfig.json                 # TypeScript config
├── vitest.config.ts              # Vitest config
└── package.json
```

## Key Components

### SDK (`src/index.ts`, `src/excel-file.ts`)

- `init(options?)` — installs DOM shims, loads SpreadJS + optional addons, returns `{ GC, ExcelFile, dispose }`
- `ExcelFile` — open/save xlsx, buffer I/O, JSON roundtrip, `batch()` for deferred calc
- `dispose()` — tears down shims, resets state
- One active `init()`/`dispose()` lifecycle per process; multiple `ExcelFile` instances within

### CLI (`src/cli/`)

Binary: `hsx`. Commands: `get`, `set`, `clear`, `copy`, `create`, `csv`, `deps`, `diff`, `eval`, `info`, `objects`, `resize`, `rows-cols`, `screenshot`, `search`, `sheet`.

Uses a background daemon by default to avoid re-initializing SpreadJS per invocation. The daemon auto-starts, caches open workbooks (LRU), and auto-exits after idle timeout.

### DOM Shims (`src/shims.ts`)

Installs `happy-dom` globals and patches `HTMLCanvasElement` with `node-canvas` for SpreadJS to run headlessly in Node.js.

## Development Commands

```bash
pnpm install             # Install dependencies
pnpm build               # Compile TypeScript (tsc)
pnpm dev                 # Watch mode (tsc --watch)
pnpm typecheck           # Type check src + examples + tests
pnpm test                # Run tests (vitest, single worker)
pnpm test:watch          # Watch mode tests
pnpm lint                # Biome check
pnpm fmt                 # Biome auto-fix + format
```

## Code Style

- Formatter/linter: Biome (2-space indent, double quotes, semicolons, trailing commas)
- No JSDoc comments on functions
- Run `pnpm fmt` before committing

## Release Workflow

Releases are triggered by pushing a version tag. CI runs quality checks, publishes to npm, and creates a GitHub release with changelog.

### Steps

1. Update `CHANGELOG.md` — move `[Unreleased]` contents to a new `[x.y.z]` section, add fresh `[Unreleased]` header
2. Bump version and tag:
   ```bash
   pnpm version patch       # or minor/major — updates package.json, creates git tag
   git push && git push --tags
   ```
3. CI (`.github/workflows/release.yml`):
   1. Installs dependencies, builds
   2. Extracts changelog section for the tagged version from `CHANGELOG.md`
   3. Publishes to npm (`npm publish --access public`)
   4. Creates GitHub release with the extracted changelog

### CI (`.github/workflows/ci.yml`)

Runs on every push/PR to `main`: typecheck → test → lint → build.

## Testing

Tests use Vitest with `--maxWorkers=1` (SpreadJS requires single-threaded init/dispose). Test files live in `test/` with a dedicated `tsconfig.json`.

## System Dependencies

`node-canvas` requires Cairo/Pango native libraries:

```bash
# macOS
brew install pkg-config cairo pango libpng jpeg giflib librsvg

# Debian/Ubuntu
sudo apt-get install build-essential libcairo2-dev libpango1.0-dev libjpeg-dev libgif-dev librsvg2-dev
```

## References

- [SpreadJS Documentation](https://developer.mescius.com/spreadjs)
- [SpreadJS API Reference](https://developer.mescius.com/spreadjs/api)
- [node-canvas](https://github.com/Automattic/node-canvas)
- [happy-dom](https://github.com/nickvdyck/happy-dom)
