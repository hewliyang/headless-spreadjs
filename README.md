# headless-spreadjs

Headless Excel engine for Node.js using [SpreadJS](https://developer.mescius.com/spreadjs).
Runs without a browser or Excel.

## Features

- Read/write `.xlsx` and `.xlsm`
- High-fidelity formulas, styles, tables, charts, pivots
- SpreadJS JSON roundtrip (`toJSON` / `fromJSON`)
- CLI (`hsx`) for quick workbook ops

## System dependencies

`canvas` requires Cairo/Pango.

```bash
# macOS
brew install pkg-config cairo pango libpng jpeg giflib librsvg

# Debian/Ubuntu
sudo apt-get install build-essential libcairo2-dev libpango1.0-dev libjpeg-dev libgif-dev librsvg2-dev
```

## Install

```bash
npm install @hewliyang/headless-spreadjs
```

## SDK quick start

```js
import { init } from "@hewliyang/headless-spreadjs";

const { ExcelFile, dispose } = await init({ licenseKey: "xxx" });

const file = new ExcelFile();
const sheet = file.workbook.getActiveSheet();
sheet.setValue(0, 0, "Name");
sheet.setValue(0, 1, "Score");
sheet.setValue(1, 0, "Alice");
sheet.setValue(1, 1, 95);
await file.save("output.xlsx");

dispose();
```

## CLI quick start

```bash
hsx create scores.xlsx
hsx set scores.xlsx A1:B3 '[["Name","Score"],["Alice",95],["Bob",87]]'
hsx get scores.xlsx A1:B3
hsx csv scores.xlsx A1:B3
hsx search scores.xlsx "Alice"
hsx diff before.xlsx after.xlsx
hsx deps scores.xlsx Sheet1!A1
```

Run `hsx --help` for all commands.

## Daemon (CLI)

`hsx` uses a background daemon by default to avoid re-initializing SpreadJS on each command.

- Auto-starts on first use
- Caches open workbooks
- Auto-exits after 5 minutes idle
- Falls back to direct mode if daemon is unavailable

Useful commands:

```bash
hsx daemon start
hsx daemon status
hsx daemon flush
hsx daemon stop
hsx --no-daemon get file.xlsx A1
hsx --timeout 120 eval file.xlsx '/* script */'
```

Environment variables:

| Variable            | Default          | Purpose                      |
| ------------------- | ---------------- | ---------------------------- |
| `HSX_SOCKET_PATH`   | platform default | Custom daemon socket/pipe    |
| `HSX_CACHE_SIZE`    | `10`             | LRU workbook cache size      |
| `HSX_WRITE_THROUGH` | `0`              | Immediate writes when truthy |

## SDK API

### `init(options?)`

Initializes runtime and returns `{ GC, ExcelFile, dispose }`.

- `licenseKey?: string` — SpreadJS license key (omit for trial)

### `ExcelFile`

- `new ExcelFile()`
- `ExcelFile.open(path)`
- `ExcelFile.openFromBuffer(buf)`
- `file.save(path)`
- `file.saveToBuffer()`
- `file.toJSON()`
- `file.fromJSON(json)`
- `file.batch(fn)`
- `file.workbook` (raw `GC.Spread.Sheets.Workbook`)

### `GC`

Full SpreadJS namespace (`GC.Spread.Sheets.*`) for advanced APIs (styles, enums, tables, formatting, etc).

### `dispose()`

Closes the DOM shim runtime. Call when all workbook work is done.

## Concurrency notes

A process supports one active `init()`/`dispose()` lifecycle at a time.
Multiple `ExcelFile` instances are fine within that lifecycle.

For isolation, use child processes (not worker threads).

## Docker

```dockerfile
FROM node:20-slim
RUN apt-get update && apt-get install -y \
    build-essential libcairo2-dev libpango1.0-dev \
    libjpeg-dev libgif-dev librsvg2-dev \
    && rm -rf /var/lib/apt/lists/*
```

## License

MIT for this package.
SpreadJS requires a separate commercial license from [MESCIUS](https://developer.mescius.com/spreadjs).
