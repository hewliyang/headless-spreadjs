# headless-spreadjs

Headless Excel workbook engine for Node.js — powered by [SpreadJS](https://developer.mescius.com/spreadjs) with DOM shims. No browser, no Excel, any platform including Linux.

## Features

- **Full Excel fidelity** — 500+ formula functions, charts, pivot tables, tables, cell styling, merging, number formatting
- **XLSX I/O** — read and write `.xlsx, .xlsm` files with full roundtrip support
- **JSON serialization** — `toJSON()` / `fromJSON()` preserves everything (formulas, styles, charts, pivots)

## System Dependencies

Cairo/Pango are required by the `canvas` npm package:

```bash
# macOS
brew install pkg-config cairo pango libpng jpeg giflib librsvg

# Linux (Debian/Ubuntu)
sudo apt-get install build-essential libcairo2-dev libpango1.0-dev libjpeg-dev libgif-dev librsvg2-dev
```

## Install

```bash
npm install headless-spreadjs
```

## Usage

```js
import { init } from "headless-spreadjs";

// Initialize (with optional SpreadJS license key)
// If you use this for commercial purposes, please get a license from the madlads!
const { ExcelFile, GC, dispose } = await init({ licenseKey: "xxx" });

// Create a workbook from scratch
const file = new ExcelFile();
const sheet = file.workbook.getActiveSheet();
sheet.setValue(0, 0, "Name");
sheet.setValue(0, 1, "Score");
sheet.setValue(1, 0, "Alice");
sheet.setValue(1, 1, 95);
sheet.setValue(2, 0, "Bob");
sheet.setValue(2, 1, 87);
sheet.setValue(3, 0, "Average");
sheet.setFormula(3, 1, "AVERAGE(B2:B3)");
await file.save("output.xlsx");

// Open an existing workbook
const file2 = await ExcelFile.open("input.xlsx");
const val = file2.workbook.getActiveSheet().getValue(0, 0);
console.log(val);

// Access raw SpreadJS API
const workbook = file.workbook; // GC.Spread.Sheets.Workbook instance

// Clean up when done
dispose();
```

## API

### `init(options?): { GC, ExcelFile, dispose }`

Initialize the headless runtime. Must be called before creating ExcelFiles.

| Option       | Type      | Description                                |
| ------------ | --------- | ------------------------------------------ |
| `licenseKey` | `string?` | SpreadJS license key. Omit for trial mode. |

### `ExcelFile`

| Method                              | Description                                    |
| ----------------------------------- | ---------------------------------------------- |
| `new ExcelFile()`                   | Create an empty workbook                       |
| `ExcelFile.open(path)`              | Open an xlsx file → `Promise<ExcelFile>`       |
| `ExcelFile.openFromBuffer(buf)`     | Open xlsx from a Buffer → `Promise<ExcelFile>` |
| `file.save(path)`                   | Save to xlsx file → `Promise<void>`            |
| `file.saveToBuffer()`               | Save to Buffer → `Promise<Buffer>`             |
| `file.toJSON()`                     | Serialize to SpreadJS JSON                     |
| `file.fromJSON(json)`               | Load from SpreadJS JSON                        |
| `file.batch(fn)`                    | Suspend calc during `fn`, resume after         |
| `file.workbook`                     | Raw `GC.Spread.Sheets.Workbook` instance       |

### `dispose()`

Close the happy-dom window to prevent memory leaks. Call when done with all workbooks.

## Concurrency

`headless-spreadjs` installs DOM shims on `globalThis` (e.g. `window`, `document`, `navigator`), so a single Node.js process supports **one `init()` / `dispose()` lifecycle at a time**. You cannot safely run multiple workbook operations concurrently within the same process.

If you need parallelism, use separate processes (e.g. worker threads or child processes) — each gets its own global scope so the shims don't collide.

## Docker

`headless-spreadjs` will run in an image as simple as follows

```dockerfile
FROM node:20-slim
RUN apt-get update && apt-get install -y \
    build-essential libcairo2-dev libpango1.0-dev \
    libjpeg-dev libgif-dev librsvg2-dev \
    && rm -rf /var/lib/apt/lists/*
```

## License

MIT (this package). SpreadJS requires a separate commercial license from [MESCIUS](https://developer.mescius.com/spreadjs).
