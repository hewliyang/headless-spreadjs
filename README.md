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
npm install @hewliyang/headless-spreadjs
```

## SDK Usage

```js
import { init } from "@hewliyang/headless-spreadjs";

// Initialize (with optional SpreadJS license key)
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

// Clean up when done
dispose();
```

## CLI Usage

The package ships a CLI called `hsx` for quick Excel operations for agents:

```bash
# Create a new workbook
hsx create scores.xlsx

# Write cells from JSON
hsx set scores.xlsx A1:B3 '[["Name","Score"],["Alice",95],["Bob",87]]'

# Read cells back
hsx get scores.xlsx A1:B3

# Dump a range as CSV
hsx csv scores.xlsx A1:B3

# Search across sheets
hsx search scores.xlsx "Alice"

# Add a formula via eval (full SpreadJS API access)
hsx eval scores.xlsx '
  sheet.setValue(3, 0, "Average");
  sheet.setFormula(3, 1, "AVERAGE(B2:B3)");
  await file.save("scores.xlsx");
'
```

Run `hsx --help` for the full list of commands and options.

## SDK API

### `init(options?): { GC, ExcelFile, dispose }`

Initialize the headless runtime. Must be called before creating ExcelFiles.

| Option       | Type      | Description                                |
| ------------ | --------- | ------------------------------------------ |
| `licenseKey` | `string?` | SpreadJS license key. Omit for trial mode. |

### `ExcelFile`

| Method                          | Description                                    |
| ------------------------------- | ---------------------------------------------- |
| `new ExcelFile()`               | Create an empty workbook                       |
| `ExcelFile.open(path)`          | Open an xlsx file → `Promise<ExcelFile>`       |
| `ExcelFile.openFromBuffer(buf)` | Open xlsx from a Buffer → `Promise<ExcelFile>` |
| `file.save(path)`               | Save to xlsx file → `Promise<void>`            |
| `file.saveToBuffer()`           | Save to Buffer → `Promise<Buffer>`             |
| `file.toJSON()`                 | Serialize to SpreadJS JSON                     |
| `file.fromJSON(json)`           | Load from SpreadJS JSON                        |
| `file.batch(fn)`                | Suspend calc during `fn`, resume after         |
| `file.workbook`                 | Raw `GC.Spread.Sheets.Workbook` instance       |

### `GC` — The SpreadJS Namespace

The `GC` object returned by `init()` is the full `GC.Spread.Sheets` namespace from SpreadJS. You need it whenever you go beyond simple get/set — styling, enums, conditional formatting, tables, etc.

```js
const { ExcelFile, GC, dispose } = await init();
const file = new ExcelFile();
const sheet = file.workbook.getActiveSheet();

const style = new GC.Spread.Sheets.Style();
style.foreColor = "white";
style.backColor = "#2563eb";
style.font = "bold 12px sans-serif";
style.hAlign = GC.Spread.Sheets.HorizontalAlign.center;
sheet.setStyle(0, 0, style);

sheet.tables.add(
  "Table1",
  0,
  0,
  5,
  3,
  GC.Spread.Sheets.Tables.TableTheme.medium2,
);

sheet.setFormatter(1, 1, "$#,##0.00");

dispose();
```

The `GC` namespace maps to the [`@mescius/spread-sheets`](https://developer.mescius.com/spreadjs/api/modules/GC.Spread.Sheets) module. Refer to the [SpreadJS API docs](https://developer.mescius.com/spreadjs/api) for the full reference. Every class, enum, and type under `GC.Spread.Sheets.*` is available.

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
