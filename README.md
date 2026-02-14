# headless-spreadjs

Headless Excel workbook engine for Node.js — powered by [SpreadJS](https://developer.mescius.com/spreadjs) with DOM shims. No browser, no Excel, any platform including Linux.

## Features

- **Full Excel fidelity** — 500+ formula functions, charts, pivot tables, tables, cell styling, merging, number formatting
- **XLSX I/O** — read and write `.xlsx` files with full roundtrip support
- **JSON serialization** — `toJSON()` / `fromJSON()` preserves everything (formulas, styles, charts, pivots)
- **Zero config** — DOM shims auto-installed, just call `init()`
- **Direct SpreadJS access** — full API available via `GC` namespace and `wb.spread`

## System Dependencies

On Linux, install Cairo/Pango (required by the `canvas` npm package):

```bash
sudo apt-get install build-essential libcairo2-dev libpango1.0-dev libjpeg-dev libgif-dev librsvg2-dev
```

macOS: `brew install pkg-config cairo pango libpng jpeg giflib librsvg`

## Install

```bash
pnpm add headless-spreadjs
```

All SpreadJS packages (core, I/O, charts, pivot tables) are bundled as dependencies — no extra installs needed.

SpreadJS is a commercial library (~$2k license). Trial mode works without a key — an "Evaluation Version" watermark sheet is auto-stripped on import.

## Usage

```js
import { init } from "headless-spreadjs";

// Initialize (with optional license key)
const { Workbook, GC, dispose } = await init({ licenseKey: "xxx" });

// Create a workbook from scratch
const wb = new Workbook();
const sheet = wb.getActiveSheet();
sheet.setValue(0, 0, "Name");
sheet.setValue(0, 1, "Score");
sheet.setValue(1, 0, "Alice");
sheet.setValue(1, 1, 95);
sheet.setValue(2, 0, "Bob");
sheet.setValue(2, 1, 87);
sheet.setValue(3, 0, "Average");
sheet.setFormula(3, 1, "AVERAGE(B2:B3)");
await wb.save("output.xlsx");

// Open an existing workbook
const wb2 = await Workbook.open("input.xlsx");
const val = wb2.getActiveSheet().getValue(0, 0);
console.log(val);

// Access raw SpreadJS API
const spread = wb.spread; // GC.Spread.Sheets.Workbook instance

// Clean up when done
dispose();
```

## API

### `init(options?): { GC, Workbook, dispose }`

Initialize the headless runtime. Must be called before creating Workbooks.

| Option       | Type      | Description                                |
| ------------ | --------- | ------------------------------------------ |
| `licenseKey` | `string?` | SpreadJS license key. Omit for trial mode. |

### `Workbook`

| Method                         | Description                                   |
| ------------------------------ | --------------------------------------------- |
| `new Workbook()`               | Create an empty workbook                      |
| `Workbook.open(path)`          | Open an xlsx file → `Promise<Workbook>`       |
| `Workbook.openFromBuffer(buf)` | Open xlsx from a Buffer → `Promise<Workbook>` |
| `wb.save(path)`                | Save to xlsx file → `Promise<void>`           |
| `wb.saveToBuffer()`            | Save to Buffer → `Promise<Buffer>`            |
| `wb.toJSON()`                  | Serialize to SpreadJS JSON                    |
| `wb.fromJSON(json)`            | Load from SpreadJS JSON                       |
| `wb.getActiveSheet()`          | Get active worksheet                          |
| `wb.getSheet(index)`           | Get worksheet by index                        |
| `wb.getSheetCount()`           | Get number of sheets                          |
| `wb.addSheet(name, index?)`    | Add a new worksheet                           |
| `wb.spread`                    | Raw `GC.Spread.Sheets.Workbook` instance      |

### `dispose()`

Close the happy-dom window to prevent memory leaks. Call when done with all workbooks.

## Docker

`headless-spreadjs` will run in an image as simple as follows

```dockerfile
FROM node:20-slim
RUN apt-get update && apt-get install -y \
    build-essential libcairo2-dev libpango1.0-dev \
    libjpeg-dev libgif-dev librsvg2-dev \
    && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY package*.json ./
RUN npm ci
COPY . .
```

## License

MIT (this package). SpreadJS requires a separate commercial license from [MESCIUS](https://developer.mescius.com/spreadjs).
