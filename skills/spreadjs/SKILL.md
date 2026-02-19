---
name: spreadjs
description: "Creating & editing Excel workbooks via CLI — powered by SpreadJS"
---

# hsx CLI

## Setup

Before using `hsx`, check it's available. If not, install the package and its system dependencies (Cairo/Pango for `canvas`):

```bash
which hsx 2>/dev/null || npm install -g @hewliyang/headless-spreadjs
```

## Commands

All operations are bash calls. Formulas recalculate instantly (no sync step needed).

All commands support a global timeout option: `--timeout <seconds>` (default: `30`).

```bash
hsx create file.xlsx                              # new workbook
hsx info file.xlsx                                # sheets, used ranges
hsx get file.xlsx "Sheet1!A1:C10"                 # read cells → JSON
hsx csv file.xlsx "Sheet1!A1:C10"                 # read cells → CSV
hsx set file.xlsx "Sheet1!A1:C3" '<json>'         # write cells
hsx clear file.xlsx "A1:C10"                      # clear values
hsx search file.xlsx "term" [--regex] [--sheet S] # find values
hsx copy file.xlsx "A1:C1" "A10:C10"              # copy range
hsx sheet file.xlsx list|create|delete|rename      # manage sheets
hsx rc file.xlsx insert|delete|hide|freeze rows|columns  # row/col ops
hsx resize file.xlsx --columns A:D --width 120    # column/row sizing
hsx objects file.xlsx                             # list charts, tables
hsx eval file.xlsx "code"                         # arbitrary JS
hsx --timeout 120 eval file.xlsx "code"           # override timeout to 120s
```

References use A1 notation: `Sheet1!A1:C10`, `A1:C10` (active sheet), or `A1` (single cell).

## Write

```bash
hsx set file.xlsx "A1:B3" '[
  [{"value":"Name","cellStyles":{"fontWeight":"bold"}}, {"value":"Qty","cellStyles":{"fontWeight":"bold"}}],
  [{"value":"Alice"}, {"value":4}],
  [{"value":"Bob"}, {"formula":"=B2+1"}]
]'
```

Each cell: `{"value": ...}`, `{"formula": "=..."}`, optional `"cellStyles": {...}`.

cellStyles: `fontWeight` (bold/normal), `fontStyle` (italic/normal), `fontSize`, `fontFamily`, `fontColor`, `backgroundColor`, `horizontalAlignment` (left/center/right), `numberFormat`.

## Read

```bash
hsx get file.xlsx "A1:B3"
# → {"cells":{"A1":{"value":"Name","styles":{"bold":true}},"B3":{"value":5,"formula":"B2+1"}}, ...}

hsx csv file.xlsx "A1:B3"
# → Name,Qty
#    Alice,4
#    Bob,5
```

## Search

```bash
hsx search file.xlsx "Alice"                    # case-insensitive across all sheets
hsx search file.xlsx "^Q[1-4]$" --regex         # regex
hsx search file.xlsx "Revenue" --sheet Summary   # single sheet
```

## Copy

```bash
hsx copy file.xlsx "A1:C1" "A1:C10"             # repeat header pattern down
hsx copy file.xlsx "Sheet1!A1:B5" "Sheet2!A1:B5" # cross-sheet
```

## Sheets

```bash
hsx sheet file.xlsx list
hsx sheet file.xlsx create "Revenue"
hsx sheet file.xlsx rename "Sheet1" "Data"
hsx sheet file.xlsx delete "OldSheet"
```

## Rows & Columns

```bash
hsx rc file.xlsx insert rows --ref 5 --count 3    # insert 3 rows at row 5
hsx rc file.xlsx delete columns --ref B --count 2  # delete columns B-C
hsx rc file.xlsx hide rows --ref 10                # hide row 10
hsx rc file.xlsx freeze rows --ref 1               # freeze top row
hsx rc file.xlsx unfreeze rows                     # unfreeze
```

## Resize

```bash
hsx resize file.xlsx --columns A:D --width 120
hsx resize file.xlsx --rows 1:1 --height 30
hsx resize file.xlsx --columns A --width 200 --sheet Revenue
```

## Eval (escape hatch)

For charts, sparklines, pivot tables, conditional formatting — anything the structured commands don't cover.

```bash
hsx eval file.xlsx "
  // Globals: workbook, sheet (active), GC, file
  sheet.charts.add('Revenue', GC.Spread.Sheets.Charts.ChartType.line,
    0, 120, 500, 300, 'A1:E6');
"
```

```bash
hsx eval file.xlsx "
  const data = workbook.getSheetFromName('Data');
  data.tables.add('SalesTable', 0, 0, 10, 4,
    GC.Spread.Sheets.Tables.TableThemes.medium2);
"
```

```bash
# Read values back
hsx eval file.xlsx "
  console.log('Sheets:', workbook.getSheetCount());
  return sheet.getValue(0, 0);
"
```

Stdin also works: `echo "return sheet.getValue(0,0)" | hsx eval file.xlsx`

## SpreadJS API Reference

When using `hsx eval`, you have access to the full SpreadJS API via the `GC` namespace. If you need to look up a specific API (enum values, method signatures, class properties), grep the type definitions:

```bash
# Find an enum or class
grep -n "enum ChartType" ./spreadjs.d.ts
grep -n "class Worksheet" ./spreadjs.d.ts

# Find methods on a class (read nearby lines for signatures)
grep -n "addRows\|deleteRows\|addColumns" ./spreadjs.d.ts

# Find available style properties
grep -n "fontWeight:\|foreColor:\|backColor:\|formatter:" ./spreadjs.d.ts

# Find conditional formatting APIs
grep -n "conditionalFormats\|ConditionalFormatting" ./spreadjs.d.ts
```

The file is at `./spreadjs.d.ts`. Use `grep -n` to find what you need, then `read` with offset to see full signatures and JSDoc examples.

## Best Practices

- Use formulas, not hardcoded computed values: `{"formula":"=SUM(A1:A9)"}` not `{"value":45}`
- Keep each `set` call focused; build incrementally across multiple calls
- Use `hsx get` or `hsx csv` to verify after writes
- Prefer uniform column widths; use empty columns for indentation
- Always specify units in headers: `Revenue ($mm)`, `Growth (%)`
