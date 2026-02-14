import { init } from "headless-spreadjs";

const { Workbook, GC, dispose } = await init();
const wb = new Workbook();

const dataSheet = wb.getActiveSheet();
dataSheet.name("Sales Data");

const headers = ["Region", "Product", "Quarter", "Sales", "Units"];
headers.forEach((value, col) => {
  dataSheet.setValue(0, col, value);
});

const records = [
  ["North", "Widget A", "Q1", 4500, 150],
  ["North", "Widget A", "Q2", 5200, 173],
  ["North", "Widget B", "Q1", 3100, 155],
  ["North", "Widget B", "Q2", 3800, 190],
  ["South", "Widget A", "Q1", 6100, 203],
  ["South", "Widget A", "Q2", 5800, 193],
  ["South", "Widget B", "Q1", 2900, 145],
  ["South", "Widget B", "Q2", 3300, 165],
  ["East", "Widget A", "Q1", 3900, 130],
  ["East", "Widget A", "Q2", 4100, 137],
  ["East", "Widget B", "Q1", 2200, 110],
  ["East", "Widget B", "Q2", 2600, 130],
  ["West", "Widget A", "Q1", 5500, 183],
  ["West", "Widget A", "Q2", 6200, 207],
  ["West", "Widget B", "Q1", 4000, 200],
  ["West", "Widget B", "Q2", 4500, 225],
];

records.forEach((row, r) => {
  row.forEach((value, c) => {
    dataSheet.setValue(r + 1, c, value);
  });
});

for (let r = 1; r <= records.length; r++) {
  dataSheet.setFormatter(r, 3, "$#,##0");
}

[80, 80, 70, 80, 60].forEach((width, col) => {
  dataSheet.setColumnWidth(col, width);
});

const pivotSheet = wb.addSheet("Pivot Analysis");
const dataRange = `'Sales Data'!A1:E${records.length + 1}`;

const pivotTable = pivotSheet.pivotTables.add(
  "SalesPivot",
  dataRange,
  1,
  0,
  GC.Spread.Pivot.PivotTableLayoutType.outline,
);

pivotTable.add(
  "Region",
  "Region",
  GC.Spread.Pivot.PivotTableFieldType.rowField,
);
pivotTable.add(
  "Product",
  "Product",
  GC.Spread.Pivot.PivotTableFieldType.rowField,
);
pivotTable.add(
  "Quarter",
  "Quarter",
  GC.Spread.Pivot.PivotTableFieldType.columnField,
);
pivotTable.add(
  "TotalSales",
  "Sales",
  GC.Spread.Pivot.PivotTableFieldType.valueField,
);
pivotTable.add(
  "TotalUnits",
  "Units",
  GC.Spread.Pivot.PivotTableFieldType.valueField,
);

const json = wb.toJSON();
const wb2 = new Workbook();
wb2.fromJSON(json);
console.log(
  `JSON roundtrip: pivot table count = ${wb2.getSheet(1).pivotTables.all().length}`,
);

await wb.save("examples/output/pivot-table.xlsx");
console.log("Saved examples/output/pivot-table.xlsx");

dispose();
