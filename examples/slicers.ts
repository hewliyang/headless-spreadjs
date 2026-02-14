import { init } from "headless-spreadjs";

const { Workbook, GC, dispose } = await init();
const wb = new Workbook();

const sheet = wb.getActiveSheet();
sheet.name("Sales Data");

// Build a data table
const headers = ["Region", "Product", "Quarter", "Revenue"];
headers.forEach((h, c) => {
  sheet.setValue(0, c, h);
});

const rows = [
  ["East", "Widget A", "Q1", 12000],
  ["East", "Widget A", "Q2", 15000],
  ["East", "Widget B", "Q1", 8000],
  ["East", "Widget B", "Q2", 9500],
  ["West", "Widget A", "Q1", 11000],
  ["West", "Widget A", "Q2", 13500],
  ["West", "Widget B", "Q1", 7500],
  ["West", "Widget B", "Q2", 10000],
  ["North", "Widget A", "Q1", 9000],
  ["North", "Widget A", "Q2", 11500],
  ["North", "Widget B", "Q1", 6000],
  ["North", "Widget B", "Q2", 7800],
  ["South", "Widget A", "Q1", 10500],
  ["South", "Widget A", "Q2", 12000],
  ["South", "Widget B", "Q1", 8200],
  ["South", "Widget B", "Q2", 9100],
];

rows.forEach((row, r) => {
  row.forEach((value, c) => {
    sheet.setValue(r + 1, c, value);
  });
});

// Format revenue
for (let r = 1; r <= rows.length; r++) {
  sheet.setFormatter(r, 3, "$#,##0");
}

// Create a table from the data range
const tableName = "SalesTable";
sheet.tables.add(
  tableName,
  0,
  0,
  rows.length + 1,
  headers.length,
  GC.Spread.Sheets.Tables.TableThemes.medium2,
);

// Add a slicer for the Region column
const regionSlicer = sheet.slicers.add("RegionSlicer", tableName, "Region");
regionSlicer.position(new GC.Spread.Sheets.Point(350, 10));
regionSlicer.width(180);
regionSlicer.height(200);
regionSlicer.style(GC.Spread.Sheets.Slicers.SlicerStyles.light1());

// Add a slicer for the Product column
const productSlicer = sheet.slicers.add(
  "ProductSlicer",
  tableName,
  "Product",
);
productSlicer.position(new GC.Spread.Sheets.Point(550, 10));
productSlicer.width(180);
productSlicer.height(150);
productSlicer.style(GC.Spread.Sheets.Slicers.SlicerStyles.light2());

sheet.setColumnWidth(0, 80);
sheet.setColumnWidth(1, 90);
sheet.setColumnWidth(2, 80);
sheet.setColumnWidth(3, 100);

await wb.save("examples/output/slicers.xlsx");
console.log("Saved examples/output/slicers.xlsx");

dispose();
