import { init } from "headless-spreadjs";

const { ExcelFile, GC, dispose } = await init();
const file = new ExcelFile();

const sheet = file.workbook.getActiveSheet();
sheet.name("Chart Data");

const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"];
const online = [12000, 15000, 13500, 17000, 19500, 22000];
const retail = [8000, 7500, 9000, 8500, 10000, 11500];

sheet.setValue(0, 0, "Month");
sheet.setValue(0, 1, "Online");
sheet.setValue(0, 2, "Retail");

months.forEach((month, i) => {
  const row = i + 1;
  sheet.setValue(row, 0, month);
  sheet.setValue(row, 1, online[i]);
  sheet.setValue(row, 2, retail[i]);
});

for (let r = 1; r <= 6; r++) {
  sheet.setFormatter(r, 1, "$#,##0");
  sheet.setFormatter(r, 2, "$#,##0");
}

sheet.setColumnWidth(0, 80);
sheet.setColumnWidth(1, 100);
sheet.setColumnWidth(2, 100);

const columnChart = sheet.charts.add(
  "RevenueComparison",
  GC.Spread.Sheets.Charts.ChartType.columnClustered,
  0,
  200,
  500,
  300,
  "A1:C7",
);
columnChart.title({ text: "Monthly Revenue: Online vs Retail" });
columnChart.legend({
  visible: true,
  position: GC.Spread.Sheets.Charts.LegendPosition.bottom,
});

const lineChart = sheet.charts.add(
  "OnlineTrend",
  GC.Spread.Sheets.Charts.ChartType.line,
  520,
  200,
  500,
  300,
  "A1:B7",
);
lineChart.title({ text: "Online Revenue Trend" });

const pieSheet = new GC.Spread.Sheets.Worksheet("Market Share");
file.workbook.addSheet(file.workbook.getSheetCount(), pieSheet);
pieSheet.setValue(0, 0, "Channel");
pieSheet.setValue(0, 1, "Total");
pieSheet.setValue(1, 0, "Online");
pieSheet.setFormula(1, 1, "SUM('Chart Data'!B2:B7)");
pieSheet.setValue(2, 0, "Retail");
pieSheet.setFormula(2, 1, "SUM('Chart Data'!C2:C7)");

const pieChart = pieSheet.charts.add(
  "MarketShare",
  GC.Spread.Sheets.Charts.ChartType.pie,
  0,
  80,
  450,
  350,
  "A1:B3",
);
pieChart.title({ text: "Revenue by Channel" });
pieChart.legend({
  visible: true,
  position: GC.Spread.Sheets.Charts.LegendPosition.right,
});

await file.save("examples/output/charts.xlsx");
console.log("Saved examples/output/charts.xlsx");

dispose();
