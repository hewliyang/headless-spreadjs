import { init } from "headless-spreadjs";

const { ExcelFile, GC, dispose } = await init();
const file = new ExcelFile();

const sheet = file.workbook.getActiveSheet();
sheet.name("Sparklines");

const headers = [
  "Rep",
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Trend",
  "Bars",
  "W/L",
];
const headerStyle = new GC.Spread.Sheets.Style();
headerStyle.font = "bold 12px Arial";
headerStyle.backColor = "#2E86AB";
headerStyle.foreColor = "#FFFFFF";
headers.forEach((value, col) => {
  sheet.setValue(0, col, value);
  sheet.setStyle(0, col, headerStyle);
});

const reps = [
  { name: "Alice", data: [4200, 3800, 5100, 4700, 6200, 5900] },
  { name: "Bob", data: [3100, 3600, 2900, 3200, 3800, 4100] },
  { name: "Carol", data: [5500, 4900, 5200, -300, 6100, 6800] },
  { name: "Dan", data: [2800, 3100, 2500, 2200, 2900, 3400] },
  { name: "Eve", data: [6000, 5400, 5800, 6300, 5100, 7200] },
];

reps.forEach((rep, i) => {
  const row = i + 1;
  sheet.setValue(row, 0, rep.name);
  rep.data.forEach((value, col) => {
    sheet.setValue(row, col + 1, value);
    sheet.setFormatter(row, col + 1, "$#,##0");
  });
});

sheet.setColumnWidth(0, 60);
for (let c = 1; c <= 6; c++) {
  sheet.setColumnWidth(c, 65);
}
sheet.setColumnWidth(7, 120);
sheet.setColumnWidth(8, 120);
sheet.setColumnWidth(9, 120);

const Sparklines = GC.Spread.Sheets.Sparklines;

const lineSettings = new Sparklines.SparklineSetting();
lineSettings.options.seriesColor = "#2E86AB";
lineSettings.options.lineWeight = 2;
lineSettings.options.showHigh = true;
lineSettings.options.showLow = true;
lineSettings.options.highMarkerColor = "#28A745";
lineSettings.options.lowMarkerColor = "#DC3545";
lineSettings.options.showMarkers = true;
lineSettings.options.markersColor = "#6C757D";

const columnSettings = new Sparklines.SparklineSetting();
columnSettings.options.seriesColor = "#17A2B8";
columnSettings.options.negativeColor = "#E8453C";
columnSettings.options.showNegative = true;

const winLossSettings = new Sparklines.SparklineSetting();
winLossSettings.options.seriesColor = "#28A745";
winLossSettings.options.negativeColor = "#DC3545";

reps.forEach((rep, i) => {
  const row = i + 1;
  const dataRange = new GC.Spread.Sheets.Range(row, 1, 1, 6);

  sheet.setSparkline(
    row,
    7,
    dataRange,
    Sparklines.DataOrientation.horizontal,
    Sparklines.SparklineType.line,
    lineSettings,
  );

  sheet.setSparkline(
    row,
    8,
    dataRange,
    Sparklines.DataOrientation.horizontal,
    Sparklines.SparklineType.column,
    columnSettings,
  );

  const helperStart = 11;
  rep.data.forEach((value, i2) => {
    sheet.setValue(row, helperStart + i2, value - 4000);
  });

  sheet.setSparkline(
    row,
    9,
    new GC.Spread.Sheets.Range(row, helperStart, 1, 6),
    Sparklines.DataOrientation.horizontal,
    Sparklines.SparklineType.winloss,
    winLossSettings,
  );
});

for (let row = 1; row <= reps.length; row++) {
  sheet.setRowHeight(row, 30);
}

await file.save("examples/output/sparklines.xlsx");
console.log("Saved examples/output/sparklines.xlsx");

dispose();
