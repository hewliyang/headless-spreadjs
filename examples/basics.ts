import { init } from "headless-spreadjs";

const { ExcelFile, GC, dispose } = await init();
const file = new ExcelFile();

// batch() suspends formula recalculation until the end,
// avoiding redundant recomputation after every setValue/setFormula call.
file.batch(() => {
  const sheet = file.workbook.getActiveSheet();
  sheet.name("Sales Report");

  sheet.addSpan(0, 0, 1, 4);
  sheet.setValue(0, 0, "Q1 Sales Report");

  const titleStyle = new GC.Spread.Sheets.Style();
  titleStyle.font = "bold 16px Arial";
  titleStyle.foreColor = "#FFFFFF";
  titleStyle.backColor = "#2E86AB";
  titleStyle.hAlign = GC.Spread.Sheets.HorizontalAlign.center;
  sheet.setStyle(0, 0, titleStyle);
  sheet.setRowHeight(0, 40);

  const headers = ["Product", "Units", "Price", "Revenue"];
  const headerStyle = new GC.Spread.Sheets.Style();
  headerStyle.font = "bold 12px Arial";
  headerStyle.backColor = "#E8E8E8";
  headers.forEach((value, col) => {
    sheet.setValue(1, col, value);
    sheet.setStyle(1, col, headerStyle);
  });

  const rows = [
    ["Widget A", 150, 29.99],
    ["Widget B", 230, 19.99],
    ["Widget C", 85, 49.99],
    ["Widget D", 310, 9.99],
  ] as const;

  rows.forEach((row, i) => {
    const r = i + 2;
    sheet.setValue(r, 0, row[0]);
    sheet.setValue(r, 1, row[1]);
    sheet.setValue(r, 2, row[2]);
    sheet.setFormula(r, 3, `B${r + 1}*C${r + 1}`);
  });

  for (let r = 2; r <= 5; r++) {
    sheet.setFormatter(r, 2, "$#,##0.00");
    sheet.setFormatter(r, 3, "$#,##0.00");
  }

  const totalsRow = 6;
  sheet.setValue(totalsRow, 0, "TOTAL");
  sheet.setFormula(totalsRow, 1, "SUM(B3:B6)");
  sheet.setFormula(totalsRow, 3, "SUM(D3:D6)");
  sheet.setFormatter(totalsRow, 3, "$#,##0.00");

  const totalsStyle = new GC.Spread.Sheets.Style();
  totalsStyle.font = "bold 12px Arial";
  totalsStyle.backColor = "#D4EDDA";
  for (let c = 0; c < 4; c++) {
    sheet.setStyle(totalsRow, c, totalsStyle);
  }

  sheet.setValue(7, 0, "AVERAGE");
  sheet.setFormula(7, 2, "AVERAGE(C3:C6)");
  sheet.setFormatter(7, 2, "$#,##0.00");
  sheet.setFormula(7, 3, "AVERAGE(D3:D6)");
  sheet.setFormatter(7, 3, "$#,##0.00");

  sheet.setColumnWidth(0, 120);
  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 120);

  const summary = new GC.Spread.Sheets.Worksheet("Summary");
  file.workbook.addSheet(file.workbook.getSheetCount(), summary);
  summary.setValue(0, 0, "Total Revenue");
  summary.setFormula(0, 1, "'Sales Report'!D7");
  summary.setFormatter(0, 1, "$#,##0.00");
  summary.setValue(1, 0, "Product Count");
  summary.setFormula(1, 1, "COUNTA('Sales Report'!A3:A6)");
  summary.setValue(2, 0, "Max Revenue Product");
  summary.setFormula(
    2,
    1,
    "INDEX('Sales Report'!A3:A6,MATCH(MAX('Sales Report'!D3:D6),'Sales Report'!D3:D6,0))",
  );
});

await file.save("examples/output/basics.xlsx");
console.log("Saved examples/output/basics.xlsx");

dispose();
