import { init } from "headless-spreadjs";

const { ExcelFile, dispose } = await init();

const file = await ExcelFile.open("examples/output/basics.xlsx");
const sheet = file.workbook.getActiveSheet();

console.log(`Opened workbook with ${file.workbook.getSheetCount()} sheets`);
console.log(`Active sheet: "${sheet.name()}"`);

console.log("\nCurrent data:");
for (let r = 1; r <= 6; r++) {
  const product = sheet.getValue(r, 0);
  const units = sheet.getValue(r, 1);
  const revenue = sheet.getText(r, 3);
  if (product) {
    console.log(`  ${product}: ${units} units, ${revenue}`);
  }
}

const newRow = 6;
const totalsRow = 7;
const averageRow = 8;

sheet.setValue(newRow, 0, "Widget E");
sheet.setValue(newRow, 1, 175);
sheet.setValue(newRow, 2, 39.99);
sheet.setFormula(newRow, 3, `B${newRow + 1}*C${newRow + 1}`);
sheet.setFormatter(newRow, 2, "$#,##0.00");
sheet.setFormatter(newRow, 3, "$#,##0.00");

sheet.setValue(totalsRow, 0, "TOTAL");
sheet.setFormula(totalsRow, 1, "SUM(B3:B7)");
sheet.setFormula(totalsRow, 3, "SUM(D3:D7)");
sheet.setFormatter(totalsRow, 3, "$#,##0.00");

sheet.setValue(averageRow, 0, "AVERAGE");
sheet.setFormula(averageRow, 2, "AVERAGE(C3:C7)");
sheet.setFormatter(averageRow, 2, "$#,##0.00");
sheet.setFormula(averageRow, 3, "AVERAGE(D3:D7)");
sheet.setFormatter(averageRow, 3, "$#,##0.00");

sheet.setValue(10, 0, "Last modified");
sheet.setValue(10, 1, new Date().toISOString().split("T")[0]);

const json = file.toJSON();
const file2 = new ExcelFile();
file2.fromJSON(json);
console.log(
  `\nJSON roundtrip check: "${file2.workbook.getActiveSheet().getValue(newRow, 0)}"`,
);

await file.save("examples/output/modified.xlsx");
console.log("Saved examples/output/modified.xlsx");

dispose();
