import { init } from "headless-spreadjs";

const { Workbook, GC, dispose } = await init();
const wb = new Workbook();

const sheet = wb.getActiveSheet();
sheet.name("Shapes");

// Add a title
sheet.addSpan(0, 0, 1, 6);
sheet.setValue(0, 0, "Shape Examples");
const titleStyle = new GC.Spread.Sheets.Style();
titleStyle.font = "bold 16px Arial";
titleStyle.hAlign = GC.Spread.Sheets.HorizontalAlign.center;
sheet.setStyle(0, 0, titleStyle);
sheet.setRowHeight(0, 40);

// Some data the shapes annotate
const data = [
  ["Status", "Count"],
  ["Complete", 42],
  ["In Progress", 18],
  ["Blocked", 5],
  ["Not Started", 12],
];
data.forEach((row, r) => {
  sheet.setValue(r + 1, 0, row[0]);
  sheet.setValue(r + 1, 1, row[1]);
});

const headerStyle = new GC.Spread.Sheets.Style();
headerStyle.font = "bold 12px Arial";
headerStyle.backColor = "#E8E8E8";
sheet.setStyle(1, 0, headerStyle);
sheet.setStyle(1, 1, headerStyle);

const { AutoShapeType } = GC.Spread.Sheets.Shapes;

// Rectangle with text
const rect = sheet.shapes.add(
  "callout",
  AutoShapeType.rectangle,
  200,
  60,
  220,
  80,
);
rect.text("42 tasks complete!\nGreat progress!");
rect.style({
  ...rect.style(),
  fill: { type: 0, color: "#D4EDDA" },
  line: { color: "#28A745", width: 2 },
  textEffect: { font: "12px Arial", color: "#155724" },
});

// Arrow pointing at the blocked row
const arrow = sheet.shapes.add(
  "alertArrow",
  AutoShapeType.rightArrow,
  200,
  170,
  80,
  30,
);
arrow.style({
  ...arrow.style(),
  fill: { type: 0, color: "#F8D7DA" },
  line: { color: "#DC3545", width: 2 },
});

// Oval highlight
const oval = sheet.shapes.add(
  "highlight",
  AutoShapeType.oval,
  200,
  220,
  160,
  60,
);
oval.text("Action needed!");
oval.style({
  ...oval.style(),
  fill: { type: 0, color: "#FFF3CD" },
  line: { color: "#FFC107", width: 2 },
  textEffect: { font: "bold 11px Arial", color: "#856404" },
});

sheet.setColumnWidth(0, 100);
sheet.setColumnWidth(1, 80);

await wb.save("examples/output/shapes.xlsx");
console.log("Saved examples/output/shapes.xlsx");

dispose();
