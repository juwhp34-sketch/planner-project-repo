const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign
} = require('docx');

const PAGE_W = 12240, PAGE_H = 15840, MARGIN = 720;
const CONTENT_W = PAGE_W - MARGIN * 2;

const COLORS = {
  dark: "2B2B2B", green: "3F6B56", greenBg: "F0F5F2",
  admin: "B8825F", bk: "4A5A8A", sm: "A15C8E", personal: "C2703D", red: "D94F4F", redBg: "FDF2F2"
};

function noBorders() {
  const b = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  return { top: b, bottom: b, left: b, right: b };
}
function boxBorders(color, size = 8) {
  const b = { style: BorderStyle.SINGLE, size, color };
  return { top: b, bottom: b, left: b, right: b };
}
function lineBorders(color = "E3E0D8", size = 4) {
  const b = { style: BorderStyle.SINGLE, size, color };
  const none = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  return { top: none, bottom: b, left: none, right: none };
}
function cell(children, { width, shading, borders, valign } = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: shading ? { type: ShadingType.CLEAR, fill: shading } : undefined,
    borders: borders || noBorders(),
    verticalAlign: valign || VerticalAlign.TOP,
    margins: { top: 100, bottom: 100, left: 140, right: 140 },
    children
  });
}
function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.after ?? 60 },
    alignment: opts.align,
    children: [new TextRun({ text, bold: opts.bold, italics: opts.italics, size: opts.size || 18, color: opts.color || "2B2B2B" })]
  });
}

// ---------- HEADER ----------
const headerTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [6800, 4000],
  rows: [ new TableRow({ children: [
    cell([ new Paragraph({ children: [ new TextRun({ text: "WEEKLY REVIEW", bold: true, size: 40 }) ] }) ], { width: 6800 }),
    cell([ p("WEEK OF: ____ / ____ / ____", { align: AlignmentType.RIGHT, size: 20, color: "777777" }),
           p("QUARTER: JUL – SEP 2026", { align: AlignmentType.RIGHT, size: 20, color: "777777" }) ], { width: 4000 })
  ]}) ]
});

const moveStrip = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [CONTENT_W],
  rows: [ new TableRow({ children: [
    cell([ new Paragraph({ children: [
      new TextRun({ text: "THIS WEEK SHOULD MOVE: ", bold: true, size: 18, color: "FFFFFF" }),
      new TextRun({ text: "________________________________________________________________________", size: 18, color: "FFFFFF" })
    ]}) ], { width: CONTENT_W, shading: COLORS.dark }) ]}) ]
});

// ---------- CATEGORY TABLE ----------
function catRow(name, color, days) {
  return new TableRow({ children: [
    cell([p(name, { bold: true, size: 18, color })], { width: 1700, borders: lineBorders() }),
    cell([p("_____________________________________________", { size: 16 })], { width: 4800, borders: lineBorders() }),
    cell([p(new Array(days).fill("☐").join(" "), { size: 16 })], { width: 2100, borders: lineBorders() }),
    cell([p("☐ Done   ☐ Carry over", { size: 16 })], { width: 2200, borders: lineBorders() })
  ]});
}

const catTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [1700, 4800, 2100, 2200],
  rows: [
    new TableRow({ children: [
      cell([p("CATEGORY", { bold: true, size: 14, color: "666666" })], { width: 1700, borders: boxBorders(COLORS.dark, 10) }),
      cell([p("THIS WEEK'S TASK(S)", { bold: true, size: 14, color: "666666" })], { width: 4800, borders: boxBorders(COLORS.dark, 10) }),
      cell([p("DAYS TO TOUCH", { bold: true, size: 14, color: "666666" })], { width: 2100, borders: boxBorders(COLORS.dark, 10) }),
      cell([p("STATUS", { bold: true, size: 14, color: "666666" })], { width: 2200, borders: boxBorders(COLORS.dark, 10) })
    ]}),
    catRow("GLS (Work)", COLORS.dark, 5),
    catRow("Book Keeping", COLORS.bk, 3),
    catRow("Administration", COLORS.admin, 3),
    catRow("Student Ministry", COLORS.sm, 3),
    catRow("Personal", COLORS.personal, 3),
    catRow("Lopez Home", COLORS.green, 3)
  ]
});

// ---------- ROTATION ROW ----------
function rotItem(day, cat, width) {
  return cell([p(day, { bold: true, size: 16, color: COLORS.green, align: AlignmentType.CENTER }), p(cat, { size: 14, align: AlignmentType.CENTER })], { width });
}
const rotationTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [1800, 1800, 1800, 1800, 1800, 1800],
  rows: [ new TableRow({ children: [
    rotItem("MON", "Book Keeping", 1800), rotItem("TUE", "Admin", 1800), rotItem("WED", "Student Min.", 1800),
    rotItem("THU", "Personal", 1800), rotItem("FRI", "Lopez Home", 1800), rotItem("SAT/SUN", "Family", 1800)
  ]}) ]
});
// wrap rotation table in a shaded/bordered outer cell for the green box look
const rotationBoxed = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [CONTENT_W],
  rows: [ new TableRow({ children: [
    new TableCell({ width: { size: CONTENT_W, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, fill: COLORS.greenBg }, borders: boxBorders(COLORS.green, 10), margins: { top: 120, bottom: 120, left: 120, right: 120 }, children: [rotationTable] })
  ]}) ]
});

const recoveryNote = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [CONTENT_W],
  rows: [ new TableRow({ children: [
    cell([ new Paragraph({ children: [
      new TextRun({ text: "IF THIS IS A RECOVERY WEEK: ", bold: true, size: 16, color: COLORS.red }),
      new TextRun({ text: "Only GLS (if cleared to work) and Personal/health tasks count. Everything else auto-carries to next week without guilt — that's the plan working, not you failing.", size: 16, color: COLORS.red })
    ]}) ], { width: CONTENT_W, shading: COLORS.redBg, borders: boxBorders(COLORS.red, 8) }) ]}) ]
});

// ---------- TWO-COL BOXES ----------
function twoColBoxes(titleL, linesL, titleR, linesR) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [5400, 5400],
    rows: [ new TableRow({ children: [
      cell([p(titleL, { bold: true, size: 18 }), ...linesL.map(t => p(t, { size: 16, after: 160 }))], { width: 5400, borders: boxBorders(COLORS.dark, 8) }),
      cell([p(titleR, { bold: true, size: 18 }), ...linesR.map(t => p(t, { size: 16, after: 160 }))], { width: 5400, borders: boxBorders(COLORS.dark, 8) })
    ]}) ]
  });
}

const carryTable = twoColBoxes(
  "CARRY-OVER FROM LAST WEEK", ["", "", ""],
  "PUSH UP TO MONTH / QUARTER PLAN", ["", "", ""]
);
const energyTable = twoColBoxes(
  "ENERGY CHECK — HOW WAS THIS WEEK?", ["Best day and why: ___________________________", "Hardest day and why: ________________________"],
  "ONE WIN THIS WEEK", ["", ""]
);

const doc = new Document({
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
    children: [
      headerTable,
      new Paragraph({ spacing: { before: 150, after: 150 }, children: [] }),
      moveStrip,
      new Paragraph({ spacing: { after: 150 }, children: [] }),
      catTable,
      new Paragraph({ spacing: { after: 150 }, children: [] }),
      rotationBoxed,
      new Paragraph({ spacing: { after: 150 }, children: [] }),
      recoveryNote,
      new Paragraph({ spacing: { after: 150 }, children: [] }),
      carryTable,
      new Paragraph({ spacing: { after: 150 }, children: [] }),
      energyTable,
      new Paragraph({ spacing: { before: 150 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "WEEKLY REVIEW — FATHER · HUSBAND · PROVIDER", size: 16, color: "BBBBBB" })] })
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  require('fs').writeFileSync('/home/claude/weekly_review.docx', buf);
  console.log('written');
});
