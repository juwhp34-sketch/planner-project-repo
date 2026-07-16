const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, HeadingLevel, AlignmentType, VerticalAlign
} = require('docx');

const PAGE_W = 12240, PAGE_H = 15840, MARGIN = 720;
const CONTENT_W = PAGE_W - MARGIN * 2; // 10800

const COLORS = {
  dark: "2B2B2B", green: "3F6B56", greenBg: "F0F5F2",
  tan: "B8825F", tanBg: "FFF9F2", blue: "4A5A8A", blueBg: "F0F1F8",
  grey: "999999", greyBg: "F7F7F5", gold: "D9A441", red: "D94F4F"
};

function noBorders() {
  const b = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
  return { top: b, bottom: b, left: b, right: b };
}
function boxBorders(color, size = 8) {
  const b = { style: BorderStyle.SINGLE, size, color };
  return { top: b, bottom: b, left: b, right: b };
}
function cell(children, { width, shading, borders, valign } = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: shading ? { type: ShadingType.CLEAR, fill: shading } : undefined,
    borders: borders || noBorders(),
    verticalAlign: valign || VerticalAlign.TOP,
    margins: { top: 120, bottom: 120, left: 140, right: 140 },
    children
  });
}
function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.after ?? 80 },
    alignment: opts.align,
    children: [new TextRun({ text, bold: opts.bold, italics: opts.italics, size: opts.size || 20, color: opts.color || "2B2B2B" })]
  });
}
function blankLine(n = 1) {
  const arr = [];
  for (let i = 0; i < n; i++) arr.push(new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "" })] }));
  return arr;
}

// ---------- HEADER ----------
const headerTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [6800, 4000],
  rows: [
    new TableRow({
      children: [
        cell([
          new Paragraph({
            children: [
              new TextRun({ text: "I AM ", bold: true, size: 40 }),
              new TextRun({ text: "____________________________", size: 28 }),
              new TextRun({ text: " .", bold: true, size: 40 })
            ]
          })
        ], { width: 6800 }),
        cell([
          p("S   M   T   W   T   F   S   (circle today)", { align: AlignmentType.RIGHT, size: 18 }),
          p("DATE: ____ / ____ / ____", { align: AlignmentType.RIGHT, size: 20, color: "777777" }),
          p("☐ RECOVERY MODE — GLS + habits only, skip B/C", { align: AlignmentType.RIGHT, size: 18, bold: true, color: COLORS.red })
        ], { width: 4000 })
      ]
    })
  ]
});

const rotationPara = new Paragraph({
  spacing: { after: 200, before: 150 },
  border: { bottom: { style: BorderStyle.DOTTED, size: 4, color: "CCCCCC" } },
  children: [
    new TextRun({ text: "TODAY'S ROTATING FOCUS:  ", bold: true, size: 18, color: "555555" }),
    new TextRun({ text: "Mon: Book Keeping   |   Tue: Admin   |   Wed: Student Ministry   |   Thu: Personal   |   Fri: Lopez Home   |   Sat/Sun: Family", size: 18, color: "888888" })
  ]
});

// ---------- PRIORITIES + HABITS ----------
function priorityCell(letter, width) {
  return cell([
    p(letter, { bold: true, size: 26, color: COLORS.tan }),
    ...blankLine(3),
    p("if this slips...", { italics: true, size: 18, color: "999999" })
  ], { width, shading: COLORS.tanBg, borders: boxBorders(COLORS.tan, 6) });
}

const prioritiesTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [3200, 3200, 3200, 1200],
  rows: [
    new TableRow({
      children: [
        priorityCell("A", 3200), priorityCell("B", 3200), priorityCell("C", 3200),
        cell([new Paragraph("")], { width: 1200 })
      ]
    })
  ]
});

function habitRow(label, dual) {
  return new Paragraph({
    spacing: { after: 160 },
    children: [
      new TextRun({ text: dual ? "☐ ☐  " : "☐  ", size: 22, color: COLORS.green }),
      new TextRun({ text: label, size: 20 })
    ]
  });
}
const habitsBox = cell([
  p("DAILY NON-NEGOTIABLES", { bold: true, size: 20, color: COLORS.green }),
  habitRow("5-min walk"),
  habitRow("60oz water + 100g protein", true),
  habitRow("Bible time"),
  habitRow("Brush teeth (AM/PM)")
], { width: 3400, shading: COLORS.greenBg, borders: boxBorders(COLORS.green, 10) });

const topGrid = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [7400, 3400],
  rows: [ new TableRow({ children: [
    cell([prioritiesRowWrapper()], { width: 7400 }),
    habitsBox
  ]}) ]
});
function prioritiesRowWrapper() {
  // returns a nested table as a single "child" isn't allowed directly in cell array; wrap differently
  return new Paragraph({ children: [] });
}

// Simpler approach: build one big table with 4 columns: A, B, C, Habits
const mainTopTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [2467, 2467, 2466, 3400],
  rows: [
    new TableRow({
      children: [
        priorityCell("A", 2467), priorityCell("B", 2467), priorityCell("C", 2466), habitsBox
      ]
    })
  ]
});

// ---------- BRAIN DUMP + FEEL/ENERGY/FUMBLE ----------
const brainDumpCell = cell([
  p("BRAIN DUMP", { bold: true, size: 20 }),
  ...blankLine(11)
], { width: 6480, shading: undefined, borders: boxBorders(COLORS.dark, 10) });

const feelBox = cell([
  p("HOW I FEEL TODAY", { bold: true, size: 20, color: COLORS.blue }),
  p("☐ drained     ☐ rough     ☐ okay     ☐ good     ☐ great", { size: 18 })
], { width: 4320, shading: COLORS.blueBg, borders: boxBorders(COLORS.blue, 10) });

const energyBox = cell([
  p("ENERGY / FOCUS TANK", { bold: true, size: 20, color: COLORS.tan }),
  p("☐ ☐ ☐ ☐ ☐   (fill in as it drops)", { size: 18 })
], { width: 4320, shading: COLORS.tanBg, borders: boxBorders(COLORS.tan, 10) });

const fumbleBox = cell([
  p("FUMBLE TRIGGERS TODAY", { bold: true, size: 18, color: "666666" }),
  ...blankLine(2)
], { width: 4320, shading: COLORS.greyBg, borders: boxBorders(COLORS.grey, 8) });

// stack feel/energy/fumble vertically inside one cell using a nested table would be ideal,
// but simplest: three separate rows in the right column via a sub-table
const rightColTable = new Table({
  width: { size: 4320, type: WidthType.DXA },
  columnWidths: [4320],
  rows: [
    new TableRow({ children: [cell([p("HOW I FEEL TODAY", { bold: true, size: 20, color: COLORS.blue }), p("☐ drained   ☐ rough   ☐ okay   ☐ good   ☐ great", { size: 18 })], { width: 4320, shading: COLORS.blueBg, borders: boxBorders(COLORS.blue, 8) })] }),
    new TableRow({ children: [cell([p("ENERGY / FOCUS TANK", { bold: true, size: 20, color: COLORS.tan }), p("☐ ☐ ☐ ☐ ☐", { size: 20 })], { width: 4320, shading: COLORS.tanBg, borders: boxBorders(COLORS.tan, 8) })] }),
    new TableRow({ children: [cell([p("FUMBLE TRIGGERS TODAY", { bold: true, size: 18, color: "666666" }), ...blankLine(2)], { width: 4320, shading: COLORS.greyBg, borders: boxBorders(COLORS.grey, 8) })] })
  ]
});

const midGrid = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [6480, 4320],
  rows: [ new TableRow({ children: [
    cell([p("BRAIN DUMP", { bold: true, size: 20 }), ...blankLine(11)], { width: 6480, borders: boxBorders(COLORS.dark, 10) }),
    cell([nestedTablePlaceholder()], { width: 4320 })
  ]}) ]
});
function nestedTablePlaceholder(){ return new Paragraph({children:[]}); }

// docx-js DOES support nesting a Table inside a TableCell's children array. Rebuild midGrid properly:
const midGridFixed = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [6480, 4320],
  rows: [
    new TableRow({
      children: [
        cell([p("BRAIN DUMP", { bold: true, size: 20 }), ...blankLine(6)], { width: 6480, borders: boxBorders(COLORS.dark, 10) }),
        new TableCell({
          width: { size: 4320, type: WidthType.DXA },
          borders: noBorders(),
          margins: { top: 0, bottom: 0, left: 0, right: 0 },
          children: [rightColTable]
        })
      ]
    })
  ]
});

// ---------- TASK LANES ----------
function taskLaneCell(title, width) {
  return cell([
    p(title, { bold: true, size: 20 }),
    p("★ ________________________________", { size: 20 }),
    p("★ ________________________________", { size: 20 }),
    p("★ ________________________________", { size: 20 })
  ], { width, borders: boxBorders(COLORS.dark, 8) });
}
const taskTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [3600, 3600, 3600],
  rows: [ new TableRow({ children: [
    taskLaneCell("GLS (Work — daily)", 3600),
    taskLaneCell("Today's Focus: ______________", 3600),
    taskLaneCell("Home, Family & Me", 3600)
  ]}) ]
});

// ---------- BLOCK IT + WIN/GRATITUDE ----------
const hourHeaderCells = [];
const hours = new Array(16).fill("");
let blockRowCells = [];
hours.forEach(h => blockRowCells.push(cell([new Paragraph({ children: [new TextRun("")] }), new Paragraph({ children: [new TextRun("")] })], { width: Math.floor(6300/hours.length), borders: boxBorders("DDDDDD",4) })));
const blockItInner = new Table({
  width: { size: 6300, type: WidthType.DXA },
  columnWidths: hours.map(()=>Math.floor(6300/hours.length)),
  rows: [ new TableRow({ children: blockRowCells }) ]
});
const gwCell = cell([
  p("WIN + GRATITUDE", { bold: true, size: 20, color: COLORS.green }),
  p("Today I'm proud I...", { size: 18 }),
  ...blankLine(1),
  p("Grateful for...", { size: 18 }),
  ...blankLine(1)
], { width: 3960, shading: COLORS.greenBg, borders: boxBorders(COLORS.green, 8) });

const bottomTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [6840, 3960],
  rows: [ new TableRow({ children: [
    new TableCell({
      width: { size: 6840, type: WidthType.DXA },
      borders: boxBorders(COLORS.dark, 10),
      margins: { top: 120, bottom: 120, left: 140, right: 140 },
      children: [ p("BLOCK IT — TIME MAP (6a–10p)", { bold: true, size: 20 }), blockItInner ]
    }),
    gwCell
  ]}) ]
});

const doc = new Document({
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
    children: [
      headerTable,
      rotationPara,
      new Paragraph({ spacing: { after: 100 }, children: [] }),
      mainTopTable,
      new Paragraph({ spacing: { after: 100 }, children: [] }),
      midGridFixed,
      new Paragraph({ spacing: { after: 100 }, children: [] }),
      taskTable,
      new Paragraph({ spacing: { after: 100 }, children: [] }),
      bottomTable,
      new Paragraph({ spacing: { before: 100 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "DAILY PLANNER — FATHER · HUSBAND · PROVIDER", size: 16, color: "BBBBBB" })] })
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  require('fs').writeFileSync('/home/claude/daily_planner.docx', buf);
  console.log('written');
});
