const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign
} = require('docx');

const PAGE_W = 12240, PAGE_H = 15840, MARGIN = 720;
const CONTENT_W = PAGE_W - MARGIN * 2;

const COLORS = {
  dark: "2B2B2B", gls: "2B2B2B", bk: "4A5A8A", admin: "B8825F", sm: "A15C8E",
  personal: "C2703D", home: "3F6B56", red: "D94F4F", redBg: "FDF2F2"
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
    margins: { top: 100, bottom: 100, left: 130, right: 130 },
    children
  });
}
function p(text, opts = {}) {
  return new Paragraph({
    spacing: { after: opts.after ?? 40 },
    alignment: opts.align,
    children: [new TextRun({ text, bold: opts.bold, italics: opts.italics, size: opts.size || 16, color: opts.color || "2B2B2B" })]
  });
}

const headerTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [6800, 4000],
  rows: [ new TableRow({ children: [
    cell([ new Paragraph({ children: [ new TextRun({ text: "QUARTER & MONTH PLAN", bold: true, size: 36 }) ] }) ], { width: 6800 }),
    cell([ p("Q3 2026 (JUL – SEP)", { align: AlignmentType.RIGHT, size: 18, color: "777777" }),
           p("REVISIT: START OF EACH MONTH", { align: AlignmentType.RIGHT, size: 18, color: "777777" }) ], { width: 4000 })
  ]}) ]
});

const yearStrip = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [CONTENT_W],
  rows: [ new TableRow({ children: [
    cell([ new Paragraph({ children: [
      new TextRun({ text: "THIS YEAR I'M WORKING TOWARD: ", bold: true, size: 16, color: "FFFFFF" }),
      new TextRun({ text: "________________________________________________________________________", size: 16, color: "FFFFFF" })
    ]}) ], { width: CONTENT_W, shading: COLORS.dark }) ]}) ]
});

const recoveryBanner = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [CONTENT_W],
  rows: [ new TableRow({ children: [
    cell([ new Paragraph({ children: [
      new TextRun({ text: "BARIATRIC SURGERY — JULY 29 (3–6 WK RECOVERY): ", bold: true, size: 15, color: COLORS.red }),
      new TextRun({ text: "July/August milestones below are intentionally light on Personal & Home. Recovering well is the milestone those months — everything else flexes around it.", size: 15, color: COLORS.red })
    ]}) ], { width: CONTENT_W, shading: COLORS.redBg, borders: boxBorders(COLORS.red, 8) }) ]}) ]
});

function planRow(name, color, goal, jul, aug, sep) {
  return new TableRow({ children: [
    cell([p(name, { bold: true, size: 15, color })], { width: 1400, borders: lineBorders() }),
    cell([p(goal, { size: 14 })], { width: 2600, borders: lineBorders() }),
    cell([p(jul, { size: 14 })], { width: 2270, borders: lineBorders() }),
    cell([p(aug, { size: 14 })], { width: 2270, borders: lineBorders() }),
    cell([p(sep, { size: 14 })], { width: 2260, borders: lineBorders() })
  ]});
}

const planTable = new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [1400, 2600, 2270, 2270, 2260],
  rows: [
    new TableRow({ children: [
      cell([p("CATEGORY", { bold: true, size: 13, color: "666666" })], { width: 1400, borders: boxBorders(COLORS.dark, 8) }),
      cell([p("QUARTER GOAL (JUL–SEP)", { bold: true, size: 13, color: "666666" })], { width: 2600, borders: boxBorders(COLORS.dark, 8) }),
      cell([p("JULY", { bold: true, size: 13, color: "999999" })], { width: 2270, borders: boxBorders(COLORS.dark, 8) }),
      cell([p("AUGUST", { bold: true, size: 13, color: "999999" })], { width: 2270, borders: boxBorders(COLORS.dark, 8) }),
      cell([p("SEPTEMBER", { bold: true, size: 13, color: "999999" })], { width: 2260, borders: boxBorders(COLORS.dark, 8) })
    ]}),
    planRow("GLS (Work)", COLORS.gls,
      "ACH transition fully completed; all SOPs & daily reporting caught up and consistent.",
      "Move $2k unapplied out before Jul 25. Begin daily ACH emails (50/day).",
      "Continue ACH emails. Catch up daily cash log + reports.",
      "ACH transition complete for all applicable customers. Finalize SOPs."),
    planRow("Book Keeping", COLORS.bk,
      "Backlog closed: Hub Reports (Jan–Jun) submitted; all SOPs written & usable by someone else.",
      "Complete Hub Reports Jan–Jun. Start learning Aplos.",
      "Learn PEX card system. Reimburse Whitney. Organize invoices/statements into binder.",
      "Submit reports. Build Sep–Dec counting schedule. Finalize SOPs."),
    planRow("Administration", COLORS.admin,
      "Recurring systems running (cleaning, supplies, phone/comms) without depending on memory.",
      "AC follow-up. Organize kitchen + cleaning supply tracking system. Set up water delivery.",
      "Create cleaning schedule (Juan/Whitney/Ikaika/Josh). Get church Google phone number.",
      "Systems running independently. Phone plan finalized."),
    planRow("Student Ministry", COLORS.sm,
      "Sep–Dec fully staffed (volunteer schedule + trained) and classroom basics operational before fall.",
      "Print lessons Jul/Aug. Clean & organize room. Plan July meetup.",
      "Build Sep–Dec volunteer schedule. Become a CORI. Plan August meetup.",
      "Volunteer training day. Check-in/out system live. September meetup."),
    planRow("Personal", COLORS.personal,
      "Recover well from surgery, finish school term, stabilize meds/routine system.",
      "Surgery Jul 29 + begin recovery. Sermon prep (Jul 19). School lessons continue.",
      "Recovery continues. Resume routines as cleared. Refill ADHD meds.",
      "Finish school (ends Sep 3). Stable daily medicine system in place."),
    planRow("Lopez Home", COLORS.home,
      "Family command center + daughter's visual routine set up; home organized at a recovery-safe pace.",
      "Set up Joelle's bed/swing. Organize storage rack. Clean bathroom/fridge (delegate if recovering).",
      "Family command center + daughter's visual chart. Organize kitchen cabinet.",
      "Clear apartment for remodel. Donate clothes. Home fully organized.")
  ]
});

function twoColBoxes(titleL, titleR) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [5400, 5400],
    rows: [ new TableRow({ children: [
      cell([p(titleL, { bold: true, size: 16 }), p("", { after: 200 }), p("", { after: 200 }), p("")], { width: 5400, borders: boxBorders(COLORS.dark, 8) }),
      cell([p(titleR, { bold: true, size: 16 }), p("", { after: 200 }), p("", { after: 200 }), p("")], { width: 5400, borders: boxBorders(COLORS.dark, 8) })
    ]}) ]
  });
}
const bottomBoxes = twoColBoxes("END OF QUARTER — CARRY TO Q4", "BIGGEST WIN THIS QUARTER");

const doc = new Document({
  sections: [{
    properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } },
    children: [
      headerTable,
      new Paragraph({ spacing: { before: 150, after: 130 }, children: [] }),
      yearStrip,
      new Paragraph({ spacing: { after: 130 }, children: [] }),
      recoveryBanner,
      new Paragraph({ spacing: { after: 130 }, children: [] }),
      planTable,
      new Paragraph({ spacing: { after: 150 }, children: [] }),
      bottomBoxes,
      new Paragraph({ spacing: { before: 150 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "QUARTER & MONTH PLAN — FATHER · HUSBAND · PROVIDER", size: 14, color: "BBBBBB" })] })
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  require('fs').writeFileSync('/home/claude/quarter_month_plan.docx', buf);
  console.log('written');
});
