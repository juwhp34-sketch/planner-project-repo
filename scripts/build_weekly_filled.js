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

function buildWeekly(cfg, outPath) {
  const headerTable = new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [6800, 4000],
    rows: [ new TableRow({ children: [
      cell([ new Paragraph({ children: [ new TextRun({ text: "WEEKLY REVIEW", bold: true, size: 40 }) ] }) ], { width: 6800 }),
      cell([ p(`WEEK OF: ${cfg.weekOf}`, { align: AlignmentType.RIGHT, size: 20, color: "777777" }),
             p("QUARTER: JUL – SEP 2026", { align: AlignmentType.RIGHT, size: 20, color: "777777" }) ], { width: 4000 })
    ]}) ]
  });

  const tagLine = cfg.tag ? new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W],
    rows: [ new TableRow({ children: [ cell([ p(cfg.tag, { bold: true, size: 15, color: COLORS.red, align: AlignmentType.CENTER }) ], { width: CONTENT_W, shading: COLORS.redBg }) ] }) ]
  }) : null;

  const moveStrip = new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [ new TableRow({ children: [
      cell([ new Paragraph({ children: [
        new TextRun({ text: "THIS WEEK SHOULD MOVE: ", bold: true, size: 18, color: "FFFFFF" }),
        new TextRun({ text: cfg.shouldMove, size: 18, color: "FFFFFF" })
      ]}) ], { width: CONTENT_W, shading: COLORS.dark }) ]}) ]
  });

  function catRow(name, color, task, boxCount) {
    return new TableRow({ children: [
      cell([p(name, { bold: true, size: 18, color })], { width: 1700, borders: lineBorders() }),
      cell([p(task, { size: 15 })], { width: 4800, borders: lineBorders() }),
      cell([p(new Array(boxCount).fill("☐").join(" "), { size: 16 })], { width: 2100, borders: lineBorders() }),
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
      catRow("GLS (Work)", COLORS.dark, cfg.gls, cfg.glsBoxes),
      catRow("Book Keeping", COLORS.bk, cfg.bk, cfg.bkBoxes),
      catRow("Administration", COLORS.admin, cfg.admin, cfg.adminBoxes),
      catRow("Student Ministry", COLORS.sm, cfg.sm, cfg.smBoxes),
      catRow("Personal", COLORS.personal, cfg.personal, cfg.personalBoxes),
      catRow("Lopez Home", COLORS.green, cfg.home, cfg.homeBoxes)
    ]
  });

  function rotItem(day, cat, width, highlight) {
    return cell([p(day, { bold: true, size: 16, color: COLORS.green, align: AlignmentType.CENTER }), p(cat, { size: 14, align: AlignmentType.CENTER, bold: highlight })], { width });
  }
  const rotationTable = new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [1800, 1800, 1800, 1800, 1800, 1800],
    rows: [ new TableRow({ children: [
      rotItem("MON", "Book Keeping", 1800, cfg.hl === 'MON'), rotItem("TUE", "Admin", 1800, cfg.hl === 'TUE'), rotItem("WED", "Student Min.", 1800, cfg.hl === 'WED'),
      rotItem("THU", "Personal", 1800, cfg.hl === 'THU'), rotItem("FRI", "Lopez Home", 1800, cfg.hl === 'FRI'), rotItem("SAT/SUN", "Family", 1800, cfg.hl === 'SAT')
    ]}) ]
  });
  const rotationBoxed = new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W],
    rows: [ new TableRow({ children: [
      new TableCell({ width: { size: CONTENT_W, type: WidthType.DXA }, shading: { type: ShadingType.CLEAR, fill: cfg.rotationDim ? "F2F2F2" : COLORS.greenBg }, borders: boxBorders(cfg.rotationDim ? "AAAAAA" : COLORS.green, 10), margins: { top: 120, bottom: 120, left: 120, right: 120 }, children: [rotationTable] })
    ]}) ]
  });

  const recoveryNote = new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W],
    rows: [ new TableRow({ children: [
      cell([ new Paragraph({ children: [
        new TextRun({ text: cfg.recoveryLabel + " ", bold: true, size: 16, color: COLORS.red }),
        new TextRun({ text: cfg.recoveryText, size: 16, color: COLORS.red })
      ]}) ], { width: CONTENT_W, shading: COLORS.redBg, borders: boxBorders(COLORS.red, 8) }) ]}) ]
  });

  function twoColBoxes(titleL, linesL, titleR, linesR) {
    return new Table({
      width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [5400, 5400],
      rows: [ new TableRow({ children: [
        cell([p(titleL, { bold: true, size: 18 }), ...linesL.map(t => p(t, { size: 15, after: 140 }))], { width: 5400, borders: boxBorders(COLORS.dark, 8) }),
        cell([p(titleR, { bold: true, size: 18 }), ...linesR.map(t => p(t, { size: 15, after: 140 }))], { width: 5400, borders: boxBorders(COLORS.dark, 8) })
      ]}) ]
    });
  }
  const carryTable = twoColBoxes("CARRY-OVER FROM LAST WEEK", cfg.carryOver, "PUSH UP TO MONTH / QUARTER PLAN", cfg.pushUp);
  const energyTable = twoColBoxes("ENERGY CHECK — HOW WAS THIS WEEK?", ["Best day and why: ___________________________", "Hardest day and why: ________________________"], "ONE WIN THIS WEEK", ["", ""]);

  const children = [headerTable];
  if (tagLine) { children.push(new Paragraph({ spacing: { before: 100, after: 100 }, children: [] })); children.push(tagLine); }
  children.push(
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
  );

  const doc = new Document({
    sections: [{ properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } }, children }]
  });
  return Packer.toBuffer(doc).then(buf => { require('fs').writeFileSync(outPath, buf); console.log('written', outPath); });
}

// ---------- WEEK 1: Jul 13–19 ----------
const week1 = {
  weekOf: "07/13 – 07/19",
  shouldMove: "Sermon prep done before Sunday (7/19). Admin, Student Ministry & Home rotation tasks progressed. Daily ACH emails on track.",
  gls: "Daily ACH emails (50/day), Mon–Fri.", glsBoxes: 5,
  bk: "No action this week — full Book Keeping push scheduled for Mon 7/20.", bkBoxes: 0,
  admin: "AC unit follow-up call. Set up Costco water delivery (Thu). [Tue focus]", adminBoxes: 3,
  sm: "Print July/Aug/Sept lessons. Clean & organize room. [Wed focus]", smBoxes: 3,
  personal: "Sermon prep for 7/19 message — priority this week. [Thu focus]", personalBoxes: 3,
  home: "Set up Joelle's bed & bed swing. Organize storage rack. [Fri focus]", homeBoxes: 3,
  hl: 'TUE',
  recoveryLabel: "NOT A RECOVERY WEEK —",
  recoveryText: "business as usual. Sermon prep is the one non-negotiable this week since it's due before Sunday service.",
  carryOver: ["Google business phone number (Admin)", "Learn PEX card system (Book Keeping)", ""],
  pushUp: ["Volunteer training day — planned for August", "Move Whitney's dresser — push to post-recovery"],
};

// ---------- WEEK 2: Jul 20–26 ----------
const week2 = {
  weekOf: "07/20 – 07/26",
  shouldMove: "Close Book Keeping backlog + move the $2k before Jul 25 deadline — this is the last full week before surgery.",
  gls: "Move $2k unapplied funds (due Jul 25 — do Monday). Continue daily ACH emails (50/day).", glsBoxes: 5,
  bk: "Complete Hub Reports Jan–Jun. Start learning Aplos. Reimburse Whitney. Submit reports. [Mon — heavy day]", bkBoxes: 3,
  admin: "Get church Google phone number. Finalize phone plan. [Tue focus]", adminBoxes: 3,
  sm: "Plan August meetup. [Wed focus]", smBoxes: 3,
  personal: "Refill ADHD meds. Keep up with school lessons. [Thu focus]", personalBoxes: 3,
  home: "Clean bathroom/fridge. Start clearing apartment for remodel. [Fri focus]", homeBoxes: 3,
  hl: 'MON',
  recoveryLabel: "NOT A RECOVERY WEEK YET —",
  recoveryText: "last full week before surgery (Jul 29). Staff meeting Sunday 5–7:30pm this week, so Sunday evening is shorter than usual.",
  carryOver: ["Anything unfinished from 7/13–7/19 week", "", ""],
  pushUp: ["Move Whitney's dresser — push to post-recovery", "Volunteer training day — August"],
};

// ---------- WEEK 3: Jul 27–Aug 2 — RECOVERY ----------
const week3 = {
  weekOf: "07/27 – 08/02",
  tag: "RECOVERY MODE WEEK — SURGERY ON WEDNESDAY, JULY 29",
  shouldMove: "Get through surgery safely and rest. Nothing else is required this week.",
  gls: "GLS only if/when cleared by your doctor — do not assume you'll work this week.", glsBoxes: 5,
  bk: "Paused — recovery mode.", bkBoxes: 0,
  admin: "Paused — recovery mode.", adminBoxes: 0,
  sm: "Paused — recovery mode.", smBoxes: 0,
  personal: "Surgery Jul 29. Rest, follow post-op instructions, take medications as prescribed.", personalBoxes: 3,
  home: "Only if Whitney/family need something specific — otherwise paused.", homeBoxes: 0,
  hl: null, rotationDim: true,
  recoveryLabel: "THIS IS A RECOVERY WEEK —",
  recoveryText: "only GLS (if cleared) and Personal/health tasks count. Everything else auto-carries to next week without guilt — that's the plan working, not you failing.",
  carryOver: ["Everything from Book Keeping, Admin, Student Ministry, Lopez Home", "not finished before surgery — all of it waits.", ""],
  pushUp: ["Resume rotation once cleared by your doctor —", "target early-to-mid August depending on recovery."],
};

Promise.all([
  buildWeekly(week1, "/home/claude/weekly_review_Jul13-19.docx"),
  buildWeekly(week2, "/home/claude/weekly_review_Jul20-26.docx"),
  buildWeekly(week3, "/home/claude/weekly_review_Jul27-Aug02_RECOVERY.docx"),
]).then(() => console.log("ALL DONE"));
