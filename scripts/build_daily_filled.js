const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, ShadingType, AlignmentType, VerticalAlign
} = require('docx');

const PAGE_W = 12240, PAGE_H = 15840, MARGIN = 720;
const CONTENT_W = PAGE_W - MARGIN * 2;

const COLORS = {
  dark: "2B2B2B", green: "3F6B56", greenBg: "F0F5F2",
  tan: "B8825F", tanBg: "FFF9F2", blue: "4A5A8A", blueBg: "F0F1F8",
  grey: "999999", greyBg: "F7F7F5", gold: "D9A441", red: "D94F4F", redBg: "FDF2F2"
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

function buildDaily(cfg, outPath) {
  const dayLetters = ["S","M","T","W","T","F","S"];
  const dayRun = dayLetters.map((d, i) => new TextRun({ text: (i === cfg.dayIndex ? `[${d}]` : ` ${d} `), bold: i === cfg.dayIndex, size: 18, color: i === cfg.dayIndex ? COLORS.red : COLORS.dark }));

  const headerTable = new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [6800, 4000],
    rows: [ new TableRow({ children: [
      cell([ new Paragraph({ children: [
        new TextRun({ text: "I AM ", bold: true, size: 40 }),
        new TextRun({ text: "____________________________", size: 28 }),
        new TextRun({ text: " .", bold: true, size: 40 })
      ]}) ], { width: 6800 }),
      cell([
        new Paragraph({ alignment: AlignmentType.RIGHT, children: dayRun }),
        p(`DATE: ${cfg.date}`, { align: AlignmentType.RIGHT, size: 20, color: "777777" }),
        p(cfg.recoveryChecked ? "☒ RECOVERY MODE — GLS + habits only, skip B/C" : "☐ RECOVERY MODE — GLS + habits only, skip B/C",
          { align: AlignmentType.RIGHT, size: 18, bold: true, color: COLORS.red })
      ], { width: 4000 })
    ]}) ]
  });

  const rotationPara = new Paragraph({
    spacing: { after: 200, before: 150 },
    children: [
      new TextRun({ text: "TODAY'S ROTATING FOCUS:  ", bold: true, size: 18, color: "555555" }),
      new TextRun({ text: cfg.rotationText, size: 18, color: cfg.rotationDim ? "AAAAAA" : "888888" })
    ]
  });

  function priorityCell(letter, width, text, consequence, dim) {
    return cell([
      p(letter, { bold: true, size: 26, color: dim ? "CCCCCC" : COLORS.tan }),
      p(text || "", { size: 16, italics: dim }),
      ...blankLine(dim ? 1 : 2),
      p(consequence || "", { italics: true, size: 16, color: "999999" })
    ], { width, shading: dim ? "F5F5F5" : COLORS.tanBg, borders: boxBorders(dim ? "DDDDDD" : COLORS.tan, 6) });
  }

  function habitRow(label, dual) {
    return new Paragraph({ spacing: { after: 160 }, children: [
      new TextRun({ text: dual ? "☐ ☐  " : "☐  ", size: 22, color: COLORS.green }),
      new TextRun({ text: label, size: 20 })
    ]});
  }
  const habitsBox = cell([
    p("DAILY NON-NEGOTIABLES", { bold: true, size: 20, color: COLORS.green }),
    habitRow(cfg.habit1 || "5-min walk"),
    habitRow("60oz water + 100g protein", true),
    habitRow("Bible time"),
    habitRow("Brush teeth (AM/PM)")
  ], { width: 3400, shading: COLORS.greenBg, borders: boxBorders(COLORS.green, 10) });

  const mainTopTable = new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [2467, 2467, 2466, 3400],
    rows: [ new TableRow({ children: [
      priorityCell("A", 2467, cfg.priorityA, cfg.consequenceA, false),
      priorityCell("B", 2467, cfg.priorityB, cfg.consequenceB, cfg.dimBC),
      priorityCell("C", 2466, cfg.priorityC, cfg.consequenceC, cfg.dimBC),
      habitsBox
    ]}) ]
  });

  const rightColTable = new Table({
    width: { size: 4320, type: WidthType.DXA }, columnWidths: [4320],
    rows: [
      new TableRow({ children: [cell([p("HOW I FEEL TODAY", { bold: true, size: 20, color: COLORS.blue }), p("☐ drained   ☐ rough   ☐ okay   ☐ good   ☐ great", { size: 18 })], { width: 4320, shading: COLORS.blueBg, borders: boxBorders(COLORS.blue, 8) })] }),
      new TableRow({ children: [cell([p("ENERGY / FOCUS TANK", { bold: true, size: 20, color: COLORS.tan }), p("☐ ☐ ☐ ☐ ☐", { size: 20 })], { width: 4320, shading: COLORS.tanBg, borders: boxBorders(COLORS.tan, 8) })] }),
      new TableRow({ children: [cell([p("FUMBLE TRIGGERS TODAY", { bold: true, size: 18, color: "666666" }), ...blankLine(2)], { width: 4320, shading: COLORS.greyBg, borders: boxBorders(COLORS.grey, 8) })] })
    ]
  });

  const midGrid = new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [6480, 4320],
    rows: [ new TableRow({ children: [
      cell([p("BRAIN DUMP", { bold: true, size: 20 }), ...blankLine(6)], { width: 6480, borders: boxBorders(COLORS.dark, 10) }),
      new TableCell({ width: { size: 4320, type: WidthType.DXA }, borders: noBorders(), margins: { top: 0, bottom: 0, left: 0, right: 0 }, children: [rightColTable] })
    ]}) ]
  });

  function taskLaneCell(title, width, lines) {
    return cell([
      p(title, { bold: true, size: 20 }),
      ...lines.map(l => p(`★ ${l}`, { size: 18 }))
    ], { width, borders: boxBorders(COLORS.dark, 8) });
  }
  const taskTable = new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [3600, 3600, 3600],
    rows: [ new TableRow({ children: [
      taskLaneCell(cfg.laneGlsTitle, 3600, cfg.laneGls),
      taskLaneCell(cfg.laneFocusTitle, 3600, cfg.laneFocus),
      taskLaneCell(cfg.laneHomeTitle, 3600, cfg.laneHome)
    ]}) ]
  });

  const hours = new Array(16).fill("");
  let blockRowCells = [];
  hours.forEach((h, i) => {
    const shade = cfg.blockShading ? cfg.blockShading[i] : null;
    blockRowCells.push(cell([new Paragraph({ children: [new TextRun("")] }), new Paragraph({ children: [new TextRun("")] })],
      { width: Math.floor(6300/hours.length), shading: shade, borders: boxBorders("DDDDDD",4) }));
  });
  const blockItInner = new Table({ width: { size: 6300, type: WidthType.DXA }, columnWidths: hours.map(()=>Math.floor(6300/hours.length)), rows: [ new TableRow({ children: blockRowCells }) ] });

  const legendPara = cfg.blockLegend ? new Paragraph({ spacing: { before: 100 }, children: cfg.blockLegend.flatMap(l => ([
    new TextRun({ text: "■ ", size: 16, color: l.color }), new TextRun({ text: l.label + "   ", size: 14, color: "777777" })
  ])) }) : new Paragraph({ children: [] });

  const gwCell = cell([
    p("WIN + GRATITUDE", { bold: true, size: 20, color: COLORS.green }),
    p("Today I'm proud I...", { size: 18 }), ...blankLine(1),
    p("Grateful for...", { size: 18 }), ...blankLine(1)
  ], { width: 3960, shading: COLORS.greenBg, borders: boxBorders(COLORS.green, 8) });

  const bottomTable = new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [6840, 3960],
    rows: [ new TableRow({ children: [
      new TableCell({ width: { size: 6840, type: WidthType.DXA }, borders: boxBorders(COLORS.dark, 10), margins: { top: 120, bottom: 120, left: 140, right: 140 },
        children: [ p("BLOCK IT — TIME MAP (6a–10p)", { bold: true, size: 20 }), blockItInner, legendPara ] }),
      gwCell
    ]}) ]
  });

  const tagPara = cfg.tag ? new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: [CONTENT_W],
    rows: [ new TableRow({ children: [ cell([ p(cfg.tag, { bold: true, size: 15, color: COLORS.red, align: AlignmentType.CENTER }) ], { width: CONTENT_W, shading: COLORS.redBg }) ] }) ]
  }) : null;

  const children = [headerTable];
  if (tagPara) { children.push(new Paragraph({ spacing: { before: 100, after: 100 }, children: [] })); children.push(tagPara); }
  children.push(
    rotationPara,
    new Paragraph({ spacing: { after: 100 }, children: [] }),
    mainTopTable,
    new Paragraph({ spacing: { after: 100 }, children: [] }),
    midGrid,
    new Paragraph({ spacing: { after: 100 }, children: [] }),
    taskTable,
    new Paragraph({ spacing: { after: 100 }, children: [] }),
    bottomTable,
    new Paragraph({ spacing: { before: 100 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: "DAILY PLANNER — FATHER · HUSBAND · PROVIDER", size: 16, color: "BBBBBB" })] })
  );

  const doc = new Document({ sections: [{ properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN } } }, children }] });
  return Packer.toBuffer(doc).then(buf => { require('fs').writeFileSync(outPath, buf); console.log('written', outPath); });
}

// ---------- TODAY: Tuesday, July 14, 2026 (Admin focus + movie night) ----------
const GLS_GREEN = "DFE7E2", ADMIN_TAN = "F4E9DD", FAMILY_GREEN = "D9E6E0";
const today = {
  dayIndex: 2, date: "07/14/2026", recoveryChecked: false,
  rotationText: "Mon: Book Keeping | [Tue: Admin] | Wed: Student Ministry | Thu: Personal | Fri: Lopez Home | Sat/Sun: Family",
  priorityA: "Email 50 customers — ACH transition (GLS)", consequenceA: "if this slips... falls behind before deadline crunch",
  priorityB: "AC unit follow-up call + set up Costco water delivery (Admin)", consequenceB: "if this slips... AC issue drags another week",
  priorityC: "Protect family movie night 6–9pm — phone away", consequenceC: "if this slips... promised time with Whitney & Joelle slides again",
  dimBC: false,
  laneGlsTitle: "GLS (Work — daily)", laneGls: ["Email 50 customers — ACH", "Process daily cash log", "________________________"],
  laneFocusTitle: "Today's Focus: Administration", laneFocus: ["Call re: AC unit follow-up", "Set up Costco water delivery (Thu)", "________________________"],
  laneHomeTitle: "Home, Family & Me", laneHome: ["Movie night 6–9pm w/ Whitney & Joelle", "________________________", "________________________"],
  blockShading: [null,null, GLS_GREEN,GLS_GREEN,GLS_GREEN,GLS_GREEN,GLS_GREEN,GLS_GREEN,GLS_GREEN,GLS_GREEN, ADMIN_TAN,ADMIN_TAN, FAMILY_GREEN,FAMILY_GREEN,FAMILY_GREEN, null],
  blockLegend: [{color: GLS_GREEN, label: "GLS (8-4)"}, {color: ADMIN_TAN, label: "Admin (4-6)"}, {color: FAMILY_GREEN, label: "Family (6-9)"}]
};

// ---------- RECOVERY DAY: Thursday, July 30, 2026 (day after surgery) ----------
const recoveryDay = {
  dayIndex: 4, date: "07/30/2026", recoveryChecked: true, tag: "RECOVERY MODE — DAY 1 AFTER SURGERY (JUL 29)",
  rotationText: "Rotation paused this week — recovery only.", rotationDim: true,
  priorityA: "Rest. Take meds/antibiotics as prescribed. Short walk only if cleared by doctor.", consequenceA: "if this slips... recovery setback — this is the only job today",
  priorityB: "", consequenceB: "", priorityC: "", consequenceC: "", dimBC: true,
  laneGlsTitle: "GLS (Work)", laneGls: ["Not working this week — on medical leave", "________________________", "________________________"],
  laneFocusTitle: "Today's Focus: Recovery", laneFocus: ["Rest", "Take meds on schedule", "Stay hydrated"],
  laneHomeTitle: "Home, Family & Me", laneHome: ["Let Whitney/family help — without guilt", "Light stretching only if cleared", "________________________"],
  blockShading: null, blockLegend: [{color: "F5F5F5", label: "No fixed schedule this week — rest as needed"}]
};

Promise.all([
  buildDaily(today, "/home/claude/daily_planner_Jul14_TODAY.docx"),
  buildDaily(recoveryDay, "/home/claude/daily_planner_Jul30_RECOVERY.docx"),
]).then(() => console.log("ALL DONE"));
