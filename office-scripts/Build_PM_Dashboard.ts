/**
 * Office Script – Build PM Dashboard  (v2: multi-WP selection + document sets)
 * --------------------------------------------------------------------------
 * Builds (or rebuilds) in THIS workbook:
 *   - "PM Dashboard"     filters, KPIs, S-curve, document flow, status per WP,
 *                        status per document set, attention list, WP selection panel
 *   - "PM S-curve data"  weekly baseline / actual / forecast series (feeds the S-curve)
 *   - two grey helper columns at the end of the table: "Dash WP selected" and "Dash in set"
 *     (calculated; they translate the dashboard filters into 1/0 per document)
 *
 * Filters on the dashboard:
 *   Work Package  : All · one WP · "Custom selection" (tick "Yes" for any WPs in the panel on the right)
 *   Document set  : All documents · Key deliverables · any other set you define in SETS below
 *   Status date   : =TODAY() by default; overwrite to look back
 *
 * Everything is a live formula on the table, so after one run the dashboard updates by itself.
 * Re-run only when work packages or document sets change, or the S-curve date range changes.
 *
 * HOW TO USE
 *   1. Documents tab: click in the list → Insert → Table (Ctrl+T) if needed.
 *      Table Design → Table Name: tblDocuments  (or change TABLE_NAME below).
 *   2. Check COL and SETS below: names must match your column headers EXACTLY.
 *   3. Automate → New Script → paste this file → Save as "Build PM Dashboard" → Run.
 *   4. Optional: run it from Power Automate (Excel Online (Business) → Run script).
 * Data never leaves the workbook: the script only reads and writes this file.
 */

// ===================== CONFIGURATION – adjust to your workbook =====================
const TABLE_NAME = "tblDocuments";

const COL = {
  wp: "Work Package",
  docId: "Document ID",
  title: "Document Title",
  owner: "Document Owner",
  status: "Status",
  statusPct: "Status %",
  baseline: "Baseline approval date",          // agreed at start (frozen)
  forecast: "Forecast approval date",          // planned date as updated by the owner
  // milestone dates, in workflow order
  msStructure: "Outline approved date",
  msDraft: "Draft in review date",
  msApprover: "Sent for approval date",
  msSignature: "Sent for sign-off date",
  msApproved: "Approved and filed date",
};

// Document sets: [name shown in the filter, column in the table, value that means "in this set"]
// Add a row per set (e.g. a Yes/No column "Procurement set"). The first entry must stay "All documents".
const SETS: [string, string, string][] = [
  ["All documents", "", ""],
  ["Key deliverables", "Key Deliverable", "Yes"],
  // ["Procurement set", "Procurement set", "Yes"],
];

// Status values exactly as in the dropdown, with their Status % weight
const STATUSES: [string, number][] = [
  ["Not started", 0],
  ["Started", 0],
  ["Outline approved", 10],
  ["Draft in review", 50],
  ["Sent for approval", 70],
  ["Sent for sign-off", 90],
  ["Approved and filed", 100],
];
const IN_REVIEW = ["Draft in review", "Sent for approval", "Sent for sign-off"];

const SCURVE_START = "2026-07-06"; // first Monday shown on the S-curve
const SCURVE_END = "2027-06-28";   // last week shown
const TARGET_DOCS = 300;           // shown as reference under the total
const DASH = "PM Dashboard";
const DATA = "PM S-curve data";
// ====================================================================================

const CUSTOM = "Custom selection";
const H_WP = "Dash WP selected";
const H_SET = "Dash in set";
const PANEL_TOP = 13, PANEL_MAX = 50;  // WP selection panel: R13:S62
const NAVY = "#1F3864", ORANGE = "#C55A11", RED = "#C00000", GREY = "#7F7F7F", LIGHT = "#F2F4F7", INPUT = "#FFF2CC";
const STATUS_COLORS = ["#D9D9D9", "#BDD7EE", "#9DC3E6", "#5B9BD5", "#2E75B6", "#1F4E79", "#2E7D32"];

function main(workbook: ExcelScript.Workbook) {
  const table = workbook.getTable(TABLE_NAME);
  if (!table) throw new Error(`Table "${TABLE_NAME}" not found. Convert the Documents list to a table (Ctrl+T) and name it ${TABLE_NAME}.`);

  // ---- validate column names
  const headers = table.getHeaderRowRange().getValues()[0].map(h => String(h).trim());
  const needed = [...Object.values(COL), ...SETS.slice(1).map(s => s[1])];
  const missing = needed.filter(c => headers.indexOf(c) < 0);
  if (missing.length) throw new Error("Column(s) not found in table: " + missing.join(" | ") + ". Adjust COL / SETS at the top of the script.");

  // ---- work packages present in the data (sorted)
  const wpCol = table.getColumnByName(COL.wp);
  if (!wpCol) throw new Error(`Column ${COL.wp} not found`);
  const wps = Array.from(new Set(wpCol.getRangeBetweenHeaderAndTotal().getValues()
    .map(r => String(r[0]).trim()).filter(v => v !== ""))).sort();
  if (wps.length > PANEL_MAX) throw new Error(`More than ${PANEL_MAX} work packages – raise PANEL_MAX.`);

  // ---- fresh sheets
  [DASH, DATA].forEach(n => { const s = workbook.getWorksheet(n); if (s) s.delete(); });
  const data = workbook.addWorksheet(DATA);
  const db = workbook.addWorksheet(DASH);
  db.setPosition(0);
  db.setShowGridlines(false);

  const T = TABLE_NAME;
  const c = (name: string) => `${T}[${name}]`;
  const D = `'${DASH}'!`;
  const PANEL_END = PANEL_TOP + PANEL_MAX - 1;

  // ================= helper columns in the table (created once, formulas refreshed every run)
  const setHelper = (name: string, formula: string) => {
    let col = table.getColumnByName(name);
    if (!col) col = table.addColumn(-1, undefined, name);
    col.getRangeBetweenHeaderAndTotal().setFormula(formula);
    col.getRange().getFormat().getFont().setColor(GREY);
  };
  setHelper(H_WP,
    `=IF(${D}$D$4="All",1,IF(${D}$D$4="${CUSTOM}",` +
    `IF(COUNTIFS(${D}$R$${PANEL_TOP}:$R$${PANEL_END},[@[${COL.wp}]],${D}$S$${PANEL_TOP}:$S$${PANEL_END},"Yes")>0,1,0),` +
    `IF([@[${COL.wp}]]=${D}$D$4,1,0)))`);
  const setParts = SETS.slice(1).map(s => `AND(${D}$H$4="${s[0]}",[@[${s[1]}]]="${s[2]}")`);
  setHelper(H_SET, setParts.length
    ? `=IF(${D}$H$4="${SETS[0][0]}",1,IF(OR(${setParts.join(",")}),1,0))`
    : `=1`);

  const FLT = `${c(H_WP)},1,${c(H_SET)},1`;
  const MASK = `(${c(H_WP)}=1)*(${c(H_SET)}=1)`;
  const PEIL = `${D}$L$4`;
  const TOTAL = `${D}$B$8`;

  // ================= Lists (on data sheet, columns K:L)
  const wpOpts = ["All", CUSTOM, ...wps];
  data.getRange("K1:L1").setValues([["WP filter", "Document set"]]);
  data.getRange(`K2:K${wpOpts.length + 1}`).setValues(wpOpts.map(w => [w]));
  data.getRange(`L2:L${SETS.length + 1}`).setValues(SETS.map(s => [s[0]]));
  data.getRange("K1:L1").getFormat().getFont().setBold(true);

  // ================= S-curve data
  data.getRange("A1").setValue("S-curve data (weekly) – follows the filters on the PM Dashboard tab");
  fmt(data.getRange("A1"), 12, true, NAVY);
  data.getRange("A2").setValue("Approved = approval date filled. Forecast = approved so far + open documents by their forecast date (overdue ones count from the status date).");
  fmt(data.getRange("A2"), 9, false, GREY, true);
  headerRow(data.getRange("A4:H4"), ["Week", "Baseline (agreed at start)", "Actual (status today)", "Forecast (planned dates updated)",
    "Baseline %", "Actual %", "Forecast %", "Earned progress % (weighted)"], NAVY);

  const weeks: number[] = [];
  for (let d = toSerial(SCURVE_START); d <= toSerial(SCURVE_END); d += 7) weeks.push(d);
  const ms = [COL.msStructure, COL.msDraft, COL.msApprover, COL.msSignature, COL.msApproved];
  const inc = [10, 40, 20, 20, 10]; // incremental weights → cumulative 10/50/70/90/100
  const rowsF: string[][] = weeks.map((w, k) => {
    const r = 5 + k;
    const earned = ms.map((m, i) => `${inc[i]}*COUNTIFS(${FLT},${c(m)},"<="&A${r})`).join("+");
    return [
      String(w),
      `=COUNTIFS(${FLT},${c(COL.baseline)},"<="&A${r})`,
      `=IF(A${r}<=${PEIL},COUNTIFS(${FLT},${c(COL.msApproved)},"<="&A${r}),NA())`,
      `=IF(A${r}+7>${PEIL},IF(A${r}<=${PEIL},C${r},COUNTIFS(${FLT},${c(COL.msApproved)},"<="&${PEIL})+COUNTIFS(${FLT},${c(COL.msApproved)},"=",${c(COL.forecast)},"<="&A${r})),NA())`,
      `=IF(ISNA(B${r}),NA(),IF(${TOTAL}=0,0,B${r}/${TOTAL}))`,
      `=IF(ISNA(C${r}),NA(),IF(${TOTAL}=0,0,C${r}/${TOTAL}))`,
      `=IF(ISNA(D${r}),NA(),IF(${TOTAL}=0,0,D${r}/${TOTAL}))`,
      `=IF(A${r}<=${PEIL},IF(${TOTAL}=0,0,(${earned})/(100*${TOTAL})),NA())`,
    ];
  });
  const lastW = 4 + weeks.length;
  data.getRange(`A5:H${lastW}`).setFormulas(rowsF);
  data.getRange(`A5:A${lastW}`).setNumberFormat("dd/mm/yy");
  data.getRange(`E5:H${lastW}`).setNumberFormat("0%");
  data.getRange(`A4:H${lastW}`).getFormat().getFont().setName("Arial");
  data.getRange("A:A").getFormat().setColumnWidth(70);
  data.getRange("B:H").getFormat().setColumnWidth(95);

  // ================= Dashboard layout
  db.getRange("A:A").getFormat().setColumnWidth(12);
  db.getRange("B:B").getFormat().setColumnWidth(210);
  db.getRange("C:P").getFormat().setColumnWidth(80);
  db.getRange("Q:Q").getFormat().setColumnWidth(18);
  db.getRange("R:R").getFormat().setColumnWidth(200);
  db.getRange("S:S").getFormat().setColumnWidth(55);
  db.getRange("A1:T140").getFormat().getFont().setName("Arial");

  db.getRange("B1").setValue("WBS Document Progress – Dashboard");
  fmt(db.getRange("B1"), 18, true, NAVY);
  db.getRange("B2").setFormula(`="Showing: "&IF($D$4="All","all work packages",IF($D$4="${CUSTOM}","custom WP selection ("&COUNTIF($S$${PANEL_TOP}:$S$${PANEL_END},"Yes")&" WPs)",$D$4))&"  ·  "&$H$4&"  ·  status date "&TEXT($L$4,"dd/mm/yyyy")`);
  fmt(db.getRange("B2"), 10, true, GREY);

  // filters
  [["B4", "Work Package"], ["G4", "Document set"], ["K4", "Status date"]].forEach(([a, t]) => {
    db.getRange(a).setValue(t); fmt(db.getRange(a), 10, true, NAVY);
    db.getRange(a).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  });
  db.getRange("D4:F4").merge(false);
  db.getRange("H4:I4").merge(false);
  db.getRange("D4").setValue("All");
  db.getRange("H4").setValue(SETS[0][0]);
  db.getRange("L4").setFormula("=TODAY()");
  db.getRange("L4").setNumberFormat("dd/mm/yyyy");
  ["D4:F4", "H4:I4", "L4"].forEach(a => {
    const r = db.getRange(a);
    r.getFormat().getFill().setColor(INPUT);
    r.getFormat().getFont().setBold(true);
    r.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  });
  db.getRange("D4").getDataValidation().setRule({ list: { inCellDropDown: true, source: `='${DATA}'!$K$2:$K$${wpOpts.length + 1}` } });
  db.getRange("H4").getDataValidation().setRule({ list: { inCellDropDown: true, source: `='${DATA}'!$L$2:$L$${SETS.length + 1}` } });
  db.getRange("M4").setValue("← yellow cells are filters");
  fmt(db.getRange("M4"), 8, false, GREY, true);

  // WP selection panel
  db.getRange(`R${PANEL_TOP - 2}`).setValue("WP selection");
  fmt(db.getRange(`R${PANEL_TOP - 2}`), 11, true, NAVY);
  db.getRange(`R${PANEL_TOP - 1}`).setValue(`Used when Work Package = "${CUSTOM}". Set "Yes" to include.`);
  fmt(db.getRange(`R${PANEL_TOP - 1}`), 8, false, GREY, true);
  if (wps.length) {
    db.getRange(`R${PANEL_TOP}:R${PANEL_TOP + wps.length - 1}`).setValues(wps.map(w => [w]));
    const inc2 = db.getRange(`S${PANEL_TOP}:S${PANEL_TOP + wps.length - 1}`);
    inc2.getFormat().getFill().setColor(INPUT);
    inc2.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    inc2.getDataValidation().setRule({ list: { inCellDropDown: true, source: "Yes" } });
    db.getRange(`R${PANEL_TOP}:S${PANEL_TOP + wps.length - 1}`).getFormat().getFont().setSize(9);
  }

  // KPI tiles
  const inReview = IN_REVIEW.map(s => `COUNTIFS(${FLT},${c(COL.status)},"${s}")`).join("+");
  const kpis: [string, string, string, string, string][] = [
    ["B", "Total documents", `=COUNTIFS(${FLT})`, "0", `="in selection (project target ${TARGET_DOCS})"`],
    ["D", "Approved & filed", `=COUNTIFS(${FLT},${c(COL.msApproved)},"<="&$L$4)`, "0", `=IF(B8=0,"",TEXT(D8/B8,"0%")&" of total")`],
    ["F", "In review / approval", `=${inReview}`, "0", `="review → sign-off"`],
    ["H", "Not started", `=COUNTIFS(${FLT},${c(COL.status)},"${STATUSES[0][0]}")`, "0", `=COUNTIFS(${FLT},${c(COL.status)},"${STATUSES[1][0]}")&" started, no milestone yet"`],
    ["J", "Weighted progress", `=IF(B8=0,0,SUMPRODUCT(${MASK}*${c(COL.statusPct)})/(100*B8))`, "0%",
      `="baseline today: "&TEXT(IF(B8=0,0,COUNTIFS(${FLT},${c(COL.baseline)},"<="&$L$4)/B8),"0%")&" approved"`],
    ["L", "Overdue", `=COUNTIFS(${FLT},${c(COL.msApproved)},"=",${c(COL.forecast)},"<"&$L$4)`, "0", `="open, forecast date passed"`],
    ["N", "Behind baseline", `=COUNTIFS(${FLT},${c(COL.baseline)},"<="&$L$4)-COUNTIFS(${FLT},${c(COL.baseline)},"<="&$L$4,${c(COL.msApproved)},"<="&$L$4)`, "0", `="docs vs. baseline plan today"`],
  ];
  kpis.forEach(([col, label, f, nf, sub]) => {
    const c2 = col === "B" ? "B" : String.fromCharCode(col.charCodeAt(0) + 1);
    [6, 8, 9].forEach(r => { if (c2 !== col) db.getRange(`${col}${r}:${c2}${r}`).merge(false); });
    db.getRange(`${col}6`).setValue(label);
    db.getRange(`${col}8`).setFormula(f);
    db.getRange(`${col}9`).setFormula(sub);
    db.getRange(`${col}8`).setNumberFormat(nf);
    fmt(db.getRange(`${col}6`), 9, true, GREY);
    fmt(db.getRange(`${col}8`), 22, true, (col === "L" || col === "N") ? RED : NAVY);
    fmt(db.getRange(`${col}9`), 8, false, GREY);
    const tile = db.getRange(`${col}6:${c2}9`);
    tile.getFormat().getFill().setColor(LIGHT);
    tile.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  });
  db.getRange("8:8").getFormat().setRowHeight(34);

  // S-curve chart
  db.getRange("B11").setValue("S-curve – documents approved (baseline vs. actual + forecast)");
  fmt(db.getRange("B11"), 12, true, NAVY);
  const sChart = db.addChart(ExcelScript.ChartType.line, data.getRange(`B4:D${lastW}`), ExcelScript.ChartSeriesBy.columns);
  sChart.setPosition("B13", "O30");
  sChart.getTitle().setVisible(false);
  sChart.getLegend().setPosition(ExcelScript.ChartLegendPosition.bottom);
  const lineStyle: [string, ExcelScript.ChartLineStyle][] = [
    ["#A6A6A6", ExcelScript.ChartLineStyle.roundDot],
    ["#1F4E79", ExcelScript.ChartLineStyle.continuous],
    ["#1F4E79", ExcelScript.ChartLineStyle.dash],
  ];
  sChart.getSeries().forEach((s, i) => {
    s.setXAxisValues(data.getRange(`A5:A${lastW}`));
    s.setMarkerStyle(ExcelScript.ChartMarkerStyle.none);
    s.getFormat().getLine().setColor(lineStyle[i][0]);
    s.getFormat().getLine().setLineStyle(lineStyle[i][1]);
    s.getFormat().getLine().setWeight(i === 1 ? 2.5 : 2);
  });
  sChart.getAxes().getCategoryAxis().setNumberFormat("mmm yy");
  sChart.getAxes().getValueAxis().getMajorGridlines().setVisible(false);

  // Document flow
  const R0 = 32;
  db.getRange(`B${R0}`).setValue("Document flow – where are the documents now?");
  fmt(db.getRange(`B${R0}`), 12, true, NAVY);
  headerRow(db.getRange(`B${R0 + 1}:E${R0 + 1}`), ["Status", "Documents", "% of total", "Status %"], NAVY);
  const rTot = R0 + 2 + STATUSES.length;
  db.getRange(`B${R0 + 2}:E${rTot - 1}`).setFormulas(STATUSES.map(([s, w], i) => {
    const r = R0 + 2 + i;
    return [s, `=COUNTIFS(${FLT},${c(COL.status)},B${r})`, `=IF($B$8=0,0,C${r}/$B$8)`, String(w)];
  }));
  db.getRange(`D${R0 + 2}:D${rTot - 1}`).setNumberFormat("0%");
  db.getRange(`B${rTot}:C${rTot}`).setFormulas([["Total", `=SUM(C${R0 + 2}:C${rTot - 1})`]]);
  db.getRange(`B${rTot}:C${rTot}`).getFormat().getFont().setBold(true);
  dataBar(db.getRange(`D${R0 + 2}:D${rTot - 1}`), "#5B9BD5");
  const fChart = db.addChart(ExcelScript.ChartType.barClustered, db.getRange(`C${R0 + 1}:C${rTot - 1}`), ExcelScript.ChartSeriesBy.columns);
  fChart.getSeries()[0].setXAxisValues(db.getRange(`B${R0 + 2}:B${rTot - 1}`));
  fChart.getSeries()[0].getFormat().getFill().setSolidColor("#2E75B6");
  fChart.getAxes().getCategoryAxis().setReversePlotOrder(true);
  fChart.getLegend().setVisible(false);
  fChart.getTitle().setVisible(false);
  fChart.setPosition(`G${R0}`, `M${rTot + 1}`);

  // Matrix 1: status per WP (all WPs side by side) – follows the document-set filter
  const R1 = rTot + 6;
  const e1 = statusMatrix(db, R1, wps, (i, r) => `${c(COL.wp)},$B${r},${c(H_SET)},1`, NAVY, "Work Package", c);
  db.getRange(`B${R1}`).setFormula(`="Status per Work Package – "&$H$4&" (all WPs side by side)"`);
  const tot = ["Total (all WPs)", ..."CDEFGHIJ".split("").map(L => `=SUM(${L}${R1 + 2}:${L}${e1 - 1})`),
    `=IF(J${e1}=0,0,I${e1}/J${e1})`, `=IF(J${e1}=0,0,SUMPRODUCT(J${R1 + 2}:J${e1 - 1},L${R1 + 2}:L${e1 - 1})/J${e1})`];
  db.getRange(`B${e1}:L${e1}`).setFormulas([tot]);
  db.getRange(`B${e1}:L${e1}`).getFormat().getFont().setBold(true);
  db.getRange(`B${e1}:L${e1}`).getFormat().getFill().setColor(LIGHT);
  db.getRange(`K${e1}:L${e1}`).setNumberFormat("0%");
  db.getRange(`C${e1}:L${e1}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  // Matrix 2: status per document set – follows the WP filter / selection
  const R2 = Math.max(e1 + 4, R1 + wps.length + 18);
  const e2 = statusMatrix(db, R2, SETS.map(s => s[0]),
    (i, r) => SETS[i][1] === "" ? `${c(H_WP)},1` : `${c(H_WP)},1,${c(SETS[i][1])},"${SETS[i][2]}"`, ORANGE, "Document set", c);
  db.getRange(`B${R2}`).setFormula(`="Status per document set – "&IF($D$4="All","all work packages",IF($D$4="${CUSTOM}","custom WP selection",$D$4))`);

  // Attention list (dynamic array – Microsoft 365)
  const R3 = Math.max(e2 + 4, R2 + SETS.length + 14);
  db.getRange(`B${R3}`).setValue("Attention list – open documents whose forecast approval date has passed (oldest first, max 15)");
  fmt(db.getRange(`B${R3}`), 12, true, RED);
  headerRow(db.getRange(`B${R3 + 1}:H${R3 + 1}`), ["Document ID", "Document Title", "Work Package", "Owner", "Status", "Forecast date", "Days late"], RED);
  const cond = `${MASK}*(${c(COL.msApproved)}="")*(${c(COL.forecast)}<$L$4)`;
  db.getRange(`B${R3 + 2}`).setFormula(
    `=IFERROR(TAKE(SORTBY(FILTER(HSTACK(${c(COL.docId)},${c(COL.title)},${c(COL.wp)},${c(COL.owner)},${c(COL.status)},${c(COL.forecast)},$L$4-${c(COL.forecast)}),${cond}),FILTER(${c(COL.forecast)},${cond}),1),15),"No overdue documents")`);
  db.getRange(`G${R3 + 2}:G${R3 + 16}`).setNumberFormat("dd/mm/yyyy");
  db.getRange(`H${R3 + 2}:H${R3 + 16}`).setNumberFormat("0");
  db.getRange(`C:C`).getFormat().setColumnWidth(170);

  db.getFreezePanes().freezeRows(4);
  db.activate();
  db.getRange("D4").select();
}

// ------------------------------------------------------------------------------------
function statusMatrix(db: ExcelScript.Worksheet, top: number, labels: string[],
  crit: (i: number, r: number) => string, color: string, firstHead: string,
  c: (n: string) => string): number {
  fmt(db.getRange(`B${top}`), 12, true, NAVY);
  headerRow(db.getRange(`B${top + 1}:L${top + 1}`), [firstHead, ...STATUSES.map(s => s[0]), "Total", "Approved %", "Weighted %"], color);
  db.getRange(`${top + 1}:${top + 1}`).getFormat().setRowHeight(48);
  const cols = "CDEFGHI";
  const rt = top + 2 + labels.length;
  if (labels.length === 0) return rt;
  const body = labels.map((lab, i) => {
    const r = top + 2 + i;
    const counts = STATUSES.map((_, j) => `=COUNTIFS(${crit(i, r)},${c(COL.status)},${cols[j]}$${top + 1})`);
    const w = STATUSES.map((s, j) => s[1] ? `${s[1]}*${cols[j]}${r}` : "").filter(x => x !== "").join("+");
    return [lab, ...counts, `=SUM(C${r}:I${r})`, `=IF(J${r}=0,0,I${r}/J${r})`, `=IF(J${r}=0,0,(${w})/(100*J${r}))`];
  });
  db.getRange(`B${top + 2}:L${rt - 1}`).setFormulas(body);
  db.getRange(`K${top + 2}:L${rt}`).setNumberFormat("0%");
  db.getRange(`C${top + 2}:L${rt - 1}`).getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  db.getRange(`B${top + 2}:L${rt - 1}`).getFormat().getFont().setSize(9);
  dataBar(db.getRange(`K${top + 2}:L${rt - 1}`), "#63BE7B");

  const ch = db.addChart(ExcelScript.ChartType.barStacked, db.getRange(`C${top + 1}:I${rt - 1}`), ExcelScript.ChartSeriesBy.columns);
  ch.getSeries().forEach((s, i) => {
    s.setXAxisValues(db.getRange(`B${top + 2}:B${rt - 1}`));
    s.getFormat().getFill().setSolidColor(STATUS_COLORS[i]);
  });
  ch.getAxes().getCategoryAxis().setReversePlotOrder(true);
  ch.getAxes().getValueAxis().getMajorGridlines().setVisible(false);
  ch.getLegend().setPosition(ExcelScript.ChartLegendPosition.bottom);
  ch.getTitle().setVisible(false);
  ch.setPosition(`N${top}`, `T${Math.max(rt + 2, top + 14)}`);
  return rt;
}

function headerRow(r: ExcelScript.Range, values: string[], color: string) {
  r.setValues([values]);
  r.getFormat().getFill().setColor(color);
  r.getFormat().getFont().setColor("#FFFFFF");
  r.getFormat().getFont().setBold(true);
  r.getFormat().getFont().setSize(9);
  r.getFormat().setWrapText(true);
  r.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  r.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
}

function fmt(r: ExcelScript.Range, size: number, bold: boolean, color: string, italic: boolean = false) {
  const f = r.getFormat().getFont();
  f.setSize(size); f.setBold(bold); f.setColor(color); f.setItalic(italic); f.setName("Arial");
}

function dataBar(r: ExcelScript.Range, color: string) {
  const cf = r.addConditionalFormat(ExcelScript.ConditionalFormatType.dataBar).getDataBar();
  if (!cf) return;
  cf.getPositiveFormat().setFillColor(color);
  cf.setLowerBoundRule({ type: ExcelScript.ConditionalFormatRuleType.number, formula: "0" });
  cf.setUpperBoundRule({ type: ExcelScript.ConditionalFormatRuleType.number, formula: "1" });
}

function toSerial(iso: string): number {
  const [y, m, d] = iso.split("-").map(Number);
  return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000);
}
