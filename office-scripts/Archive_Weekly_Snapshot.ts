/**
 * Office Script – Archive Weekly Snapshot
 * --------------------------------------------------------------------------
 * Appends this week's status counts (all documents + one row per Work Package)
 * to the table "tblHistory" on the tab "PM History". Creates the tab the first time.
 * Run it manually each week, or from Power Automate:
 *   Recurrence (weekly, e.g. Friday 17:00) → Excel Online (Business) "Run script"
 *   → this script → (optional) send the returned summary by Outlook / Teams.
 * Running it twice on the same day replaces that day's rows (no duplicates).
 */
const TABLE_NAME = "tblDocuments";
const COL_WP = "Work Package";
const COL_STATUS = "Status";
const COL_KEY = "Key Deliverable";
const STATUSES: [string, number][] = [
  ["Not started", 0], ["Started", 0], ["Outline approved", 10],
  ["Draft in review", 50], ["Sent for approval", 70],
  ["Sent for sign-off", 90], ["Approved and filed", 100],
];
const HIST_SHEET = "PM History";
const HIST_TABLE = "tblHistory";

function main(workbook: ExcelScript.Workbook): string {
  const t = workbook.getTable(TABLE_NAME);
  if (!t) throw new Error(`Table ${TABLE_NAME} not found`);
  const colVals = (n: string): string[] => {
    const col = t.getColumnByName(n);
    if (!col) throw new Error(`Column ${n} not found`);
    return col.getRangeBetweenHeaderAndTotal().getValues().map(r => String(r[0]).trim());
  };
  const wp = colVals(COL_WP), st = colVals(COL_STATUS), ps = colVals(COL_KEY);

  const now = new Date();
  const today = Math.floor((Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()) - Date.UTC(1899, 11, 30)) / 86400000);

  const groups: string[] = ["All", "All – Key deliverables", ...Array.from(new Set(wp.filter(v => v !== ""))).sort()];
  const rows: (string | number)[][] = groups.map(g => {
    const idx = wp.map((w, i) => i).filter(i =>
      wp[i] !== "" && (g === "All" || (g === "All – Key deliverables" ? ps[i] === "Yes" : wp[i] === g)));
    const counts = STATUSES.map(([s]) => idx.filter(i => st[i] === s).length);
    const total = idx.length;
    const weighted = total === 0 ? 0 :
      idx.reduce((a, i) => a + (STATUSES.find(x => x[0] === st[i]) || ["", 0])[1], 0) / (100 * total);
    return [today, g, total, ...counts, total === 0 ? 0 : counts[6] / total, weighted];
  });

  let sh = workbook.getWorksheet(HIST_SHEET);
  let ht = workbook.getTable(HIST_TABLE);
  if (!sh) sh = workbook.addWorksheet(HIST_SHEET);
  if (!ht) {
    // first run: write header + this week's rows together, then make it a table
    // (a table created from a header row only gets an empty first data row)
    const head = ["Snapshot date", "Scope", "Total", ...STATUSES.map(s => s[0]), "Approved %", "Weighted %"];
    const block = sh.getRangeByIndexes(0, 0, rows.length + 1, head.length);
    block.setValues([head, ...rows]);
    ht = sh.addTable(block, true);
    ht.setName(HIST_TABLE);
  } else {
    // remove rows already written today (re-run safe)
    const body = ht.getRangeBetweenHeaderAndTotal().getValues();
    for (let i = body.length - 1; i >= 0; i--) {
      if (Number(body[i][0]) === today) ht.deleteRowsAt(i, 1);
    }
    ht.addRows(-1, rows);
    // a table can't be empty: if all old rows were today's, Excel kept one blank row – drop it
    const after = ht.getRangeBetweenHeaderAndTotal().getValues();
    for (let i = after.length - 1; i >= 0; i--) {
      if (after[i].every(v => String(v).trim() === "")) ht.deleteRowsAt(i, 1);
    }
  }
  const n = rows.length;
  const rng = ht.getRangeBetweenHeaderAndTotal();
  rng.getColumn(0).setNumberFormat("dd/mm/yyyy");
  rng.getColumn(10).setNumberFormat("0%");
  rng.getColumn(11).setNumberFormat("0%");

  const all = rows[0];
  return `Snapshot ${n} rows · ${all[2]} documents · ${all[9]} approved (${Math.round(Number(all[10]) * 100)}%) · weighted progress ${Math.round(Number(all[11]) * 100)}%`;
}
