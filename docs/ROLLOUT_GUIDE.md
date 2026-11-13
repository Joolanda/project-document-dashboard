# PM Dashboard – Rollout Guide

A live dashboard tab for the WBS document tracker, built inside your own workbook in about one hour. No data leaves your Microsoft 365 environment.

## What you receive

| File | What it is | When you need it |
| --- | --- | --- |
| `demo/WBS_Document_Dashboard_DEMO.xlsx` | Example workbook with 300 fictitious documents and the finished dashboard | Step 1 – to see the result before touching real data |
| `office-scripts/Build_PM_Dashboard.ts` | Office Script that builds the dashboard tab in your real workbook | Step 3 – run once, re-run when WPs or document sets change |
| `office-scripts/Archive_Weekly_Snapshot.ts` | Office Script that stores the weekly status counts in a history tab | Step 6 – weekly, manually or via Power Automate |
| `copilot/Copilot_prompt_PM_Dashboard.md` | Prompt that asks Copilot in Excel to build the same dashboard | Alternative to Step 3, or for ad-hoc questions in meetings |
| `docs/images/` | Screenshots of the dashboard and the S-curve with the demo data | Reference |

The dashboard shows KPI tiles, an S-curve (baseline vs. actual + forecast), the document flow per status, status per work package and per document set, and a list of overdue documents. Everything follows three filters: work package (one, several or all), document set (e.g. Key deliverable) and status date.

## Before you start

Check these five points first; each takes a minute and prevents most problems later.

- [ ] **Excel for Microsoft 365** (desktop or Excel on the web). Office Scripts and the attention list need it; Excel 2019/2021 will not work.
- [ ] **The workbook is saved in OneDrive or SharePoint.** Scripts and Power Automate only run on files stored there.
- [ ] **The Automate tab is visible** in the Excel ribbon. If not, ask IT to enable Office Scripts for your account.
- [ ] **Work on a copy first.** File → Save a copy, e.g. `WBS Tracker – dashboard test.xlsx`. Switch to the real file only after Step 4.
- [ ] **Compliance:** the scripts read and write only inside the workbook. If your organisation requires approval for scripts or Copilot, get it before Step 3.

## Step 1 – Explore the demo (15 min)

Open **WBS_Document_Dashboard_DEMO.xlsx** and try the filters before building anything; it shows exactly what Step 3 will produce. All names and figures in it are fictitious.

1. Tab **Dashboard**: the yellow cells in row 4 are the filters.
2. Set **Work Package** to `WP06 - Commissioning Preparation`. All tiles, the S-curve and the lists now show WP06 only.
3. Set **Work Package** to `Custom selection`. The panel on the right (*WP selection*) decides which WPs count; WP02 and WP03 are marked "Yes". Add or remove a "Yes" and watch the numbers change.
4. Set **Document set** to `Key deliverables`. Combined with step 3 you now see the key deliverables of WP02 + WP03.
5. Overwrite the **Status date** with an earlier date to see the dashboard as it looked then. Type `=TODAY()` to reset. Only the date-based figures go back in time (approved, overdue, behind baseline, S-curve, attention list); tiles and tables based on the *Status* column keep showing today's status.
6. Tab **Documents**: change one document's *Status* or *Approved and filed date* and check that the dashboard follows.

The grey columns *Dash WP selected* and *Dash in set* at the end of the Documents table are calculated helper columns. Nobody types in them.

## Step 2 – Prepare the Documents tab (30 min)

The dashboard reads one Excel table. It needs the columns below, and **two dates per document**: the baseline agreed at the start and the forecast the owner keeps up to date.

1. **Make it a table.** Click any cell in the document list → *Insert → Table* (Ctrl+T) → tick *My table has headers*. If it already is a table, skip this.
2. **Name the table.** *Table Design* tab → *Table Name* (far left) → type `tblDocuments` → Enter.
3. **Add the baseline date.** The tracker has one *Target approval date*. Decide which it is (see *Decisions to agree*):
   - It is the agreed baseline → rename it `Baseline approval date`, freeze it, and add a new column `Forecast approval date` (copy the values in once).
   - It is the latest planning → rename it `Forecast approval date` and add `Baseline approval date` with the dates agreed at the start.
4. **Check the column names.** They must match the list below exactly (spaces and capitals included). Either rename the headers, or change the names in the `COL` section at the top of the script.

| Column header the script expects | Content |
| --- | --- |
| Work Package | WP name, e.g. WP00 - Project Setup |
| Document ID / Document Title / Document Owner | As now |
| Status | Dropdown with the 7 statuses |
| Status % | 0 / 0 / 10 / 50 / 70 / 90 / 100, from the status |
| Baseline approval date | Agreed at start, never changed |
| Forecast approval date | Planned date, updated by the owner |
| Outline approved date | Milestone date |
| Draft in review date | Milestone date |
| Sent for approval date | Milestone date |
| Sent for sign-off date | Milestone date |
| Approved and filed date | Milestone date; filled = approved |
| Key Deliverable | Yes / No |

Any other columns (comments, links, version) can stay; the script ignores them. Each further document set (e.g. a Procurement set) needs its own Yes/No column.

## Step 3 – Run the script (10 min)

One run creates the tabs *PM Dashboard* and *PM S-curve data* and adds the two grey helper columns to the table. Your existing tabs and data stay as they are.

1. Open **Build_PM_Dashboard.ts** in Notepad (Windows) or TextEdit (Mac) and copy all text (Ctrl+A, Ctrl+C).
2. In Excel: **Automate → New Script**. A code pane opens on the right.
3. Select everything in that pane and paste (Ctrl+A, Ctrl+V).
4. Check the settings at the top of the script and change only the text between quotes:
   - `COL` – your column names (Step 2).
   - `SETS` – document sets; Key deliverables is there. Add a line per extra set, e.g. `["Procurement set", "Procurement set", "Yes"],`
   - `SCURVE_START` / `SCURVE_END` – first and last week on the S-curve, as `YYYY-MM-DD`.
   - `TARGET_DOCS` – the agreed number of documents (300).
5. Rename the script (click its name at the top) to `Build PM Dashboard` and click **Save script**.
6. Click **Run**. After 10–30 seconds the *PM Dashboard* tab opens.

If it stops with a red message, it names the cause, usually a column name that does not match. Fix it in the `COL` section and run again. Re-running is always safe: it rebuilds the two dashboard tabs from scratch.

## Step 4 – Check the result (15 min)

Compare a few dashboard numbers with a plain filter on the Documents tab before you show it to anyone.

- [ ] **Total documents** (filters on *All*) equals the number of rows in the table.
- [ ] **Approved & filed** equals the rows with a filled *Approved and filed date*.
- [ ] Pick one WP: **Not started** matches a filter on that WP + status *Not started*.
- [ ] **Overdue**: filter *Approved* empty and *Forecast approval date* before today; same count, same documents as the attention list.
- [ ] **S-curve**: the grey baseline ends at the total number of documents; the blue line stops at today and continues dashed.
- [ ] Set **Document set** to *Key deliverables*: the total equals the rows with *Key Deliverable* = Yes.

All good → run the script once more on the real file (Step 3) and delete the test copy.

## Step 5 – Use it in meetings

The dashboard answers most stakeholder questions with two clicks; share the screen on the *PM Dashboard* tab and change the yellow filters live.

| Question in the meeting | What you do |
| --- | --- |
| "Where are we overall?" | Filters on *All* / *All documents*; read the tiles and the S-curve |
| "How is WP06 doing?" | Work Package → `WP06 …` |
| "And WP02 and WP03 together?" | Work Package → `Custom selection`; "Yes" for WP02 and WP03 in the panel on the right |
| "What about the key deliverables?" | Document set → `Key deliverables` (combines with any WP choice) |
| "Which WP is behind?" | Scroll to *Status per Work Package*: all WPs side by side |
| "Who needs to act?" | *Attention list*: overdue documents with owner and days late |
| "What did it look like at the last steering meeting?" | Type that date in *Status date* for approvals and the S-curve; use *PM History* for the status counts of that week |
| A question the dashboard doesn't cover | Ask Copilot in Excel (see the prompt file), e.g. "average days from draft review to approval per WP" |

Reading the S-curve: grey dotted = baseline agreed at the start; solid blue = documents approved up to today; dashed blue = forecast from the owners' planned dates. A gap between grey and blue is the delay in documents; where the dashed line meets 100% is the forecast completion date.

## Step 6 – Weekly routine and Power Automate

The dashboard always shows today's status; to see a trend over the weeks, store a snapshot every week. Start manually, automate once the routine works.

**Every week (owners, before the deadline you set):** update *Status*, the milestone dates and the *Forecast approval date* for their documents. The *Baseline approval date* is never changed.

**Every week (you, 2 min):** Automate → *Archive Weekly Snapshot* → Run. The first run creates a tab *PM History*; each run adds one row per WP plus totals. Paste the script the same way as in Step 3.

**Automate it** (once, 15 min):

1. Go to [make.powerautomate.com](https://make.powerautomate.com) → *Create* → *Scheduled cloud flow*.
2. Name it `WBS weekly snapshot`, repeat every 1 week, e.g. Friday 16:00.
3. *New step* → *Excel Online (Business)* → **Run script**. Choose the location (OneDrive or the SharePoint site), the file and the script *Archive Weekly Snapshot*.
4. Optional: *New step* → *Office 365 Outlook* → **Send an email (V2)** to the stakeholders. Put the dynamic value *result* in the body; it reads e.g. "300 documents · 49 approved (16%) · weighted progress 35%". Add a link to the workbook.
5. *Save* → *Test* → *Manually* to check it once.

Later options, when the first reports run smoothly: a reminder mail to each owner with overdue documents, and a Power BI report on the same table for a larger audience.

## Decisions to agree with the team

The dashboard is only as good as these agreements; settle them before the first stakeholder report.

| Decision | Proposal in the current setup |
| --- | --- |
| Which date is the baseline, and when is it frozen? | The approval dates agreed at kick-off; changed only through a formal re-baseline |
| Who updates the forecast dates, and how often? | Each document owner, weekly, before the snapshot |
| When does a document count as done? | *Approved and filed date* is filled |
| Weights per status (Status %) | 0 / 0 / 10 / 50 / 70 / 90 / 100 – adjust in the Lists if the team prefers others |
| Which document sets to report on? | Key deliverables; others (e.g. Procurement) as Yes/No columns |
| Target number of documents | 300 (currently in the script as `TARGET_DOCS`) |
| Who receives the weekly update? | To decide: steering group, WP leads, document owners |

## Troubleshooting

| What you see | Cause | Fix |
| --- | --- | --- |
| No *Automate* tab in Excel | Office Scripts not enabled, or the file is not in OneDrive/SharePoint | Save the file in OneDrive; otherwise ask IT to enable Office Scripts |
| Script stops: *Table "tblDocuments" not found* | The list is not a table, or has another name | Step 2, points 1–2 |
| Script stops: *Column(s) not found* | A header differs from the script (space, capital, typo) | Rename the header or adjust `COL` / `SETS` |
| Dashboard shows 0 everywhere | Filter set to a WP or set with no documents, or helper columns broken | Set filters to *All*; re-run the script |
| New WP missing from the dropdown | WP list is read when the script runs | Re-run the script |
| S-curve flat or empty | Date columns hold text instead of dates | Select the column → *Data → Text to Columns → Finish*, or retype as dd/mm/yyyy |
| Attention list shows #NAME? | Excel version without FILTER/SORTBY | Use Excel for Microsoft 365 (desktop or web) |
| *Unable to refresh – linked workbook* banner | Existing link to another workbook | Not caused by the dashboard; fix or break the link via *Data → Edit Links* |
