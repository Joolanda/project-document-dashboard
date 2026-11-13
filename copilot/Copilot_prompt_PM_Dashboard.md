# Copilot in Excel – prompt for the WBS document dashboard

Use in Excel (Microsoft 365) with the workbook open, the Documents list formatted as a table (`tblDocuments`). Paste the prompt in one go; if Copilot only does part of it, continue with the follow-up prompts below.

---

**Prompt**

> Using the table **tblDocuments** on the Documents tab (one row per project document), create a new sheet called **PM Dashboard** for stakeholders and team members. Use formulas that reference the table, so the dashboard updates when the table changes. Do not change the existing data on the Documents tab (adding the two helper columns described below is fine).
>
> 1. **Filters** at the top: a *Work Package* dropdown (All, Custom selection, and each work package), a *Document set* dropdown (All documents, Key deliverables = documents with *Key Deliverable* = Yes, plus any other Yes/No set column I name) and a *Status date* cell (default `=TODAY()`). Next to the dashboard, add a **WP selection panel**: each work package with an "Include" cell (Yes/blank) that is used when Work Package = *Custom selection*, so I can pick several WPs at once (e.g. WP02 + WP03). Make all figures below follow these filters. If that is easiest, add two helper columns at the end of tblDocuments (1/0: "WP selected", "in set") and base the formulas on them.
> 2. **KPI tiles:** Total documents · Approved (column *Approved and filed date* filled, on or before the status date) with % of total · In review/approval (status *Draft in review*, *Sent for approval* or *Sent for sign-off*) · Not started · Weighted progress (average of *Status %*) · Overdue (not approved and *Forecast approval date* before the status date) · Behind baseline (documents due by *Baseline approval date* up to the status date minus documents approved).
> 3. **S-curve**, weekly from 6 July 2026 to end of June 2027, cumulative number of documents approved:
>    - **Baseline** (agreed at start) = count of *Baseline approval date* ≤ week date.
>    - **Actual** (status today) = count of *Approved and filed date* ≤ week date, only up to the status date.
>    - **Forecast** (planned dates updated) = approved up to the status date + open documents with *Forecast approval date* ≤ week date, only from the status date onward.
>    Actual and forecast must look like **one line**: same colour, actual solid, forecast dashed, starting where actual ends. Baseline in grey dotted. Use `NA()` for empty weeks so the lines don't drop to zero.
> 4. **Document flow:** count and % per status in this order: Not started, Started, Outline approved, Draft in review, Sent for approval, Sent for sign-off, Approved and filed. Add a horizontal bar chart.
> 5. **Status per Work Package:** a matrix with a row per work package and a column per status, plus Total, Approved % and Weighted %. Add a stacked bar chart. This matrix shows all work packages side by side and follows the document-set filter. Make a second matrix **per document set** (All documents, Key deliverables, other sets) that follows the work-package filter/selection.
> 6. **Attention list:** the 15 oldest open documents whose forecast approval date has passed: Document ID, Title, Work Package, Owner, Status, Forecast date, Days late.
>
> Use a clean layout: white background, navy headers, Arial, no gridlines.

**Follow-up prompts**
- "Add a document set 'Procurement set' based on the Yes/No column *Procurement set* and include it in the set filter and the set matrix."
- "Add a filter for Document Owner and make all figures respond to it."
- "Add a KPI for the average number of days between *Draft in review date* and *Approved and filed date* per work package."
- "Show the S-curve for WP02 only and compare it with WP03 in one chart."

**Note:** Copilot works on the data in your own Microsoft 365 environment, so the compliance rules that apply to your workbook stay in force. Check the formulas Copilot writes once; after that the tab updates by itself.
