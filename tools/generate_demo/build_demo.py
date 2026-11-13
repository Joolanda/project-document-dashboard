"""Demo workbook v2: WP filter (All / single / custom multi-selection) + document-set filter."""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import LineChart, BarChart, Reference
from openpyxl.formatting.rule import DataBarRule
from openpyxl.comments import Comment
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.utils import get_column_letter as CL

import os, pathlib
S = str(pathlib.Path(__file__).resolve().parent) + "/"
exec(open(S + "gen_data.py").read().split("print(len(rows))")[0])

OUT = str(pathlib.Path(__file__).resolve().parents[2] / "demo" / "WBS_Document_Dashboard_DEMO.xlsx")
F = "Arial"
NAVY, GREEN, ORANGE, GREY, LIGHT, RED, PURPLE = "1F3864", "2E7D32", "C55A11", "7F7F7F", "F2F4F7", "C00000", "7030A0"
STATUSES = ["Not started", "Started"] + MILESTONES
WEIGHTS = [0, 0, 10, 50, 70, 90, 100]
STATUS_COLORS = ["D9D9D9", "BDD7EE", "9DC3E6", "5B9BD5", "2E75B6", "1F4E79", "2E7D32"]
# Document sets: (name shown in filter, column in Documents, value meaning "in set")
SETS = [("All documents", None, None),
        ("Key deliverables", "Key Deliverable", "Yes"),
        ("Design review set (demo)", "Design Review Set", "Yes")]
CUSTOM = "Custom selection"
thin = Side(style="thin", color="D0D5DD")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)


def font(sz=10, bold=False, color="000000", italic=False):
    return Font(name=F, size=sz, bold=bold, color=color, italic=italic)


def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


def head(ws, row, col, values, color, size=9):
    for j, h in enumerate(values):
        c = ws.cell(row, col + j, h)
        c.font = font(size, True, "FFFFFF"); c.fill = fill(color)
        c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")


wb = Workbook()

# ================================================================ Lists
ls = wb.active
ls.title = "Lists"
ls["A1"], ls["B1"] = "Status", "Status %"
for i, (s, w) in enumerate(zip(STATUSES, WEIGHTS), start=2):
    ls.cell(i, 1, s); ls.cell(i, 2, w)
ls["D1"] = "Work Package filter options"
wp_opts = ["All", CUSTOM] + [w[0] for w in WPS]
for i, v in enumerate(wp_opts, start=2):
    ls.cell(i, 4, v)
ls["F1"], ls["G1"], ls["H1"] = "Document set", "Column in Documents", "Value = in set"
for i, (n, c, v) in enumerate(SETS, start=2):
    ls.cell(i, 6, n); ls.cell(i, 7, c or "–"); ls.cell(i, 8, v or "–")
ls["F7"] = "To add a set: add a Yes/No column to Documents, add a row here and extend the 'Dash in set' formula (the Office Script does this automatically from its SETS setting)."
ls["F7"].font = font(8, italic=True, color=GREY)
for c in ls[1]:
    c.font = font(bold=True)
for col, w in zip("ABCDEFGH", [46, 9, 2, 36, 2, 26, 20, 14]):
    ls.column_dimensions[col].width = w

# ================================================================ Dashboard (created early: helper columns refer to it)
db = wb.create_sheet("Dashboard", 0)
PANEL_TOP = 13                      # WP selection panel rows (R:S)
PANEL_ROWS = range(PANEL_TOP, PANEL_TOP + len(WPS))
PANEL_END = PANEL_TOP + 29         # room for up to 30 work packages

# ================================================================ Documents
ws = wb.create_sheet("Documents", 1)
HEAD = ["Work Package", "Work Package Reference", "Document ID", "Document Title", "Document Owner",
        "Status", "Status %", "Baseline approval date", "Forecast approval date", "Start date",
        "Outline approved date", "Draft in review date",
        "Sent for approval date", "Sent for sign-off date",
        "Approved and filed date", "Key Deliverable", "Design Review Set", "Comment / Note",
        "Dash WP selected", "Dash in set"]
col_of = {h: CL(i + 1) for i, h in enumerate(HEAD)}
ws.append(HEAD)
for r in rows:
    ws.append([r["wp"], r["ref"], r["id"], r["title"], r["owner"], r["status"], None,
               r["baseline"], r["forecast"], r["start"], *r["ms"], r["key"], r["dr"], None, None, None])
last = len(rows) + 1
SEL_D4 = "Dashboard!$D$4"
SEL_SET = "Dashboard!$H$4"
for i in range(2, last + 1):
    ws.cell(i, 7, f'=IFERROR(INDEX(Lists!$B$2:$B$8,MATCH(F{i},Lists!$A$2:$A$8,0)),0)')
    ws.cell(i, 19, (f'=IF({SEL_D4}="All",1,IF({SEL_D4}="{CUSTOM}",'
                    f'IF(COUNTIFS(Dashboard!$R${PANEL_TOP}:$R${PANEL_END},A{i},Dashboard!$S${PANEL_TOP}:$S${PANEL_END},"Yes")>0,1,0),'
                    f'IF(A{i}={SEL_D4},1,0)))'))
    parts = ",".join(f'AND({SEL_SET}="{n}",{col_of[c]}{i}="{v}")' for n, c, v in SETS[1:])
    ws.cell(i, 20, f'=IF({SEL_SET}="{SETS[0][0]}",1,IF(OR({parts}),1,0))')
LASTCOL = CL(len(HEAD))
tab = Table(displayName="tblDocuments", ref=f"A1:{LASTCOL}{last}")
tab.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
ws.add_table(tab)
colors = [NAVY] * 5 + [GREEN] * 10 + [ORANGE, ORANGE, PURPLE, GREY, GREY]
widths = [30, 12, 18, 30, 14, 34, 9, 13, 13, 12, 14, 14, 16, 14, 16, 9, 10, 24, 10, 10]
for j, (colr, w) in enumerate(zip(colors, widths), start=1):
    c = ws.cell(1, j)
    c.font = font(bold=True, color="FFFFFF"); c.fill = fill(colr)
    c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.column_dimensions[CL(j)].width = w
ws.row_dimensions[1].height = 45
for row in ws.iter_rows(min_row=2, max_row=last):
    for c in row:
        c.font = font(9)
    for c in row[7:15]:
        c.number_format = "dd/mm/yyyy"
    for c in row[18:20]:
        c.font = font(9, color=GREY)
ws.freeze_panes = "D2"
dv = DataValidation(type="list", formula1="=Lists!$A$2:$A$8", allow_blank=True)
ws.add_data_validation(dv); dv.add(f"F2:F{last}")
dv2 = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
ws.add_data_validation(dv2); dv2.add(f"P2:Q{last}")
ws.conditional_formatting.add(f"G2:G{last}", DataBarRule(start_type="num", start_value=0, end_type="num",
                                                         end_value=100, color="63BE7B"))
ws["A1"].comment = Comment("FICTITIOUS DEMO DATA, same structure as a typical document tracker. "
                           "'Baseline approval date' = agreed at start; 'Forecast approval date' = planned date "
                           "as updated by the owner. The two grey 'Dash' columns are calculated helper columns "
                           "that drive the dashboard filters – don't type in them.", "Claude")

# ================================================================ shared formula pieces
T = "tblDocuments"
FLT = f'{T}[Dash WP selected],1,{T}[Dash in set],1'
MASK = f'({T}[Dash WP selected]=1)*({T}[Dash in set]=1)'
PEIL = "Dashboard!$L$4"
TOTAL = "Dashboard!$B$8"

# ================================================================ S-curve data
sc = wb.create_sheet("S-curve data")
sc["A1"] = "S-curve data (weekly) – follows the filters on the Dashboard tab"
sc["A1"].font = font(12, True, NAVY)
sc["A2"] = ("Approved = 'Approved and filed date' filled. Forecast = approved so far + open documents "
            "by their forecast date (overdue ones count from the status date).")
sc["A2"].font = font(9, italic=True, color=GREY)
head(sc, 4, 1, ["Week", "Baseline (agreed at start)", "Actual (status today)", "Forecast (planned dates updated)",
                "Baseline %", "Actual %", "Forecast %", "Earned progress % (weighted)"], NAVY)
sc.row_dimensions[4].height = 42
weeks, d = [], dt.date(2026, 7, 6)
while d <= dt.date(2027, 6, 28):
    weeks.append(d); d += dt.timedelta(days=7)
MS_COLS = ["Outline approved date", "Draft in review date",
           "Sent for approval date", "Sent for sign-off date",
           "Approved and filed date"]
INC = [10, 40, 20, 20, 10]
for k, wk in enumerate(weeks):
    r = 5 + k
    sc.cell(r, 1, wk).number_format = "dd/mm/yy"
    sc.cell(r, 2, f'=COUNTIFS({FLT},{T}[Baseline approval date],"<="&A{r})')
    sc.cell(r, 3, f'=IF(A{r}<={PEIL},COUNTIFS({FLT},{T}[Approved and filed date],"<="&A{r}),NA())')
    sc.cell(r, 4, (f'=IF(A{r}+7>{PEIL},IF(A{r}<={PEIL},C{r},'
                   f'COUNTIFS({FLT},{T}[Approved and filed date],"<="&{PEIL})'
                   f'+COUNTIFS({FLT},{T}[Approved and filed date],"=",{T}[Forecast approval date],"<="&A{r})),NA())'))
    for j, src in ((5, "B"), (6, "C"), (7, "D")):
        sc.cell(r, j, f'=IF(ISNA({src}{r}),NA(),IF({TOTAL}=0,0,{src}{r}/{TOTAL}))').number_format = "0%"
    parts = "+".join(f'{w}*COUNTIFS({FLT},{T}[{m}],"<="&A{r})' for w, m in zip(INC, MS_COLS))
    sc.cell(r, 8, f'=IF(A{r}<={PEIL},IF({TOTAL}=0,0,({parts})/(100*{TOTAL})),NA())').number_format = "0%"
    for j in range(1, 9):
        sc.cell(r, j).font = font(9)
LAST_W = 4 + len(weeks)
for col, w in zip("ABCDEFGH", [11, 13, 13, 15, 10, 10, 10, 14]):
    sc.column_dimensions[col].width = w
sc.freeze_panes = "B5"

# ================================================================ Dashboard content
db.sheet_view.showGridLines = False
for j in range(1, 21):
    db.column_dimensions[CL(j)].width = 13
db.column_dimensions["A"].width = 2
db.column_dimensions["B"].width = 34
db.column_dimensions["Q"].width = 3
db.column_dimensions["R"].width = 32
db.column_dimensions["S"].width = 9

db["B1"] = "WBS Document Progress – Dashboard"
db["B1"].font = font(18, True, NAVY)
db["B2"] = ('="Showing: "&IF($D$4="All","all work packages",IF($D$4="' + CUSTOM + '",'
            '"custom WP selection ("&COUNTIF($S$' + str(PANEL_TOP) + ':$S$' + str(PANEL_END) + ',"Yes")&" WPs)",$D$4))'
            '&"  ·  "&$H$4&"  ·  status date "&TEXT($L$4,"dd/mm/yyyy")')
db["B2"].font = font(10, True, GREY)

for cell, label in (("B4", "Work Package"), ("G4", "Document set"), ("K4", "Status date")):
    db[cell] = label
    db[cell].font = font(10, True, NAVY)
    db[cell].alignment = Alignment(horizontal="right")
db["D4"] = "All"; db["H4"] = SETS[0][0]; db["L4"] = "=TODAY()"
db.merge_cells("D4:F4"); db.merge_cells("H4:I4")
for c in ("D4", "H4", "L4"):
    db[c].fill = fill("FFF2CC"); db[c].font = font(10, True); db[c].border = BOX
    db[c].alignment = Alignment(horizontal="center")
db["L4"].number_format = "dd/mm/yyyy"
db["M4"] = "← yellow cells are filters"
db["M4"].font = font(8, italic=True, color=GREY)
v1 = DataValidation(type="list", formula1=f"=Lists!$D$2:$D${len(wp_opts)+1}"); db.add_data_validation(v1); v1.add("D4")
v2 = DataValidation(type="list", formula1=f"=Lists!$F$2:$F${len(SETS)+1}"); db.add_data_validation(v2); v2.add("H4")

# WP selection panel
db[f"R{PANEL_TOP-2}"] = "WP selection"
db[f"R{PANEL_TOP-2}"].font = font(11, True, NAVY)
db[f"R{PANEL_TOP-2+1}"] = f'Used when Work Package = "{CUSTOM}". Set "Yes" to include.'
db[f"R{PANEL_TOP-2+1}"].font = font(8, italic=True, color=GREY)
dv3 = DataValidation(type="list", formula1='"Yes"', allow_blank=True); db.add_data_validation(dv3)
for i, (wp, _, _) in enumerate(WPS):
    r = PANEL_TOP + i
    db.cell(r, 18, wp).font = font(9)
    c = db.cell(r, 19, "Yes" if wp.startswith(("WP02", "WP03")) else None)
    c.fill = fill("FFF2CC"); c.font = font(9, True); c.alignment = Alignment(horizontal="center")
    db.cell(r, 18).border = BOX; c.border = BOX
dv3.add(f"S{PANEL_TOP}:S{PANEL_END}")

# KPI tiles
kpis = [
    ("B", "Total documents", f"=COUNTIFS({FLT})", "0", '="in selection (project target 300)"'),
    ("D", "Approved & filed", f'=COUNTIFS({FLT},{T}[Approved and filed date],"<="&$L$4)', "0",
     '=IF(B8=0,"",TEXT(D8/B8,"0%")&" of total")'),
    ("F", "In review / approval", "=" + "+".join(f'COUNTIFS({FLT},{T}[Status],"{s}")' for s in MILESTONES[1:4]), "0",
     '="review → sign-off"'),
    ("H", "Not started", f'=COUNTIFS({FLT},{T}[Status],"Not started")', "0",
     f'=COUNTIFS({FLT},{T}[Status],"Started")&" started, no milestone yet"'),
    ("J", "Weighted progress", f'=IF(B8=0,0,SUMPRODUCT({MASK}*{T}[Status %])/(100*B8))', "0%",
     f'="baseline today: "&TEXT(IF(B8=0,0,COUNTIFS({FLT},{T}[Baseline approval date],"<="&$L$4)/B8),"0%")&" approved"'),
    ("L", "Overdue", f'=COUNTIFS({FLT},{T}[Approved and filed date],"=",{T}[Forecast approval date],"<"&$L$4)',
     "0", '="open, forecast date passed"'),
    ("N", "Behind baseline", f'=COUNTIFS({FLT},{T}[Baseline approval date],"<="&$L$4)-COUNTIFS({FLT},{T}[Baseline approval date],"<="&$L$4,{T}[Approved and filed date],"<="&$L$4)', "0",
     '="docs vs. baseline plan today"'),
]
for col, label, f, fmt, sub in kpis:
    c2 = chr(ord(col) + 1)
    cols = (col,) if col == "B" else (col, c2)
    if col != "B":
        for r in (6, 7, 8, 9):
            db.merge_cells(f"{col}{r}:{c2}{r}")
    db[f"{col}6"], db[f"{col}8"], db[f"{col}9"] = label, f, sub
    db[f"{col}6"].font = font(9, True, GREY)
    db[f"{col}8"].font = font(22, True, RED if col in "LN" else NAVY)
    db[f"{col}8"].number_format = fmt
    db[f"{col}9"].font = font(8, color=GREY)
    for r in (6, 7, 8, 9):
        db[f"{col}{r}"].alignment = Alignment(horizontal="center", vertical="center")
        for cc in cols:
            db[f"{cc}{r}"].fill = fill(LIGHT)
db.row_dimensions[7].height = 4
db.row_dimensions[8].height = 34

# S-curve chart
db["B11"] = "S-curve – documents approved (baseline vs. actual + forecast)"
db["B11"].font = font(12, True, NAVY)
lc = LineChart()
lc.height, lc.width = 9, 31
lc.y_axis.title = "Documents approved (cumulative)"
lc.y_axis.majorGridlines = None
lc.x_axis.number_format = "mmm yy"
lc.legend.position = "b"
lc.add_data(Reference(sc, min_col=2, max_col=4, min_row=4, max_row=LAST_W), titles_from_data=True)
lc.set_categories(Reference(sc, min_col=1, min_row=5, max_row=LAST_W))
for s, (colr, dash, w) in zip(lc.series, [("A6A6A6", "sysDash", 22000), ("1F4E79", "solid", 32000),
                                          ("1F4E79", "dash", 26000)]):
    s.graphicalProperties.line.solidFill = colr
    s.graphicalProperties.line.prstDash = dash
    s.graphicalProperties.line.width = w
    s.smooth = False
    s.marker.symbol = "none"
lc.x_axis.delete = False; lc.y_axis.delete = False
lc.display_blanks = "gap"
db.add_chart(lc, "B13")

# Document flow
R0 = 32
db[f"B{R0}"] = "Document flow – where are the documents now?"
db[f"B{R0}"].font = font(12, True, NAVY)
head(db, R0 + 1, 2, ["Status", "Documents", "% of total", "Status %"], NAVY)
for i, (s, w) in enumerate(zip(STATUSES, WEIGHTS)):
    r = R0 + 2 + i
    db.cell(r, 2, s).font = font(9)
    db.cell(r, 3, f'=COUNTIFS({FLT},{T}[Status],B{r})').font = font(9, True)
    c = db.cell(r, 4, f"=IF($B$8=0,0,C{r}/$B$8)"); c.number_format = "0%"; c.font = font(9)
    db.cell(r, 5, w).font = font(9, color="0000FF")
    for j in range(2, 6):
        db.cell(r, j).border = BOX
r_tot = R0 + 2 + len(STATUSES)
db.cell(r_tot, 2, "Total").font = font(9, True)
db.cell(r_tot, 3, f"=SUM(C{R0+2}:C{r_tot-1})").font = font(9, True)
db.conditional_formatting.add(f"D{R0+2}:D{r_tot-1}", DataBarRule(start_type="num", start_value=0,
                                                                  end_type="num", end_value=1, color="5B9BD5"))
bc = BarChart(); bc.type = "bar"
bc.height, bc.width = 6.5, 17
bc.add_data(Reference(db, min_col=3, min_row=R0 + 1, max_row=r_tot - 1), titles_from_data=True)
bc.set_categories(Reference(db, min_col=2, min_row=R0 + 2, max_row=r_tot - 1))
bc.legend = None
bc.y_axis.majorGridlines = None
bc.x_axis.scaling.orientation = "maxMin"
bc.series[0].graphicalProperties.solidFill = "2E75B6"
bc.x_axis.delete = False; bc.y_axis.delete = False
db.add_chart(bc, f"G{R0}")


def status_matrix(top, title, labels, crit_fn, color, first_head):
    """labels: row labels; crit_fn(row_index, r) -> COUNTIFS criteria string (without status)."""
    db[f"B{top}"] = title
    db[f"B{top}"].font = font(12, True, NAVY)
    head(db, top + 1, 2, [first_head] + STATUSES + ["Total", "Approved %", "Weighted %"], color, 8)
    db.row_dimensions[top + 1].height = 48
    for i, lab in enumerate(labels):
        r = top + 2 + i
        db.cell(r, 2, lab).font = font(9)
        crit = crit_fn(i, r)
        for j in range(len(STATUSES)):
            L = CL(3 + j)
            db.cell(r, 3 + j, f'=COUNTIFS({crit},{T}[Status],{L}${top+1})')
        db.cell(r, 10, f"=SUM(C{r}:I{r})")
        db.cell(r, 11, f"=IF(J{r}=0,0,I{r}/J{r})").number_format = "0%"
        w = "+".join(f'{wt}*{CL(3+j)}{r}' for j, wt in enumerate(WEIGHTS) if wt)
        db.cell(r, 12, f"=IF(J{r}=0,0,({w})/(100*J{r}))").number_format = "0%"
        for j in range(2, 13):
            db.cell(r, j).border = BOX
            if j > 2:
                db.cell(r, j).font = font(9, j == 10)
                db.cell(r, j).alignment = Alignment(horizontal="center")
    rt = top + 2 + len(labels)
    db.conditional_formatting.add(f"K{top+2}:L{rt-1}", DataBarRule(start_type="num", start_value=0,
                                                                    end_type="num", end_value=1, color="63BE7B"))
    ch = BarChart(); ch.type = "bar"; ch.grouping = "stacked"; ch.overlap = 100
    ch.height, ch.width = max(5.5, 1.0 * len(labels) + 3.5), 17
    ch.add_data(Reference(db, min_col=3, max_col=9, min_row=top + 1, max_row=rt - 1), titles_from_data=True)
    ch.set_categories(Reference(db, min_col=2, min_row=top + 2, max_row=rt - 1))
    for s, colr in zip(ch.series, STATUS_COLORS):
        s.graphicalProperties.solidFill = colr
        s.graphicalProperties.line.noFill = True
    ch.x_axis.scaling.orientation = "maxMin"
    ch.y_axis.majorGridlines = None
    ch.legend.position = "b"
    ch.x_axis.delete = False; ch.y_axis.delete = False
    db.add_chart(ch, f"N{top}")
    return rt


# matrix 1: per WP (all WPs side by side), follows the document-set filter
R1 = r_tot + 6
wp_names = [w[0] for w in WPS]
e1 = status_matrix(R1, None, wp_names,
                   lambda i, r: f'{T}[Work Package],$B{r},{T}[Dash in set],1', NAVY, "Work Package")
db[f"B{R1}"] = '="Status per Work Package – "&$H$4&" (all WPs side by side)"'
db[f"B{R1}"].font = font(12, True, NAVY)
rt = e1
db.cell(rt, 2, "Total (all WPs)").font = font(9, True)
for j in range(3, 11):
    L = CL(j)
    db.cell(rt, j, f"=SUM({L}{R1+2}:{L}{rt-1})").font = font(9, True)
    db.cell(rt, j).alignment = Alignment(horizontal="center")
db.cell(rt, 11, f"=IF(J{rt}=0,0,I{rt}/J{rt})").number_format = "0%"
db.cell(rt, 12, f"=IF(J{rt}=0,0,SUMPRODUCT(J{R1+2}:J{rt-1},L{R1+2}:L{rt-1})/J{rt})").number_format = "0%"
for j in range(2, 13):
    db.cell(rt, j).border = BOX; db.cell(rt, j).fill = fill(LIGHT)

# matrix 2: per document set, follows the WP selection
R2 = max(e1 + 4, R1 + 26)


def set_crit(i, r):
    n, c, v = SETS[i]
    base = f'{T}[Dash WP selected],1'
    return base if c is None else f'{base},{T}[{c}],"{v}"'


e2 = status_matrix(R2, None, [s[0] for s in SETS], set_crit, ORANGE, "Document set")
db[f"B{R2}"] = ('="Status per document set – "&IF($D$4="All","all work packages",IF($D$4="' + CUSTOM +
                '","custom WP selection",$D$4))')
db[f"B{R2}"].font = font(12, True, NAVY)
for k in range(len(SETS)):
    db.cell(R2 + 2 + k, 2).value = f"=Lists!F{2+k}"

# Attention list
R3 = max(e2 + 4, R2 + 16)
db[f"B{R3}"] = "Attention list – open documents whose forecast approval date has passed (oldest first, max 15)"
db[f"B{R3}"].font = font(12, True, RED)
ocols = [2, 3, 5, 7, 8, 11, 12]
for h, c in zip(["Document ID", "Document Title", "Work Package", "Owner", "Status", "Forecast date", "Days late"], ocols):
    x = db.cell(R3 + 1, c, h); x.font = font(9, True, "FFFFFF"); x.fill = fill(RED)
for a, b in ((3, 4), (5, 6), (8, 10)):
    db.merge_cells(start_row=R3 + 1, start_column=a, end_row=R3 + 1, end_column=b)
HROW = f"ROW({T}[#Headers])"
COND = f'{MASK}*({T}[Approved and filed date]="")*({T}[Forecast approval date]<$L$4)'
for k in range(1, 16):
    r = R3 + 1 + k
    key = f"$N{r}"
    db[f"N{r}"] = ArrayFormula(f"N{r}", f'=IFERROR(SMALL(IF({COND},{T}[Forecast approval date]*10000+ROW({T}[Document ID])-{HROW}),{k}),"")')
    db[f"N{r}"].font = font(7, color="FFFFFF")
    idx = f"MOD({key},10000)"

    def g(col):
        return f'=IF({key}="","",INDEX({T}[{col}],{idx}))'
    db.cell(r, 2, g("Document ID"))
    db.cell(r, 3, g("Document Title")); db.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    db.cell(r, 5, g("Work Package")); db.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
    db.cell(r, 7, g("Document Owner"))
    db.cell(r, 8, g("Status")); db.merge_cells(start_row=r, start_column=8, end_row=r, end_column=10)
    db.cell(r, 11, g("Forecast approval date")).number_format = "dd/mm/yyyy"
    db.cell(r, 12, f'=IF({key}="","",$L$4-INT({key}/10000))').number_format = "0"
    for c in ocols:
        db.cell(r, c).font = font(9); db.cell(r, c).border = BOX
db.cell(R3 + 17, 2, "Hidden helper keys in column N (white text) – do not delete.").font = font(7, italic=True, color=GREY)

db.freeze_panes = "A5"
wb.move_sheet("S-curve data", offset=0)
wb.save(OUT)
print("saved", R0, R1, R2, R3)
