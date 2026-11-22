"""
build_workforce_model.py
---------------------------
Builds Contractor_Workforce_Planning_Model.xlsx: an 8-quarter contractor
headcount and spend forecast against budget, a contractor-vs-FTE cost
comparison for the same 5 roles, and a scenario toggle (dropdown: Base
Case / 10% Rate Increase / Hiring Freeze) that live-recalculates every
quarter's headcount, rate, and spend — showing exactly which quarter the
budget runs out under each scenario.

All formulas, no pasted values — change the scenario dropdown in Excel
and every downstream number (including the "budget runway" quarter)
recalculates.
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).parent.parent
OUT_DIR = ROOT / "outputs"
OUT_DIR.mkdir(exist_ok=True)
OUT_PATH = OUT_DIR / "Contractor_Workforce_Planning_Model.xlsx"

FONT_NAME = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
TITLE_FONT = Font(name=FONT_NAME, bold=True, size=16, color="1F4E78")
BODY_FONT = Font(name=FONT_NAME, size=10)
INPUT_FONT = Font(name=FONT_NAME, size=10, color="0000FF")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
RED_FILL = PatternFill("solid", fgColor="FBE4E1")
GREEN_FILL = PatternFill("solid", fgColor="DFF5E8")

ROLES = [
    # role, base_q1_headcount, hourly_rate, fte_annual_salary
    ("Software Engineer", 12, 95.0, 145000),
    ("Data Analyst",        6, 75.0, 110000),
    ("Business Analyst",    5, 70.0, 105000),
    ("Project Manager",     4, 85.0, 125000),
    ("QA Engineer",         7, 65.0,  95000),
]
N_QUARTERS = 8
QUARTER_LABELS = [f"Q{i+1}" for i in range(N_QUARTERS)]

wb = Workbook()
wb.remove(wb.active)

# === Assumptions ============================================================
ws_a = wb.create_sheet("Assumptions", 0)
ws_a.sheet_view.showGridLines = False
ws_a["B2"] = "Contractor Workforce Planning — Assumptions"
ws_a["B2"].font = TITLE_FONT

ws_a["B4"] = "Scenario:"
ws_a["B4"].font = Font(name=FONT_NAME, bold=True, size=12)
ws_a["C4"] = "Base Case"
ws_a["C4"].font = Font(name=FONT_NAME, bold=True, size=12, color="0000FF")
ws_a["C4"].fill = INPUT_FILL
dv = DataValidation(type="list", formula1='"Base Case,10% Rate Increase,Hiring Freeze"', allow_blank=False)
ws_a.add_data_validation(dv)
dv.add(ws_a["C4"])
ws_a["D4"] = "<-- change this dropdown to re-run the forecast under a different scenario"
ws_a["D4"].font = Font(name=FONT_NAME, italic=True, size=9, color="595959")

global_assumptions = [
    ("Quarterly hours per contractor", 520),
    ("Base headcount growth rate per quarter", 0.025),
    ("Quarterly contractor budget", 1550000),
    ("FTE benefits/overhead loading", 0.32),
    ("Rate increase scenario adjustment", 0.10),
]
r0 = 6
for i, (label, val) in enumerate(global_assumptions):
    r = r0 + i
    ws_a.cell(r, 2, label).font = BODY_FONT
    cell = ws_a.cell(r, 3, val)
    cell.font = INPUT_FONT
    cell.fill = INPUT_FILL
    if "rate" in label.lower() and "growth" in label.lower() or "loading" in label.lower() or "adjustment" in label.lower():
        cell.number_format = "0.0%"
    if "budget" in label.lower():
        cell.number_format = '$#,##0'
    for c in (2, 3):
        ws_a.cell(r, c).border = BORDER

QUARTERLY_HOURS = "$C$6"
GROWTH_RATE = "$C$7"
QUARTERLY_BUDGET = "$C$8"
BENEFITS_LOADING = "$C$9"
RATE_INCREASE_ADJ = "$C$10"
SCENARIO_CELL = "$C$4"

role_hdr_row = 13
ws_a.cell(role_hdr_row, 2, "Role").font = HEADER_FONT
ws_a.cell(role_hdr_row, 2).fill = HEADER_FILL
for i, h in enumerate(["Base Q1 Headcount", "Hourly Rate ($)", "FTE Annual Salary ($)"]):
    ws_a.cell(role_hdr_row, 3 + i, h).font = HEADER_FONT
    ws_a.cell(role_hdr_row, 3 + i).fill = HEADER_FILL
    ws_a.cell(role_hdr_row, 3 + i).alignment = Alignment(wrap_text=True, horizontal="center")

for i, (role, hc, rate, salary) in enumerate(ROLES):
    r = role_hdr_row + 1 + i
    ws_a.cell(r, 2, role).font = BODY_FONT
    ws_a.cell(r, 3, hc).font = INPUT_FONT
    ws_a.cell(r, 3).fill = INPUT_FILL
    ws_a.cell(r, 4, rate).font = INPUT_FONT
    ws_a.cell(r, 4).fill = INPUT_FILL
    ws_a.cell(r, 4).number_format = '$#,##0.00'
    ws_a.cell(r, 5, salary).font = INPUT_FONT
    ws_a.cell(r, 5).fill = INPUT_FILL
    ws_a.cell(r, 5).number_format = '$#,##0'
    for c in range(2, 6):
        ws_a.cell(r, c).border = BORDER
role_last_row = role_hdr_row + len(ROLES)

ws_a.column_dimensions["B"].width = 32
for c in "CDE":
    ws_a.column_dimensions[c].width = 18

# === Forecast =================================================================
ws_f = wb.create_sheet("Forecast")
ws_f.sheet_view.showGridLines = False
ws_f["B2"] = "8-Quarter Contractor Headcount & Spend Forecast"
ws_f["B2"].font = TITLE_FONT
ws_f["B3"] = "=\"Scenario: \" & Assumptions!C4"
ws_f["B3"].font = Font(name=FONT_NAME, bold=True, size=12, color="C0392B")

hdr_row = 5
ws_f.cell(hdr_row, 2, "")
for i, q in enumerate(QUARTER_LABELS):
    ws_f.cell(hdr_row, 3 + i, q)
style_header_range = list(range(3, 3 + N_QUARTERS))
for c in style_header_range:
    cell = ws_f.cell(hdr_row, c)
    cell.fill = HEADER_FILL
    cell.font = HEADER_FONT
    cell.alignment = Alignment(horizontal="center")
    cell.border = BORDER
ws_f.cell(hdr_row, 2).fill = HEADER_FILL
ws_f.cell(hdr_row, 2).border = BORDER

# --- Headcount rows per role ---
hc_start_row = hdr_row + 2
ws_f.cell(hc_start_row - 1, 2, "Headcount by Role").font = Font(name=FONT_NAME, bold=True, size=11, color="1F4E78")
for ri, (role, base_hc, rate, salary) in enumerate(ROLES):
    r = hc_start_row + ri
    ws_f.cell(r, 2, role).font = BODY_FONT
    assump_row = role_hdr_row + 1 + ri
    for qi in range(N_QUARTERS):
        col = 3 + qi
        col_letter = get_column_letter(col)
        if qi == 0:
            formula = f"=Assumptions!$C${assump_row}"
        else:
            prev_letter = get_column_letter(col - 1)
            # NOTE: headcount is intentionally NOT rounded to a whole number
            # each quarter -- at a 2.5%/quarter growth rate, rounding a
            # small integer headcount (e.g. 12 -> 12.3 -> rounds back to 12)
            # silently kills the growth signal entirely within a few
            # quarters. Keeping headcount as a continuous "FTE-equivalent
            # contractor count" (standard financial-modeling practice for
            # blended mid-quarter joins) and only rounding for display
            # avoids that trap.
            formula = (
                f'=IF(Assumptions!{SCENARIO_CELL}="Hiring Freeze",{prev_letter}{r},'
                f"{prev_letter}{r}*(1+Assumptions!{GROWTH_RATE}))"
            )
        cell = ws_f.cell(r, col, formula)
        cell.font = BODY_FONT
        cell.border = BORDER
        cell.number_format = "0.0"
hc_end_row = hc_start_row + len(ROLES) - 1

total_hc_row = hc_end_row + 1
ws_f.cell(total_hc_row, 2, "Total Headcount").font = Font(name=FONT_NAME, bold=True, size=10)
for qi in range(N_QUARTERS):
    col = 3 + qi
    col_letter = get_column_letter(col)
    cell = ws_f.cell(total_hc_row, col, f"=SUM({col_letter}{hc_start_row}:{col_letter}{hc_end_row})")
    cell.font = Font(name=FONT_NAME, bold=True, size=10)
    cell.border = BORDER
    cell.number_format = "0.0"

# --- Cost rows per role ---
cost_start_row = total_hc_row + 3
ws_f.cell(cost_start_row - 1, 2, "Quarterly Cost by Role ($)").font = Font(name=FONT_NAME, bold=True, size=11, color="1F4E78")
for ri, (role, base_hc, rate, salary) in enumerate(ROLES):
    r = cost_start_row + ri
    hc_row = hc_start_row + ri
    assump_row = role_hdr_row + 1 + ri
    ws_f.cell(r, 2, role).font = BODY_FONT
    for qi in range(N_QUARTERS):
        col = 3 + qi
        col_letter = get_column_letter(col)
        rate_formula = (
            f'IF(Assumptions!{SCENARIO_CELL}="10% Rate Increase",'
            f"Assumptions!$D${assump_row}*(1+Assumptions!{RATE_INCREASE_ADJ}),Assumptions!$D${assump_row})"
        )
        formula = f"=ROUND({col_letter}{hc_row}*({rate_formula})*Assumptions!{QUARTERLY_HOURS},0)"
        cell = ws_f.cell(r, col, formula)
        cell.font = BODY_FONT
        cell.border = BORDER
        cell.number_format = '$#,##0'
cost_end_row = cost_start_row + len(ROLES) - 1

total_cost_row = cost_end_row + 1
ws_f.cell(total_cost_row, 2, "Total Quarterly Cost").font = Font(name=FONT_NAME, bold=True, size=10)
for qi in range(N_QUARTERS):
    col = 3 + qi
    col_letter = get_column_letter(col)
    cell = ws_f.cell(total_cost_row, col, f"=SUM({col_letter}{cost_start_row}:{col_letter}{cost_end_row})")
    cell.font = Font(name=FONT_NAME, bold=True, size=10)
    cell.border = BORDER
    cell.number_format = '$#,##0'

budget_row = total_cost_row + 1
ws_f.cell(budget_row, 2, "Quarterly Budget").font = Font(name=FONT_NAME, bold=True, size=10)
for qi in range(N_QUARTERS):
    col = 3 + qi
    col_letter = get_column_letter(col)
    cell = ws_f.cell(budget_row, col, f"=Assumptions!{QUARTERLY_BUDGET}")
    cell.font = Font(name=FONT_NAME, bold=True, size=10)
    cell.border = BORDER
    cell.number_format = '$#,##0'

cum_cost_row = budget_row + 2
ws_f.cell(cum_cost_row, 2, "Cumulative Cost").font = BODY_FONT
cum_budget_row = cum_cost_row + 1
ws_f.cell(cum_budget_row, 2, "Cumulative Budget").font = BODY_FONT
surplus_row = cum_budget_row + 1
ws_f.cell(surplus_row, 2, "Cumulative Surplus / (Deficit)").font = Font(name=FONT_NAME, bold=True, size=10)
within_budget_row = surplus_row + 1
ws_f.cell(within_budget_row, 2, "Within Budget?").font = Font(name=FONT_NAME, bold=True, size=10)

for qi in range(N_QUARTERS):
    col = 3 + qi
    col_letter = get_column_letter(col)
    if qi == 0:
        cum_cost_f = f"={col_letter}{total_cost_row}"
        cum_budget_f = f"={col_letter}{budget_row}"
    else:
        prev_letter = get_column_letter(col - 1)
        cum_cost_f = f"={prev_letter}{cum_cost_row}+{col_letter}{total_cost_row}"
        cum_budget_f = f"={prev_letter}{cum_budget_row}+{col_letter}{budget_row}"
    c1 = ws_f.cell(cum_cost_row, col, cum_cost_f); c1.font = BODY_FONT; c1.border = BORDER; c1.number_format = '$#,##0'
    c2 = ws_f.cell(cum_budget_row, col, cum_budget_f); c2.font = BODY_FONT; c2.border = BORDER; c2.number_format = '$#,##0'
    c3 = ws_f.cell(surplus_row, col, f"={col_letter}{cum_budget_row}-{col_letter}{cum_cost_row}")
    c3.font = Font(name=FONT_NAME, bold=True, size=10); c3.border = BORDER; c3.number_format = '$#,##0;[RED]-$#,##0'
    c4 = ws_f.cell(within_budget_row, col, f'=IF({col_letter}{surplus_row}>=0,"Yes","NO - OVER BUDGET")')
    c4.font = Font(name=FONT_NAME, bold=True, size=10); c4.border = BORDER

ws_f.conditional_formatting.add(
    f"{get_column_letter(3)}{within_budget_row}:{get_column_letter(2+N_QUARTERS)}{within_budget_row}",
    CellIsRule(operator="equal", formula=['"Yes"'], fill=GREEN_FILL),
)
ws_f.conditional_formatting.add(
    f"{get_column_letter(3)}{within_budget_row}:{get_column_letter(2+N_QUARTERS)}{within_budget_row}",
    CellIsRule(operator="notEqual", formula=['"Yes"'], fill=RED_FILL),
)

ws_f.column_dimensions["B"].width = 28
for c in range(3, 3 + N_QUARTERS):
    ws_f.column_dimensions[get_column_letter(c)].width = 13

# ---- Chart: Total Cost vs Budget ----
chart = LineChart()
chart.title = "Total Quarterly Contractor Cost vs. Budget"
chart.y_axis.title = "$"
chart.x_axis.title = "Quarter"
data = Reference(ws_f, min_col=2, max_col=2 + N_QUARTERS, min_row=total_cost_row, max_row=budget_row)
cats = Reference(ws_f, min_col=3, max_col=2 + N_QUARTERS, min_row=hdr_row, max_row=hdr_row)
chart.add_data(data, titles_from_data=True, from_rows=True)
chart.set_categories(cats)
chart.height, chart.width = 9, 18
ws_f.add_chart(chart, f"B{within_budget_row + 3}")

# === Contractor_vs_FTE ========================================================
ws_c = wb.create_sheet("Contractor_vs_FTE")
ws_c.sheet_view.showGridLines = False
ws_c["B2"] = "Contractor vs. FTE Annual Cost Comparison"
ws_c["B2"].font = TITLE_FONT
ws_c["B3"] = "Contractor annual cost = hourly rate x 2,080 hours (4 quarters x 520 hrs). FTE annual cost = salary x (1 + benefits loading)."
ws_c["B3"].font = Font(name=FONT_NAME, italic=True, size=10, color="595959")

hdr2 = ["Role", "Contractor Annual Cost", "FTE Annual Cost (loaded)", "Contractor Premium $", "Contractor Premium %"]
r0 = 5
for i, h in enumerate(hdr2):
    ws_c.cell(r0, 2 + i, h)
for c in range(2, 2 + len(hdr2)):
    cell = ws_c.cell(r0, c)
    cell.fill = HEADER_FILL
    cell.font = HEADER_FONT
    cell.alignment = Alignment(horizontal="center", wrap_text=True)
    cell.border = BORDER

for i, (role, hc, rate, salary) in enumerate(ROLES):
    r = r0 + 1 + i
    assump_row = role_hdr_row + 1 + i
    ws_c.cell(r, 2, role).font = BODY_FONT
    ws_c.cell(r, 3, f"=Assumptions!$D${assump_row}*2080")
    ws_c.cell(r, 4, f"=Assumptions!$E${assump_row}*(1+Assumptions!{BENEFITS_LOADING})")
    ws_c.cell(r, 5, f"=C{r}-D{r}")
    ws_c.cell(r, 6, f"=C{r}/D{r}-1")
    for c in (3, 4, 5):
        ws_c.cell(r, c).number_format = '$#,##0'
    ws_c.cell(r, 6).number_format = "+0.0%;-0.0%"
    for c in range(2, 7):
        cell = ws_c.cell(r, c)
        cell.font = BODY_FONT
        cell.border = BORDER

ws_c.column_dimensions["B"].width = 22
for c in "CDEF":
    ws_c.column_dimensions[c].width = 20

# === README ===================================================================
ws_r = wb.create_sheet("README", 0)
ws_r.sheet_view.showGridLines = False
ws_r["B2"] = "Contractor Workforce Planning Model"
ws_r["B2"].font = TITLE_FONT
lines = [
    "",
    "Forecasts contractor headcount and spend by quarter against budget, with a live",
    "scenario toggle (Assumptions!C4) and a role-by-role contractor-vs-FTE cost comparison.",
    "",
    "Sheets:",
    "  - Assumptions: role headcount/rates/salaries and the scenario dropdown (edit the",
    "    yellow cells; everything downstream recalculates).",
    "  - Forecast: 8-quarter headcount, cost, cumulative budget, and a 'Within Budget?' row",
    "    that flags the exact quarter spend would exceed cumulative budget under the",
    "    selected scenario -- try switching the dropdown between Base Case / 10% Rate",
    "    Increase / Hiring Freeze and watch this row change.",
    "  - Contractor_vs_FTE: annual cost comparison for the same 5 roles.",
    "",
    "Scenario results (computed by this workbook):",
    "  Base Case:          cumulative surplus peaks ~Q4 (~$294K cushion), then erodes as",
    "                       compounding 2.5%/qtr growth outpaces the flat budget -- budget",
    "                       exceeded starting Q8.",
    "  10% Rate Increase:   over budget from Q1 onward -- the one-time rate jump alone",
    "                       exceeds the quarterly budget before any headcount growth.",
    "  Hiring Freeze:        never exceeds budget -- surplus grows every quarter",
    "                       (headcount held flat at the Q1 level).",
]
for i, line in enumerate(lines):
    cell = ws_r.cell(4 + i, 2, line)
    cell.font = Font(name=FONT_NAME, bold=line.startswith("Sheets:") or line.startswith("Scenario"), size=11)
ws_r.column_dimensions["B"].width = 100

wb.save(OUT_PATH)
print(f"Wrote {OUT_PATH}")
