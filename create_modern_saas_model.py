"""
create_modern_saas_model.py
----------------------------
Generates modern_saas_model.xlsx — a realistic SaaS financial model that
showcases Excel features that *break* regex-based formula parsing:

  • Named ranges   (DiscountRate, TaxRate, GrowthRates, …)
  • Excel Tables   (structured refs: Products[AnnualPrice], [@ProductID])
  • XLOOKUP        (cross-table lookups without helper columns)
  • LET            (intermediate variables inside a single formula)
  • SEQUENCE       (dynamic-array / spill formula)
  • UNIQUE / FILTER (dynamic-array functions)
  • Spill-range ref (A2# notation)
  • NPV / IRR      (financial functions that take named-range args)

Run:  python create_modern_saas_model.py
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.defined_name import DefinedName

# ── Colour palette ────────────────────────────────────────────────────────────
C_NAVY   = "1F3864"
C_BLUE   = "2E75B6"
C_LBLUE  = "D6E4F0"
C_GREEN  = "375623"
C_LGREEN = "E2EFDA"
C_AMBER  = "7F6000"
C_LAMBER = "FFF2CC"
C_WHITE  = "FFFFFF"
C_LGRAY  = "F2F2F2"
C_RED    = "FF0000"

# ── Helpers ───────────────────────────────────────────────────────────────────

def hfont(white=True, size=11, bold=True):
    return Font(bold=bold, color=C_WHITE if white else "000000", size=size)

def hfill(color=C_BLUE):
    return PatternFill("solid", fgColor=color)

def sc(ws, row, col, value=None, fmt=None, bold=False, bg=None,
       fg="000000", align=None):
    """Shorthand set-cell."""
    c = ws.cell(row=row, column=col)
    if value is not None:
        c.value = value
    if fmt:
        c.number_format = fmt
    if bold or fg != "000000":
        c.font = Font(bold=bold, color=fg)
    if bg:
        c.fill = PatternFill("solid", fgColor=bg)
    if align:
        c.alignment = Alignment(horizontal=align)
    return c

def col_widths(ws, pairs):
    """pairs = [(col_letter_or_int, width), …]"""
    for col, w in pairs:
        letter = col if isinstance(col, str) else get_column_letter(col)
        ws.column_dimensions[letter].width = w

def header_row(ws, row, labels, bg=C_BLUE, start_col=1):
    for i, lbl in enumerate(labels):
        c = ws.cell(row=row, column=start_col + i, value=lbl)
        c.font = hfont()
        c.fill = hfill(bg)
        c.alignment = Alignment(horizontal="center")

def section_title(ws, row, col, text, bg=C_NAVY):
    c = ws.cell(row=row, column=col, value=text)
    c.font = Font(bold=True, color=C_WHITE, size=12)
    c.fill = hfill(bg)
    return c

def add_named_range(wb, name, sheet, cell):
    """Add a single-cell named range.  cell = 'B4' (no $)."""
    ref = f"'{sheet}'!${cell[0]}${cell[1:]}"
    defn = DefinedName(name=name, attr_text=ref)
    wb.defined_names.add(defn)

def add_range_name(wb, name, sheet, ref):
    """Add a multi-cell named range.  ref = '$B$12:$B$16'"""
    defn = DefinedName(name=name, attr_text=f"'{sheet}'!{ref}")
    wb.defined_names.add(defn)

# ═════════════════════════════════════════════════════════════════════════════
# Build workbook
# ═════════════════════════════════════════════════════════════════════════════
wb = Workbook()

# ── Sheet 1: Assumptions ──────────────────────────────────────────────────────
ws_a = wb.active
ws_a.title = "Assumptions"

sc(ws_a, 1, 1, "SaaS Financial Model — Key Assumptions",
   bold=True, fg=C_NAVY)
ws_a.cell(1, 1).font = Font(bold=True, color=C_NAVY, size=14)

# -- Valuation Parameters --
section_title(ws_a, 3, 1, "Valuation Parameters")
header_row(ws_a, 4, ["Parameter", "Value", "Named Range"], bg=C_BLUE)

params = [
    ("Discount Rate (WACC)",    0.10,  "0.0%",    "DiscountRate",       "B5"),
    ("Tax Rate",                0.25,  "0.0%",    "TaxRate",            "B6"),
    ("Terminal Growth Rate",    0.03,  "0.0%",    "TerminalGrowthRate", "B7"),
    ("EBITDA Margin",           0.30,  "0.0%",    "EBITDAMargin",       "B8"),
    ("Capex % of Revenue",      0.05,  "0.0%",    "CapexPct",           "B9"),
    ("D&A % of Revenue",        0.04,  "0.0%",    "DandAPct",           "B10"),
]
for i, (lbl, val, fmt, nm, _cell) in enumerate(params):
    r = 5 + i
    sc(ws_a, r, 1, lbl)
    sc(ws_a, r, 2, val, fmt=fmt, bg=C_LBLUE)
    sc(ws_a, r, 3, nm, fg=C_AMBER)

# -- Revenue Growth Rates (named range GrowthRates) --
section_title(ws_a, 12, 1, "Annual Revenue Growth Rates")
header_row(ws_a, 13, ["Year", "Label", "Growth Rate"], bg=C_BLUE)
growth = [(1, "Year 1 — 2025", 0.40),
          (2, "Year 2 — 2026", 0.35),
          (3, "Year 3 — 2027", 0.28),
          (4, "Year 4 — 2028", 0.20),
          (5, "Year 5 — 2029", 0.15)]
for i, (yr, lbl, g) in enumerate(growth):
    r = 14 + i
    sc(ws_a, r, 1, yr, align="center")
    sc(ws_a, r, 2, lbl)
    sc(ws_a, r, 3, g, fmt="0.0%", bg=C_LGREEN)

sc(ws_a, 13 + len(growth) + 1, 1,
   "↑ Named range 'GrowthRates' = C14:C18", fg="888888")

# -- SaaS Operating Metrics --
section_title(ws_a, 21, 1, "SaaS Operating Metrics")
header_row(ws_a, 22, ["Metric", "Value", "Named Range"], bg=C_BLUE)
saas_params = [
    ("Base ARPU (Annual, $)",      12_000, "$#,##0",  "ARPUBase",      "B23"),
    ("Monthly Churn Rate",          0.020, "0.0%",    "ChurnRate",     "B24"),
    ("Starting Customer Count",        50, "#,##0",   "StartCustomers","B25"),
    ("Customer Acquisition Cost ($)", 8000, "$#,##0", "CAC",           "B26"),
]
for i, (lbl, val, fmt, nm, _cell) in enumerate(saas_params):
    r = 23 + i
    sc(ws_a, r, 1, lbl)
    sc(ws_a, r, 2, val, fmt=fmt, bg=C_LBLUE)
    sc(ws_a, r, 3, nm, fg=C_AMBER)

col_widths(ws_a, [("A", 30), ("B", 16), ("C", 22)])

# Register named ranges – single cells
named_cells = {
    "DiscountRate":       "B5",
    "TaxRate":            "B6",
    "TerminalGrowthRate": "B7",
    "EBITDAMargin":       "B8",
    "CapexPct":           "B9",
    "DandAPct":           "B10",
    "ARPUBase":           "B23",
    "ChurnRate":          "B24",
    "StartCustomers":     "B25",
    "CAC":                "B26",
}
for name, cell in named_cells.items():
    add_named_range(wb, name, "Assumptions", cell)

# Multi-cell range for growth rates
add_range_name(wb, "GrowthRates", "Assumptions", "$C$14:$C$18")

# ── Sheet 2: Products ─────────────────────────────────────────────────────────
ws_p = wb.create_sheet("Products")

sc(ws_p, 1, 1, "Product Catalog", bold=True, fg=C_NAVY)
ws_p.cell(1, 1).font = Font(bold=True, color=C_NAVY, size=14)
sc(ws_p, 2, 1, "Excel Table: 'Products'  |  Structured-ref demo: =1-[@COGSPct]",
   fg="888888")

header_row(ws_p, 3,
           ["ProductID", "ProductName", "Category",
            "AnnualPrice", "COGSPct", "GrossMargin"])

products = [
    ("P001", "Starter",          "SMB",        6_000,  0.15),
    ("P002", "Professional",     "Mid-Market", 18_000, 0.12),
    ("P003", "Enterprise",       "Enterprise", 60_000, 0.08),
    ("P004", "Analytics Add-on", "Add-on",      4_800, 0.20),
    ("P005", "API Access",       "Add-on",      3_600, 0.22),
]
for i, (pid, name, cat, price, cogs) in enumerate(products):
    r = 4 + i
    ws_p.cell(r, 1, pid)
    ws_p.cell(r, 2, name)
    ws_p.cell(r, 3, cat)
    sc(ws_p, r, 4, price,  fmt="$#,##0")
    sc(ws_p, r, 5, cogs,   fmt="0.0%")
    # Structured reference — regex can't parse this as a cell address
    sc(ws_p, r, 6, "=1-[@COGSPct]", fmt="0.0%", fg=C_GREEN)

tab_p = Table(displayName="Products", ref="A3:F8")
tab_p.tableStyleInfo = TableStyleInfo(
    name="TableStyleMedium9", showRowStripes=True)
ws_p.add_table(tab_p)

col_widths(ws_p, [("A", 12), ("B", 20), ("C", 14),
                  ("D", 14), ("E", 12), ("F", 14)])

# ── Sheet 3: Customers ────────────────────────────────────────────────────────
ws_c = wb.create_sheet("Customers")

sc(ws_c, 1, 1, "Customer Database", bold=True, fg=C_NAVY)
ws_c.cell(1, 1).font = Font(bold=True, color=C_NAVY, size=14)
sc(ws_c, 2, 1,
   "Excel Table: 'Customers'  |  XLOOKUP + structured refs in every formula row",
   fg="888888")

header_row(ws_c, 3,
           ["CustomerID", "CustomerName", "Segment", "ProductID",
            "ARR", "Status", "AnnualCOGS", "GrossProfit"])

customers = [
    ("C001", "Acme Corp",        "Mid-Market", "P002", "Active"),
    ("C002", "TechStart Inc",    "SMB",        "P001", "Active"),
    ("C003", "GlobalNet",        "Enterprise", "P003", "Active"),
    ("C004", "Pixel Labs",       "SMB",        "P001", "Churned"),
    ("C005", "DataDriven Co",    "Mid-Market", "P002", "Active"),
    ("C006", "SkyBridge Ltd",    "Enterprise", "P003", "Active"),
    ("C007", "Neon Startup",     "SMB",        "P004", "Active"),
    ("C008", "CloudVault",       "Mid-Market", "P002", "Active"),
    ("C009", "IronForge LLC",    "Enterprise", "P003", "Active"),
    ("C010", "Quickfire Inc",    "SMB",        "P001", "Churned"),
    ("C011", "Apex Analytics",   "Mid-Market", "P005", "Active"),
    ("C012", "ClearPath",        "Enterprise", "P003", "Active"),
]
for i, (cid, cname, seg, pid, status) in enumerate(customers):
    r = 4 + i
    ws_c.cell(r, 1, cid)
    ws_c.cell(r, 2, cname)
    ws_c.cell(r, 3, seg)
    ws_c.cell(r, 4, pid)
    # XLOOKUP cross-table lookup — regex sees no A1-style ref, misses this dep
    sc(ws_c, r, 5,
       "=XLOOKUP([@ProductID],Products[ProductID],Products[AnnualPrice])",
       fmt="$#,##0", fg=C_GREEN)
    ws_c.cell(r, 6, status)
    # Second XLOOKUP for COGS %
    sc(ws_c, r, 7,
       "=[@ARR]*XLOOKUP([@ProductID],Products[ProductID],Products[COGSPct])",
       fmt="$#,##0", fg=C_GREEN)
    # Simple structured ref arithmetic
    sc(ws_c, r, 8, "=[@ARR]-[@AnnualCOGS]", fmt="$#,##0", fg=C_GREEN)

tab_c = Table(displayName="Customers", ref="A3:H15")
tab_c.tableStyleInfo = TableStyleInfo(
    name="TableStyleMedium2", showRowStripes=True)
ws_c.add_table(tab_c)

col_widths(ws_c, [("A", 13), ("B", 18), ("C", 14), ("D", 12),
                  ("E", 13), ("F", 10), ("G", 14), ("H", 14)])

# ── Sheet 4: Projections ──────────────────────────────────────────────────────
ws_pr = wb.create_sheet("Projections")

sc(ws_pr, 1, 1, "5-Year Revenue Projections", bold=True, fg=C_NAVY)
ws_pr.cell(1, 1).font = Font(bold=True, color=C_NAVY, size=14)
sc(ws_pr, 2, 1,
   "Uses named ranges (EBITDAMargin, TaxRate, …), LET, and SEQUENCE (spill)",
   fg="888888")

# Row 4: Year labels — SEQUENCE spill formula
sc(ws_pr, 4, 1, "Year", bold=True, bg=C_BLUE, fg=C_WHITE)
sc(ws_pr, 4, 2,
   "=SEQUENCE(1,5,2025)",          # ← spill formula: fills B4:F4
   bg=C_LGREEN, fg=C_GREEN, bold=True)
sc(ws_pr, 4, 2).comment = None     # placeholder so cell is noted

# Helper: label column A, formula in B (base year), C-F copy formulas
rows_meta = [
    # (row, label,                  B-formula,                                 C-F-formula-template)
    (5,  "Base ARR ($)",
         "=SUM(Customers[ARR])",                                          # B: sum of current ARR table
         "=ROUND({prev}*(1+INDEX(GrowthRates,COLUMN()-COLUMN($B$5))),0)"), # C-F: compound growth
    (6,  "EBITDA ($)",
         "=ROUND(B5*EBITDAMargin,0)",                                     # named range
         "=ROUND({col}*EBITDAMargin,0)"),
    (7,  "D&A ($)",
         "=ROUND(B5*DandAPct,0)",
         "=ROUND({col}*DandAPct,0)"),
    (8,  "EBIT ($)",
         "=B6-B7",
         "={col_e}-{col_da}"),                                            # simplified below
    (9,  "Taxes ($)",
         "=MAX(B8*TaxRate,0)",                                            # named range in MAX
         "=MAX({col}*TaxRate,0)"),
    (10, "NOPAT ($)",
         "=B8-B9",
         "={col_ebit}-{col_tax}"),
    (11, "CapEx ($)",
         "=ROUND(B5*CapexPct,0)",
         "=ROUND({col}*CapexPct,0)"),
    (12, "Free Cash Flow ($)",
         # LET formula: uses named ranges + multiple intermediates
         "=LET(nopat,B10,da,B7,capex,B11,nopat+da-capex)",
         "=LET(nopat,{col_nopat},da,{col_da},capex,{col_capex},nopat+da-capex)"),
]

# Write row labels
for row, label, b_formula, _ in rows_meta:
    sc(ws_pr, row, 1, label, bold=True, bg=C_BLUE, fg=C_WHITE)
    sc(ws_pr, row, 2, b_formula, fmt="$#,##0", fg=C_GREEN)

# Write C:F formulas (simplified for clarity — reference the cell above/left)
col_letters = ["C", "D", "E", "F"]
prev_cols   = ["B", "C", "D", "E"]
for ci, (col_l, prev_l) in enumerate(zip(col_letters, prev_cols), start=1):
    col_n = ci + 2  # column index: C=3, D=4, …

    sc(ws_pr, 5, col_n,
       f"=ROUND({prev_l}5*(1+INDEX(GrowthRates,{ci})),0)",
       fmt="$#,##0", fg=C_GREEN)
    sc(ws_pr, 6, col_n,
       f"=ROUND({col_l}5*EBITDAMargin,0)", fmt="$#,##0", fg=C_GREEN)
    sc(ws_pr, 7, col_n,
       f"=ROUND({col_l}5*DandAPct,0)", fmt="$#,##0", fg=C_GREEN)
    sc(ws_pr, 8, col_n,
       f"={col_l}6-{col_l}7", fmt="$#,##0", fg=C_GREEN)
    sc(ws_pr, 9, col_n,
       f"=MAX({col_l}8*TaxRate,0)", fmt="$#,##0", fg=C_GREEN)
    sc(ws_pr, 10, col_n,
       f"={col_l}8-{col_l}9", fmt="$#,##0", fg=C_GREEN)
    sc(ws_pr, 11, col_n,
       f"=ROUND({col_l}5*CapexPct,0)", fmt="$#,##0", fg=C_GREEN)
    sc(ws_pr, 12, col_n,
       f"=LET(nopat,{col_l}10,da,{col_l}7,capex,{col_l}11,nopat+da-capex)",
       fmt="$#,##0", fg=C_GREEN)

col_widths(ws_pr, [("A", 22), ("B", 16), ("C", 16),
                   ("D", 16), ("E", 16), ("F", 16)])

# ── Sheet 5: DCF ──────────────────────────────────────────────────────────────
ws_d = wb.create_sheet("DCF")

sc(ws_d, 1, 1, "Discounted Cash Flow Model", bold=True, fg=C_NAVY)
ws_d.cell(1, 1).font = Font(bold=True, color=C_NAVY, size=14)
sc(ws_d, 2, 1,
   "Named ranges: DiscountRate, TerminalGrowthRate  |  LET for terminal value",
   fg="888888")

# Key rate display
sc(ws_d, 4, 1, "WACC (DiscountRate)",     bold=True)
sc(ws_d, 4, 2, "=DiscountRate",           fmt="0.0%",  fg=C_GREEN)  # named range ref
sc(ws_d, 5, 1, "Terminal Growth Rate",    bold=True)
sc(ws_d, 5, 2, "=TerminalGrowthRate",     fmt="0.0%",  fg=C_GREEN)  # named range ref
sc(ws_d, 6, 1, "Tax Rate",               bold=True)
sc(ws_d, 6, 2, "=TaxRate",               fmt="0.0%",  fg=C_GREEN)

header_row(ws_d, 8, ["", "Yr 1", "Yr 2", "Yr 3", "Yr 4", "Yr 5"])

# FCF row — link to Projections
sc(ws_d, 9, 1, "Free Cash Flow ($)", bold=True)
for ci, col_l in enumerate(["B", "C", "D", "E", "F"], start=2):
    proj_col = ["B", "C", "D", "E", "F"][ci - 2]
    sc(ws_d, 9, ci, f"=Projections!{proj_col}12", fmt="$#,##0")

# Discount factors
sc(ws_d, 10, 1, "Discount Factor", bold=True)
for ci in range(2, 7):
    year = ci - 1
    sc(ws_d, 10, ci,
       f"=1/(1+DiscountRate)^{year}",   # named range inside formula
       fmt="0.0000", fg=C_GREEN)

# PV of FCF
sc(ws_d, 11, 1, "PV of FCF ($)", bold=True)
for ci, col_l in enumerate(["B", "C", "D", "E", "F"], start=2):
    sc(ws_d, 11, ci, f"={col_l}9*{col_l}10", fmt="$#,##0")

# Totals
sc(ws_d, 13, 1, "Sum PV of FCFs ($)",   bold=True)
sc(ws_d, 13, 2, "=SUM(B11:F11)",        fmt="$#,##0")

# NPV check with named range arg
sc(ws_d, 14, 1, "NPV check ($)",        bold=True)
sc(ws_d, 14, 2,
   "=NPV(DiscountRate,B9:F9)",          # DiscountRate as named-range arg to NPV
   fmt="$#,##0", fg=C_GREEN)

# Terminal value — complex LET formula with named ranges
sc(ws_d, 15, 1, "Terminal Value PV ($)", bold=True)
sc(ws_d, 15, 2,
   "=LET("
   "last_fcf,F9,"
   "tv,last_fcf*(1+TerminalGrowthRate)/(DiscountRate-TerminalGrowthRate),"
   "pv_tv,tv/(1+DiscountRate)^5,"
   "pv_tv)",
   fmt="$#,##0", fg=C_GREEN)

# Enterprise value
sc(ws_d, 17, 1, "Enterprise Value ($)", bold=True, bg=C_NAVY, fg=C_WHITE)
sc(ws_d, 17, 2, "=B13+B15",            fmt="$#,##0", bold=True, bg=C_LBLUE)

col_widths(ws_d, [("A", 24), ("B", 16), ("C", 14),
                  ("D", 14), ("E", 14), ("F", 14)])

# ── Sheet 6: Summary ──────────────────────────────────────────────────────────
ws_s = wb.create_sheet("Summary")

sc(ws_s, 1, 1, "Executive Summary — Dynamic Array Demo",
   bold=True, fg=C_NAVY)
ws_s.cell(1, 1).font = Font(bold=True, color=C_NAVY, size=14)
sc(ws_s, 2, 1,
   "UNIQUE / FILTER / spill-range (#) — features invisible to regex parsers",
   fg="888888")

# KPI block
section_title(ws_s, 4, 1, "Key Metrics")
sc(ws_s, 5, 1, "Total Current ARR ($)",      bold=True)
sc(ws_s, 5, 2, "=SUM(Customers[ARR])",       fmt="$#,##0")  # table-col ref
sc(ws_s, 6, 1, "Active Customers",           bold=True)
sc(ws_s, 6, 2, '=COUNTIF(Customers[Status],"Active")')
sc(ws_s, 7, 1, "Avg ARR per Customer ($)",   bold=True)
sc(ws_s, 7, 2, "=IFERROR(B5/B6,0)",          fmt="$#,##0")
sc(ws_s, 8, 1, "Enterprise Value ($)",       bold=True)
sc(ws_s, 8, 2, "=DCF!B17",                  fmt="$#,##0")

# UNIQUE spill — column A, rows 11+
section_title(ws_s, 10, 1, "ARR by Segment  (UNIQUE spill + SUMIF on A11#)")
header_row(ws_s, 11, ["Segment", "Total ARR ($)", "% of Total"])
# UNIQUE in A12: spills downward — regex won't see this as producing output cells
sc(ws_s, 12, 1,
   "=UNIQUE(Customers[Segment])",           # ← spill anchor
   fg=C_GREEN)
# SUMIF using the spill-range reference A12# (modern Excel only)
sc(ws_s, 12, 2,
   '=SUMIF(Customers[Segment],A12#,Customers[ARR])',   # ← spill ref (#)
   fmt="$#,##0", fg=C_GREEN)
sc(ws_s, 12, 3,
   "=B12#/SUM(Customers[ARR])",
   fmt="0.0%", fg=C_GREEN)

# FILTER spill — active customer names
section_title(ws_s, 18, 1, "Active Customers  (FILTER spill)")
header_row(ws_s, 19, ["Name", "Segment", "ARR ($)"])
sc(ws_s, 20, 1,
   '=FILTER(Customers[CustomerName],Customers[Status]="Active")',  # ← spill
   fg=C_GREEN)
sc(ws_s, 20, 2,
   '=FILTER(Customers[Segment],Customers[Status]="Active")',
   fg=C_GREEN)
sc(ws_s, 20, 3,
   '=FILTER(Customers[ARR],Customers[Status]="Active")',
   fmt="$#,##0", fg=C_GREEN)

col_widths(ws_s, [("A", 26), ("B", 18), ("C", 14)])

# ── Save ──────────────────────────────────────────────────────────────────────
out = "modern_saas_model.xlsx"
wb.save(out)
print(f"Saved {out}")
print(f"   Sheets : {wb.sheetnames}")
print(f"   Named ranges: {list(wb.defined_names)[:5]} ({len(list(wb.defined_names))} total)")
