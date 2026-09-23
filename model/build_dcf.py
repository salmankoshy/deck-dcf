# -*- coding: utf-8 -*-
"""Build model/deck_dcf.xlsx from inputs/assumptions.csv, inputs/historicals.csv, inputs/beta.csv.
Live Excel formulas only. Blue = hardcoded input (Assumptions/Historical/Comps inputs),
black = same-sheet formula, green = cross-sheet link. Run: py -3.13 model/build_dcf.py
"""
import csv, sys, os
sys.stdout.reconfigure(encoding="utf-8")
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
A_CSV = os.path.join(ROOT, "inputs", "assumptions.csv")
H_CSV = os.path.join(ROOT, "inputs", "historicals.csv")
B_CSV = os.path.join(ROOT, "inputs", "beta.csv")
OUT = os.path.join(ROOT, "model", "deck_dcf.xlsx")

BLUE, GREEN, BLACK = "0000FF", "008000", "000000"
F_IN = Font(color=BLUE)
F_LINK = Font(color=GREEN)
F_CALC = Font(color=BLACK)
F_BOLD = Font(bold=True)
F_TITLE = Font(bold=True, size=14)
F_HDR = Font(bold=True, color="FFFFFF")
FILL_HDR = PatternFill("solid", fgColor="1F3864")
FILL_NOTE = PatternFill("solid", fgColor="FFF2CC")
FMT_K = '#,##0;(#,##0);"-"'
FMT_PCT = '0.0%'
FMT_PCT2 = '0.00%'
FMT_PX = '$#,##0.00'
FMT_X = '0.00"x"'
FMT_B = '0.0000'
FMT_SH = '#,##0'

# ------------------------------------------------------------------ inputs
def read_csv(p):
    with open(p, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))

A = read_csv(A_CSV)
H = read_csv(H_CSV)
B = read_csv(B_CSV)

def num(s):
    if s is None or s == "":
        return None
    try:
        v = float(s)
        return int(v) if v == int(v) and "." not in s else v
    except ValueError:
        return s  # text

# ------------------------------------------------------------------ helpers
wb = Workbook()
wb.remove(wb.active)
SHEETS = ["Cover", "Assumptions", "Historical", "WACC", "DCF", "Sensitivity", "Comps", "MC"]
ws = {n: wb.create_sheet(n) for n in SHEETS}

def put(sheet, ref, value, kind="calc", fmt=None, bold=False, wrap=False):
    """kind: 'in' blue input | 'calc' black same-sheet formula | 'link' green cross-sheet | 'text' plain label"""
    c = sheet[ref]
    c.value = value
    if kind == "in":
        c.font = Font(color=BLUE, bold=bold)
    elif kind == "link":
        c.font = Font(color=GREEN, bold=bold)
    elif kind == "calc":
        c.font = Font(color=BLACK, bold=bold)
    else:
        c.font = Font(bold=bold)
    if fmt:
        c.number_format = fmt
    if wrap:
        c.alignment = Alignment(wrap_text=True, vertical="top")
    return c

def hdr_row(sheet, row, labels, start_col=1):
    for i, lab in enumerate(labels):
        c = sheet.cell(row=row, column=start_col + i, value=lab)
        c.font = F_HDR
        c.fill = FILL_HDR

def name(nm, sheet, ref):
    """Define a workbook-level name pointing at sheet!$col$row."""
    col = "".join(ch for ch in ref if ch.isalpha())
    row = "".join(ch for ch in ref if ch.isdigit())
    wb.defined_names[nm] = DefinedName(nm, attr_text=f"{sheet}!${col}${row}")

def widths(sheet, d):
    for col, w in d.items():
        sheet.column_dimensions[col].width = w

# ================================================================== ASSUMPTIONS
sa = ws["Assumptions"]
put(sa, "A1", "Assumptions: all rows from inputs/assumptions.csv (blue = hardcoded input)", "text", bold=True)
hdr_row(sa, 3, ["key", "value", "unit", "source", "status"])
ALIAS = {  # csv key -> extra short defined name(s) requested
    "terminal_growth": ["tg"], "exit_multiple_ev_ebitda": ["exit_mult"],
    "shares_diluted_current": ["shares"], "tax_rate": ["tax"],
    "rev_growth_fy27": ["g_fy27"], "rev_growth_fy28": ["g_fy28"], "rev_growth_fy29": ["g_fy29"],
    "rev_growth_fy30": ["g_fy30"], "rev_growth_fy31": ["g_fy31"],
    "da_pct_rev": ["da_pct"], "capex_pct_rev": ["capex_pct"], "nwc_pct_rev": ["nwc_pct"],
    "cost_debt_pretax": ["kd_pretax"], "weight_equity": ["we_input"],
}
FORMULA_ROWS = {"cost_equity", "wacc", "weight_equity", "beta_regression_crosscheck"}  # status=formula in csv -> green link to WACC tab, never typed
a_row = {}
r = 4
for rec in A:
    k = rec["key"]
    a_row[k] = r
    put(sa, f"A{r}", k, "text")
    unit = rec["unit"]
    if k in FORMULA_ROWS:
        pass  # filled after WACC tab rows are known
    else:
        v = num(rec["value"])
        fmt = None
        if isinstance(v, (int, float)):
            if unit == "pct":
                fmt = FMT_PCT2
            elif unit == "USD":
                fmt = FMT_PX
            elif unit == "USD thousands":
                fmt = FMT_K
            elif unit == "shares":
                fmt = FMT_SH
            elif unit == "x":
                fmt = FMT_X
            elif unit == "n":
                fmt = FMT_SH
        put(sa, f"B{r}", v, "in", fmt=fmt)
    put(sa, f"C{r}", unit, "text")
    put(sa, f"D{r}", rec["source"], "text")
    put(sa, f"E{r}", rec["status"], "text")
    name(k, "Assumptions", f"B{r}")
    for al in ALIAS.get(k, []):
        name(al, "Assumptions", f"B{r}")
    r += 1

# beta regression inputs (from inputs/beta.csv), inputs live here so WACC tab has zero typed numbers
r += 1
put(sa, f"A{r}", "Beta regression inputs from inputs/beta.csv (blue)", "text", bold=True)
r += 1
hdr_row(sa, r, ["key", "value", "unit", "source", "status"])
r += 1
beta_keys = []
for rec in B:
    tag = "5y" if rec["window"].startswith("5Y") else "2y"
    src = (f"inputs/beta.csv, {rec['window']} ({rec['interval']}), {rec['start']}..{rec['end']}, "
           f"n={rec['n_obs']}, R2={rec['r_squared']}; {rec['source']}")
    for col, nm in (("raw_beta", f"beta_{tag}_raw"), ("blume_adj_beta", f"beta_{tag}_blume")):
        a_row[nm] = r
        put(sa, f"A{r}", nm, "text")
        put(sa, f"B{r}", float(rec[col]), "in", fmt=FMT_B)
        put(sa, f"C{r}", "x", "text")
        put(sa, f"D{r}", src, "text")
        put(sa, f"E{r}", "verified", "text")
        name(nm, "Assumptions", f"B{r}")
        beta_keys.append(nm)
        r += 1
widths(sa, {"A": 30, "B": 16, "C": 14, "D": 95, "E": 22})
sa.freeze_panes = "A4"

# ================================================================== HISTORICAL
sh = ws["Historical"]
YEARS = ["FY2022", "FY2023", "FY2024", "FY2025", "FY2026"]
YCOL = {y: L(2 + i) for i, y in enumerate(YEARS)}  # B..F
put(sh, "A1", "Historical: fs_value from inputs/historicals.csv, $ thousands (shares in thousands); blue = input, blank = not in source", "text", bold=True)
hdr_row(sh, 3, ["Line item"] + YEARS + ["source_file(s)", "note"])
ITEMS = ["Net sales", "Cost of sales", "Gross profit", "SG&A", "Income from operations",
         "Total other income net", "Income before income taxes", "Income tax expense", "Net income",
         "Depreciation amortization and accretion", "Stock-based compensation",
         "Purchases of property and equipment", "Cash and cash equivalents",
         "Trade accounts receivable net", "Inventories", "Trade accounts payable",
         "Operating lease liabilities current", "Long-term operating lease liabilities",
         "Total stockholders' equity", "Long-term debt", "Diluted weighted-average shares"]
hidx = {(x["item"], x["period"]): x for x in H}
h_row = {}
r = 4
for it in ITEMS:
    h_row[it] = r
    put(sh, f"A{r}", it, "text")
    srcs, notes = [], []
    for y in YEARS:
        rec = hidx.get((it, y))
        fv = rec["fs_value"] if rec else ""
        if fv != "":
            v = num(fv)
            put(sh, f"{YCOL[y]}{r}", v, "in", fmt=FMT_K)
            if rec["source_file"] and rec["source_file"] not in srcs:
                srcs.append(rec["source_file"])
        if rec and rec.get("note"):
            notes.append(f"{y}: {rec['note']}")
    IS_ROWS = {"Net sales", "Cost of sales", "Gross profit", "SG&A", "Income from operations", "Total other income net",
               "Income before income taxes", "Income tax expense", "Net income"}
    CF_ROWS = {"Depreciation amortization and accretion", "Stock-based compensation", "Purchases of property and equipment"}
    if it in IS_ROWS:
        cite = "10-K FY2024 (FY22-23), 10-K FY2026 (FY24-26), Consolidated Statements of Comprehensive Income"
    elif it in CF_ROWS:
        cite = "10-K FY2024 (FY22-23), 10-K FY2026 (FY24-26), Consolidated Statements of Cash Flows"
    elif it == "Diluted weighted-average shares":
        cite = "10-K FY2024 (FY22-23), 10-K FY2026 (FY24-26), Note: Basic and Diluted Shares"
    elif it == "Total stockholders' equity":
        cite = "10-K FY2026, Consolidated Balance Sheets"
    else:
        cite = "10-K FY2024 (FY23-24), 10-K FY2026 (FY25-26), Consolidated Balance Sheets"
    put(sh, f"G{r}", cite, "text")
    if it == "Total other income net":
        notes.append("as printed: negative = net other income, added to operating income")
    if it == "Purchases of property and equipment":
        notes.append("as printed in cash-flow statement: negative = cash outflow")
    if it == "Long-term debt":
        notes.append("no debt line item on DECK balance sheet; undrawn revolver only")
    put(sh, f"H{r}", "; ".join(notes), "text")
    r += 1

# ratios (formulas); no spacer row, so the ratio rows keep their positions after the equity row was added
put(sh, f"A{r}", "Ratios (formulas)", "text", bold=True)
r += 1
R = h_row
def ycell(item, y):
    return f"{YCOL[y]}{R[item]}"
ratio_row = {}
def ratio(label, key, fn, fmt=FMT_PCT):
    global r
    ratio_row[key] = r
    put(sh, f"A{r}", label, "text")
    for i, y in enumerate(YEARS):
        f = fn(y, YEARS[i - 1] if i > 0 else None)
        if f is not None:
            put(sh, f"{YCOL[y]}{r}", f, "calc", fmt=fmt)
    r += 1

ratio("Revenue growth", "growth",
      lambda y, py: None if py is None else f"=IF(COUNT({ycell('Net sales', py)},{ycell('Net sales', y)})=2,{ycell('Net sales', y)}/{ycell('Net sales', py)}-1,\"\")")
ratio("Gross margin", "gm", lambda y, py: f"=IF(COUNT({ycell('Gross profit', y)},{ycell('Net sales', y)})=2,{ycell('Gross profit', y)}/{ycell('Net sales', y)},\"\")")
ratio("EBIT margin (income from operations / revenue)", "ebitm", lambda y, py: f"=IF(COUNT({ycell('Income from operations', y)},{ycell('Net sales', y)})=2,{ycell('Income from operations', y)}/{ycell('Net sales', y)},\"\")")
ratio("D&A % revenue", "da", lambda y, py: f"=IF(COUNT({ycell('Depreciation amortization and accretion', y)},{ycell('Net sales', y)})=2,{ycell('Depreciation amortization and accretion', y)}/{ycell('Net sales', y)},\"\")")
ratio("Capex % revenue (ABS of cash-flow line)", "capex", lambda y, py: f"=IF(COUNT({ycell('Purchases of property and equipment', y)},{ycell('Net sales', y)})=2,ABS({ycell('Purchases of property and equipment', y)})/{ycell('Net sales', y)},\"\")")
ratio("NWC = AR + Inventories - AP ($ thousands)", "nwc",
      lambda y, py: f"=IF(COUNT({ycell('Trade accounts receivable net', y)},{ycell('Inventories', y)},{ycell('Trade accounts payable', y)})=3,{ycell('Trade accounts receivable net', y)}+{ycell('Inventories', y)}-{ycell('Trade accounts payable', y)},\"\")",
      fmt=FMT_K)
NWC_ROW = ratio_row["nwc"]
ratio("NWC % revenue", "nwcp", lambda y, py: f"=IF(AND(ISNUMBER({YCOL[y]}{NWC_ROW}),ISNUMBER({ycell('Net sales', y)})),{YCOL[y]}{NWC_ROW}/{ycell('Net sales', y)},\"\")")
ratio("Effective tax rate", "tax", lambda y, py: f"=IF(COUNT({ycell('Income tax expense', y)},{ycell('Income before income taxes', y)})=2,{ycell('Income tax expense', y)}/{ycell('Income before income taxes', y)},\"\")")
ratio("SBC % revenue", "sbc", lambda y, py: f"=IF(COUNT({ycell('Stock-based compensation', y)},{ycell('Net sales', y)})=2,{ycell('Stock-based compensation', y)}/{ycell('Net sales', y)},\"\")")

# FY26 ROIC (formulas, tax from Assumptions): NOPAT / (equity - cash + operating lease liabilities), and excluding leases
F26 = YCOL["FY2026"]
EQ, CASH26, OLC, OLL, EBIT26 = h_row["Total stockholders' equity"], h_row["Cash and cash equivalents"], h_row["Operating lease liabilities current"], h_row["Long-term operating lease liabilities"], h_row["Income from operations"]
ratio_row["roic"] = r
put(sh, f"A{r}", "ROIC FY26 = income from operations x (1 - tax) / (equity - cash + operating lease liabilities, current + long-term)", "text")
put(sh, f"{F26}{r}", f"={F26}{EBIT26}*(1-tax)/({F26}{EQ}-{F26}{CASH26}+{F26}{OLC}+{F26}{OLL})", "link", FMT_PCT2)
name("roic_fy26", "Historical", f"{F26}{r}")
r += 1
ratio_row["roic_ex"] = r
put(sh, f"A{r}", "ROIC FY26 excluding leases = same numerator / (equity - cash)", "text")
put(sh, f"{F26}{r}", f"={F26}{EBIT26}*(1-tax)/({F26}{EQ}-{F26}{CASH26})", "link", FMT_PCT2)
name("roic_fy26_ex_leases", "Historical", f"{F26}{r}")
r += 1

# LTM to Jun 30 2026 = FY2026 + Q1 FY2027 - Q1 FY2026. Quarters are blue inputs from inputs/historicals.csv; LTM cells are formulas.
r += 1
put(sh, f"A{r}", "LTM to Jun 30 2026 ($ thousands): FY2026 + Q1 FY2027 - Q1 FY2026", "text", bold=True)
r += 1
hdr_row(sh, r, ["Line item", "FY2026", "Q1 FY2027", "Q1 FY2026", "LTM Jun 2026", "", "source_file(s)", "note"])
r += 1
ltm_row = {}
for it, nm in (("Income from operations", "ltm_ebit"), ("Depreciation amortization and accretion", "ltm_da")):
    ltm_row[it] = r
    q27, q26, ltm = hidx[(it, "Q1FY27")], hidx[(it, "Q1FY26")], hidx[(it, "LTM_Jun26")]
    put(sh, f"A{r}", it, "text")
    put(sh, f"B{r}", f"={YCOL['FY2026']}{h_row[it]}", "calc", FMT_K)
    put(sh, f"C{r}", num(q27["fs_value"]), "in", FMT_K)
    put(sh, f"D{r}", num(q26["fs_value"]), "in", FMT_K)
    put(sh, f"E{r}", f"=B{r}+C{r}-D{r}", "calc", FMT_K)
    put(sh, f"G{r}", "10-K FY2026 (FY2026 column); 10-Q Q1 FY2027 condensed statements, current-period column (Q1 FY2027) and prior-period column (Q1 FY2026)", "text")
    put(sh, f"H{r}", f"inputs/historicals.csv period LTM_Jun26 fs_value {num(ltm['fs_value']):,} ({ltm['note']}); must equal column E", "text")
    name(nm, "Historical", f"E{r}")
    r += 1
ltm_row["ebitda"] = r
put(sh, f"A{r}", "LTM EBITDA = income from operations + D&A and accretion", "text", bold=True)
for col in "BCDE":
    put(sh, f"{col}{r}", f"={col}{ltm_row['Income from operations']}+{col}{ltm_row['Depreciation amortization and accretion']}", "calc", FMT_K, bold=(col == "E"))
put(sh, f"G{r}", "formula; feeds Comps tab (defined name ltm_ebitda)", "text")
name("ltm_ebitda", "Historical", f"E{r}")
r += 1
widths(sh, {"A": 46, "B": 14, "C": 14, "D": 14, "E": 14, "F": 14, "G": 70, "H": 80})
sh.freeze_panes = "B4"

# ================================================================== WACC
sw = ws["WACC"]
put(sw, "A1", "WACC: every cell is a formula or a green link; no typed numbers on this sheet", "text", bold=True)
hdr_row(sw, 3, ["Component", "Value", "Formula / source"])
w_row = {}
r = 4
def wrow(key, label, formula, kind, fmt, note=""):
    global r
    w_row[key] = r
    put(sw, f"A{r}", label, "text")
    put(sw, f"B{r}", formula, kind, fmt=fmt)
    put(sw, f"C{r}", note, "text")
    r += 1
wrow("rf", "Risk-free rate (rf)", "=rf", "link", FMT_PCT2, "Assumptions!rf")
wrow("beta", "Beta: Damodaran shoe, cash-corrected (operating) unlevered (used)", "=beta", "link", FMT_B, "Assumptions!beta; cash-corrected, so compare with the operating beta rows below")
wrow("b5r", "Beta: regression 5Y monthly, raw (includes cash drag, not cash-corrected)", "=beta_5y_raw", "link", FMT_B, "Assumptions!beta_5y_raw (inputs/beta.csv)")
wrow("b5b", "Beta: regression 5Y monthly, Blume (0.67*raw+0.33)", "=beta_5y_blume", "link", FMT_B, "Assumptions!beta_5y_blume")
wrow("b2r", "Beta: regression 2Y weekly, raw", "=beta_2y_raw", "link", FMT_B, "Assumptions!beta_2y_raw")
wrow("b2b", "Beta: regression 2Y weekly, Blume", "=beta_2y_blume", "link", FMT_B, "Assumptions!beta_2y_blume")
wrow("erp", "Equity risk premium (ERP)", "=erp", "link", FMT_PCT2, "Assumptions!erp")
wrow("ke", "Cost of equity  Ke = rf + beta × ERP", f"=B{w_row['rf']}+B{w_row['beta']}*B{w_row['erp']}", "calc", FMT_PCT2, "CAPM")
wrow("kd", "Pre-tax cost of debt (Kd)", "=kd_pretax", "link", FMT_PCT2, "Assumptions!cost_debt_pretax")
wrow("tax", "Tax rate (t)", "=tax", "link", FMT_PCT2, "Assumptions!tax_rate")
wrow("kdat", "After-tax cost of debt  Kd × (1 − t)", f"=B{w_row['kd']}*(1-B{w_row['tax']})", "calc", FMT_PCT2, "")
wrow("px", "Share price ($)", "=price", "link", FMT_PX, "Assumptions!price")
wrow("sh", "Diluted shares (count)", "=shares", "link", FMT_SH, "Assumptions!shares_diluted_current")
wrow("mcap", "Market cap ($ thousands) = price × shares / 1000", f"=B{w_row['px']}*B{w_row['sh']}/1000", "calc", FMT_K, "")
wrow("debt", "Debt ($ thousands)", "=debt", "link", FMT_K, "Assumptions!debt")
wrow("cap", "Total capital ($ thousands)", f"=B{w_row['mcap']}+B{w_row['debt']}", "calc", FMT_K, "")
wrow("we", "Equity weight  We", f"=B{w_row['mcap']}/B{w_row['cap']}", "calc", FMT_PCT2, "")
wrow("wd", "Debt weight  Wd", f"=B{w_row['debt']}/B{w_row['cap']}", "calc", FMT_PCT2, "")
wrow("wacc", "WACC = Ke×We + Kd×(1−t)×Wd (used; feeds DCF)", f"=B{w_row['ke']}*B{w_row['we']}+B{w_row['kdat']}*B{w_row['wd']}", "calc", FMT_PCT2, "defined name: wacc_calc")
wrow("wacc5", "WACC at 5Y regression equity beta (includes cash drag, not comparable to the cash-corrected beta; see row 24)", f"=(B{w_row['rf']}+B{w_row['b5r']}*B{w_row['erp']})*B{w_row['we']}+B{w_row['kdat']}*B{w_row['wd']}", "calc", FMT_PCT2, "not used downstream")
sw[f"A{w_row['wacc']}"].font = F_BOLD
sw[f"B{w_row['wacc']}"].font = Font(color=BLACK, bold=True)
name("wacc_calc", "WACC", f"B{w_row['wacc']}")
name("ke_calc", "WACC", f"B{w_row['ke']}")
widths(sw, {"A": 60, "B": 16, "C": 48})

# Assumptions rows with status=formula -> green links back to WACC tab (never typed)
put(sa, f"B{a_row['cost_equity']}", f"=WACC!B{w_row['ke']}", "link", fmt=FMT_PCT2)
put(sa, f"B{a_row['wacc']}", f"=WACC!B{w_row['wacc']}", "link", fmt=FMT_PCT2)
put(sa, f"B{a_row['weight_equity']}", f"=WACC!B{w_row['we']}", "link", fmt=FMT_PCT2)
put(sa, f"B{a_row['beta_regression_crosscheck']}", "=beta_5y_raw", "calc", fmt=FMT_B)

# ================================================================== DCF
sd = ws["DCF"]
put(sd, "A1", "DCF: unlevered FCFF, mid-year convention, $ thousands (price in $/share). Every numeric cell is a formula or a link; labels, notes and check ranges are text.", "text", bold=True)
COLS = ["FY2026A", "FY2027E", "FY2028E", "FY2029E", "FY2030E", "FY2031E"]
C = {c: L(2 + i) for i, c in enumerate(COLS)}  # B..G
FC = COLS[1:]  # forecast columns
hdr_row(sd, 2, ["($ thousands)"] + COLS + ["Notes"])
d_row = {}
r = 3
def drow(key, label):
    global r
    d_row[key] = r
    put(sd, f"A{r}", label, "text")
    r += 1
for k, lab in [("g", "Revenue growth"), ("rev", "Revenue"), ("m", "EBIT margin"), ("ebit", "EBIT"),
               ("taxes", "Taxes on EBIT  (EBIT × tax)"), ("nopat", "NOPAT"), ("da", "+ D&A"),
               ("capex", "− Capex"), ("nwc", "NWC level  (AR + Inv − AP)"), ("dnwc", "− Change in NWC  (increase = outflow)"),
               ("fcff", "FCFF  = NOPAT + D&A − Capex − ΔNWC"), ("ebitda", "EBITDA  (EBIT + D&A)")]:
    drow(k, lab)
G_MAP = {"FY2027E": "g_fy27", "FY2028E": "g_fy28", "FY2029E": "g_fy29", "FY2030E": "g_fy30", "FY2031E": "g_fy31"}
hist_col = YCOL["FY2026"]  # Historical column for FY2026
b = C["FY2026A"]
# FY2026A (base), green links to Historical
put(sd, f"{b}{d_row['g']}", f"=Historical!{hist_col}{ratio_row['growth']}", "link", FMT_PCT)
put(sd, f"{b}{d_row['rev']}", f"=Historical!{hist_col}{h_row['Net sales']}", "link", FMT_K)
put(sd, f"{b}{d_row['ebit']}", f"=Historical!{hist_col}{h_row['Income from operations']}", "link", FMT_K)
put(sd, f"{b}{d_row['m']}", f"={b}{d_row['ebit']}/{b}{d_row['rev']}", "calc", FMT_PCT)
put(sd, f"{b}{d_row['da']}", f"=Historical!{hist_col}{h_row['Depreciation amortization and accretion']}", "link", FMT_K)
put(sd, f"{b}{d_row['capex']}", f"=ABS(Historical!{hist_col}{h_row['Purchases of property and equipment']})", "link", FMT_K)
put(sd, f"{b}{d_row['nwc']}", f"=Historical!{hist_col}{NWC_ROW}", "link", FMT_K)
put(sd, f"{b}{d_row['ebitda']}", f"={b}{d_row['ebit']}+{b}{d_row['da']}", "calc", FMT_K)
# forecast columns
for i, c in enumerate(FC):
    col = C[c]; prev = C[COLS[i]]
    put(sd, f"{col}{d_row['g']}", f"={G_MAP[c]}", "link", FMT_PCT)
    put(sd, f"{col}{d_row['rev']}", f"={prev}{d_row['rev']}*(1+{col}{d_row['g']})", "calc", FMT_K)
    put(sd, f"{col}{d_row['m']}", "=ebit_margin", "link", FMT_PCT)
    put(sd, f"{col}{d_row['ebit']}", f"={col}{d_row['rev']}*{col}{d_row['m']}", "calc", FMT_K)
    put(sd, f"{col}{d_row['taxes']}", f"={col}{d_row['ebit']}*tax", "link", FMT_K)
    put(sd, f"{col}{d_row['nopat']}", f"={col}{d_row['ebit']}-{col}{d_row['taxes']}", "calc", FMT_K)
    put(sd, f"{col}{d_row['da']}", f"=da_pct*{col}{d_row['rev']}", "link", FMT_K)
    put(sd, f"{col}{d_row['capex']}", f"=capex_pct*{col}{d_row['rev']}", "link", FMT_K)
    put(sd, f"{col}{d_row['nwc']}", f"=nwc_pct*{col}{d_row['rev']}", "link", FMT_K)
    put(sd, f"{col}{d_row['dnwc']}", f"={col}{d_row['nwc']}-{prev}{d_row['nwc']}", "calc", FMT_K)
    put(sd, f"{col}{d_row['fcff']}", f"={col}{d_row['nopat']}+{col}{d_row['da']}-{col}{d_row['capex']}-{col}{d_row['dnwc']}", "calc", FMT_K)
    put(sd, f"{col}{d_row['ebitda']}", f"={col}{d_row['ebit']}+{col}{d_row['da']}", "calc", FMT_K)
NOTE_COL = "H"
put(sd, f"{NOTE_COL}{d_row['g']}", "FY26A growth from Historical; FY27 to FY31 from Assumptions g_fy27..g_fy31", "text")
put(sd, f"{NOTE_COL}{d_row['m']}", "FY26A actual; forecast = Assumptions ebit_margin (flat)", "text")
put(sd, f"{NOTE_COL}{d_row['nwc']}", "FY26A = actual AR+Inv−AP from Historical; forecast = nwc_pct × revenue", "text")
put(sd, f"{NOTE_COL}{d_row['capex']}", "Historical capex line is negative in the cash-flow statement; shown here as a positive outflow", "text")
r += 1
d_row["note_sbc"] = r
c = put(sd, f"A{r}", "SBC: expense stays inside EBIT and is not added back to FCFF (SBC treated as a cost). Existing unvested awards are captured in the share count via the 296K dilutive effect (10-Q Note 9).", "text", wrap=True)
c.fill = FILL_NOTE
sd.merge_cells(f"A{r}:{NOTE_COL}{r}")
r += 1
d_row["note_lease"] = r
c = put(sd, f"A{r}", "Leases: operating lease cost stays inside EBIT. No operating lease liability is deducted in the equity bridge (leases excluded from the bridge because lease cost is already in EBIT).", "text", wrap=True)
c.fill = FILL_NOTE
sd.merge_cells(f"A{r}:{NOTE_COL}{r}")
r += 2

# discounting: period derived from header text so no typed numbers
for k, lab in [("yr", "Fiscal year (numeric, from header)"), ("n", "Year offset n  (FY − base FY)"),
               ("t", "Mid-year period t = n − 0.5"), ("df", "Discount factor  1 / (1+WACC)^t"), ("pv", "PV of FCFF")]:
    drow(k, lab)
for c in COLS:
    col = C[c]
    put(sd, f"{col}{d_row['yr']}", f"=VALUE(MID({col}$2,3,4))", "calc", "0")
    put(sd, f"{col}{d_row['n']}", f"={col}{d_row['yr']}-VALUE(MID(base_year,3,4))", "link", "0")
for c in FC:
    col = C[c]
    put(sd, f"{col}{d_row['t']}", f"={col}{d_row['n']}-0.5", "calc", "0.0")
    put(sd, f"{col}{d_row['df']}", f"=1/(1+wacc_calc)^{col}{d_row['t']}", "link", "0.0000")
    put(sd, f"{col}{d_row['pv']}", f"={col}{d_row['fcff']}*{col}{d_row['df']}", "calc", FMT_K)
put(sd, f"{NOTE_COL}{d_row['n']}", "base FY from Assumptions!base_year", "text")
r += 1

# terminal value & bridge, two methods side by side
hdr_row(sd, r, ["Terminal value & equity bridge", "Perpetuity (Gordon)", "Exit multiple", "", "", "", "", "Convention"])
r += 1
P, X = "B", "C"
last = C["FY2031E"]
def brow(key, label, fp, fx, fmt, kp="calc", kx="calc", note=""):
    global r
    d_row[key] = r
    put(sd, f"A{r}", label, "text")
    put(sd, f"{P}{r}", fp, kp, fmt)
    put(sd, f"{X}{r}", fx, kx, fmt)
    if note:
        put(sd, f"{NOTE_COL}{r}", note, "text")
    r += 1
brow("tv", "Terminal value (undiscounted)",
     f"={last}{d_row['fcff']}*(1+tg)/(wacc_calc-tg)", f"={last}{d_row['ebitda']}*exit_mult", FMT_K, "link", "link",
     "Perp: FCFF_FY31×(1+g)/(WACC−g).  Exit: EBITDA_FY31 × exit multiple")
brow("tvt", "Discount period for TV (years)", f"={last}{d_row['t']}", f"={last}{d_row['n']}", "0.0",
     note="Perp discounted at t=4.5 (mid-year consistent); Exit discounted at t=5.0 (end of FY31)")
brow("pvtv", "PV of terminal value", f"={P}{d_row['tv']}/(1+wacc_calc)^{P}{d_row['tvt']}",
     f"={X}{d_row['tv']}/(1+wacc_calc)^{X}{d_row['tvt']}", FMT_K, "link", "link")
brow("sumpv", "Sum of PV FCFF (FY27 to FY31)", f"=SUM({C['FY2027E']}{d_row['pv']}:{last}{d_row['pv']})",
     f"=SUM({C['FY2027E']}{d_row['pv']}:{last}{d_row['pv']})", FMT_K)
brow("ev", "Enterprise value", f"={P}{d_row['sumpv']}+{P}{d_row['pvtv']}", f"={X}{d_row['sumpv']}+{X}{d_row['pvtv']}", FMT_K)
brow("cash", "+ Cash", "=cash", "=cash", FMT_K, "link", "link", "Assumptions!cash: Jun 30 balance sheet 1,602,589 less 32,168 repurchased to Jul 9 (10-Q subsequent events)")
brow("debt", "− Debt", "=debt", "=debt", FMT_K, "link", "link", "Assumptions!debt (0); leases excluded")
brow("eq", "Equity value", f"={P}{d_row['ev']}+{P}{d_row['cash']}-{P}{d_row['debt']}", f"={X}{d_row['ev']}+{X}{d_row['cash']}-{X}{d_row['debt']}", FMT_K)
brow("sh", "Diluted shares (count)", "=shares", "=shares", FMT_SH, "link", "link", "Assumptions!shares_diluted_current")
brow("px", "Implied price per share ($)  = equity ×1000 / shares", f"={P}{d_row['eq']}*1000/{P}{d_row['sh']}", f"={X}{d_row['eq']}*1000/{X}{d_row['sh']}", FMT_PX)
brow("mkt", "Current share price ($)", "=price", "=price", FMT_PX, "link", "link", "Assumptions!price")
brow("up", "Upside / (downside) vs current", f"={P}{d_row['px']}/{P}{d_row['mkt']}-1", f"={X}{d_row['px']}/{X}{d_row['mkt']}-1", FMT_PCT)
brow("tvpct", "PV TV as % of EV", f"={P}{d_row['pvtv']}/{P}{d_row['ev']}", f"={X}{d_row['pvtv']}/{X}{d_row['ev']}", FMT_PCT)
for k in ("px", "up"):
    sd[f"A{d_row[k]}"].font = F_BOLD
    sd[f"{P}{d_row[k]}"].font = Font(color=BLACK, bold=True)
    sd[f"{X}{d_row[k]}"].font = Font(color=BLACK, bold=True)
name("price_perp", "DCF", f"{P}{d_row['px']}")
name("price_exit", "DCF", f"{X}{d_row['px']}")
name("upside_perp", "DCF", f"{P}{d_row['up']}")
name("upside_exit", "DCF", f"{X}{d_row['up']}")
name("ev_perp", "DCF", f"{P}{d_row['ev']}")
name("ev_exit", "DCF", f"{X}{d_row['ev']}")
r += 1

# checks
# market-implied EBIT margin: value is linear in margin, so two repricings at the Assumptions probe margins pin it exactly
fc0 = C["FY2027E"]
def price_at_margin(m):
    """Perpetuity-method price with EBIT margin m swapped in: Sensitivity-grid style SUMPRODUCT, defined names for every input."""
    rng = lambda k: f"${fc0}${d_row[k]}:${last}${d_row[k]}"
    c31 = lambda k: f"${last}${d_row[k]}"
    return (f"=(SUMPRODUCT({m}*{rng('rev')}*(1-tax)+{rng('da')}-{rng('capex')}-{rng('dnwc')},1/(1+wacc_calc)^{rng('t')})"
            f"+({m}*{c31('rev')}*(1-tax)+{c31('da')}-{c31('capex')}-{c31('dnwc')})*(1+tg)/(wacc_calc-tg)/(1+wacc_calc)^{c31('t')}"
            f"+cash-debt)/shares*1000")
hdr_row(sd, r, ["Market-implied EBIT margin at current price", "Value", "", "", "", "", "", "Convention"])
r += 1
d_row["pm_lo"] = r
put(sd, f"A{r}", "Price at probe margin lo ($)", "text")
put(sd, f"B{r}", price_at_margin("implied_margin_probe_lo"), "link", FMT_PX)
put(sd, f"{NOTE_COL}{r}", "Same SUMPRODUCT repricing as the Sensitivity grids with the margin swapped in; probe margins are Assumptions inputs", "text")
r += 1
d_row["pm_hi"] = r
put(sd, f"A{r}", "Price at probe margin hi ($)", "text")
put(sd, f"B{r}", price_at_margin("implied_margin_probe_hi"), "link", FMT_PX)
r += 1
d_row["im"] = r
put(sd, f"A{r}", "Market-implied EBIT margin at current price", "text", bold=True)
put(sd, f"B{r}", f"=implied_margin_probe_lo+(price-B{d_row['pm_lo']})*(implied_margin_probe_hi-implied_margin_probe_lo)/(B{d_row['pm_hi']}-B{d_row['pm_lo']})", "link", FMT_PCT2)
sd[f"B{r}"].font = Font(color=GREEN, bold=True)
put(sd, f"{NOTE_COL}{r}", "Value per share is linear in margin, so linear interpolation between the two probes is exact; the check row below reprices at this margin and must return the current price", "text")
name("implied_margin", "DCF", f"B{r}")
r += 2
# terminal reinvestment sensitivity rows (not in the pass count)
hdr_row(sd, r, ["Terminal reinvestment sensitivity", "Value", "", "", "", "", "", "Convention"])
r += 1
NP31, FC31, T31c = f"${last}${d_row['nopat']}", f"${last}${d_row['fcff']}", f"${last}${d_row['t']}"
d_row["ronic_implied"] = r
put(sd, f"A{r}", "Implied return on new capital  tg / (net reinvestment FY31 / NOPAT FY31)", "text")
put(sd, f"B{r}", f"=tg/(({NP31}-{FC31})/{NP31})", "link", FMT_PCT2)
put(sd, f"{NOTE_COL}{r}", "net reinvestment = NOPAT - FCFF (capex - D&A + change in NWC)", "text")
name("ronic_implied", "DCF", f"B{r}")
r += 1
d_row["px_ronic"] = r
put(sd, f"A{r}", "Price at Assumptions!ronic ($)  terminal value = NOPAT_FY31 x (1+g) x (1 - g/RONIC) / (WACC - g)", "text")
put(sd, f"B{r}", f"=(SUMPRODUCT(${fc0}${d_row['fcff']}:${last}${d_row['fcff']},1/(1+wacc_calc)^${fc0}${d_row['t']}:${last}${d_row['t']})+{NP31}*(1+tg)*(1-tg/ronic)/(wacc_calc-tg)/(1+wacc_calc)^{T31c}+cash-debt)/shares*1000", "link", FMT_PX)
put(sd, f"{NOTE_COL}{r}", "same SUMPRODUCT repricing as the Sensitivity grids with a value-driver terminal value", "text")
name("px_at_ronic", "DCF", f"B{r}")
r += 1
d_row["g_ronic"] = r
TX = f"{X}{d_row['tv']}/(1+wacc_calc)^({X}{d_row['tvt']}-{P}{d_row['tvt']})"
QA_, QB_, QC_ = f"(-{NP31}/ronic)", f"({NP31}*(1-1/ronic)+{TX})", f"({NP31}-{TX}*wacc_calc)"
put(sd, f"A{r}", "Implied g from exit TV at Assumptions!ronic, timed to the exit convention", "text")
put(sd, f"B{r}", f"=(-{QB_}+SQRT({QB_}^2-4*{QA_}*{QC_}))/(2*{QA_})", "link", FMT_PCT2)
put(sd, f"{NOTE_COL}{r}", "solves T = NOPAT x (1+g) x (1 - g/RONIC) / (WACC - g), with T the exit TV pulled back to the mid-year point", "text")
name("g_exit_at_ronic", "DCF", f"B{r}")
r += 2

# ---- Block A: integrity checks (must all pass)
hdr_row(sd, r, ["Integrity checks (must all pass)", "Value", "Test", "Result"])
r += 1
chk_row = {}
def check(key, label, value_f, rng_text, pass_f, fmt, vkind="calc", rkind="calc"):
    global r
    chk_row[key] = r
    put(sd, f"A{r}", label, "text")
    put(sd, f"B{r}", value_f, vkind, fmt)
    put(sd, f"C{r}", rng_text, "text")
    put(sd, f"D{r}", pass_f, rkind)
    r += 1
fc_first = C["FY2027E"]
check("df1", "Y1 discount factor − 1/(1+WACC)^0.5", f"={fc_first}{d_row['df']}-1/(1+wacc_calc)^0.5", "exactly 0",
      f"=IF(ABS(B{r})<1E-9,\"PASS\",\"FAIL\")", "0.00E+00", "link")
check("pxim", "Price at market-implied margin (must equal current price)", price_at_margin(f"$B${d_row['im']}"), "equals current price",
      f"=IF(ABS(B{r}-price)<0.000001,\"PASS\",\"FAIL\")", FMT_PX, "link")
name("px_at_implied_margin", "DCF", f"B{chk_row['pxim']}")
check("sg1", "Sensitivity grid 1 center − price_perp", "=sens_diff_perp", "exactly 0 (Sensitivity tab)", "=sens_check_perp", "0.00E+00", "link", "link")
check("sg2", "Sensitivity grid 2 center − price_exit", "=sens_diff_exit", "exactly 0 (Sensitivity tab)", "=sens_check_exit", "0.00E+00", "link", "link")
check("mcr", "Monte Carlo deterministic run − price_perp", "=mc_recon_diff", "< $0.005 (MC tab)", "=mc_recon_check", "0.00E+00", "link", "link")
check("mcp", "Monte Carlo current_price − Assumptions price", "=mc_price_diff", "< $0.005 (MC tab)", "=mc_price_check", "0.00E+00", "link", "link")
a_first, a_last = chk_row["df1"], chk_row["mcp"]
d_row["integrity"] = r
put(sd, f"A{r}", "Integrity checks passing", "text", bold=True)
put(sd, f"D{r}", f"=COUNTIF(D{a_first}:D{a_last},\"PASS\")&\" of \"&ROWS(D{a_first}:D{a_last})&\" pass\"", "calc", bold=True)
name("integrity_summary", "DCF", f"D{r}")
r += 2

# ---- Block B: cross-checks (value vs reference range; outside range is a flag, not an error)
hdr_row(sd, r, ["Cross-checks (value vs reference range; outside range is a flag, not an error)", "Value", "Reference range", "Result"])
r += 1
x_row = {}
def xcheck(key, label, value_f, rng, fmt, vkind, flag_f=None):
    global r
    x_row[key] = r
    put(sd, f"A{r}", label, "text")
    put(sd, f"B{r}", value_f, vkind, fmt)
    put(sd, f"C{r}", rng, "calc" if str(rng).startswith("=") else "text")
    put(sd, f"D{r}", flag_f if flag_f else "context", "calc" if flag_f else "text")
    r += 1
def in_range(lo, hi):
    return f"=IF(AND(B{r}>={lo},B{r}<={hi}),\"in range\",\"outside range\")"
xcheck("tvpct", "TV % of EV (perpetuity)", f"={P}{d_row['tvpct']}", "60% to 80%", FMT_PCT, "calc", in_range(0.6, 0.8))
xcheck("impx", "Implied exit multiple from perpetuity TV / EBITDA_FY31, timed to the exit convention",
       f"={P}{d_row['tv']}*(1+wacc_calc)^({X}{d_row['tvt']}-{P}{d_row['tvt']})/{last}{d_row['ebitda']}", "8x to 12x", FMT_X, "link", in_range(8, 12))
xcheck("impg", "Implied g from exit multiple (timed to exit convention)",
       f"=({X}{d_row['tv']}/(1+wacc_calc)^({X}{d_row['tvt']}-{P}{d_row['tvt']})*wacc_calc-{last}{d_row['fcff']})/({X}{d_row['tv']}/(1+wacc_calc)^({X}{d_row['tvt']}-{P}{d_row['tvt']})+{last}{d_row['fcff']})",
       "1.5% to 3.5%", FMT_PCT2, "link", in_range(0.015, 0.035))
xcheck("meth", "Perpetuity vs exit gap, abs(price_perp / price_exit - 1)", f"=ABS({P}{d_row['px']}/{X}{d_row['px']}-1)", "<= 20%", FMT_PCT, "calc",
       f"=IF(B{r}<=0.2,\"in range\",\"outside range\")")
xcheck("fcfm", "Terminal FCFF margin (FY31 FCFF / revenue)", f"={last}{d_row['fcff']}/{last}{d_row['rev']}", "12% to 18%", FMT_PCT, "calc", in_range(0.12, 0.18))
xcheck("evltm", "Implied EV / LTM EBITDA  (perpetuity EV / Historical!ltm_ebitda)", f"={P}{d_row['ev']}/ltm_ebitda",
       "=\"peer median \"&TEXT(comps_med,\"0.0\")&\"x, DECK current \"&TEXT(deck_ev_ebitda,\"0.00\")&\"x\"", FMT_X, "link")
put(sd, f"{NOTE_COL}{x_row['impg']}", "(TV×WACC − FCFF)/(TV + FCFF), with the exit TV pulled back half a year to the mid-year point", "text")
name("implied_ev_ltm_ebitda", "DCF", f"B{x_row['evltm']}")
xcheck("evfy27", "Implied EV / FY27E EBITDA  (perpetuity EV / EBITDA FY2027E)", f"={P}{d_row['ev']}/{C['FY2027E']}{d_row['ebitda']}", "context", FMT_X, "calc")
name("implied_ev_fy27_ebitda", "DCF", f"B{x_row['evfy27']}")
b_first, b_last = x_row["tvpct"], x_row["evfy27"]
d_row["xflags"] = r
put(sd, f"A{r}", "Cross-checks outside reference range", "text", bold=True)
put(sd, f"B{r}", f"=IF(COUNTIF(D{b_first}:D{b_last},\"outside range\")=0,\"none\",COUNTIF(D{b_first}:D{b_last},\"outside range\")&\" outside range: \"&INDEX(A{b_first}:A{b_last},MATCH(\"outside range\",D{b_first}:D{b_last},0)))", "calc", bold=True)
name("cross_check_flags", "DCF", f"B{r}")
r += 1
widths(sd, {"A": 58, "B": 16, "C": 16, "D": 14, "E": 14, "F": 14, "G": 14, "H": 90})
sd.freeze_panes = "B3"

# ================================================================== SENSITIVITY
ss = ws["Sensitivity"]
put(ss, "A1", "Sensitivity: live formula grids. Each cell re-prices the DCF for the row WACC and the column g / exit multiple, "
              "using the DCF FCFF and EBITDA rows, the DCF t and n rows, and the defined names cash, debt, shares.", "text", bold=True)
put(ss, "A2", "Axes are centered on the base case (wacc_calc, tg, exit_mult) ± 2 steps; steps are Assumptions inputs "
              "sens_wacc_step / sens_g_step / sens_mult_step. The highlighted center cell must equal the DCF price.", "text")
FILL_CTR = PatternFill("solid", fgColor="FFFF00")
fc_, lc_ = C["FY2027E"], C["FY2031E"]
FCFF_RNG = f"DCF!${fc_}${d_row['fcff']}:${lc_}${d_row['fcff']}"
T_RNG = f"DCF!${fc_}${d_row['t']}:${lc_}${d_row['t']}"
FCFF31 = f"DCF!${lc_}${d_row['fcff']}"
EBITDA31 = f"DCF!${lc_}${d_row['ebitda']}"
T31 = f"DCF!${lc_}${d_row['t']}"      # 4.5 (mid-year)  -> perpetuity TV discount period
N31 = f"DCF!${lc_}${d_row['n']}"      # 5.0 (end-year)  -> exit TV discount period
KS = [-2, -1, 0, 1, 2]
grid_hdr = {}   # price_name -> header row of that grid (inner 3x3 = rows hr+2..hr+4, columns D..F)

def axis_formula(center_name, step_name, k):
    if k == 0:
        return f"={center_name}"
    sgn = "+" if k > 0 else "-"
    mag = abs(k)
    return f"={center_name}{sgn}{'' if mag == 1 else str(mag) + '*'}{step_name}"

def perp_body(w, g):
    return (f"=(SUMPRODUCT({FCFF_RNG},1/(1+{w})^{T_RNG})"
            f"+{FCFF31}*(1+{g})/({w}-{g})/(1+{w})^{T31}+cash-debt)/shares*1000")

def exit_body(w, m):
    return (f"=(SUMPRODUCT({FCFF_RNG},1/(1+{w})^{T_RNG})"
            f"+{EBITDA31}*{m}/(1+{w})^{N31}+cash-debt)/shares*1000")

def live_grid(top, title, row_label, col_label, col_center, col_step, col_fmt, body_fn, price_name, check_label):
    put(ss, f"A{top}", title, "text", bold=True)
    put(ss, f"A{top + 1}", f"rows: {row_label}   |   columns: {col_label}", "text")
    hr = top + 2
    grid_hdr[price_name] = hr
    put(ss, f"B{hr}", "WACC ↓  /  →", "text", bold=True)
    for j, k in enumerate(KS):
        put(ss, f"{L(3 + j)}{hr}", axis_formula(col_center, col_step, k), "link", col_fmt, bold=True)
    for i, k in enumerate(KS):
        rr = hr + 1 + i
        put(ss, f"B{rr}", axis_formula("wacc_calc", "sens_wacc_step", k), "link", FMT_PCT2, bold=True)
        for j in range(len(KS)):
            col = L(3 + j)
            c = put(ss, f"{col}{rr}", body_fn(f"$B{rr}", f"{col}${hr}"), "link", FMT_PX)
            c.border = Border(*(Side(style="hair"),) * 4)
            if i == 2 and j == 2:
                c.fill = FILL_CTR
                c.font = Font(color=GREEN, bold=True)
    center = f"E{hr + 3}"
    cr = hr + 7
    put(ss, f"A{cr}", check_label, "text")
    put(ss, f"B{cr}", f"={center}-{price_name}", "link", "0.00E+00")
    name("sens_diff_perp" if price_name == "price_perp" else "sens_diff_exit", "Sensitivity", f"B{cr}")
    put(ss, f"C{cr}", f'=IF(ABS(B{cr})<0.000001,"PASS","FAIL")', "calc", bold=True)
    return cr

chk1 = live_grid(4, "Grid 1: implied price ($), perpetuity method. WACC (rows) × terminal growth g (columns)",
                 "WACC = wacc_calc ± 2 × sens_wacc_step", "g = tg ± 2 × sens_g_step",
                 "tg", "sens_g_step", FMT_PCT2, perp_body, "price_perp",
                 "Check: center cell − DCF price_perp (must be 0)")
chk2 = live_grid(chk1 + 2, "Grid 2: implied price ($), exit-multiple method. WACC (rows) × EV/EBITDA exit multiple (columns)",
                 "WACC = wacc_calc ± 2 × sens_wacc_step", "multiple = exit_mult ± 2 × sens_mult_step",
                 "exit_mult", "sens_mult_step", FMT_X, exit_body, "price_exit",
                 "Check: center cell − DCF price_exit (must be 0)")
name("sens_check_perp", "Sensitivity", f"C{chk1}")
name("sens_check_exit", "Sensitivity", f"C{chk2}")
widths(ss, {"A": 50, "B": 16, "C": 13, "D": 13, "E": 13, "F": 13, "G": 13})

# ================================================================== WACC, continued: operating-beta comparison (needs perp_body from the Sensitivity section)
rw = w_row["wacc5"] + 1
def wrow2(key, label, formula, kind, fmt, note=""):
    global rw
    w_row[key] = rw
    put(sw, f"A{rw}", label, "text")
    put(sw, f"B{rw}", formula, kind, fmt=fmt)
    put(sw, f"C{rw}", note, "text")
    rw += 1
wrow2("bop", "Operating beta equivalent of the 5Y regression  beta_5y_raw / (1 − cash / market cap)", f"=B{w_row['b5r']}/(1-cash/B{w_row['mcap']})", "link", FMT_B,
      "cash drag removed, so it compares like for like with the cash-corrected Damodaran beta")
wrow2("keop", "Ke at operating beta  rf + beta_op × ERP", f"=B{w_row['rf']}+B{w_row['bop']}*B{w_row['erp']}", "calc", FMT_PCT2, "")
wrow2("waccop", "WACC at operating beta  Ke_op × We + Kd × (1 − t) × Wd", f"=B{w_row['keop']}*B{w_row['we']}+B{w_row['kdat']}*B{w_row['wd']}", "calc", FMT_PCT2, "equals Ke_op because Wd = 0")
wrow2("pxop", "Value per share at that Ke ($, perpetuity method)", perp_body(f"$B${w_row['waccop']}", "tg"), "link", FMT_PX,
      "same repricing as Sensitivity grid 1 with the operating-beta WACC swapped in")
name("beta_operating", "WACC", f"B{w_row['bop']}")
name("ke_operating", "WACC", f"B{w_row['keop']}")
name("wacc_operating", "WACC", f"B{w_row['waccop']}")
name("px_at_ke_operating", "WACC", f"B{w_row['pxop']}")

# ================================================================== COMPS
sc = ws["Comps"]
put(sc, "A1", "Comps: EV / LTM EBITDA (blue inputs from stockanalysis.com and GuruFocus), DECK implied value from peer multiples, football field", "text", bold=True)
hdr_row(sc, 3, ["Ticker", "EV / LTM EBITDA", "Source"])
COMPS = [("CROX", 7.72), ("BIRK", 9.50), ("NKE", 11.35), ("WWW", 11.42), ("SHOO", 12.54), ("ONON", 15.71)]
r = 4
comp_row = {}
for tkr, mult in COMPS:
    comp_row[tkr] = r
    put(sc, f"A{r}", tkr, "text")
    put(sc, f"B{r}", mult, "in", FMT_X)
    src = {"CROX": "stockanalysis.com / GuruFocus", "WWW": "GuruFocus"}.get(tkr, "stockanalysis.com") + ", retrieved Sep 2026"
    if tkr == "BIRK":
        src = "stockanalysis.com 9.50x vs TipRanks 18.46x. Kept: yfinance same-currency recompute 8.89x (EUR items at EURUSD 1.1465, 2026-09-22) is within 15% of 9.50x"
    put(sc, f"C{r}", src, "text")
    r += 1
rng = f"B4:B{r - 1}"
r += 1
stat_row = {}
for key, lab, f in [("min", "Minimum", f"=MIN({rng})"), ("q1", "25th percentile", f"=QUARTILE({rng},1)"), ("med", "Median", f"=MEDIAN({rng})"),
                    ("q3", "75th percentile", f"=QUARTILE({rng},3)"), ("max", "Maximum", f"=MAX({rng})")]:
    stat_row[key] = r
    put(sc, f"A{r}", lab, "text", bold=(key == "med"))
    put(sc, f"B{r}", f, "calc", FMT_X, bold=(key == "med"))
    r += 1
name("comps_q1", "Comps", f"B{stat_row['q1']}")
name("comps_med", "Comps", f"B{stat_row['med']}")
name("comps_q3", "Comps", f"B{stat_row['q3']}")

# DECK implied value from peer multiples (all formulas; LTM EBITDA is a green link to Historical)
r += 1
put(sc, f"A{r}", "DECK implied value from peer EV / LTM EBITDA multiples", "text", bold=True)
r += 1
hdr_row(sc, r, ["($ thousands unless stated)", "25th percentile", "Median", "75th percentile", "Source / formula"])
r += 1
imp_row = {}
def irow(key, label, fB, fC, fD, kind, fmt, note, bold=False):
    global r
    imp_row[key] = r
    put(sc, f"A{r}", label, "text", bold=bold)
    for col, f in zip("BCD", (fB, fC, fD)):
        put(sc, f"{col}{r}", f, kind, fmt, bold=bold)
    put(sc, f"E{r}", note, "text")
    r += 1
irow("mult", "Peer EV / LTM EBITDA (x)", f"=B{stat_row['q1']}", f"=B{stat_row['med']}", f"=B{stat_row['q3']}", "calc", FMT_X, "peer statistics above")
irow("ebitda", "DECK LTM EBITDA to Jun 30 2026", "=ltm_ebitda", "=ltm_ebitda", "=ltm_ebitda", "link", FMT_K,
     "Historical!ltm_ebitda = FY2026 + Q1 FY2027 - Q1 FY2026 (income from operations + D&A and accretion)")
irow("ev", "Implied enterprise value", *(f"={c}{imp_row['mult']}*{c}{imp_row['ebitda']}" for c in "BCD"), "calc", FMT_K, "peer multiple x DECK LTM EBITDA")
irow("cash", "+ Cash", "=cash", "=cash", "=cash", "link", FMT_K, "Assumptions!cash: Jun 30 balance sheet 1,602,589 less 32,168 repurchased to Jul 9 (10-Q subsequent events)")
irow("debt", "- Debt", "=debt", "=debt", "=debt", "link", FMT_K, "Assumptions!debt (0); leases excluded, same bridge as DCF")
irow("eq", "Implied equity value", *(f"={c}{imp_row['ev']}+{c}{imp_row['cash']}-{c}{imp_row['debt']}" for c in "BCD"), "calc", FMT_K, "EV + cash - debt")
irow("sh", "Diluted shares (count)", "=shares", "=shares", "=shares", "link", FMT_SH, "Assumptions!shares_diluted_current")
irow("px", "Implied price per share ($)", *(f"={c}{imp_row['eq']}*1000/{c}{imp_row['sh']}" for c in "BCD"), "calc", FMT_PX, "equity x 1000 / shares", bold=True)
name("px_comps_q1", "Comps", f"B{imp_row['px']}")
name("px_comps_med", "Comps", f"C{imp_row['px']}")
name("px_comps_q3", "Comps", f"D{imp_row['px']}")
r += 1
put(sc, f"A{r}", "DECK current EV / LTM EBITDA (x)", "text")
put(sc, f"B{r}", "=(price*shares/1000+debt-cash)/ltm_ebitda", "link", FMT_X)
put(sc, f"E{r}", "context: (price x shares / 1000 + debt - cash) / LTM EBITDA, at Assumptions!price", "text")
name("deck_ev_ebitda", "Comps", f"B{r}")
deck_mult_row = r
r += 1
put(sc, f"A{r}", "Median ex-IFRS peers (CROX, NKE, WWW, SHOO)", "text")
put(sc, f"B{r}", "=MEDIAN(" + ",".join(f"B{comp_row[t_]}" for t_ in ("CROX", "NKE", "WWW", "SHOO")) + ")", "calc", FMT_X)
put(sc, f"E{r}", "US GAAP reporters only; BIRK (EUR, IFRS) and ONON (CHF, IFRS) excluded", "text")
name("comps_med_ex_ifrs", "Comps", f"B{r}")
r += 1
put(sc, f"A{r}", "Value per share at DECK's current EV / LTM EBITDA multiple ($)", "text")
put(sc, f"B{r}", exit_body("wacc_calc", f"$B${deck_mult_row}"), "link", FMT_PX)
put(sc, f"E{r}", "DCF exit-multiple method with the market's current multiple in place of exit_mult", "text")
name("px_exit_at_current_mult", "Comps", f"B{r}")
r += 1
for key_, lab_, val_, src_ in (("deck_ev_ebitda_stockanalysis", "DECK EV / EBITDA per stockanalysis (x)", 7.39, "stockanalysis.com, Sep 21 2026"),
                               ("deck_ev_ebitda_gurufocus", "DECK EV / EBITDA per GuruFocus (x)", 7.62, "GuruFocus, Sep 1 2026"),
                               ("deck_ev_ebitda_10y_median_gurufocus", "DECK 10-year median EV / EBITDA per GuruFocus (x)", 13.49, "GuruFocus, Sep 1 2026")):
    put(sc, f"A{r}", lab_, "text")
    put(sc, f"B{r}", val_, "in", FMT_X)
    put(sc, f"E{r}", src_, "text")
    name(key_, "Comps", f"B{r}")
    r += 1
comps_next_row = r

# ================================================================== MC
sm = ws["MC"]
MC_CSV = os.path.join(ROOT, "mc", "mc_summary.csv")
MC_PNGS = [os.path.join(ROOT, "mc", "hist.png"), os.path.join(ROOT, "mc", "tornado.png")]
a_val = {rec["key"]: rec["value"] for rec in A}
mc_seed_txt = str(int(float(a_val["mc_seed"])))
put(sm, "A1", "Monte Carlo: perpetuity-method value per share. Written by mc/deck_mc.py (blue = values from mc_summary.csv).", "text", bold=True)
if os.path.exists(MC_CSV):
    MC = read_csv(MC_CSV)
    mcv = {rec["metric"]: rec for rec in MC}
    src_txt = f"mc/mc_summary.csv, Python output, seed {mc_seed_txt}"
    hdr_row(sm, 3, ["metric", "value", "source"])
    FMT_MC = {"runs": "#,##0", "seed": "0", "prob_value_gt_price": FMT_PCT,
              "guardrail_runs_violating_initially": "#,##0", "guardrail_total_redraws": "#,##0", "guardrail_redraw_passes": "0",
              "wacc_capm_reference": FMT_PCT2, "deterministic_diff": "0.00E+00", "corner_wacc_at_price": FMT_PCT2, "corner_wacc_sigma": "0.00"}
    MC_LABEL = {"prob_value_gt_price": "share of draws above current price (within modeled ranges)",
                "worst_corner_value": "worst corner at WACC +2 sd (growth, margin, g at floors)",
                "corner_wacc_at_price": "WACC at which that corner reprices to the current price",
                "corner_wacc_sigma": "that WACC in sd above the mc_wacc_normal mean"}
    r = 4
    mc_rows = {}
    def mc_section(title, keys):
        global r
        put(sm, f"A{r}", title, "text", bold=True); r += 1
        for k in keys:
            if k not in mcv:
                continue
            rec = mcv[k]
            v = float(rec["value"])
            fmt = FMT_MC.get(k, FMT_B if k.startswith("spearman_") else FMT_PX)
            put(sm, f"A{r}", MC_LABEL.get(k, k), "text")
            put(sm, f"B{r}", v, "in", fmt=fmt)
            put(sm, f"C{r}", src_txt, "text")
            mc_rows[k] = r
            r += 1
        r += 1
    mc_section("Distribution (value per share, $)",
               ["runs", "seed", "mean", "std", "P5", "P10", "P25", "P50", "P75", "P90", "P95", "min", "max", "worst_corner_value", "corner_wacc_at_price", "corner_wacc_sigma",
                "prob_value_gt_price", "current_price"])
    mc_section("Guardrail: WACC - g >= mc_guardrail_wacc_minus_g_min",
               ["guardrail_runs_violating_initially", "guardrail_total_redraws", "guardrail_redraw_passes"])
    mc_section("Tornado: Spearman rank correlation of driver draw vs value",
               [k for k in mcv if k.startswith("spearman_")])
    mc_section("Deterministic reconciliation (all drivers at mode)",
               ["deterministic_python", "deterministic_excel_price_perp", "deterministic_diff", "wacc_capm_reference"])
    # live reconciliation against the workbook as it stands now
    put(sm, f"A{r}", "Live check", "text", bold=True); r += 1
    rec_rows = {}
    for k in ("deterministic_python",):
        for rr in range(4, r):
            if sm[f"A{rr}"].value == k:
                rec_rows[k] = rr
    put(sm, f"A{r}", "price_perp (live link to DCF)", "text")
    put(sm, f"B{r}", "=price_perp", "link", fmt=FMT_PX)
    live_row = r; r += 1
    put(sm, f"A{r}", "|deterministic_python - live price_perp| < $0.005", "text")
    put(sm, f"B{r}", f"=ABS(B{rec_rows['deterministic_python']}-B{live_row})", "calc", fmt="0.00E+00")
    name("mc_recon_diff", "MC", f"B{r}")
    put(sm, f"C{r}", f'=IF(B{r}<0.005,"PASS","FAIL")', "calc", bold=True)
    name("mc_recon_check", "MC", f"C{r}")
    r += 1
    put(sm, f"A{r}", "|MC current_price - Assumptions price| < $0.005", "text")
    put(sm, f"B{r}", f"=ABS(B{mc_rows['current_price']}-price)", "link", fmt="0.00E+00")
    name("mc_price_diff", "MC", f"B{r}")
    put(sm, f"C{r}", f'=IF(B{r}<0.005,"PASS","FAIL")', "calc", bold=True)
    name("mc_price_check", "MC", f"C{r}")
    r += 2
    try:
        from openpyxl.drawing.image import Image as XLImage
        anchor_row = 3
        for png in MC_PNGS:
            if os.path.exists(png):
                img = XLImage(png)
                scale = 800 / img.width
                img.width, img.height = int(img.width * scale), int(img.height * scale)
                sm.add_image(img, f"E{anchor_row}")
                anchor_row += int(img.height / 20) + 3
    except Exception as ex:
        put(sm, f"A{r}", f"images not embedded: {ex}", "text")
    widths(sm, {"A": 48, "B": 16, "C": 46})
else:
    mc_rows = {}
    put(sm, "A3", "mc/mc_summary.csv not found: run py -3.13 mc/deck_mc.py, then rebuild.", "text")
    for _nm, _ref in (("mc_recon_check", "C5"), ("mc_price_check", "C6"), ("mc_recon_diff", "B5"), ("mc_price_diff", "B6")):
        put(sm, _ref, "n/a", "text")
        name(_nm, "MC", _ref)

# ================================================================== FOOTBALL FIELD (Comps tab; formulas and links, blue = typed inputs from the sources named beside them)
r = comps_next_row + 1
put(sc, f"A{r}", "Football field ($ per share): low and high per bar", "text", bold=True)
r += 1
hdr_row(sc, r, ["Bar", "Low", "High", "Source / formula"])
r += 1
ff_row = {}
g1, g2 = grid_hdr["price_perp"], grid_hdr["price_exit"]
inner1 = f"Sensitivity!$D${g1 + 2}:$F${g1 + 4}"
inner2 = f"Sensitivity!$D${g2 + 2}:$F${g2 + 4}"
def frow(key, label, lo, hi, kind, note):
    global r
    ff_row[key] = r
    put(sc, f"A{r}", label, "text")
    put(sc, f"B{r}", lo, kind, FMT_PX)
    put(sc, f"C{r}", hi, kind, FMT_PX)
    put(sc, f"D{r}", note, "text")
    name(f"ff_{key}_lo", "Comps", f"B{r}")
    name(f"ff_{key}_hi", "Comps", f"C{r}")
    r += 1
frow("dcf_perp", "DCF perpetuity (WACC +/-1 step, g +/-1 step)", f"=MIN({inner1})", f"=MAX({inner1})", "link",
     f"MIN / MAX of Sensitivity grid 1 inner 3x3, {inner1}")
frow("dcf_exit", "DCF exit multiple (WACC +/-1 step, multiple +/-1 step)", f"=MIN({inner2})", f"=MAX({inner2})", "link",
     f"MIN / MAX of Sensitivity grid 2 inner 3x3, {inner2}")
if "P10" in mc_rows and "P90" in mc_rows:
    frow("mc", "Monte Carlo P10 to P90", f"=MC!B{mc_rows['P10']}", f"=MC!B{mc_rows['P90']}", "link", "MC tab (mc/mc_summary.csv, seed per Assumptions!mc_seed)")
else:
    frow("mc", "Monte Carlo P10 to P90", "n/a", "n/a", "text", "MC tab empty: run py -3.13 mc/deck_mc.py, then rebuild")
frow("comps", "Trading comps (25th to 75th percentile peer multiple)", f"=B{imp_row['px']}", f"=D{imp_row['px']}", "calc", "implied price per share, this tab")
frow("pt", "Analyst price targets", 70, 184, "in", "stockanalysis.com, Sep 2026, 27 analysts")
frow("w52", "52-week range", 77.36, 122.29, "in", "stockanalysis.com, Sep 2026")
r += 1
put(sc, f"A{r}", "Base DCF price, perpetuity ($)", "text")
put(sc, f"B{r}", "=price_perp", "link", FMT_PX)
put(sc, f"D{r}", "DCF!price_perp", "text")
name("ff_base_dcf", "Comps", f"B{r}")
r += 1
put(sc, f"A{r}", "Current price ($)", "text")
put(sc, f"B{r}", "=price", "link", FMT_PX)
put(sc, f"D{r}", "Assumptions!price", "text")
name("ff_price", "Comps", f"B{r}")
widths(sc, {"A": 58, "B": 18, "C": 18, "D": 18, "E": 70})

# ================================================================== COVER
sv = ws["Cover"]
put(sv, "A1", "Deckers Outdoor Corporation (DECK): DCF valuation", "text", bold=True).font = F_TITLE
rows = [
    ("Company", "=company", None), ("Ticker", "=ticker", None), ("Base fiscal year", "=base_year", None),
    ("Price date", "=price_date", None), ("Current share price ($)", "=price", FMT_PX),
    ("Diluted shares (count)", "=shares", FMT_SH), ("WACC (live formula, WACC tab)", "=wacc_calc", FMT_PCT2),
    ("Implied price, perpetuity method ($)", "=price_perp", FMT_PX), ("Implied price, exit multiple method ($)", "=price_exit", FMT_PX),
    ("Upside, perpetuity", "=upside_perp", FMT_PCT), ("Upside, exit multiple", "=upside_exit", FMT_PCT),
    ("Integrity checks", "=integrity_summary", None),
    ("Cross-checks outside reference range", "=cross_check_flags", None),
]
r = 3
for lab, f, fmt in rows:
    put(sv, f"A{r}", lab, "text")
    c = put(sv, f"B{r}", f, "link", fmt)
    if "Implied price" in lab or "Upside" in lab:
        c.font = Font(color=GREEN, bold=True)
    r += 1
put(sv, f"A{r + 1}", "Color key: blue = hardcoded input (Assumptions / Historical / Comps inputs), black = same-sheet formula, green = cross-sheet link.", "text")
put(sv, f"A{r + 2}", "Units: $ thousands throughout except price per share ($) and share counts (count).", "text")
widths(sv, {"A": 42, "B": 22})

def autofit(sheet, cap=160):
    """Widen columns so no text label/header is truncated (text cells only; merged note rows skipped)."""
    merged = set()
    for mr in sheet.merged_cells.ranges:
        for row in sheet.iter_rows(min_row=mr.min_row, max_row=mr.max_row, min_col=mr.min_col, max_col=mr.max_col):
            for c in row:
                merged.add(c.coordinate)
    need = {}
    for row in sheet.iter_rows():
        for c in row:
            if not (isinstance(c.value, str) and not c.value.startswith("=")) or c.coordinate in merged:
                continue
            # text only truncates when the cell to its right is occupied; otherwise it overflows into the empty cell
            right = sheet.cell(row=c.row, column=c.column + 1).value
            if right is None or right == "":
                continue
            need[c.column_letter] = max(need.get(c.column_letter, 0), len(c.value))
    for col, n in need.items():
        cur = sheet.column_dimensions[col].width or 8.43
        sheet.column_dimensions[col].width = min(max(cur, n * 1.1 + 2), cap)
for _s in wb.worksheets:
    autofit(_s)
wb.properties.creator = "Salman Koshy"
wb.properties.lastModifiedBy = "Salman Koshy"
wb.properties.title = "DECK DCF"
wb.save(OUT)
print("saved", OUT)

def set_doc_props(path, creator="Salman Koshy", title="DECK DCF"):
    """Rewrite docProps/core.xml inside the saved xlsx: creator, lastModifiedBy and title (Excel resets them on Save)."""
    import zipfile, shutil, re as _re
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "docProps/core.xml":
                x = data.decode("utf-8")
                x = _re.sub(r"<dc:creator>.*?</dc:creator>", f"<dc:creator>{creator}</dc:creator>", x, flags=_re.S)
                x = _re.sub(r"<cp:lastModifiedBy>.*?</cp:lastModifiedBy>", f"<cp:lastModifiedBy>{creator}</cp:lastModifiedBy>", x, flags=_re.S)
                if "<dc:title>" in x:
                    x = _re.sub(r"<dc:title>.*?</dc:title>", f"<dc:title>{title}</dc:title>", x, flags=_re.S)
                else:
                    x = x.replace("<dc:creator>", f"<dc:title>{title}</dc:title><dc:creator>", 1)
                data = x.encode("utf-8")
            zout.writestr(item, data)
    shutil.move(tmp, path)
    print("document properties set: creator, lastModifiedBy, title")

def excel_recalc_and_save(path):
    """Open in Excel via COM, CalculateFullRebuild, Save: the file then carries calculated values."""
    try:
        import win32com.client as w32, pythoncom
        pythoncom.CoInitialize()
        xl = w32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
        wb_ = xl.Workbooks.Open(os.path.abspath(path))
        xl.CalculateFullRebuild()
        wb_.Save()
        wb_.Close(SaveChanges=False); xl.Quit()
        print("Excel COM: CalculateFullRebuild + Save done")
        return True
    except Exception as ex:
        print("Excel COM recalc/save skipped:", repr(ex))
        return False

def football_png(wbv, out_png):
    """Football field drawn only from Excel-cached values (data_only) of the named Comps-tab cells."""
    def dv(nm):
        dn = wbv.defined_names[nm].attr_text
        s_, ref_ = dn.split("!")
        return wbv[s_][ref_.replace("$", "")].value
    bars = [("DCF perpetuity", "dcf_perp"), ("DCF exit multiple", "dcf_exit"), ("Monte Carlo P10 to P90", "mc"),
            ("Trading comps 25th to 75th", "comps"), ("Analyst price targets", "pt"), ("52-week range", "w52")]
    rows_ = [(lab, float(dv(f"ff_{k}_lo")), float(dv(f"ff_{k}_hi"))) for lab, k in bars]
    base, px = float(dv("ff_base_dcf")), float(dv("ff_price"))
    print(f"{'football field ($ per share, Excel data_only)':48}{'low':>10}{'high':>10}")
    for lab, lo, hi in rows_:
        print(f"{lab:48}{lo:>10.2f}{hi:>10.2f}")
    print(f"{'base DCF price (perpetuity)':48}{base:>10.2f}")
    print(f"{'current price':48}{px:>10.2f}")
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.ticker import FuncFormatter
    BLUE_, BLUE_LT, ORANGE = "#2a78d6", "#86b6ef", "#eb6834"
    INK, INK2, MUTED, GRID, SURF = "#0b0b0b", "#52514e", "#898781", "#e1e0d9", "#fcfcfb"
    plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10, "axes.edgecolor": MUTED, "axes.labelcolor": INK2,
                         "xtick.color": INK2, "ytick.color": INK2, "axes.spines.top": False, "axes.spines.right": False})
    fig, ax = plt.subplots(figsize=(6, 3.7), dpi=200, facecolor=SURF)
    ax.tick_params(labelsize=11.5)
    ax.set_facecolor(SURF)
    n_ = len(rows_)
    ys = list(range(n_))[::-1]
    for y, (lab, lo, hi) in zip(ys, rows_):
        ax.barh(y, hi - lo, left=lo, height=0.55, color=BLUE_LT, edgecolor=BLUE_, linewidth=0.8)
        ax.text(lo - 1.5, y, f"${lo:,.2f}", va="center", ha="right", color=INK, fontsize=10.5, parse_math=False)
        ax.text(hi + 1.5, y, f"${hi:,.2f}", va="center", ha="left", color=INK, fontsize=10.5, parse_math=False)
    ax.set_yticks(ys)
    ax.set_yticklabels([x[0] for x in rows_])
    ax.tick_params(axis="y", length=0)
    ax.axvline(px, color=INK, linestyle=":", linewidth=1.6, zorder=1)
    ax.text(px, -0.62, f" Current price ${px:,.2f}", color=INK, fontsize=10.5, va="center", ha="left", parse_math=False)
    ax.plot([base], [ys[0]], marker="D", color=ORANGE, markersize=8, linestyle="none", zorder=5)
    ax.text(base, ys[0] + 0.40, f"Base DCF ${base:,.2f}", color=ORANGE, fontsize=11, ha="center", va="bottom", parse_math=False)
    ax.xaxis.grid(True, color=GRID, linewidth=0.8)
    ax.set_axisbelow(True)
    ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f"${v:,.0f}"))
    lo_all = min(min(x[1] for x in rows_), px)
    hi_all = max(max(x[2] for x in rows_), base)
    span = hi_all - lo_all
    ax.set_xlim(lo_all - 0.3 * span, hi_all + 0.24 * span)
    ax.set_ylim(-1.0, n_ + 0.05)
    fig.suptitle("DECK: valuation summary ($ per share)", color=INK, fontsize=13.5, x=0.02, y=0.985, ha="left")
    ax.set_xlabel("$ per share", fontsize=12)
    fig.subplots_adjust(left=0.40, right=0.975, top=0.87, bottom=0.17)
    fig.savefig(out_png, dpi=200, facecolor=SURF)
    plt.close(fig)
    print("wrote", out_png)

if excel_recalc_and_save(OUT):
    set_doc_props(OUT)
    from openpyxl import load_workbook
    _wbv = load_workbook(OUT, data_only=True)
    _dn = _wbv.defined_names["price_perp"].attr_text          # e.g. DCF!$B$35
    _sh, _ref = _dn.split("!")
    print("data_only price_perp:", _wbv[_sh][_ref.replace("$", "")].value)
    os.makedirs(os.path.join(ROOT, "onepager"), exist_ok=True)
    football_png(_wbv, os.path.join(ROOT, "onepager", "football.png"))
print("DCF rows:", {k: v for k, v in d_row.items() if k in ("rev", "fcff", "pv", "tv", "pvtv", "ev", "eq", "px", "up")})
print("defined names:", len(wb.defined_names))
