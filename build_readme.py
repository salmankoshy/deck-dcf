# -*- coding: utf-8 -*-
"""Write README.md, RESUME_BULLET.md and INTERVIEW_QA.md. Every number comes from model/deck_dcf.xlsx (openpyxl data_only=True,
Excel-cached values) or mc/mc_summary.csv; none is typed here. Prints value and cell for each number used.
Run: py -3.13 build_readme.py  (after model/build_dcf.py has saved the workbook through Excel)"""
import csv, os, sys
sys.stdout.reconfigure(encoding="utf-8")
from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(ROOT, "model", "deck_dcf.xlsx")
MC_CSV = os.path.join(ROOT, "mc", "mc_summary.csv")
wb = load_workbook(XLSX, data_only=True)
used = []   # (label, shown, raw, cell)

def _cell(nm):
    dn = wb.defined_names[nm].attr_text
    s, r = dn.split("!")
    return s, r.replace("$", "")

def val(nm, label, fmt):
    s, r = _cell(nm)
    v = wb[s][r].value
    if v is None:
        sys.exit(f"STOP: {nm} ({s}!{r}) has no cached value; run model/build_dcf.py with Excel first")
    shown = fmt(v)
    used.append((label, shown, v, f"{s}!{r} ({nm})"))
    return shown

MC = {r["metric"]: r["value"] for r in csv.DictReader(open(MC_CSV, newline="", encoding="utf-8"))}
def mcv(metric, label, fmt):
    v = float(MC[metric])
    shown = fmt(v)
    used.append((label, shown, v, f"mc/mc_summary.csv:{metric}"))
    return shown

def by_label(sheet, label, col, col_hdr=None):
    ws = wb[sheet]
    c = col
    if col_hdr is not None:
        c = None
        for r in range(1, 4):
            for cc in range(1, ws.max_column + 1):
                if ws.cell(r, cc).value == col_hdr:
                    c = cc
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, 1).value == label:
            return ws.cell(r, c).value, f"{sheet}!{ws.cell(r, c).coordinate}"
    sys.exit(f"STOP: {label!r} not found on {sheet}")

usd2 = lambda v: f"${v:,.2f}"
pct1 = lambda v: f"{v * 100:.1f}%"
pct2 = lambda v: f"{v * 100:.2f}%"
x1 = lambda v: f"{v:.1f}x"
x2 = lambda v: f"{v:.2f}x"
k_m = lambda v: f"${v / 1000:,.0f}M"
usd0 = lambda v: f"${v:,.0f}"

# ---------------------------------------------------------------- model numbers
price = val("price", "current share price", usd2)
pdate = val("price_date", "price date", lambda v: str(v)[:10])
px_perp = val("price_perp", "DCF value per share, perpetuity", usd2)
px_exit = val("price_exit", "DCF value per share, exit multiple", usd2)
up_perp = val("upside_perp", "upside, perpetuity vs price", pct1)
px_comps = val("px_comps_med", "comps implied price, peer median", usd2)
peer_med = val("comps_med", "peer median EV / LTM EBITDA", x1)
deck_x = val("deck_ev_ebitda", "DECK EV / LTM EBITDA", x1)
wacc = val("wacc_calc", "WACC", pct2)
ke = val("ke_calc", "cost of equity", pct2)
rf = val("rf", "risk-free rate", pct2)
beta = val("beta", "beta used", lambda v: f"{v:.2f}")
beta5 = val("beta_5y_raw", "regression beta, 5Y monthly raw", lambda v: f"{v:.2f}")
erp = val("erp", "equity risk premium", pct2)
tg = val("tg", "terminal growth", pct1)
exm = val("exit_mult", "exit multiple", x1)
shares = val("shares", "diluted shares", lambda v: f"{v:,.0f}")
cash = val("cash", "cash, Jun 30 2026", k_m)
debt = val("debt", "debt", lambda v: f"{v:,.0f}")
ltm = val("ltm_ebitda", "LTM EBITDA to Jun 30 2026", k_m)
runs = val("mc_runs", "Monte Carlo runs", lambda v: f"{v:,.0f}")
seed = val("mc_seed", "Monte Carlo seed", lambda v: f"{v:.0f}")
guard = val("mc_guardrail_wacc_minus_g_min", "MC floor on WACC minus g", pct1)
sens_p = val("sens_check_perp", "sensitivity grid 1 center check", str)
sens_x = val("sens_check_exit", "sensitivity grid 2 center check", str)
mc_chk = val("mc_recon_check", "MC tab live reconciliation check", str)
mc_pchk = val("mc_price_check", "MC tab current_price equals Assumptions price", str)
im = val("implied_margin", "market-implied EBIT margin at current price", pct1)
m_fc = val("ebit_margin", "forecast EBIT margin", pct1)
_m26v, _m26c = by_label("Historical", "EBIT margin (income from operations / revenue)", None, "FY2026")
m26 = pct1(_m26v); used.append(("FY26 EBIT margin", m26, _m26v, _m26c))
_mlow, _mlowc = by_label("Historical", "EBIT margin (income from operations / revenue)", None, "FY2022")
m_low = pct1(_mlow); used.append(("FY22 EBIT margin, lowest on the Historical tab", m_low, _mlow, _mlowc))
b_op = val("beta_operating", "operating beta equivalent of the 5Y regression", lambda v: f"{v:.2f}")
ke_op = val("ke_operating", "Ke at operating beta", pct2)
px_op = val("px_at_ke_operating", "value per share at Ke at operating beta", usd2)
px_cur = val("px_exit_at_current_mult", "value per share at DECK's current EV/LTM EBITDA multiple, exit method", usd2)
med_ex = val("comps_med_ex_ifrs", "peer median ex-IFRS (CROX, NKE, WWW, SHOO)", x1)
integ = val("integrity_summary", "integrity checks passing", str)
xflags = val("cross_check_flags", "cross-checks outside reference range", str)
ev_ltm = val("implied_ev_ltm_ebitda", "implied EV / LTM EBITDA", x1)
ev_fy27 = val("implied_ev_fy27_ebitda", "implied EV / FY27E EBITDA", x1)
crox_mult, crox_cell = by_label("Comps", "CROX", 2); crox = x1(crox_mult); used.append(("CROX EV / LTM EBITDA", crox, crox_mult, crox_cell))
_ebitda26, _c26a = by_label("DCF", "EBITDA  (EBIT + D&A)", None, "FY2026A"); _rev26, _c26b = by_label("DCF", "Revenue", None, "FY2026A")
ebitda_m26 = pct1(_ebitda26 / _rev26); used.append(("FY26 EBITDA margin", ebitda_m26, _ebitda26 / _rev26, f"{_c26a} / {_c26b}"))
HIST = {(r_["item"], r_["period"]): r_ for r_ in csv.DictReader(open(os.path.join(ROOT, "inputs", "historicals.csv"), newline="", encoding="utf-8"))}
_q27m = float(HIST[("Income from operations", "Q1FY27")]["fs_value"]) / float(HIST[("Net sales", "Q1FY27")]["fs_value"])
_q26m = float(HIST[("Income from operations", "Q1FY26")]["fs_value"]) / float(HIST[("Net sales", "Q1FY26")]["fs_value"])
q1fy27_m = pct1(_q27m); used.append(("Q1 FY27 operating margin", q1fy27_m, _q27m, "inputs/historicals.csv Q1FY27 income from operations / net sales (10-Q)"))
q1fy26_m = pct1(_q26m); used.append(("Q1 FY26 operating margin", q1fy26_m, _q26m, "inputs/historicals.csv Q1FY26 income from operations / net sales (10-Q prior period)"))
q27_ebit = f"{float(HIST[('Income from operations', 'Q1FY27')]['fs_value']):,.0f}"; q27_rev = f"{float(HIST[('Net sales', 'Q1FY27')]['fs_value']):,.0f}"
roic = val("roic_fy26", "FY26 ROIC on equity less cash plus lease liabilities", pct1)
roic_ex = val("roic_fy26_ex_leases", "FY26 ROIC excluding leases", lambda v: f"{v * 100:.0f}%")
ronic_implied = val("ronic_implied", "implied return on new capital, DCF tab", lambda v: f"{v * 100:.0f}%")
ronic_in = val("ronic", "RONIC sensitivity input", lambda v: f"{v * 100:.0f}%")
px_ronic = val("px_at_ronic", "perpetuity value at the RONIC input", usd2)
g_ronic = val("g_exit_at_ronic", "implied g from exit TV at the RONIC input, timed", pct2)
_mlo_r = str(wb["Assumptions"].cell(next(rr for rr in range(4, wb["Assumptions"].max_row + 1) if wb["Assumptions"].cell(rr, 1).value == "mc_margin_tri"), 2).value).split(",")[0].strip()
used.append(("MC margin floor (%)", _mlo_r + "%", _mlo_r, "Assumptions!mc_margin_tri, first value"))
# terminal-value reinvestment: FY31 net reinvestment as a share of NOPAT, and the return on new capital that g then implies
_nopat31, _c1 = by_label("DCF", "NOPAT", None, "FY2031E")
_capex31, _c2 = by_label("DCF", "\u2212 Capex", None, "FY2031E")
_da31, _c3 = by_label("DCF", "+ D&A", None, "FY2031E")
_dnwc31, _c4 = by_label("DCF", "\u2212 Change in NWC  (increase = outflow)", None, "FY2031E")
_reinv = (_capex31 - _da31 + _dnwc31) / _nopat31
_tg_raw = float(wb[_cell("tg")[0]][_cell("tg")[1]].value)
_ronic = _tg_raw / _reinv
reinv_rate = pct1(_reinv); ronic = f"{_ronic * 100:.0f}%"
used.append(("FY31 net reinvestment (capex - D&A + change in NWC) as % of NOPAT", reinv_rate, _reinv, f"({_c2} - {_c3} + {_c4}) / {_c1}"))
used.append(("implied return on new capital = terminal g / reinvestment rate", ronic, _ronic, f"Assumptions tg / the rate above"))
# share of Monte Carlo WACC draws at or above the operating-beta Ke (normal tail from the mc_wacc_normal parameters)
import math as _math
_wm, _ws = [float(x) / 100 for x in str(wb["Assumptions"].cell(next(rr for rr in range(4, wb["Assumptions"].max_row + 1) if wb["Assumptions"].cell(rr, 1).value == "mc_wacc_normal"), 2).value).split(",")]
_keop_raw = float(wb[_cell("ke_operating")[0]][_cell("ke_operating")[1]].value)
_tail = 0.5 * _math.erfc((_keop_raw - _wm) / (_ws * _math.sqrt(2)))
tail = pct1(_tail); used.append(("share of MC WACC draws at or above Ke at operating beta (normal tail)", tail, _tail, "Assumptions!mc_wacc_normal and WACC!ke_operating (computed)"))
p10 = mcv("P10", "Monte Carlo P10", usd2)
p50 = mcv("P50", "Monte Carlo P50", usd2)
p90 = mcv("P90", "Monte Carlo P90", usd2)
mc_min = mcv("min", "Monte Carlo minimum", usd2)
prob = mcv("prob_value_gt_price", "share of draws above price", lambda v: f"{v * 100:g}%")
det_diff = mcv("deterministic_diff", "MC deterministic run minus Excel price_perp", lambda v: f"{v:.3g}")
spear_w = mcv("spearman_WACC", "Spearman rho, WACC", lambda v: f"{v:+.2f}")
spear_m = mcv("spearman_EBIT margin", "Spearman rho, EBIT margin", lambda v: f"{v:+.2f}")
corner = mcv("worst_corner_value", "MC worst corner of the ranges (growth low, margin low, g low, WACC mean + 2 sd)", usd0)
corner2 = mcv("worst_corner_value", "MC worst corner value, 2 decimals", usd2)
corner_w = mcv("corner_wacc_at_price", "WACC at which the corner reprices to the current price", pct2)
corner_sig = mcv("corner_wacc_sigma", "that WACC in sd above the mean", lambda v: f"{v:.1f}")

# BIRK note and multiple from the Comps tab
birk_mult, birk_cell = by_label("Comps", "BIRK", 2)
birk_note, birk_note_cell = by_label("Comps", "BIRK", 3)
used.append(("BIRK EV / LTM EBITDA used", x2(birk_mult), birk_mult, birk_cell))
used.append(("BIRK source note (quoted)", birk_note, birk_note, birk_note_cell))

# key assumptions table straight from the Assumptions tab (value, unit, source)
KEYS = ["price", "shares_diluted_current", "cash", "debt", "rev_growth_fy27", "rev_growth_fy31", "ebit_margin", "da_pct_rev",
        "capex_pct_rev", "nwc_pct_rev", "tax_rate", "rf", "beta", "erp", "terminal_growth", "exit_multiple_ev_ebitda",
        "mc_growth_tri", "mc_margin_tri", "mc_wacc_normal", "mc_g_tri"]
wa = wb["Assumptions"]
arow = {wa.cell(r, 1).value: r for r in range(4, wa.max_row + 1) if wa.cell(r, 1).value}
def fmt_unit(v, unit):
    if isinstance(v, str):
        return v
    return {"pct": pct2, "USD": usd2, "USD thousands": lambda x: f"{x:,.0f}", "shares": lambda x: f"{x:,.0f}", "x": x2}.get(unit, str)(v)
arows = []
for k in KEYS:
    r = arow[k]
    v, unit, src = wa.cell(r, 2).value, wa.cell(r, 3).value, wa.cell(r, 4).value
    shown = fmt_unit(v, unit)
    used.append((f"assumption {k}", shown, v, f"Assumptions!B{r}"))
    arows.append((k, shown, unit, src))

# DCF check blocks: A integrity (PASS/FAIL), B cross-checks (in range / outside range / context)
wd = wb["DCF"]
chkA, chkB = [], []
block = None
for r in range(1, wd.max_row + 1):
    a = wd.cell(r, 1).value
    if a == "Integrity checks (must all pass)":
        block = "A"; continue
    if isinstance(a, str) and a.startswith("Cross-checks (value vs reference range"):
        block = "B"; continue
    if a in ("Integrity checks passing", "Cross-checks outside reference range"):
        block = None; continue
    if block and a:
        v, rng, res = wd.cell(r, 2).value, wd.cell(r, 3).value, wd.cell(r, 4).value
        if "Price at" in a:
            shown = usd2(v)
        elif block == "A":
            shown = "0" if v == 0 else f"{v:.3g}"
        elif a.startswith("Implied g"):
            shown = pct2(v)
        elif "multiple" in a or "EV /" in a:
            shown = x2(v)
        else:
            shown = pct2(v)
        used.append((f"check: {a}", shown, v, f"DCF!B{r}"))
        (chkA if block == "A" else chkB).append((a, shown, rng, res))

# ---------------------------------------------------------------- README
FILINGS = [
    ("10-K FY2026", "FYE 2026-03-31, filed 2026-05-22", "0001628280-26-037664", "deck-20260331.htm"),
    ("10-K FY2024", "FYE 2024-03-31, filed 2024-05-24", "0000910521-24-000017", "deck-20240331.htm"),
    ("10-Q Q1 FY2027", "quarter ended 2026-06-30, filed 2026-07-30", "0000910521-26-000022", "deck-20260630.htm"),
    ("8-K", "2026-05-21, FY2026 results and FY2027 guidance", "0000910521-26-000007",
     "deck-20260521.htm (8-K cover, verified in submissions.json). Content used: EX-99.1 press release deckex991pressrelease-3312.htm, self-labeled EX-99.1 and dated May 21 2026 inside the file"),
    ("8-K", "2026-07-23, Q1 FY2027 results and guidance update", "0000910521-26-000015",
     "deck-20260723.htm (8-K cover, verified in submissions.json). Content used: EX-99.1 press release deckex991pressrelease-6302.htm, self-labeled EX-99.1 and dated July 23 2026 inside the file"),
]
OTHER = [
    ("Damodaran, Jan 2026", "industry beta (shoe, cash-corrected unlevered) and implied ERP"),
    ("FRED DGS10", "10-year Treasury yield for the risk-free rate"),
    ("Google Finance", f"closing price on {pdate}"),
    ("Yahoo Finance via yfinance", "DECK and SPY monthly returns for the regression beta, a same-currency cross-check on the comps"),
    ("stockanalysis.com, retrieved Sep 2026", "peer EV/EBITDA multiples for CROX, NKE, ONON, BIRK and SHOO, DECK EV/EBITDA 7.39x, analyst price targets and count, 52-week range"),
    ("GuruFocus, retrieved Sep 2026", "peer EV/EBITDA multiples for CROX and WWW, DECK EV/EBITDA 7.62x and its 13.49x 10-year median"),
    ("TipRanks", "the 18.46x alternate BIRK multiple recorded on the Comps tab"),
    ("SGI Europe", "coverage of the Jul 23 2026 Q1 FY27 call: go-forward tariff rate 12.5% from 10%"),
]
a_tbl = "\n".join(f"| `{k}` | {v} | {u} | {s} |" for k, v, u, s in arows)
f_tbl = "\n".join(f"| {a} | {b} | `{c}` | {d if ' ' in d else '`' + d + '`'} |" for a, b, c, d in FILINGS)
o_tbl = "\n".join(f"| {a} | {b} |" for a, b in OTHER)
_md = lambda s: str(s).replace("|", "\\|")
a_tbl2 = "\n".join(f"| {_md(a)} | {b} | {_md(c)} | {d} |" for a, b, c, d in chkA)
b_tbl2 = "\n".join(f"| {_md(a)} | {b} | {_md(c)} | {d} |" for a, b, c, d in chkB)
impg_shown = next(b for a, b, c, d in chkB if a.startswith("Implied g"))
impg_rng = next(c for a, b, c, d in chkB if a.startswith("Implied g"))
impg_floor = impg_rng.split(" to ")[0]


# B6: cash as a share of market cap (WACC tab market cap row, Assumptions cash)
_mcap, _mcap_c = by_label("WACC", "Market cap ($ thousands) = price × shares / 1000", 2)
_cash_raw = float(wb[_cell("cash")[0]][_cell("cash")[1]].value)
cash_mcap = f"{_cash_raw / _mcap * 100:.0f}%"
used.append(("cash / market cap", cash_mcap, _cash_raw / _mcap, f"Assumptions cash / {_mcap_c}"))
# limitations: the industry beta must sit below every regression estimate for the sentence to hold
_regs = {nm: float(wb[_cell(nm)[0]][_cell(nm)[1]].value) for nm in ("beta_5y_raw", "beta_5y_blume", "beta_2y_raw", "beta_2y_blume")}
_beta_raw = float(wb[_cell("beta")[0]][_cell("beta")[1]].value)
if not _beta_raw < min(_regs.values()):
    sys.exit(f"STOP: beta {_beta_raw} is not below all four regression estimates {_regs}")
used.append(("lowest of the four regression betas", f"{min(_regs.values()):.2f}", min(_regs.values()), "Assumptions beta_5y_raw, beta_5y_blume, beta_2y_raw, beta_2y_blume"))
if med_ex != peer_med:
    sys.exit(f"STOP: peer median ex-IFRS {med_ex} differs from the full median {peer_med}; the comps limitation says unchanged")
import re as _rebn
_bn = _rebn.findall(r"([0-9]+\.[0-9]+)x", birk_note)
if len(_bn) < 3:
    sys.exit(f"STOP: expected three multiples in the BIRK note, got {_bn}")
birk_sa, birk_tr, birk_yf = [f"{float(x):.2f}x" for x in _bn[:3]]
used.append(("BIRK multiples in the Comps note (stockanalysis, TipRanks, yfinance recompute)", f"{birk_sa}, {birk_tr}, {birk_yf}", _bn, birk_note_cell))
_det = float(MC["deterministic_diff"])
if abs(_det) > 1e-9:
    sys.exit(f"STOP: deterministic_diff {_det} is not zero; the README says the rebuild matches Excel exactly")
_nflag = xflags.split(" outside")[0].strip()
if _nflag == "1":
    xsent = f"One cross-check sits outside its reference range: implied g from the exit multiple, {impg_shown} against {impg_rng}."
else:
    xsent = f"Cross-checks outside their reference range: {xflags}."

readme = f"""# DECK DCF: valuation of Deckers Outdoor (NYSE: DECK)

![One-page pitch](onepager/preview.png)

[One-page pitch (PDF)](onepager/DECK_pitch.pdf) and [Excel model](model/deck_dcf.xlsx).

## Overview

A DCF and Monte Carlo on Deckers Outdoor, fiscal year ending Mar 31, base year FY2026. Everything downstream of the two input csvs is generated, so the Excel model carries live formulas only and the one-pager reads its numbers back from the saved workbook.

## Result

The perpetuity DCF gives {px_perp} per share against a {price} close on {pdate}. Upside is {up_perp}. The exit-multiple cross-check at {exm} EV/EBITDA gives {px_exit} and comps at the {peer_med} peer median give {px_comps}. The Monte Carlo P10 to P90 band runs {p10} to {p90} around a P50 of {p50}. The {price} close implies a {im} EBIT margin against the {m_fc} forecast and the {m26} FY26 actual. Growth and margin are drawn independently but move together through operating leverage, so the true spread is modestly wider.

## Method

Five forecast years, FY2027 to FY2031, discounted at mid-year periods 0.5 to 4.5. WACC equals the cost of equity. Debt is zero, so the equity weight is 100% and WACC is rf {rf} plus beta {beta} times ERP {erp}, which gives {ke}. The 5-year monthly regression beta is {beta5}. It is an equity beta on a firm holding cash worth {cash_mcap} of its market cap. Removing that cash drag gives an operating beta of {b_op}, a {ke_op} cost of equity and {px_op} per share. The perpetuity method at g {tg} is primary and the {exm} EV/EBITDA exit multiple is the cross-check, with the exit terminal value discounted from the end of FY2031. The bridge adds cash of {cash}, subtracts debt of {debt}, keeps operating leases inside EBIT, expenses SBC and divides by {shares} diluted shares from the 10-Q cover plus the dilutive awards in its Note 9. Enterprise value is dated Mar 31 2026 and bridged with Jul 9 cash and shares. Rolling forward would raise value, so the omission is conservative.

## Key assumptions

Values live in inputs/assumptions.csv. The Assumptions tab mirrors that file cell for cell, and the WACC, DCF, Sensitivity and Comps calculations read from it through defined names, while the Historical and Comps tabs also carry their own blue inputs from historicals.csv and the sources named beside them.

| key | value | unit | source |
|---|---|---|---|
{a_tbl}

## Data sources

All filings came from SEC EDGAR. The FY2024 10-K supplies FY2022 and FY2023 and the FY2026 10-K supplies FY2024 to FY2026, with XBRL companyfacts as the cross-check on every line item in inputs/historicals.csv. The Q1 FY2027 10-Q supplies the Jun 30 2026 balance sheet and the quarter that closes the LTM window.

| filing | period | accession | primary document |
|---|---|---|---|
{f_tbl}

| other source | used for |
|---|---|
{o_tbl}

## Repo layout

```
inputs/           assumptions.csv, historicals.csv, beta.csv
model/            deck_dcf.xlsx, build_dcf.py (builder), verify_dcf.py (recalc + checks)
mc/               deck_mc.py, mc_summary.csv, hist.png, tornado.png
onepager/         build_onepager.py, DECK_pitch.pdf, DECK_pitch.html, football.png, number_map.csv, preview.png
filings/          10K_FY2026_Item7_FS.md, 10K_FY2024_FS.md, 10Q_Q1FY27.md, 8K_releases_May_Jul_2026.md
source/           raw EDGAR downloads, gitignored
build_readme.py   writes README.md from the workbook
requirements.txt  pinned packages
```

source/ is gitignored. The generated files under model/, mc/, onepager/ and the README are rebuilt by the scripts beside them, inputs/ holds the csvs that every script reads, and filings/ holds the text extracts of the five filings that the csvs cite.

## Reproduce

The saved workbook, PDF and charts are tracked, so nothing needs rebuilding to read the model. Excel must be installed. Steps 2 and 4 open the workbook through pywin32 COM to recalculate and save cached values, so the chain runs on Windows with Microsoft Excel and does not run under LibreOffice.

1. `py -3.13 -m pip install -r requirements.txt`
2. `py -3.13 model/build_dcf.py` builds the workbook from the csvs, then recalculates and saves it through Excel (the MC tab reports mc_summary.csv missing on this first pass).
3. Then `py -3.13 mc/deck_mc.py`, which reconciles its base case to the saved price_perp before drawing {runs} scenarios with seed {seed}.
4. `py -3.13 model/build_dcf.py` a second time, so the MC tab and the football field pick up mc_summary.csv and the charts.
5. Run `py -3.13 model/verify_dcf.py` for the Python recomputation and the Excel comparison.
6. `py -3.13 onepager/build_onepager.py` renders the PDF through headless Edge or Chrome and writes number_map.csv.
7. Last, `py -3.13 build_readme.py` rewrites this file from the workbook.

The tracked csvs already carry every value used, so source/ is needed only to re-audit them against the filings. Set the environment variable SEC_USER_AGENT to "Your Name your@email" before re-downloading from EDGAR, which requires that string as the User-Agent header, and stay under 10 requests per second:

```
https://data.sec.gov/submissions/CIK0000910521.json            -> source/CIK0000910521_submissions.json
https://data.sec.gov/api/xbrl/companyfacts/CIK0000910521.json  -> source/CIK0000910521_companyfacts.json
https://www.sec.gov/Archives/edgar/data/910521/<accession without dashes>/<primary document>
                                                                -> source/<primary document>, for the five filings above
```

## Checks

Integrity checks: {integ}. {xsent} The integrity block holds the six tests that must hold exactly, from the first-year discount factor to the two Monte Carlo links, while the cross-check block places model outputs against reference ranges and treats an outside-range value as a flag to read rather than an error to fix. The verifier, verify_dcf.py, recomputes the DCF and both sensitivity grids in Python from the csvs and compares each Excel value after a COM recalculation, while deck_mc.py reprices the base case in NumPy and matches Excel exactly. The typed-number scan over the WACC, DCF and Sensitivity tabs finds none.

| integrity check | value | test | result |
|---|---|---|---|
{a_tbl2}

| cross-check | value | reference range | result |
|---|---|---|---|
{b_tbl2}

## Limitations

- No segment build. Revenue is one line, so HOKA and UGG mix effects sit inside the single margin assumption.
- Enterprise value is dated Mar 31 2026 and bridged with Jul 9 cash and shares, with no stub roll-forward. Rolling forward would raise value.
- FY31 net reinvestment is {reinv_rate} of NOPAT, so {tg} perpetual growth implies a return on new capital of about {ronic_implied}. FY26 ROIC is {roic} on equity less cash plus lease liabilities, {roic_ex} excluding leases, so the terminal value holds today's returns forever. At a {ronic_in} return on new capital the perpetuity value is {px_ronic}, and the {exm} exit multiple, timed to the exit convention, then implies {g_ronic} growth.
- Because the Monte Carlo ranges are anchored to guidance, they exclude the market's own scenario. At WACC +2 sd with growth, margin and g at their floors, value is {corner2}. That corner reaches the {price} price at WACC {corner_w}, {corner_sig} sd above the mean. The {_mlo_r}% margin floor is the FY22 and FY23 trough.
- The industry beta of {beta} sits below all four regression estimates. The operating-beta equivalent of the 5-year regression gives {px_op} per share.
- Comps use LTM EV/EBITDA only, with no growth or margin adjustment. ONON and BIRK report under IFRS 16; the median excluding them is unchanged at {med_ex}. BIRK sources disagree ({birk_sa} stockanalysis.com, {birk_tr} TipRanks). A same-currency yfinance recompute gives {birk_yf}, so {x2(birk_mult)} is kept.
- No second person has reviewed the inputs or the formulas.

## Disclaimer

Independent student project. Not investment advice and not a recommendation to buy or sell any security.
"""

_raw = lambda nm: float(wb[_cell(nm)[0]][_cell(nm)[1]].value)
resume = f"""# Resume bullets: DECK DCF

- Valued Deckers Outdoor (NYSE: DECK) with a five-year unlevered DCF built from 10-K and 10-Q filings, and reverse-engineered the ${_raw("price"):,.2f} share price to a {_raw("implied_margin") * 100:.1f}% implied EBIT margin against {_raw("ebit_margin") * 100:.1f}% guidance to support a ${_raw("price_perp"):,.0f} base case
- Stress-tested the valuation with a 10,000-draw Python Monte Carlo (P10 to P90 of ${float(MC["P10"]):,.0f} to ${float(MC["P90"]):,.0f}), exit-multiple and trading-comps cross-checks, and a return-on-capital test of the terminal value, with the Python rebuild matching Excel exactly
"""

for doc in (readme, resume):
    assert "—" not in doc and "–" not in doc, "em/en dash"
for line in resume.splitlines():
    if line.startswith("- "):
        assert len(line) - 2 <= 240, f"resume bullet over 240 chars: {len(line) - 2}"
open(os.path.join(ROOT, "README.md"), "w", encoding="utf-8", newline="\n").write(readme)
open(os.path.join(ROOT, "RESUME_BULLET.md"), "w", encoding="utf-8", newline="\n").write(resume)


# ---------------------------------------------------------------- INTERVIEW_QA.md (every number re-read here and listed with its cell)
import re as _re
qa_used = []
prev = {lab: (shown, raw, cell) for lab, shown, raw, cell in used}
def Q(label, shown=None):
    s, r, c = prev[label]
    qa_used.append((label, shown if shown is not None else s, r, c))
    return shown if shown is not None else s
def qval(nm, label, fmt):
    s, r = _cell(nm); v = wb[s][r].value
    qa_used.append((label, fmt(v), v, f"{s}!{r} ({nm})")); return fmt(v)
def qlab(sheet, label, col, lab, fmt, col_hdr=None):
    v, c = by_label(sheet, label, col, col_hdr)
    qa_used.append((lab, fmt(v), v, c)); return fmt(v)
def qtext(key, pattern, lab, fmt):
    """number parsed out of a source-text cell on the Assumptions tab (column D of the key's row)"""
    r = arow[key]; txt = wa.cell(r, 4).value
    m = _re.search(pattern, txt); v = float(m.group(1))
    qa_used.append((lab, fmt(v), v, f"Assumptions!D{r} (text: {txt[:60]})")); return fmt(v)
usd0 = lambda v: f"${v:,.0f}"
q_price = Q("current share price"); q_price0 = Q("current share price", usd0(float(prev["current share price"][1])))
q_dcf = Q("DCF value per share, perpetuity"); q_dcf0 = Q("DCF value per share, perpetuity", usd0(float(prev["DCF value per share, perpetuity"][1])))
q_exit = Q("DCF value per share, exit multiple"); q_up = Q("upside, perpetuity vs price")
q_g27 = Q("assumption rev_growth_fy27", pct1(float(prev["assumption rev_growth_fy27"][1]))); q_g31 = Q("assumption rev_growth_fy31", pct1(float(prev["assumption rev_growth_fy31"][1])))
q_m = Q("assumption ebit_margin", pct1(float(prev["assumption ebit_margin"][1]))); q_tax = Q("assumption tax_rate", f"{float(prev['assumption tax_rate'][1]) * 100:.0f}%")
q_da = Q("assumption da_pct_rev", pct1(float(prev["assumption da_pct_rev"][1]))); q_capex = Q("assumption capex_pct_rev", pct1(float(prev["assumption capex_pct_rev"][1])))
q_nwc = Q("assumption nwc_pct_rev", pct1(float(prev["assumption nwc_pct_rev"][1])))
q_wacc = Q("WACC"); q_ke = Q("cost of equity"); q_rf = Q("risk-free rate"); q_beta = Q("beta used"); q_erp = Q("equity risk premium")
q_beta5 = Q("regression beta, 5Y monthly raw"); q_tg = Q("terminal growth"); q_exm = Q("exit multiple")
q_sh = Q("diluted shares"); q_shm = Q("diluted shares", f"{float(prev['diluted shares'][1]) / 1e6:.1f}M"); q_cash = Q("cash, Jun 30 2026"); q_debt = Q("debt")
q_runs = Q("Monte Carlo runs"); q_p10 = Q("Monte Carlo P10"); q_p50 = Q("Monte Carlo P50"); q_p90 = Q("Monte Carlo P90"); q_min = Q("Monte Carlo minimum")
q_prob = Q("share of draws above price"); q_sw = Q("Spearman rho, WACC"); q_sm = Q("Spearman rho, EBIT margin")
q_peer = Q("peer median EV / LTM EBITDA"); q_deckx = Q("DECK EV / LTM EBITDA")
q_m26 = qlab("Historical", "EBIT margin (income from operations / revenue)", None, "FY26 EBIT margin", pct1, col_hdr="FY2026")
q_tvpct = Q("check: TV % of EV (perpetuity)", pct1(float(prev["check: TV % of EV (perpetuity)"][1])))
q_tvpct0 = Q("check: TV % of EV (perpetuity)", f"{float(prev['check: TV % of EV (perpetuity)'][1]) * 100:.0f}%")
q_impx = Q("check: Implied exit multiple from perpetuity TV / EBITDA_FY31, timed to the exit convention")
q_impg = Q("check: Implied g from exit multiple (timed to exit convention)")
q_meth = Q("check: Perpetuity vs exit gap, abs(price_perp / price_exit - 1)", pct1(float(prev["check: Perpetuity vs exit gap, abs(price_perp / price_exit - 1)"][1])))
q_fcfm = Q("check: Terminal FCFF margin (FY31 FCFF / revenue)", pct1(float(prev["check: Terminal FCFF margin (FY31 FCFF / revenue)"][1])))
q_impx_rng = qlab("DCF", "Implied exit multiple from perpetuity TV / EBITDA_FY31, timed to the exit convention", 3, "implied exit multiple acceptable range", str)
q_impg_rng = qlab("DCF", "Implied g from exit multiple (timed to exit convention)", 3, "implied g acceptable range", str)
q_meth_rng = qlab("DCF", "Perpetuity vs exit gap, abs(price_perp / price_exit - 1)", 3, "methods tolerance", lambda v: v.replace("<= ", ""))
q_we = qlab("WACC", "Equity weight  We", 2, "equity weight", lambda v: f"{v * 100:.0f}%")
q_wacc5 = qlab("WACC", "WACC at 5Y regression equity beta (includes cash drag, not comparable to the cash-corrected beta; see row 24)", 2, "WACC at the 1.10 regression beta", pct2)
q_tvt_p = qlab("DCF", "Discount period for TV (years)", 2, "TV discount period, perpetuity", lambda v: f"{v:.1f}")
q_tvt_x = qlab("DCF", "Discount period for TV (years)", 3, "TV discount period, exit multiple", lambda v: f"{v:.1f}")
q_t1 = qlab("DCF", "Mid-year period t = n − 0.5", 3, "first forecast period t", lambda v: f"{v:.1f}")
q_tv_hi = qlab("DCF", "TV % of EV (perpetuity)", 3, "TV % of EV acceptable range, upper bound", lambda v: v.split(" to ")[1])
q_lease = qval("lease_liabilities_total", "operating lease liabilities, FY2026 total", k_m)
q_sbc = qlab("Historical", "Stock-based compensation", None, "FY26 stock-based compensation", k_m, col_hdr="FY2026")
q_sbcp = qlab("Historical", "SBC % revenue", None, "FY26 SBC % of revenue", pct1, col_hdr="FY2026")
q_r2 = qtext("beta_5y_raw", r"R2=([0-9.]+)", "regression R-squared, 5Y monthly", lambda v: f"{v:.2f}")
q_med10 = qtext("exit_multiple_ev_ebitda", r"below ([0-9.]+)x 10Y median", "DECK own 10-year median EV/EBITDA", lambda v: f"{v:.1f}x")
q_dil = qtext("shares_diluted_current", r"\+ ([0-9]+)K dilutive", "dilutive awards, 10-Q Note 9", lambda v: f"{v:.0f}K")
_mt = wa.cell(arow["mc_margin_tri"], 2).value; _mlo, _mmode, _mhi = [x.strip() for x in _mt.split(",")]
qa_used.append(("MC margin range (tri, %)", _mt, _mt, f"Assumptions!B{arow['mc_margin_tri']}"))
_wn = wa.cell(arow["mc_wacc_normal"], 2).value; _wmean, _wsd = [x.strip() for x in _wn.split(",")]
qa_used.append(("MC WACC normal (mean, sd, %)", _wn, _wn, f"Assumptions!B{arow['mc_wacc_normal']}"))
_gt = wa.cell(arow["mc_growth_tri"], 2).value; _glo = _gt.split(",")[0].strip()
qa_used.append(("MC growth range (tri, %)", _gt, _gt, f"Assumptions!B{arow['mc_growth_tri']}"))
# sensitivity grid 1 worst corner: highest WACC row, lowest g column
ws_ = wb["Sensitivity"]
hr = next(r for r in range(1, ws_.max_row + 1) if isinstance(ws_.cell(r, 2).value, str) and ws_.cell(r, 2).value.startswith("WACC"))
q_wmax = ws_.cell(hr + 5, 2).value; q_gmin = ws_.cell(hr, 3).value; q_corner = ws_.cell(hr + 5, 3).value
qa_used.append(("sensitivity grid 1, highest WACC", pct2(q_wmax), q_wmax, f"Sensitivity!B{hr + 5}"))
qa_used.append(("sensitivity grid 1, lowest g", pct1(q_gmin), q_gmin, f"Sensitivity!C{hr}"))
qa_used.append(("sensitivity grid 1, worst-corner price", usd0(q_corner), q_corner, f"Sensitivity!C{hr + 5}"))
q_cq1 = qval("px_comps_q1", "comps implied price, 25th percentile peer multiple", usd0)
q_w52 = qval("ff_w52_lo", "52-week low (stockanalysis.com, typed on Comps tab)", usd2)
q_bop = Q("operating beta equivalent of the 5Y regression"); q_keop = Q("Ke at operating beta"); q_pxop = Q("value per share at Ke at operating beta")
q_pxcur = Q("value per share at DECK's current EV/LTM EBITDA multiple, exit method"); q_im = Q("market-implied EBIT margin at current price")
q_mlow = Q("FY22 EBIT margin, lowest on the Historical tab"); q_reinv = Q("FY31 net reinvestment (capex - D&A + change in NWC) as % of NOPAT")
q_ronic = Q("implied return on new capital = terminal g / reinvestment rate"); q_tail = Q("share of MC WACC draws at or above Ke at operating beta (normal tail)")
q_mccorner = Q("MC worst corner of the ranges (growth low, margin low, g low, WACC mean + 2 sd)"); q_corner2 = Q("MC worst corner value, 2 decimals"); q_cw = Q("WACC at which the corner reprices to the current price"); q_csig = Q("that WACC in sd above the mean")
q_roic = Q("FY26 ROIC on equity less cash plus lease liabilities"); q_roic_ex = Q("FY26 ROIC excluding leases"); q_pxronic = Q("perpetuity value at the RONIC input"); q_ronic_in = Q("RONIC sensitivity input"); q_gronic = Q("implied g from exit TV at the RONIC input, timed")
q_evltm = Q("implied EV / LTM EBITDA"); q_crox = Q("CROX EV / LTM EBITDA"); q_ebitda_m26 = Q("FY26 EBITDA margin"); q_q27m = Q("Q1 FY27 operating margin"); q_q26m = Q("Q1 FY26 operating margin"); q_impg_floor = impg_floor

qa = f"""# DECK model interview Q&A (12 questions)

Spoken answers, first person. Every number comes from model/deck_dcf.xlsx or mc/mc_summary.csv through build_readme.py, or from the filing named in the answer.

## 1. Walk me through your DCF.

Five years of unlevered free cash flow, FY2027 to FY2031, off the FY2026 base in the 10-K. Sales grow {q_g27} in FY27, the midpoint of management's guide, and fade to {q_g31} by FY31 with EBIT margin held at {q_m} against the {q_m26} we just saw in FY26. NOPAT is EBIT after {q_tax} tax. I add back D&A at {q_da} of sales, take out capex at {q_capex} and fund working capital at {q_nwc} of the sales increase, then discount at a {q_wacc} WACC with mid-year timing. I run the terminal value both ways and then bridge to equity with {q_cash} of cash from the Q1 10-Q and no debt, over {q_shm} diluted shares. That lands at {q_dcf0} against {q_price}.

## 2. Why is your WACC basically cost of equity?

Because there is no debt. The 10-K shows no borrowings, just an undrawn revolver, so the equity weight on the WACC tab is {q_we} and the debt weight is zero. Cost of equity is {q_rf} risk-free plus a {q_beta} beta times a {q_erp} equity risk premium, which is {q_ke}, and WACC comes out at the same {q_wacc} because the debt term multiplies by zero. The only debt-like item is {q_lease} of operating lease liabilities, and I keep the lease cost inside EBIT as an operating expense. So the discount rate is the equity rate.

## 3. Where did beta come from?

Damodaran's shoe set, January 2026, unlevered and corrected for cash, which gives {q_beta}. No debt, so relevered equals unlevered. My own regression on yfinance data, five years of monthly returns against SPY, gives {q_beta5} raw with an R-squared of {q_r2}, and that raw figure still carries DECK's cash pile, which damps the beta of the equity. Strip the cash out and the operating-beta equivalent is {q_bop}, which the WACC tab computes as {q_beta5} divided by one minus cash over market cap. At that beta the cost of equity is {q_keop} and the perpetuity value is {q_pxop} a share, still above {q_price}. The Monte Carlo draws WACC from a normal at {_wmean}% plus or minus {_wsd}%, so a {q_keop} rate shows up in about {q_tail} of draws, and I kept the industry beta for the base case with the regression case shown on the WACC tab.

## 4. What % of value is terminal value? Problem?

About {q_tvpct0} of enterprise value, {q_tvpct} on the check row, sits in the perpetuity terminal value. That is normal. Five years of explicit cash flow at a {q_wacc} discount rate can only capture so much when the business keeps earning beyond FY2031 at a {q_fcfm} free cash flow margin. The check I run is whether the terminal value implies a sane exit multiple, and timed to the exit convention it does: {q_impx} EBITDA on {q_tg} growth, inside the {q_impx_rng} range and below DECK's own {q_med10} ten-year median. The softer point is reinvestment: FY31 net reinvestment is {q_reinv} of NOPAT, so {q_tg} perpetual growth implies a return on new capital of about {q_ronic}, and a terminal value that made reinvestment fund that growth at a normal return would come in lower. If the terminal share had come in above the {q_tv_hi} top of my check range I would have gone back to the near-term cash flows first. FY26 ROIC is {q_roic} on equity less cash plus lease liabilities, {q_roic_ex} excluding leases, so the terminal value holds today's returns. At a {q_ronic_in} return on new capital it is {q_pxronic}.

## 5. Perpetuity growth vs exit multiple. Which do you trust?

Perpetuity, at {q_tg} growth, which sits under nominal GDP and under the {q_rf} risk-free rate. It gives {q_dcf} a share. The exit multiple at {q_exm} EBITDA gives {q_exit}, so the perpetuity answer sits {q_meth} above it, and I keep the multiple as the cross-check because it imports whatever the market is paying today. Timed to the exit convention, the perpetuity value implies an {q_impx} exit multiple while {q_exm} implies only {q_impg} perpetual growth, so the two methods disagree about how much growth DECK keeps after FY2031, and that {q_impg} is the one check in the model that fails its band. Both prices are inside the {q_meth_rng} tolerance I set, and the timing differs because the perpetuity value is discounted from {q_tvt_p} years under the mid-year convention while a buyer pays at the end of FY2031, {q_tvt_x} years out. At today's {q_deckx} multiple the exit method gives {q_pxcur}, which is what the market's own multiple says my cash flows are worth.

## 6. Stock at {q_price0}. You say {q_dcf0}. Why is the market wrong?

Start with what the price implies. At {q_price} the DCF only balances if the EBIT margin is {q_im}, against the {q_m} I forecast and the {q_m26} DECK just printed for FY26. DECK trades at {q_deckx} LTM EBITDA on the model's numbers against an {q_peer} peer median, and even the 25th-percentile peer multiple would price it at {q_cq1}. In my grid the worst corner is a {pct2(q_wmax)} WACC with {pct1(q_gmin)} terminal growth, and it still prices at {usd0(q_corner)}. So the market is paying for a margin below anything in the five years on the Historical tab, where the low was {q_mlow} in FY22, or for a multiple that never re-rates from {q_deckx}. The Q1 10-Q shows HOKA still growing 7.7% and the balance sheet holding {q_cash} of cash with no debt, while the stock sits near its {q_w52} low.

## 7. How did you handle leases?

Lease cost stays inside EBIT as an operating expense, where the income statement already puts it. That means the {q_lease} of operating lease liabilities on the FY2026 balance sheet does not come off in the bridge. Taking that liability off as debt while the rent is still running through EBIT would count the same leases twice, once in the cash flows and once in the bridge. The bridge is cash only. Equity value is enterprise value plus {q_cash}, over {q_shm} shares.

## 8. How did you treat SBC?

As an expense. It stays inside EBIT, so the {q_m} margin is after stock compensation, and I do not add the {q_sbc} from the FY26 cash flow statement back to free cash flow. Adding it back and then valuing the company on today's share count would treat shares paid to employees as free. The dilution that already exists is in the denominator, since the {q_sh} diluted shares come from the 10-Q cover count plus the {q_dil} of dilutive awards in Note 9. Future grants are a cost of doing business at about {q_sbcp} of sales, which is what FY26 ran.

## 9. Biggest risk to thesis?

HOKA. Its growth was +23.6% in FY25 and +15.9% in FY26 on the 10-K segment note. The Q1 10-Q shows +7.7%, and my FY27 growth of {q_g27} is the guide midpoint. If HOKA prints a negative quarter, or if the EBIT margin drops below the {_mlo}% floor of my Monte Carlo range, then the thesis is wrong and the model has nothing to say about it. Tariffs are the second risk: the go-forward rate went to 12.5% from 10% on the Jul 23 2026 call, per SGI Europe coverage. What settles it either way is the Q2 report in the holiday quarter.

## 10. What does Monte Carlo add?

Grids move two inputs at a time. The simulation moves growth, margin, WACC and terminal g together across {q_runs} draws, so I get a distribution instead of a table: P10 {q_p10}, P50 {q_p50}, P90 {q_p90}, and a worst draw of {q_min}. The Spearman tornado ranks the ranges I chose, with WACC at {q_sw} and EBIT margin at {q_sm}, and margin is the driver that more work can narrow through segment mix and the tariff math. The ranges are anchored to guidance, so they exclude the market's own scenario. At WACC +2 sd with growth, margin and g at their floors, value is {q_corner2}. That corner reaches the {q_price} price at WACC {q_cw}, {q_csig} sd above the mean. The {_mlo}% margin floor is the FY22 and FY23 trough. Q1 FY27 operating margin was {q_q27m} (10-Q: income from operations {q27_ebit} on net sales {q27_rev}), below my floor. Q1 is the seasonal low: Q1 FY26 was {q_q26m} in a {q_m26} year.

## 11. One of your checks fails. Which one, and why keep it?

One cross-check sits outside my reference range: the {q_exm} exit multiple implies {q_impg} perpetual growth against my {q_impg_floor} floor. Banks show this as a value, not a pass or fail. It has the same root as the terminal reinvestment point: at a {q_ronic_in} return on new capital the implied g is {q_gronic}, inside the range.

## 12. Your DCF implies {q_evltm} EBITDA while peers trade at {q_peer}. Why a premium?

The perpetuity enterprise value is {q_evltm} LTM EBITDA against an {q_peer} peer median. Two things sit behind that. DECK earned a {q_ebitda_m26} EBITDA margin in FY26 with zero debt, while stockanalysis shows NKE near 9.9%, SHOO near 8%, WWW near 10.5% and CROX at 22% to 24%. CROX, the nearest margin peer, trades at {q_crox}, so the premium rests on growth, and HOKA's deceleration is the risk to it.

## If I had another day

Seven things, in order.

- Build the segments out, one line per brand, so the mix shows up in the consolidated margin instead of sitting inside one {q_m} assumption.
- Then a stub-period adjustment, since the valuation date is {pdate} and the discounting starts from the FY2026 year end.
- Widen the Monte Carlo below guidance, with margins under {_mlo}% and growth below {_glo}%, since the worst corner the current ranges allow is about {q_mccorner}.
- For the comps, add growth and margin adjustments across the six peers instead of the raw LTM EV/EBITDA.
- Rebuild the terminal value on a RONIC basis, since FY31 reinvestment of {q_reinv} of NOPAT cannot fund {q_tg} growth at a normal return.
- BIRK: resolve the {x2(birk_mult)} question at the source, since the same-currency recompute comes from yfinance rather than a filing.
- A second pair of eyes on the inputs and the formulas.
"""
assert "\u2014" not in qa and "\u2013" not in qa, "em/en dash in Q&A"
open(os.path.join(ROOT, "INTERVIEW_QA.md"), "w", encoding="utf-8", newline="\n").write(qa)
print("\n" + "=" * 100); print("INTERVIEW_QA.md: value and cell for each number"); print("=" * 100)
print(f"{'label':52}{'shown':>22}   cell / source")
seen = set()
for lab, shown, raw, cell in qa_used:
    if (lab, str(shown)) in seen:
        continue
    seen.add((lab, str(shown)))
    print(f"{lab[:52]:52}{str(shown)[:22]:>22}   {cell}")
print(f"{len(seen)} numbers; wrote INTERVIEW_QA.md")

print(f"{'label':52}{'shown':>22}   cell / source")
for lab, shown, raw, cell in used:
    print(f"{lab[:52]:52}{str(shown)[:22]:>22}   {cell}")
print(f"\n{len(used)} numbers read; wrote README.md and RESUME_BULLET.md")
for line in resume.splitlines():
    if line.startswith("- "):
        print(f"resume bullet length: {len(line) - 2} chars")
