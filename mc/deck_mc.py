# -*- coding: utf-8 -*-
"""DECK Monte Carlo on the perpetuity DCF. Run: py -3.13 mc/deck_mc.py

All inputs come from inputs/assumptions.csv; base FY2026 revenue and NWC come from
inputs/historicals.csv (fs_value). FCFF logic mirrors the Excel DCF tab.
Outputs: mc/mc_summary.csv, mc/hist.png, mc/tornado.png.
"""
import csv, os, sys, shutil, tempfile
sys.stdout.reconfigure(encoding="utf-8")
import numpy as np
from scipy import stats
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MC_DIR = os.path.join(ROOT, "mc")
XLSX = os.path.join(ROOT, "model", "deck_dcf.xlsx")

# ------------------------------------------------------------------ inputs
def read_csv(p):
    with open(p, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))

A = {r["key"]: r["value"] for r in read_csv(os.path.join(ROOT, "inputs", "assumptions.csv"))}
H = {(r["item"], r["period"]): r for r in read_csv(os.path.join(ROOT, "inputs", "historicals.csv"))}
f = lambda k: float(A[k])
tri = lambda k: tuple(float(x) / 100 for x in A[k].split(","))   # "3,6,9" -> (0.03, 0.06, 0.09)
fs = lambda item, p="FY2026": float(H[(item, p)]["fs_value"])

price, shares, cash, debt = f("price"), f("shares_diluted_current"), f("cash"), f("debt")
tax, m_base, da_pct, capex_pct, nwc_pct = f("tax_rate"), f("ebit_margin"), f("da_pct_rev"), f("capex_pct_rev"), f("nwc_pct_rev")
g_path = np.array([f(f"rev_growth_fy{y}") for y in (27, 28, 29, 30, 31)])
tg_base = f("terminal_growth")
N, SEED = int(float(A["mc_runs"])), int(float(A["mc_seed"]))
G_LO, G_MODE, G_HI = tri("mc_growth_tri")
M_LO, M_MODE, M_HI = tri("mc_margin_tri")
W_MEAN, W_SD = (float(x) / 100 for x in A["mc_wacc_normal"].split(","))
TG_LO, TG_MODE, TG_HI = tri("mc_g_tri")
GUARD = f("mc_guardrail_wacc_minus_g_min")
parts = lambda k: ", ".join(x.strip() for x in A[k].split(","))     # "3,6,9" -> "3, 6, 9" (raw csv text for labels)
W_TXT = [x.strip() for x in A["mc_wacc_normal"].split(",")]
RANGE_LABEL = {"WACC": f"WACC N({W_TXT[0]}%, {W_TXT[1]}%)",
               "EBIT margin": f"EBIT margin tri({parts('mc_margin_tri')})%",
               "Revenue growth shift": f"Growth shift tri({parts('mc_growth_tri')})%",
               "Terminal growth g": f"Terminal g tri({parts('mc_g_tri')})%"}
wacc_capm = (f("rf") + f("beta") * f("erp"))   # for reference only; debt weight is 0

rev26 = fs("Net sales")
m26 = fs("Income from operations") / rev26        # FY26 actual EBIT margin, for the chart box
nwc26 = fs("Trade accounts receivable net") + fs("Inventories") - fs("Trade accounts payable")
T = np.arange(5) + 0.5                          # mid-year periods 0.5 .. 4.5
T_TV = T[-1]                                    # perpetuity TV discounted at t = 4.5

# ------------------------------------------------------------------ vectorized DCF
def value_per_share(g_shift, margin, wacc, tg):
    """Perpetuity-method value per share for arrays of drivers (shape (n,)). Same maths as DCF tab."""
    g_shift, margin, wacc, tg = (np.atleast_1d(np.asarray(x, dtype=float)) for x in (g_shift, margin, wacc, tg))
    growth = g_path[None, :] + g_shift[:, None]                 # level shift applied to every year
    rev = rev26 * np.cumprod(1 + growth, axis=1)               # (n, 5)
    ebit = margin[:, None] * rev
    nopat = ebit * (1 - tax)
    da, capex = da_pct * rev, capex_pct * rev
    nwc = nwc_pct * rev
    dnwc = np.diff(np.concatenate([np.full((rev.shape[0], 1), nwc26), nwc], axis=1), axis=1)
    fcff = nopat + da - capex - dnwc
    df = 1 / (1 + wacc[:, None]) ** T[None, :]
    pv = (fcff * df).sum(axis=1)
    tv = fcff[:, -1] * (1 + tg) / (wacc - tg)
    pv_tv = tv / (1 + wacc) ** T_TV
    equity = pv + pv_tv + cash - debt
    return equity * 1000 / shares

# ------------------------------------------------------------------ deterministic reconciliation
det = float(value_per_share(0.0, M_MODE, W_MEAN, TG_MODE)[0])
print("=" * 72)
print("DETERMINISTIC RUN (all drivers at mode: growth shift 0, margin mode, WACC mean, g mode)")
print("=" * 72)
print(f"  growth path   : {', '.join(f'{x:.1%}' for x in g_path)}   (shift 0)")
print(f"  EBIT margin   : {M_MODE:.2%}   WACC: {W_MEAN:.2%} (mc_wacc_normal mean; CAPM WACC {wacc_capm:.4%})   g: {TG_MODE:.2%}")
print(f"  base FY26 rev : {rev26:,.0f}   NWC (AR+Inv-AP): {nwc26:,.0f}   cash {cash:,.0f}   debt {debt:,.0f}   shares {shares:,.0f}")
print(f"  Python value  : ${det:.4f} / share")

excel_px = None
try:
    import win32com.client as w32, pythoncom
    pythoncom.CoInitialize()
    tmp = os.path.join(tempfile.mkdtemp(), "deck_dcf_read.xlsx")
    shutil.copy(XLSX, tmp)
    xl = w32.DispatchEx("Excel.Application"); xl.Visible = False; xl.DisplayAlerts = False
    wb = xl.Workbooks.Open(tmp); xl.CalculateFullRebuild()
    excel_px = float(wb.Names("price_perp").RefersToRange.Value)
    excel_im = float(wb.Names("implied_margin").RefersToRange.Value)   # market-implied EBIT margin at the current price (DCF tab)
    wb.Close(SaveChanges=False); xl.Quit()
except Exception as ex:
    print("  could not read price_perp from Excel via COM:", repr(ex))
if excel_px is None:
    sys.exit("STOP: Excel price_perp unavailable, cannot reconcile.")
diff = det - excel_px
print(f"  Excel price_perp (model/deck_dcf.xlsx via COM): ${excel_px:.4f}")
print(f"  difference    : ${diff:+.6f}   -> {'OK (< $0.005)' if abs(diff) < 0.005 else 'FAIL'}")
if abs(diff) >= 0.005:
    sys.exit("STOP: deterministic run does not reconcile to Excel within $1.")

# ------------------------------------------------------------------ simulation
rng = np.random.default_rng(SEED)
g_draw = rng.triangular(G_LO, G_MODE, G_HI, N)
m_draw = rng.triangular(M_LO, M_MODE, M_HI, N)
w_draw = rng.normal(W_MEAN, W_SD, N)
tg_draw = rng.triangular(TG_LO, TG_MODE, TG_HI, N)

# guardrail: WACC - g >= GUARD; redraw WACC and g for violating runs until none violate
viol = (w_draw - tg_draw) < GUARD
n_viol_initial = int(viol.sum()); n_redraws = 0; passes = 0
while viol.any():
    k = int(viol.sum()); n_redraws += k; passes += 1
    w_draw[viol] = rng.normal(W_MEAN, W_SD, k)
    tg_draw[viol] = rng.triangular(TG_LO, TG_MODE, TG_HI, k)
    viol = (w_draw - tg_draw) < GUARD
g_shift = g_draw - G_MODE
values = value_per_share(g_shift, m_draw, w_draw, tg_draw)

pct = {p: float(np.percentile(values, p)) for p in (5, 10, 25, 50, 75, 90, 95)}
S = {
    "runs": N, "seed": SEED,
    "mean": float(values.mean()), "std": float(values.std(ddof=1)),
    "P5": pct[5], "P10": pct[10], "P25": pct[25], "P50": pct[50], "P75": pct[75], "P90": pct[90], "P95": pct[95],
    "min": float(values.min()), "max": float(values.max()),
    "worst_corner_value": float(value_per_share(G_LO - G_MODE, M_LO, W_MEAN + 2 * W_SD, TG_LO)[0]),
    "prob_value_gt_price": float((values > price).mean()),
    "current_price": price,
    "guardrail_runs_violating_initially": n_viol_initial,
    "guardrail_total_redraws": n_redraws, "guardrail_redraw_passes": passes,
}
from scipy.optimize import brentq
corner_f = lambda w_: float(value_per_share(G_LO - G_MODE, M_LO, w_, TG_LO)[0]) - price   # corner: growth shift -3pp, margin 18%, g 1.5%
S["corner_wacc_at_price"] = float(brentq(corner_f, W_MEAN, W_MEAN + 0.10))
S["corner_wacc_sigma"] = (S["corner_wacc_at_price"] - W_MEAN) / W_SD
print("\n" + "=" * 72); print(f"SIMULATION ({N:,} runs, seed {SEED}, independent draws)"); print("=" * 72)
print(f"  growth tri({G_LO:.0%},{G_MODE:.0%},{G_HI:.0%}) as level shift | margin tri({M_LO:.1%},{M_MODE:.1%},{M_HI:.1%}) flat | "
      f"WACC N({W_MEAN:.2%},{W_SD:.2%}) | g tri({TG_LO:.1%},{TG_MODE:.1%},{TG_HI:.1%})")
print(f"  guardrail WACC - g >= {GUARD:.1%}: {n_viol_initial:,} runs violated initially; {n_redraws:,} total (WACC, g) redraws over {passes} pass(es)")
for k in ("mean", "P5", "P10", "P25", "P50", "P75", "P90", "P95", "std", "min", "max"):
    print(f"  {k:6} ${S[k]:>10.2f}")
print(f"  share of draws above ${price:.2f}: {S['prob_value_gt_price']:.2%}")
print(f"  worst corner of the ranges (growth {G_LO:.0%}, margin {M_LO:.0%}, g {TG_LO:.1%}, WACC mean + 2 sd {W_MEAN + 2 * W_SD:.2%}): ${S['worst_corner_value']:.2f}")
print(f"  market-implied EBIT margin at ${price:.2f} (DCF tab): {excel_im:.2%}")
print(f"  corner reprices to ${price:.2f} at WACC {S['corner_wacc_at_price']:.2%} = {S['corner_wacc_sigma']:.2f} sd above the mean")

# ------------------------------------------------------------------ tornado (Spearman)
drivers = {"Revenue growth shift": g_shift, "EBIT margin": m_draw, "WACC": w_draw, "Terminal growth g": tg_draw}
rho = {k: float(stats.spearmanr(v, values).correlation) for k, v in drivers.items()}
rho_sorted = sorted(rho.items(), key=lambda kv: -abs(kv[1]))
print("\nSpearman rank correlation vs value (sorted by |rho|):")
for k, v in rho_sorted:
    print(f"  {k:22} {v:+.4f}")

# ------------------------------------------------------------------ mc_summary.csv
os.makedirs(MC_DIR, exist_ok=True)
csv_path = os.path.join(MC_DIR, "mc_summary.csv")
with open(csv_path, "w", newline="", encoding="utf-8") as fh:
    w = csv.writer(fh, lineterminator="\n")
    w.writerow(["metric", "value", "note"])
    for k, v in S.items():
        note = "value per share ($)" if k in ("mean", "std", "P5", "P10", "P25", "P50", "P75", "P90", "P95", "min", "max") else ""
        if k == "worst_corner_value":
            note = "value per share ($) at the corner of the ranges: growth low, margin low, g low, WACC mean + 2 sd"
        if k == "corner_wacc_at_price":
            note = "WACC at which the corner (growth low, margin low, g low) reprices to Assumptions price (brentq)"
        if k == "corner_wacc_sigma":
            note = "(corner_wacc_at_price - mc_wacc_normal mean) / sd"
        w.writerow([k, v, note])
    for k, v in rho_sorted:
        w.writerow([f"spearman_{k}", round(v, 6), "Spearman rho, driver draw vs value"])
    w.writerow(["deterministic_python", det, "all drivers at mode"])
    w.writerow(["deterministic_excel_price_perp", excel_px, "model/deck_dcf.xlsx via Excel COM"])
    w.writerow(["deterministic_diff", diff, "python minus excel, must be < 1"])
    w.writerow(["wacc_capm_reference", wacc_capm, "rf + beta x erp (debt weight 0)"])
print("\nwrote", csv_path)

# ------------------------------------------------------------------ charts
BLUE, BLUE_LT, ORANGE = "#2a78d6", "#86b6ef", "#eb6834"
INK, INK2, MUTED, GRID, SURF = "#0b0b0b", "#52514e", "#898781", "#e1e0d9", "#fcfcfb"
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10, "axes.edgecolor": MUTED, "axes.labelcolor": INK2,
                     "xtick.color": INK2, "ytick.color": INK2, "axes.spines.top": False, "axes.spines.right": False})

# hist.png
fig, ax = plt.subplots(figsize=(6, 4.2), dpi=200, facecolor=SURF)
ax.tick_params(labelsize=11.5)
ax.set_facecolor(SURF)
ax.hist(values, bins=80, color=BLUE_LT, edgecolor=BLUE, linewidth=0.4)
ax.yaxis.grid(True, color=GRID, linewidth=0.8); ax.set_axisbelow(True)
ymax = ax.get_ylim()[1] * 1.18   # headroom so the reference-line labels sit above the bars
ax.set_ylim(0, ymax)
lines = [(pct[10], "P10", INK2, "--"), (pct[50], "P50", INK, "-"), (pct[90], "P90", INK2, "--"),
         (excel_px, "Base DCF", ORANGE, "-"), (price, "Current price", INK, ":")]
xr = ax.get_xlim()[1] - ax.get_xlim()[0]
placed = []   # (x, y_frac) of labels already drawn; step a label down when it would sit on a neighbour
for i_, (x, lab, col, ls) in enumerate(sorted(lines, key=lambda z: z[0])):
    ax.axvline(x, color=col, linestyle=ls, linewidth=1.6)
    yf = 0.985 if i_ % 2 == 0 else 0.855   # alternate rows left to right, then the loop below resolves any remaining overlap
    while any(abs(x - px_) < 0.13 * xr and abs(yf - py_) < 0.12 for px_, py_ in placed):
        yf -= 0.13
    placed.append((x, yf))
    ax.text(x, ymax * yf, f" {lab}\n ${x:,.2f}", color=col, fontsize=10, va="top", ha="left", parse_math=False,
            bbox=dict(boxstyle="square,pad=0.1", facecolor=SURF, edgecolor="none", alpha=0.85))
# market-implied margin sits under the current-price label, where the left of the chart is empty
ax.text(price, ymax * 0.80, f" implies {excel_im:.1%}\n EBIT margin vs\n {M_MODE:.1%} forecast,\n {m26:.1%} FY26", color=INK, fontsize=9,
        va="top", ha="left", parse_math=False, bbox=dict(boxstyle="square,pad=0.1", facecolor=SURF, edgecolor="none", alpha=0.85))
ax.set_title(f"DECK: Monte Carlo value per share ({N:,} runs)", color=INK, fontsize=14, loc="left", pad=10)
ax.set_xlabel("Value per share ($)", fontsize=12); ax.set_ylabel("Runs", fontsize=12)
ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f"${v:,.0f}"))
box = (       f"P10  ${pct[10]:,.2f}\nP50  ${pct[50]:,.2f}\nP90  ${pct[90]:,.2f}\n"
       f"mean ${S['mean']:,.2f}   std ${S['std']:,.2f}\n"
       f"worst run ${S['min']:,.2f}")
ax.text(0.985, 0.66, box, transform=ax.transAxes, ha="right", va="top", fontsize=9.5, color=INK, family="DejaVu Sans Mono",
        bbox=dict(boxstyle="round,pad=0.5", facecolor="white", edgecolor=GRID), parse_math=False)
fig.tight_layout()
hist_path = os.path.join(MC_DIR, "hist.png"); fig.savefig(hist_path, dpi=200, facecolor=SURF); plt.close(fig)

# tornado.png
labels = [RANGE_LABEL[k] for k, _ in rho_sorted][::-1]; vals = [v for _, v in rho_sorted][::-1]   # largest on top
fig, ax = plt.subplots(figsize=(10, 4.8), dpi=200, facecolor=SURF)
ax.set_facecolor(SURF)
cols = [BLUE if v >= 0 else ORANGE for v in vals]
bars = ax.barh(labels, vals, color=cols, height=0.55)
ax.axvline(0, color=MUTED, linewidth=1)
ax.xaxis.grid(True, color=GRID, linewidth=0.8); ax.set_axisbelow(True)
lim = max(abs(v) for v in vals) * 1.25
ax.set_xlim(-lim, lim)
for b, v in zip(bars, vals):
    ax.text(v + (0.02 if v >= 0 else -0.02) * lim, b.get_y() + b.get_height() / 2, f"{v:+.3f}",
            va="center", ha="left" if v >= 0 else "right", color=INK, fontsize=10)
ax.set_title(f"DECK: value per share sensitivity, Spearman rank correlation ({N:,} runs)", color=INK, fontsize=12.5, loc="left", pad=12)
ax.set_xlabel("Spearman rho (driver draw vs value per share)")
ax.tick_params(axis="y", length=0)
fig.tight_layout()
torn_path = os.path.join(MC_DIR, "tornado.png"); fig.savefig(torn_path, dpi=200, facecolor=SURF); plt.close(fig)
print("wrote", hist_path); print("wrote", torn_path)
