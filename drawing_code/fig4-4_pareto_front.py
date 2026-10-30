#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Figure 4.4 (fig:pareto): the consensus-fairness Pareto front of the
seven-expert Case I fairness panel, its concave majorant, the unsupported
stretch below it, and the weighted-sum solutions.

Data source: ../simulation_data/fig4-4_pareto_front.xlsx, written by the
FIGURE 4.4 block of build_workbooks.py, which solves the front by the
epsilon-constraint method (max fairness subject to a floor on the mean
overlap, multi-start SLSQP on a grid of floors refined around the deepest
point of the dent), solves the weighted sums, and computes the concave
majorant of all those efficient points and the unsupported stretch below it.
This script only reads and draws.  It re-derives from the stored points, and
asserts, the properties the caption and Section 4.3 state:

  * the front is nonincreasing in fairness as the consensus floor rises;
  * the stored majorant is the upper hull of the stored efficient points
    (epsilon points and weighted sums together), every weighted sum lies on
    it, and the marked point is the deepest point of the stretch below it;
  * the stretch ends are the weighted sums on either side of lambda*, and no
    weighted-sum solution lies inside the stretch;
  * the number drawn on the figure is the one chapter4.tex prints.

The main panel draws the front against its majorant.  Because the dent is
shallow on the scale of the whole front, the inset plots the stored vertical
gap between the majorant and the efficient points on its own scale, so the
extent and the depth of the unsupported stretch can be read directly.
"""
from _common import load, config, save, PALETTE, plt
import numpy as np
from matplotlib.text import Text
from matplotlib.ticker import MaxNLocator, FuncFormatter

s = load("fig4-4_pareto_front.xlsx")
c = config(s)
f = s["front"]
hull = s["hull"]
ws = s["weighted_sum"]
summ = dict(zip(s["summary"]["key"].astype(str), s["summary"]["value"]))
printed = dict(zip(s["printed"]["key"].astype(str), s["printed"]["value"]))

A = f["Abar"].to_numpy(float)
P = f["Phi"].to_numpy(float)
gap = f["gap_below_majorant"].to_numpy(float)
main = f["in_main_stretch"].to_numpy(int) == 1
k = int(summ["max_gap_index"])
wsA = ws["Abar"].to_numpy(float)
wsP = ws["Phi"].to_numpy(float)
wsgap = ws["gap_below_majorant"].to_numpy(float)
lam = ws["lambda_consensus"].to_numpy(float)
hx, hy = hull["Abar"].to_numpy(float), hull["Phi"].to_numpy(float)
a_lo, a_hi = float(summ["unsupported_Abar_lo"]), float(summ["unsupported_Abar_hi"])
p_lo, p_hi = float(summ["unsupported_Phi_at_lo"]), float(summ["unsupported_Phi_at_hi"])
v_lo, v_hi = int(summ["unsupported_lo_hull_vertex"]), int(summ["unsupported_hi_hull_vertex"])
lamstar = float(summ["lambda_star"])

# ------------------------------------------------------------------ checks --
assert np.all(np.diff(A) >= -1e-12) and np.all(np.diff(P) <= 0.0), "the stored front is not a nonincreasing front"
assert np.all(np.diff(hx) > 0), "hull abscissae not ascending"


def _upper_hull(x, y):
    h = []
    for i in range(len(x)):
        while len(h) >= 2:
            i1, i2 = h[-2], h[-1]
            if (x[i2] - x[i1]) * (y[i] - y[i1]) - (y[i2] - y[i1]) * (x[i] - x[i1]) >= 0:
                h.pop()
            else:
                break
        h.append(i)
    return h


# the stored hull majorizes every stored efficient point and is itself concave
assert list(_upper_hull(hx, hy)) == list(range(len(hx))), "the stored hull is not concave"
assert np.allclose(np.interp(A, hx, hy) - P, gap, atol=1e-12), "stored front gaps disagree with the stored hull"
assert np.allclose(np.interp(wsA, hx, hy) - wsP, wsgap, atol=1e-12), "stored weighted-sum gaps disagree with the hull"
assert gap.min() >= -1e-12 and wsgap.min() >= -1e-12, "a stored point lies above the majorant"
assert wsgap.max() < 1e-9, "a weighted-sum solution lies below the majorant"
for src, sheet in (("epsilon", f), ("weighted sum", ws)):
    rows = hull.loc[hull["source"] == src, "row"].astype(int)
    assert np.allclose(hull.loc[hull["source"] == src, "Abar"].to_numpy(float),
                       sheet["Abar"].to_numpy(float)[rows], atol=0), "hull vertex does not point at its row"
assert v_hi == v_lo + 1 and abs(hx[v_lo] - a_lo) < 1e-15 and abs(hx[v_hi] - a_hi) < 1e-15, \
    "stretch ends are not consecutive hull vertices"
assert list(np.where(main)[0]) == list(np.where((A > a_lo + 1e-9) & (A < a_hi - 1e-9))[0])
assert k == int(np.argmax(gap)) and main[k], "marked point is not the deepest unsupported point"
assert abs(gap[k] - float(summ["max_gap"])) < 1e-12
inside = (wsA > a_lo + 1e-7) & (wsA < a_hi - 1e-7)
assert not inside.any(), "a weighted sum lands in the stretch"
jl = int(np.max(np.where(wsA <= a_lo + 1e-7)[0]))
jr = jl + 1
assert lam[jl] < lamstar < lam[jr] and wsA[jr] >= a_hi - 1e-7
assert max(abs(wsA[jl] - a_lo), abs(wsP[jl] - p_lo), abs(wsA[jr] - a_hi), abs(wsP[jr] - p_hi)) < 1e-7, \
    "the stretch ends are not the weighted sums on either side of lambda*"
# the one number drawn on the figure is the one the text prints
gap_label = "%.4f" % gap[k]
assert gap_label == "%.4f" % float(printed["max_gap_4dp"]), \
    "inset label %s differs from the printed %s" % (gap_label, printed["max_gap_4dp"])

# ------------------------------------------------------------------ drawing --
fig, ax = plt.subplots(figsize=(6.6, 4.6))
# the stretch: its two ends (weighted-sum solutions on the hull) and the epsilon points between
sA = np.r_[a_lo, A[main], a_hi]
sP = np.r_[p_lo, P[main], p_hi]
ax.fill_between(sA, sP, np.interp(sA, hx, hy), color=PALETTE["blue"], alpha=0.35, lw=0, zorder=1)
h_arc, = ax.plot(sA, sP, color=PALETTE["blue"], lw=4.2, alpha=0.9,
                 solid_capstyle="round", label=str(c["arc_legend"]), zorder=2)
h_front, = ax.plot(A, P, color=PALETTE["ink"], lw=1.8, label=str(c["front_legend"]), zorder=3)
h_hull, = ax.plot([a_lo, a_hi], [p_lo, p_hi], color=PALETTE["slate"], lw=1.5, ls="--",
                  label=str(c["hull_legend"]), zorder=4)
h_ws, = ax.plot(wsA, wsP, "o", ms=5.2, mfc="white", mec=PALETTE["ink"], mew=1.0,
                label=str(c["ws_legend"]), zorder=5)
ax.plot([A[k]], [P[k]], "o", color=PALETTE["red"], ms=6, mec="white", mew=0.6, zorder=6)
t_mark = ax.annotate(str(c["marked_label"]), (A[k], P[k]),
                     xytext=(-28, -85), textcoords="offset points", fontsize=9.5, color=PALETTE["text"],
                     ha="left", va="center", zorder=6,
                     arrowprops=dict(arrowstyle="->", color=PALETTE["red"], lw=0.9, shrinkA=2, shrinkB=4))

padx = 0.04 * (A.max() - A.min())
pady = 0.06 * (P.max() - P.min())
xlim = (A.min() - padx, A.max() + padx)
ax.set_xlim(*xlim)
ax.set_ylim(P.min() - pady, P.max() + pady)
ax.set_xlabel(c["x_label"]); ax.set_ylabel(c["y_label"])
ax.legend(handles=[h_front, h_arc, h_hull, h_ws], loc="upper right", fontsize=9.5)

# inset: the gap between the majorant and the efficient points, on its own
# scale and in the units of the fairness axis
uA = np.r_[A, wsA]
ugap = np.r_[gap, wsgap]
o = np.argsort(uA, kind="stable")
uA, ugap = uA[o], ugap[o]
axi = ax.inset_axes([0.10, 0.10, 0.36, 0.27])
axi.fill_between(sA, 0.0, np.r_[0.0, gap[main], 0.0], color=PALETTE["blue"], alpha=0.35, lw=0)
axi.plot(uA, ugap, color=PALETTE["ink"], lw=1.1)
axi.plot([A[k]], [gap[k]], "o", color=PALETTE["red"], ms=4.5, mec="white", mew=0.5, zorder=4)
t_val = axi.annotate(gap_label, (A[k], gap[k]), xytext=(0, 6), textcoords="offset points",
                     fontsize=9, color=PALETTE["text"], ha="center", va="bottom")
axi.set_xlim(*xlim)
axi.set_ylim(0.0, 1.40 * gap[k])
axi.xaxis.set_major_locator(MaxNLocator(nbins=3))
axi.yaxis.set_major_locator(MaxNLocator(nbins=3, steps=[1, 2, 5]))
axi.yaxis.set_major_formatter(FuncFormatter(lambda v, _: "0" if abs(v) < 1e-15 else "%.3f" % v))
axi.tick_params(labelsize=8.5, length=2.5, pad=1.5)
axi.set_title(str(c["inset_title"]), fontsize=9.5, pad=3)
for sp in ("top", "right"):
    axi.spines[sp].set_visible(False)
fig.tight_layout()

# rendered checks: the marker label keeps >= 6 pt from the front, the chord and
# every weighted-sum marker, and neither it, the inset nor the legend overlap
fig.canvas.draw()
pt = fig.dpi / 72.0


def _clearance(bb, px, py, n=4000):
    u = np.linspace(0, 1, n)
    t = np.linspace(0, 1, len(px))
    Q = ax.transData.transform(np.column_stack([np.interp(u, t, px), np.interp(u, t, py)]))
    dx = np.maximum(np.maximum(bb.x0 - Q[:, 0], Q[:, 0] - bb.x1), 0.0)
    dy = np.maximum(np.maximum(bb.y0 - Q[:, 1], Q[:, 1] - bb.y1), 0.0)
    return float(np.min(np.hypot(dx, dy))) / pt


bb_mark = Text.get_window_extent(t_mark)
bb_leg = ax.get_legend().get_window_extent()
bb_in = axi.get_tightbbox(fig.canvas.get_renderer())
for name, (px, py) in (("front", (A, P)),
                       ("chord", (np.array([a_lo, a_hi]), np.array([p_lo, p_hi]))),
                       ("weighted-sum markers", (wsA, wsP))):
    d = _clearance(bb_mark, px, py)
    assert d >= 6.0, "marker label only %.1f pt from the %s" % (d, name)
    d = _clearance(bb_in, px, py)
    assert d >= 6.0, "inset only %.1f pt from the %s" % (d, name)
u = np.linspace(0, 1, 4000)
Qi = axi.transData.transform(np.column_stack([np.interp(u, np.linspace(0, 1, len(uA)), uA),
                                              np.interp(u, np.linspace(0, 1, len(uA)), ugap)]))
bb_val = Text.get_window_extent(t_val)
dxi = np.maximum(np.maximum(bb_val.x0 - Qi[:, 0], Qi[:, 0] - bb_val.x1), 0.0)
dyi = np.maximum(np.maximum(bb_val.y0 - Qi[:, 1], Qi[:, 1] - bb_val.y1), 0.0)
assert float(np.min(np.hypot(dxi, dyi))) / pt >= 2.0, "inset value label touches the gap curve"
assert bb_val.y1 <= axi.get_window_extent().y1 + 0.5 * pt, "inset value label leaves the inset"
assert not bb_mark.overlaps(bb_leg), "marker label overlaps the legend"
assert not bb_mark.overlaps(bb_in), "marker label overlaps the inset"
assert not bb_in.overlaps(bb_leg), "inset overlaps the legend"

save(fig, "fig4-4_pareto_front")

if __name__ == "__main__":
    print("front: %d epsilon points (%d grid floors, %d refinement floors), g1 in [%.4f, %.4f], g2 in [%.4f, %.4f]"
          % (len(A), int(summ["grid_floors"]), int(summ["refinement_floors"]), A.min(), A.max(), P.min(), P.max()))
    print("unsupported stretch: g1 in (%.6f, %.6f), deepest %.6f below the majorant at g1 = %.6f"
          % (a_lo, a_hi, gap[k], A[k]))
    print("weighted sums: %d weightings, none inside the stretch; lambda = %.2f returns its lower end and"
          " lambda = %.2f its upper end; lambda* = %.4f" % (len(ws), lam[jl], lam[jr], lamstar))
