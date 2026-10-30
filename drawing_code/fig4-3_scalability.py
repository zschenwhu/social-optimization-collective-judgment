#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
fig4-3_scalability.py -- computational-cost curves of the flat pairwise
program versus the representative-level hierarchy (Section 4.2).

All numbers are read from ../simulation_data/fig4-3_scalability.xlsx, where
build_workbooks.py computes them from the stated interaction-count
formulas (no timings anywhere).  The vertical axis is the pairwise
interaction count, in normalized cost units.

Output: fig4-3_scalability.png/.pdf in ../figures_out/
"""
from _common import PALETTE, save, load, config
import matplotlib.pyplot as plt

WB = "fig4-3_scalability.xlsx"
S = load(WB)
C = config(S)
D = S["series"]

N = D["N"].to_numpy(dtype=float)
flat = D["flat_pairwise"].to_numpy(dtype=float)
hier = D["hier_pairwise"].to_numpy(dtype=float)
kstar = D["K_star"].to_numpy(dtype=int)

fig, ax = plt.subplots(figsize=(7.0, 4.4))
ax.loglog(N, flat, color=PALETTE["slate"], lw=2.2, ls="-", marker="o", ms=4,
          markevery=3, label=str(C["flat_legend"]), zorder=3)
ax.loglog(N, hier, color=PALETTE["ink"], lw=2.2, ls="--", marker="s", ms=4,
          markevery=3, label=str(C["hier_legend"]), zorder=3)
ax.fill_between(N, hier, flat, color=PALETTE["gray_l"], alpha=0.55, zorder=1)

# annotate the saving at the largest panel, computed from the data; both
# annotations live in the empty region below the two-stage curve, with
# leaders up to their data points, so neither text crosses a series
i = len(N) - 1
factor = flat[i] / hier[i]
# the factor written on the figure is the one the text prints (summary sheet)
SUMM = dict(zip(S["summary"]["key"].astype(str), S["summary"]["value"]))
assert int(SUMM["N_max"]) == int(N[i]) and abs(float(SUMM["ratio_flat_over_hier_at_N_max"]) - factor) < 1e-9
assert "%.0f" % factor == "%d" % int(SUMM["ratio_printed"])
ax.annotate("saving factor %.0f$\\times$ at $N$ = %d" % (factor, int(N[i])),
            xy=(N[i], hier[i] * 0.82), xytext=(N[i] * 0.085, hier[i] * 0.00040),
            fontsize=8.5, color=PALETTE["text"], va="center",
            arrowprops=dict(arrowstyle="->", color=PALETTE["ink"], lw=0.9,
                            shrinkA=4, shrinkB=3))
# annotate K* at a mid grid point, also from the data
j = len(N) // 2
ax.annotate("$K^{*}$ = %d subgroups at $N$ = %d" % (kstar[j], int(N[j])),
            xy=(N[j], hier[j] * 0.80), xytext=(N[j] * 0.50, hier[j] * 0.10),
            fontsize=8.5, color=PALETTE["slate"], va="center",
            arrowprops=dict(arrowstyle="->", color=PALETTE["faint"], lw=0.9,
                            shrinkA=4, shrinkB=3))

ax.set_xlabel(C["x_label"])
ax.set_ylabel(C["y_label"])
ax.legend(loc="upper left", fontsize=9)
fig.tight_layout()
save(fig, "fig4-3_scalability")

if __name__ == "__main__":
    print("\n=== NUMBERS READ FROM THE WORKBOOK (data %s) ===" % WB)
    for r in range(0, len(N), 9):
        print("  N=%6d  flat=%12d  K*=%3d  two-stage=%10d  factor=%.2f"
              % (N[r], flat[r], kstar[r], hier[r], flat[r] / hier[r]))
    print("  N=%6d  flat=%12d  K*=%3d  two-stage=%10d  factor=%.2f"
          % (N[i], flat[i], kstar[i], hier[i], factor))
