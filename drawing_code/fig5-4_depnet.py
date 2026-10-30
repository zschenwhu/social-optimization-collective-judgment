#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
fig5-4_depnet.py -- the expert-dependence network of Example
ex:corr-lineage (Section 5.6).  A stipulated block of five weakly
correlated human experts and five fine-tuned variants of one base model
with within-lineage error correlation about 0.85.

All numbers are read from ../simulation_data/fig5-4_depnet.xlsx.  Node coordinates
are fixed in the workbook (deterministic layout).  Every pair is drawn as
an edge whose width (base + slope x r, in pt) and opacity (r clipped to
[floor, cap]) increase with the pairwise correlation r; the four encoding
constants are workbook config keys, so the weak edges stay faintly
visible and the lineage block reads as a dense clique.

Output: fig5-4_depnet.png/.pdf in ../figures_out/
"""
from _common import PALETTE, save, load, config
import numpy as np
import matplotlib.pyplot as plt

WB = "fig5-4_depnet.xlsx"
S = load(WB)
C = config(S)
ND = S["nodes"]
ids = list(ND["id"].astype(str))
grp = list(ND["group"].astype(str))
X = ND["x"].to_numpy(dtype=float)
Y = ND["y"].to_numpy(dtype=float)
R = S["correlation"][ids].to_numpy(dtype=float)
n = len(ids)

# edge encoding from the workbook: width and opacity increase with r
LW0, LW1 = float(C["edge_width_base_pt"]), float(C["edge_width_slope_pt"])
A_LO, A_HI = float(C["edge_alpha_floor"]), float(C["edge_alpha_cap"])

fig, ax = plt.subplots(figsize=(7.2, 4.2))
for i in range(n):
    for j in range(i + 1, n):
        r = R[i, j]
        ax.plot([X[i], X[j]], [Y[i], Y[j]], color=PALETTE["slate"],
                alpha=min(A_HI, max(A_LO, r)), lw=LW0 + LW1 * r, zorder=1,
                solid_capstyle="round")

# nodes are light tints with a hue outline and dark labels: sky circles for
# the humans, sage squares for the lineage
for i in range(n):
    if grp[i] == "human":
        ax.scatter(X[i], Y[i], s=420, marker="o", color=PALETTE["blue_l"],
                   edgecolor=PALETTE["blue"], linewidth=1.4, zorder=3)
    else:
        ax.scatter(X[i], Y[i], s=380, marker="s", color=PALETTE["teal_l"],
                   edgecolor=PALETTE["teal"], linewidth=1.4, zorder=3)
    ax.text(X[i], Y[i], ids[i], ha="center", va="center", fontsize=8.5,
            color=PALETTE["text"], zorder=4, fontweight="bold")

# annotations carry the workbook's effective-number accounting
neff_lin = float(C["neff_lineage_equalweight"])
r_lin = float(C["lineage_mean_corr"])
ax.text(1.75, 1.32, "one lineage, correlation about %.2f\n"
        "five variants count as $N_{\\mathrm{eff}}\\approx %.1f$ voices"
        % (r_lin, neff_lin),
        ha="center", fontsize=8.5, color=PALETTE["text"])
ax.text(-1.75, 1.62, str(C["human_note"]).replace("\\n", "\n"),
        ha="center", fontsize=8.5, color=PALETTE["text"])
ax.text(3.25, -1.55, str(C["edge_note"]), ha="right", fontsize=8,
        color=PALETTE["slate"])

# legend proxies (marker shape doubles the colour coding for grayscale print)
h1 = plt.Line2D([], [], ls="none", marker="o", ms=10, markerfacecolor=PALETTE["blue_l"],
                markeredgecolor=PALETTE["blue"], markeredgewidth=1.2,
                label=str(C["human_legend"]))
h2 = plt.Line2D([], [], ls="none", marker="s", ms=9, markerfacecolor=PALETTE["teal_l"],
                markeredgecolor=PALETTE["teal"], markeredgewidth=1.2,
                label=str(C["lineage_legend"]))
ax.legend(handles=[h1, h2], loc="lower left", fontsize=8.5)
ax.set_xlim(-3.3, 3.3)
ax.set_ylim(-1.75, 1.95)
ax.set_aspect("equal")
ax.axis("off")
fig.tight_layout()
save(fig, "fig5-4_depnet")

if __name__ == "__main__":
    print("\n=== NUMBERS READ FROM THE WORKBOOK (synthetic, seed %d, data %s) ===" % (C["seed"], WB))
    lin = [i for i in range(n) if grp[i] == "lineage"]
    hum = [i for i in range(n) if grp[i] == "human"]
    pl = [R[i, j] for i in lin for j in lin if i < j]
    ph = [R[i, j] for i in hum for j in hum if i < j]
    print("  lineage block mean correlation %.3f (range %.3f to %.3f)"
          % (np.mean(pl), np.min(pl), np.max(pl)))
    print("  human block mean correlation  %.3f" % np.mean(ph))
    print("  N_eff lineage (equal weights) %.2f" % neff_lin)
    print("  N_eff panel (equal weights)   %.2f" % float(C["neff_panel_equalweight"]))
