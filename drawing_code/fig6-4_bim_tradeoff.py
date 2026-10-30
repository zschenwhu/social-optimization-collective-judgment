#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Figure 6.4 (fig:bim-tradeoff): fairness-consensus trade-off as the fairness
weight delta is swept over the grid {0, 0.1, ..., 0.6}. Data source:
../simulation_data/fig6-4_bim_tradeoff.xlsx (the sweep SOLVED by build_workbooks.py)"""
from _common import load, config, save, PALETTE, plt

s = load("fig6-4_bim_tradeoff.xlsx")
c = config(s)
d = s["series"]

fig, ax = plt.subplots(figsize=(6.8, 4.2))
ax.grid(axis="y", color="#DDDDDD", alpha=0.5, lw=0.6, zorder=0)
ax.plot(d["delta"], d["consensus_Abar"], color=PALETTE["blue"], marker="o",
        lw=2.2, label=r"consensus $\bar{A}$",
        markeredgecolor="white", markeredgewidth=0.6)
ax.plot(d["delta"], d["fairness_Phibar"], color=PALETTE["amber"], marker="s",
        lw=2.2, ls="--", label=r"fairness $\bar{\Phi}$",
        markeredgecolor="white", markeredgewidth=0.6)
ax.plot(d["delta"], d["minimum_overlap"], color=PALETTE["teal"], marker="^",
        lw=2.2, ls="-.", label="minimum overlap",
        markeredgecolor="white", markeredgewidth=0.6)

ax.set_xlim(0, 0.6); ax.set_ylim(c["ymin"], c["ymax"])
ax.set_xlabel(r"fairness weight $\delta$"); ax.set_ylabel(c["y_label"])
ax.legend(loc="center left", fontsize=9)
fig.tight_layout()
save(fig, "fig6-4_bim_tradeoff")
