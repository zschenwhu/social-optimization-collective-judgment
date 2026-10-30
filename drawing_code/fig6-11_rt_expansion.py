#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
fig6-11_rt_expansion.py -- the figure and the solved numbers of Case Study IV
(urban rail-transit station accessibility with stakeholder equity),
adapted from the authorized source paper chen2024railtransit at the
level of MECHANISM only.

All data are read from ../simulation_data/fig6-11_rt_case.xlsx, the
authoritative workbook of this case, where build_workbooks.py generates the
24-stakeholder panel under seed 202603 and solves, per station, the
consensus-only program and the Sen-welfare program over the simplex by
multi-start SQP (the confidence objective is constant on this panel and is
dropped), then stores instrument, panel parameters, solver certificates,
robustness checks, and results side by side.  The values are ILLUSTRATIVE
SYNTHETIC ones that exhibit the source study's MECHANISM (equity weighting
reorders priorities), NOT its empirical findings.

Outputs (../figures_out/, then copy *.pdf into book/figures/):
  fig6-11_rt_equity.pdf -- per-station collective accessibility under the
                        two regimes, the priority swap visible

Run: python fig6-11_rt_expansion.py
"""
from _common import PALETTE, save, load, config
import numpy as np
import matplotlib.pyplot as plt

WB = "fig6-11_rt_case.xlsx"
S = load(WB)
C = config(S)
SEED = int(C["seed"])

R = S["results"]
ids = list(R["station"].astype(str))
names = list(R["name"].astype(str))
m_eff = R["mean_eff"].to_numpy(dtype=float)
m_eq = R["mean_eq"].to_numpy(dtype=float)
p_eff = R["priority_eff"].to_numpy(dtype=int)
p_eq = R["priority_eq"].to_numpy(dtype=int)
swapped = [t.strip() for t in str(C["swapped"]).split(",")]

# the recorded swap must match what the stored means give
assert [ids[j] for j in range(len(ids)) if p_eff[j] != p_eq[j]] == swapped


def fig_equity():
    fig, ax = plt.subplots(figsize=(7.2, 4.9))
    x0, x1 = 0.0, 1.0
    for j in range(len(ids)):
        hot = ids[j] in swapped
        col = PALETTE["amber"] if hot else PALETTE["faint"]
        lw = 2.6 if hot else 1.2
        ax.plot([x0, x1], [m_eff[j], m_eq[j]], color=col, lw=lw, zorder=3 if hot else 2)
        mcol = PALETTE["red"] if hot else PALETTE["slate"]
        ax.plot(x0, m_eff[j], marker="o", ms=6 if hot else 4.5, color=mcol, zorder=4)
        ax.plot(x1, m_eq[j], marker="s", ms=6 if hot else 4.5, color=mcol, zorder=4)
        ax.text(x0 - 0.04, m_eff[j], "%s  %.2f" % (ids[j], m_eff[j]),
                ha="right", va="center", fontsize=8,
                color=PALETTE["text"] if hot else PALETTE["slate"],
                fontweight="bold" if hot else "normal")
        ax.text(x1 + 0.04, m_eq[j], "%.2f  %s" % (m_eq[j], ids[j]),
                ha="left", va="center", fontsize=8,
                color=PALETTE["text"] if hot else PALETTE["slate"],
                fontweight="bold" if hot else "normal")
    # annotate the swap with the priority ranks, all from the data
    j4 = ids.index(swapped[0])
    j5 = ids.index(swapped[1])
    ax.text(0.5, max(m_eff[j4], m_eq[j4]) + 0.30,
            "the equity-aware regime moves %s ahead of %s\n"
            "in upgrading priority (rank %d to rank %d)"
            % (swapped[0], swapped[1], p_eff[j4], p_eq[j4]),
            ha="center", fontsize=8.3, color=PALETTE["text"])
    ax.set_xlim(-0.62, 1.62)
    ax.set_xticks([x0, x1])
    ax.set_xticklabels([str(C["eff_tick"]), str(C["eq_tick"])], fontsize=10)
    ax.set_xlabel(C["x_label"])
    ax.set_ylabel(C["y_label"])
    fig.tight_layout()
    save(fig, "fig6-11_rt_equity")


if __name__ == "__main__":
    fig_equity()
    print("\n=== NUMBERS READ FROM THE WORKBOOK (synthetic, seed %d, data %s) ===" % (SEED, WB))
    print("Per-station collective accessibility and upgrading priority")
    for j in range(len(ids)):
        print("  %-4s %-22s eff %.3f (rank %2d)   eq %.3f (rank %2d)%s"
              % (ids[j], names[j], m_eff[j], p_eff[j], m_eq[j], p_eq[j],
                 "   <- swap" if ids[j] in swapped else ""))
    print("swapped pair", swapped)
    print("average weight Gini  eff %.3f  vs  eq %.3f  (shared nearest-to-uniform tie-break)"
          % (float(C["weight_gini_eff_avg"]), float(C["weight_gini_eq_avg"])))
    wg_eff = R["weight_gini_eff"].to_numpy(dtype=float)
    wg_eq = R["weight_gini_eq"].to_numpy(dtype=float)
    a_eff = R["adv_share_eff"].to_numpy(dtype=float)
    a_eq = R["adv_share_eq"].to_numpy(dtype=float)
    j = ids.index(str(C["contested_station"]))
    print("contested station %s  weight gini %.2f -> %.2f  reported advocate share %.2f -> %.2f"
          % (ids[j], wg_eff[j], wg_eq[j], a_eff[j], a_eq[j]))
    print("level-set range of the advocate share (what every optimum shares; the reported"
          " shares are tie-break representatives inside it)")
    for k in range(len(ids)):
        print("  %-4s cons [%.3f, %.3f]   equity [%.3f, %.3f]"
              % (ids[k], R["adv_share_min_eff"].iloc[k], R["adv_share_max_eff"].iloc[k],
                 R["adv_share_min_eq"].iloc[k], R["adv_share_max_eq"].iloc[k]))
    print("near-ties: %s" % C["near_ties"])
