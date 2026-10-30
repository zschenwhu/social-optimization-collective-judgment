#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
fig5-3_loo.py -- per-expert leave-one-out contribution chart for the
synthetic human-machine panel of Example ex:hmpanel (Section 5.5).

All numbers are read from ../simulation_data/fig5-3_loo.xlsx, where
build_workbooks.py generates the panel, solves the convex program by
Algorithm 4.2 on the full panel, on the human subpanel, and on every
leave-one-out subpanel, and splits each contribution into its
calibration and resolution channels under the Hersbach convention
recorded in the workbook config.  The marked total is the step-score
total (the sum of the two channels); the workbook's results sheet
carries the solved values behind Table tab:hmpanel.

Output: fig5-3_loo.png/.pdf in ../figures_out/
"""
from _common import PALETTE, save, load, config
import numpy as np
import matplotlib.pyplot as plt

WB = "fig5-3_loo.xlsx"
S = load(WB)
C = config(S)
L = S["loo"]

ids = list(L["id"].astype(str))
groups = list(L["group"].astype(str))
scale = 1000.0                        # display styling only, axis says so
d_cal = L["delta_cal"].to_numpy(dtype=float) * scale
d_res = L["delta_res"].to_numpy(dtype=float) * scale
d_tot = L["delta_total"].to_numpy(dtype=float) * scale

# consistency gate: the plotted total is the sum of the two channels (exact
# under the one fixed binning convention of the workbook)
assert np.allclose(d_cal + d_res, d_tot, atol=1e-9)

# tick labels stay in the neutral text gray; the dotted rule separates the
# human block from the machine block and the annotation names the
# overconfident cell, so no colour is needed on the axis
GCOL = {"human": PALETTE["text"], "machine calibrated": PALETTE["text"],
        "machine overconfident": PALETTE["text"]}

n = len(ids)
ys = np.arange(n)[::-1]               # H1 at the top
fig, ax = plt.subplots(figsize=(7.2, 4.8))
off = 0.19
ax.barh(ys + off, d_cal, height=0.34, color=PALETTE["slate"], zorder=3,
        label=str(C["cal_legend"]))
ax.barh(ys - off, d_res, height=0.34, color=PALETTE["amber_l"], hatch="///",
        edgecolor=PALETTE["amber"], lw=0.4, zorder=3, label=str(C["res_legend"]))
ax.plot(d_tot, ys, ls="none", marker="D", ms=5.5, color=PALETTE["ink"],
        zorder=4, label=str(C["tot_legend"]))
ax.axvline(0.0, color=PALETTE["slate"], lw=1.0, zorder=2)

# separate the human block from the machine block
ax.axhline(ys[4] - 0.5, color=PALETTE["faint"], lw=0.8, ls=":", zorder=1)
ax.set_yticks(ys)
ax.set_yticklabels(ids, fontsize=9.5)
for tick, gname in zip(ax.get_yticklabels(), groups):
    tick.set_color(GCOL[gname])
ax.set_xlabel(str(C["x_label"]) + "  (units of $10^{-3}$)")

# annotate the overconfident cell, from the data (last two experts)
over = [k for k, gname in enumerate(groups) if gname == "machine overconfident"]
ax.annotate("beneficial resolution\nbut miscalibrated\n"
            "($\\Delta^{\\mathrm{cal}}_i<0$, $\\Delta^{\\mathrm{res}}_i>0$)",
            xy=(d_res[over[-1]], ys[over[-1]] - off),
            xytext=(d_res[over[-1]] + 1.4, ys[over[-1]] + 1.1),
            fontsize=8, color=PALETTE["text"],
            arrowprops=dict(arrowstyle="->", color=PALETTE["red"], lw=0.9))
ax.legend(loc="lower right", fontsize=8.5)
fig.tight_layout()
save(fig, "fig5-3_loo")

if __name__ == "__main__":
    R = S["results"]
    print("\n=== NUMBERS READ FROM THE WORKBOOK (synthetic, seed %d, data %s) ===" % (C["seed"], WB))
    for _, r in R.iterrows():
        print("  %-46s %10.4f  (printed %s)" % (r["quantity"], r["value"], r["printed"]))
    for k in range(n):
        print("  %-3s %-22s  cal %+7.4f  res %+7.4f  tot %+7.4f  w %.3f"
              % (ids[k], groups[k], d_cal[k] / scale, d_res[k] / scale,
                 d_tot[k] / scale, L["weight_full"][k]))
