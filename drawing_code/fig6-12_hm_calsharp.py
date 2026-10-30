#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
fig6-12_hm_calsharp.py -- the diagnostic figure for the Section 6.6 exploratory
vignette on human-machine forecast aggregation. It places each forecaster of the
seed-202603 record on the reliability-resolution plane: humans and calibrated
machines cluster at low reliability error inside the well-calibrated band, while
the two overconfident (narrow, biased) machines stand far above it, M4 also at
the highest resolution -- the central lesson that machine confidence must not be
mistaken for reliability. (Part of the plotted resolution of M3 and M4 is an
outer-bin artifact of their narrow spreads and bias, chiefly the spread for M3,
and it shrinks on finer representations; the recalibrated resolution the text
reports is the 'resolution_recalibrated' column of the same sheet.)

All data are read from ../simulation_data/fig6-12_hm_calsharp.xlsx; nothing numeric
is hard-coded here. The workbook is GENERATED AND SOLVED by the vignette block of
build_workbooks.py (target series, eight forecasters, convex CRPS program,
leave-one-out re-solves, Hersbach split, design level, re-seeding study); the
plotted columns 'resolution' (RES_i = UNC - POT_i) and 'reliability_error'
(REL_i) of the 'forecasters' sheet are solved outputs, and the axis limits, the
band, the axis labels and the callout text and position in the config sheet are
derived from them. The values are ILLUSTRATIVE SYNTHETIC ones; the vignette
reports no empirical magnitudes.

Note on terminology: the two channels of the CRPS decomposition are called
CALIBRATION and RESOLUTION throughout the book. The file name is legacy.

Output: ../figures_out/fig6-12_hm_calsharp.pdf  (copy into book/figures/)
Run: python fig6-12_hm_calsharp.py
"""
from _common import PALETTE, save, load, config
import numpy as np
import matplotlib.pyplot as plt

WB = "fig6-12_hm_calsharp.xlsx"
S = load(WB)
C = config(S)
SEED = int(C["seed"])

F = S["forecasters"]
# preserve workbook order of the groups
GROUP_ORDER = list(dict.fromkeys(F["group"].astype(str)))


def fig():
    fig, ax = plt.subplots(figsize=(7.0, 4.4))
    for name in GROUP_ORDER:
        sub = F[F["group"].astype(str) == name]
        col = PALETTE[str(sub["color"].iloc[0])]
        mk = str(sub["marker"].iloc[0])
        ax.scatter(list(sub["resolution"].astype(float)),
                   list(sub["reliability_error"].astype(float)),
                   c=col, marker=mk, s=70, edgecolor="white", lw=0.8,
                   label=name, zorder=3)
        # label each point by its id so the prose can refer to it; a label whose
        # default place (above right) would sit on a close neighbor up and to the
        # right goes below right instead
        for _, r in sub.iterrows():
            x0, y0 = float(r["resolution"]), float(r["reliability_error"])
            crowded = any(0 < xn - x0 < 0.02 and 0 < yn - y0 < 0.03
                          for xn, yn in zip(F["resolution"].astype(float),
                                            F["reliability_error"].astype(float)))
            ax.annotate(str(r["id"]), (x0, y0),
                        xytext=(5, -9) if crowded else (5, 3), textcoords="offset points",
                        fontsize=7.5, color=PALETTE["text"], zorder=4)
    # a guide: "well calibrated" region near the x-axis
    ax.axhspan(0, C["well_calibrated_cut"], color=PALETTE["band"], zorder=0)
    ax.text(C["band_label_x"], C["band_label_y"], C["band_label"],
            fontsize=8, color=PALETTE["text"])
    # the callout text sits where the config places it, with one arrow to each
    # overconfident machine (the two need not lie close together)
    # (every arrow is annotated with the same text so that it starts at the
    # text's edge; the text itself is drawn once)
    over = F[F["group"].astype(str) == "overconfident machines"]
    for k, (_, r) in enumerate(over.iterrows()):
        ax.annotate(C["callout_text"].replace("\\n", "\n"),
                    xy=(float(r["resolution"]), float(r["reliability_error"])),
                    xytext=(C["callout_text_x"], C["callout_text_y"]),
                    fontsize=8.5, color=PALETTE["text"] if k == 0 else "none",
                    ha="center", va="center",
                    arrowprops=dict(arrowstyle="->", color=PALETTE["red"], lw=1.0, shrinkB=6))
    ax.set_xlabel(C["x_label"])
    ax.set_ylabel(C["y_label"])
    ax.set_xlim(C["x_min"], C["x_max"]); ax.set_ylim(C["y_min"], C["y_max"])
    ax.legend(loc="upper left", fontsize=9)
    fig.tight_layout(); save(fig, "fig6-12_hm_calsharp")


if __name__ == "__main__":
    fig()
    n_h = int(C["n_human"]); n_m = int(C["n_machine"])
    print("saved fig6-12_hm_calsharp (seed %d, data %s); illustrative synthetic, "
          "no empirical magnitudes" % (SEED, WB))
    print("panel: %d human + %d machine forecasters over T=%d periods"
          % (n_h, n_m, int(C["horizon_T"])))
    for name in GROUP_ORDER:
        sub = F[F["group"].astype(str) == name]
        print("   %-24s n=%d  resolution %s  reliability error %s"
              % (name, len(sub),
                 [round(float(v), 3) for v in sub["resolution"]],
                 [round(float(v), 3) for v in sub["reliability_error"]]))
