#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
fig6-5_dt_expansion.py -- the three figures and the printed numbers of Case Study II
(Section 6.3, digital-transformation barriers in SMCEs), adapted from the
published application chen2024socialwelfare.

All data are read from ../simulation_data/fig6-5_dt_case.xlsx, which is built
(generated AND solved) by build_workbooks.py: a 30 x 14 truncated-Weibull panel
under seed 202603 and the 70 solutions of the welfare-confidence program
eq:dt-welfare-conf (Sen-form complete welfare W_s = Ubar (1 - I_s) against the
aggregate variance, both min-max normalized, multi-start SLSQP with an edge
and face search, checked on a doubled discretization).  Nothing
numeric is hard-coded here; this script reads and plots only.  The values are
SYNTHETIC and reproduce the qualitative lesson of the source study, not its
empirical estimates.

Outputs (../figures_out/, then copy *.pdf into book/figures/):
  fig6-5_dt_ranking14.pdf   -- the 14 barriers ranked by the mean of the collective
                             distribution, averaged over the five equity indices
  fig6-6_dt_crossswf.pdf    -- per-objective collective means across the 14 barriers
  fig6-7_dt_triad.pdf       -- solved quantile-averaged collective densities of the
                             leading triad under the objective named in config

Run: python fig6-5_dt_expansion.py
"""
from _common import PALETTE, save, load, config
import numpy as np
import matplotlib.pyplot as plt

WB = "fig6-5_dt_case.xlsx"
S = load(WB)
C = config(S)
SEED = int(C["seed"])

# 14 barriers under the TOE framework, in workbook order
B = S["barriers"]
BARRIERS = list(zip(B["id"].astype(str), B["label"].astype(str), B["dimension"].astype(str)))

DIM = S["dimensions"]
DIMCOL = {c: PALETTE[k] for c, k in zip(DIM["code"].astype(str), DIM["color"].astype(str))}
# bars are filled with the pale companion of each dimension's tint and
# outlined in the tint, so the labels can sit on them in dark text
DIMFILL = {c: PALETTE[k + "_l"] for c, k in zip(DIM["code"].astype(str), DIM["color"].astype(str))}
DIMNAME = dict(zip(DIM["code"].astype(str), DIM["name"].astype(str)))
DIMORDER = list(DIM["code"].astype(str))

W = S["equity_objectives"]
EQ_IDS = list(W["eq"].astype(str))
EQ_NAMES = list(W["name"].astype(str))
EQ_COLS = [PALETTE[k] for k in W["color"].astype(str)]

# the solved 14 x 5 matrix of collective means (eq:dt-expectation), the
# five-way average (eq:dt-average) and the rank, as build_workbooks.py wrote them
MX = S["matrix"].set_index(S["matrix"]["id"].astype(str))
TRIAD = [t.strip() for t in str(C["triad"]).split(",")]
TRIAD_EQ = str(C["triad_objective"])


def eq_matrix():
    """Return (ids, 14 x 5 collective-mean matrix, avg, rank order) from the workbook."""
    ids = [b[0] for b in BARRIERS]
    M = np.array([[float(MX.loc[bid, sid]) for sid in EQ_IDS] for bid in ids])
    avg = np.array([float(MX.loc[bid, "avg"]) for bid in ids])
    order = sorted(range(len(ids)), key=lambda r: -avg[r])
    return ids, M, avg, order


def fig_ranking():
    ids, M, avg, order = eq_matrix()
    labels = {b[0]: (b[1], b[2]) for b in BARRIERS}
    ys = list(range(len(order)))[::-1]   # highest at top
    fig, ax = plt.subplots(figsize=(7.4, 5.4))
    for y, r in zip(ys, order):
        bid = ids[r]; short, dim = labels[bid]
        ax.barh(y, avg[r], color=DIMFILL[dim], edgecolor=DIMCOL[dim], lw=0.9,
                height=0.66, zorder=3)
        ax.text(avg[r] + 0.05, y, f"{avg[r]:.2f}", va="center", fontsize=8,
                color=PALETTE["text"], zorder=6,
                bbox=dict(facecolor="white", edgecolor="none", pad=0.8))
        ax.text(0.06, y, f"{bid}  {short}", va="center", fontsize=8.2,
                color=PALETTE["text"], zorder=4)
    for b in (C["band_lo"], C["band_hi"]):
        ax.axvline(b, color=PALETTE["faint"], lw=0.8, ls=":", zorder=1)
    # the bands carry the book's level symbols ell_4 = [3, 4) and ell_5 = [4, 5]
    ax.text((C["band_lo"] + C["band_hi"]) / 2, len(order) - 0.3, r"level $\ell_4$",
            ha="center", fontsize=7.5, color=PALETTE["slate"])
    ax.text((C["band_hi"] + C["scale_max"]) / 2, len(order) - 0.3, r"level $\ell_5$",
            ha="center", fontsize=7.5, color=PALETTE["slate"])
    ax.set_yticks([]); ax.set_xlim(C["scale_min"], C["scale_max"])
    ax.set_xlabel(C["x_label_ranking"])
    handles = [plt.Rectangle((0, 0), 1, 1, facecolor=DIMFILL[d], edgecolor=DIMCOL[d], lw=0.9)
               for d in DIMORDER]
    ax.legend(handles, [DIMNAME[d] for d in DIMORDER], loc="lower right", fontsize=8.5)
    fig.tight_layout(); save(fig, "fig6-5_dt_ranking14")


def fig_triad():
    """The solved quantile-averaged collective densities of the leading triad
    under the objective TRIAD_EQ, read from the triad_density sheet (computed
    in build_workbooks.py by the inverse-function rule from the solved
    aggregate quantile function)."""
    D = S["triad_density"]
    x = D["x"].to_numpy(dtype=float)
    fig, ax = plt.subplots(figsize=(7.0, 4.0))
    cols = [PALETTE["ink"], PALETTE["amber"], PALETTE["blue"]]
    lss = ["-", "--", "-."]
    for bid, col, ls in zip(TRIAD, cols, lss):
        y = D["f_" + bid].to_numpy(dtype=float)
        m = float(MX.loc[bid, TRIAD_EQ])
        ax.plot(x, y, color=col, lw=2.4, ls=ls,
                label=f"{bid}  (mean {m:.2f} under {TRIAD_EQ})", zorder=3)
        if col == PALETTE["ink"]:
            ax.fill_between(x, 0, y, color=PALETTE["gray_l"], alpha=0.45, zorder=2)
    for b in (C["band_lo"], C["band_hi"]):
        ax.axvline(b, color=PALETTE["faint"], lw=0.8, ls=":", zorder=1)
    ax.set_xlim(C["scale_min"], C["scale_max"]); ax.set_ylim(0, None)
    ax.set_xlabel(C["x_label_triad"]); ax.set_ylabel(C["y_label_triad"])
    ax.legend(loc="upper left", fontsize=9, title="leading triad")
    fig.tight_layout(); save(fig, "fig6-7_dt_triad")


def fig_crossswf():
    ids, M, avg, order = eq_matrix()
    xs = list(range(len(order)))
    fig, ax = plt.subplots(figsize=(7.6, 4.2))
    marks = ["o", "s", "^", "D", "v"]     # shapes double the tint coding
    for s in range(len(EQ_IDS)):
        ys = [M[r, s] for r in order]
        ax.plot(xs, ys, marker=marks[s % len(marks)], ms=3.4, lw=1.2, color=EQ_COLS[s],
                label=EQ_NAMES[s], zorder=3)
    ax.set_xticks(xs); ax.set_xticklabels([ids[r] for r in order], fontsize=7.5, rotation=0)
    ax.set_ylabel("mean of the collective distribution"); ax.set_xlabel("barrier (in final ranked order)")
    ax.legend(loc="upper right", fontsize=7.8, ncol=2)
    fig.tight_layout(); save(fig, "fig6-6_dt_crossswf")


if __name__ == "__main__":
    fig_ranking(); fig_triad(); fig_crossswf()
    ids, M, avg, order = eq_matrix()
    SOL = S["solves"]
    print("\n=== SOLVED NUMBERS OF CASE II (seed %d, data %s) ===" % (SEED, WB))
    print("tab:dt-swf-results (2 dp): EQ1..EQ5, average, rank")
    for rank, r in enumerate(order, 1):
        print(f"  {rank:2d}. {ids[r]:4s} " + " ".join(f"{M[r,s]:.2f}" for s in range(len(EQ_IDS)))
              + f"  avg {avg[r]:.2f}  spread {M[r].max()-M[r].min():.3f}")
    print("ranking:", ">".join(ids[r] for r in order))
    for _, row in S["per_objective_order"].iterrows():
        print("  %s: %s" % (row["eq"], row["order"]))
    print("leading triad %s, bottom three %s" % (TRIAD, [ids[r] for r in order[-3:]]))
    i14, i6 = ids.index("B14"), ids.index("B6")
    print("objectives under which B6 > B14:", [EQ_IDS[s] for s in range(5) if M[i6, s] > M[i14, s]])
    print("max |EQ2 - EQ5| over barriers: %.3f" % np.abs(M[:, 1] - M[:, 4]).max())
    print("active experts per solve: min %d, max %d; mean overlap at the optimum %.2f..%.2f "
          "(equal weights %.2f..%.2f)" % (SOL["n_active"].min(), SOL["n_active"].max(),
                                          SOL["Abar"].min(), SOL["Abar"].max(),
                                          SOL["Abar_equal_weights"].min(), SOL["Abar_equal_weights"].max()))
    print("Ubar at the optimum %.2f..%.2f; W~ at the optimum %.2f..%.3f; V~ %.2f..%.2f"
          % (SOL["Ubar"].min(), SOL["Ubar"].max(), SOL["W_tilde"].min(), SOL["W_tilde"].max(),
             SOL["V_tilde"].min(), SOL["V_tilde"].max()))
    EL = S["elicitation_b1"]
    print("tab:dt-elicit (3 dp): %s expert %d, profile %s -> grouped-MLE truncated Weibull "
          "kappa %.2f, lambda %.2f, mean %.2f"
          % (C["elicit_barrier"], int(C["elicit_expert"]),
             " ".join(f"{p:.3f}" for p in EL["probability"].to_numpy(dtype=float)),
             C["elicit_fit_kappa"], C["elicit_fit_lambda"], C["elicit_fit_mean"]))
    print("Figure 6.7 shows the triad under %s: means %s" % (
        TRIAD_EQ, ", ".join("%s %.2f" % (b, float(MX.loc[b, TRIAD_EQ])) for b in TRIAD)))
    print("adjacent pairs tied at two decimals (margin, and on the doubled discretization):")
    for _, row in S["near_ties"].iterrows():
        print("  %s: %s over %s by %.4f (%.4f)" % (row["eq"], row["upper"], row["lower"], row["margin"],
                                                  row["margin_fine"]))
    print("restart stability: interior starts (the barycentre and the Dirichlet(1) draws) reaching the "
          "retained value (of %d) median %g, min %d; "
          "mean range over reaching starts <= %.4f; largest mean deviation among near-optimal starts %.3f"
          % (int(SOL["n_interior_starts"].iloc[0]), SOL["n_interior_reach"].median(), SOL["n_interior_reach"].min(),
             SOL["mu_spread_reach"].max(), SOL["mu_dev_near"].max()))
    print("doubled-discretization check:")
    for _, row in S["fine_check"].iterrows():
        print("  %-4s %s (%s): mean %.4f -> %.4f, experts %s -> %s" % (
            row["barrier"], row["eq"], str(row["reason"]).strip(), row["mu"], row["mu_fine"],
            row["active_experts"], row["active_experts_fine"]))
    print("numbers printed in the text of Section 6.3 (sheet text_numbers):")
    for _, row in S["text_numbers"].iterrows():
        print("  %-34s %-14s %s" % (row["key"], row["printed"], row["where"]))
