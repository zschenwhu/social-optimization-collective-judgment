#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
fig6-8_bdt_expansion.py -- the three figures of Case Study III (Section 6.4,
building digital twin maturity via FA-LSCOG), adapted from the authorized source
paper chen2024bdt.

All data are read from ../simulation_data/fig6-8_bdt_case.xlsx, which is the
authoritative source for this case and is GENERATED AND SOLVED by
build_workbooks.py under seed 202603 (a 15-expert x 22-indicator truncated-Weibull
panel, k-means + TOPSIS inter-group weights, best-worst dimension weights, and the
intra-group fairness-confidence program with the weight-Gini regularizer solved by
multi-start SLSQP for every subgroup, indicator and fairness weight). Nothing
numeric is hard-coded here. The values are ILLUSTRATIVE SYNTHETIC ones that
reproduce the *qualitative* lessons of the source study, NOT its empirical
estimates.

Outputs (../figures_out/, then copy *.pdf into book/figures/):
  fig6-8_bdt_q1agg.pdf        -- inter-group aggregation of indicator q1 (Response and Control):
                                 the three group representatives and the collective drawn as
                                 quantile-average densities 1/(dQ/dp) from the solved weights
  fig6-9_bdt_profile.pdf      -- the seven-dimension maturity profile + the global maturity
  fig6-10_bdt_sensitivity.pdf -- sensitivity of the aggregation to the fairness weight

Run: python fig6-8_bdt_expansion.py
"""
from _common import PALETTE, save, load, config
import numpy as np
import matplotlib.pyplot as plt

WB = "fig6-8_bdt_case.xlsx"
S = load(WB)
C = config(S)
SEED = int(C["seed"])
_trapz = getattr(np, "trapezoid", np.trapz)

# The between-group reliability weights are written Omega throughout the book
# (canonical notation: between-group weights Omega, within-group w^{(k)}).
WEIGHT_SYMBOL = r"\Omega"

# --- expert groups (k-means strata) and inter-group (TOPSIS) weights -----------
G = S["groups"]
GROUPS = list(G["group"].astype(str))
GSIZE = list(G["size"].astype(int))
OMEGA = list(G["Omega"].astype(float))
GCOL = [PALETTE[c] for c in G["color"].astype(str)]

# --- indicator q1 (Response and Control): solved group representatives ----------
# q1_mu is the mean of each group's quantile-averaged representative at the
# operating point (the mean of its stored quantile function); the collective
# mean is the Omega-weighted mean of the group means, an identity of quantile
# averaging.
q1_group_mu = list(G["q1_mu"].astype(float))
q1_agg = float(np.dot(OMEGA, q1_group_mu))
DENS = S["q1_densities"]
X = DENS["x"].to_numpy(float)
DX = float(X[1] - X[0])
q1_group_f = [DENS["f_G%d" % (k + 1)].to_numpy(float) for k in range(len(GROUPS))]
q1_agg_f = DENS["f_collective"].to_numpy(float)
# the legend means must be the means of the curves drawn
for k in range(len(GROUPS)):
    assert abs(_trapz(X * q1_group_f[k], dx=DX) - q1_group_mu[k]) < 0.01, \
        "legend mean of %s differs from the drawn density" % GROUPS[k]
assert abs(_trapz(X * q1_agg_f, dx=DX) - q1_agg) < 0.01, \
    "legend mean of the collective differs from the drawn density"

# --- seven dimensions: solved mean maturity and BWM dimension weights -----------
D = S["dimensions"]
DIMS = list(D["dimension"].astype(str))
DKIND = list(D["kind"].astype(str))
dim_mu = list(D["mu"].astype(float))
dim_w = list(D["weight"].astype(float))
glob_mu = float(np.dot(dim_w, dim_mu))

# --- fairness-weight sensitivity sweep (solved) ----------------------------------
SW = S["sensitivity"]
DELTA = list(SW["delta"].astype(float))
glob_by_delta = list(SW["global_maturity"].astype(float))
fairness_util = list(SW["fairness_Phi"].astype(float))        # collective fairness Phi
variance_conf = list(SW["aggregate_variance"].astype(float))  # aggregate variance of F_G
weight_gini = list(SW["weight_gini"].astype(float))           # influence concentration

# the global maturity at the operating point is, by the quantile-averaging
# mean identity, the dimension-weighted mean of the dimension means
op = float(C["operating_delta"])
assert abs(glob_by_delta[DELTA.index(op)] - glob_mu) < 1e-6, (
    "sweep at delta=%.1f gives %.4f but the dimension weights give %.4f"
    % (op, glob_by_delta[DELTA.index(op)], glob_mu))

LEVEL_NAMES = list(S["levels"]["name"].astype(str))


def fig_q1agg():
    fig, ax = plt.subplots(figsize=(7.0, 4.0))
    for f, mu, c, g, om in zip(q1_group_f, q1_group_mu, GCOL, GROUPS, OMEGA):
        ax.plot(X, f, color=c, lw=1.9, ls="--",
                label=f"{g} (${WEIGHT_SYMBOL}={om:.2f}$), mean {mu:.2f}", zorder=2)
    # the collective is the one dark object; the groups are light tints
    ax.fill_between(X, 0, q1_agg_f, color=PALETTE["gray_l"], alpha=0.55, zorder=1)
    ax.plot(X, q1_agg_f, color=PALETTE["ink"], lw=2.6,
            label=f"inter-group collective (mean {q1_agg:.2f})", zorder=3)
    for b in (C["band_lo"], C["band_hi"]):
        ax.axvline(b, color=PALETTE["faint"], lw=0.8, ls=":", zorder=0)
    ax.set_xlim(C["scale_min"], C["scale_max"]); ax.set_ylim(0, None)
    ax.set_xlabel(C["x_label_q1"])
    ax.set_ylabel(C["y_label_q1"])
    ax.legend(loc="upper left", fontsize=8.5)
    fig.tight_layout(); save(fig, "fig6-8_bdt_q1agg")


def fig_profile():
    fig, ax = plt.subplots(figsize=(7.4, 4.4))
    ys = list(range(len(DIMS)))[::-1]
    # bars: pale fill of each perspective's tint, outlined in the tint, dark labels
    kcol = {"tech": PALETTE["blue"], "system": PALETTE["purple"], "org": PALETTE["green"]}
    kfill = {"tech": PALETTE["blue_l"], "system": PALETTE["purple_l"], "org": PALETTE["green_l"]}
    for y, name, mu, kind in zip(ys, DIMS, dim_mu, DKIND):
        ax.barh(y, mu, color=kfill[kind], edgecolor=kcol[kind], lw=0.9, height=0.62, zorder=3)
        # white halo so the value stays legible where the global-maturity
        # reference line passes right behind it
        ax.text(mu + 0.06, y, f"{mu:.2f}", va="center", fontsize=8.5,
                color=PALETTE["text"], zorder=6,
                bbox=dict(facecolor="white", edgecolor="none", pad=1.0))
        ax.text(0.05, y, name, va="center", fontsize=8.5, color=PALETTE["text"], zorder=4)
    # a dedicated headroom band for the maturity-level names: the reference
    # line and the level gridlines stop below it, so no label is grazed
    y_top = len(DIMS) + 0.55
    y_line_top = len(DIMS) - 0.05
    frac = (y_line_top + 0.6) / (y_top + 0.6)
    ax.axvline(glob_mu, color=PALETTE["ink"], lw=1.6, ls="-", zorder=5,
               ymax=frac, label=f"global maturity {glob_mu:.2f}")
    for b in range(1, int(C["n_levels"])):
        ax.axvline(b, color=PALETTE["faint"], lw=0.8, ls=":", zorder=1, ymax=frac)
    for j, nm in enumerate(LEVEL_NAMES):
        ax.text(j + 0.5, y_top - 0.08, nm, ha="center", va="top", fontsize=6.8,
                color=PALETTE["slate"])
    ax.set_yticks([]); ax.set_xlim(C["scale_min"], C["scale_max"]); ax.set_ylim(-0.6, y_top)
    ax.set_xlabel(C["x_label_profile"])
    ax.legend(loc="lower right", fontsize=8.5)
    fig.tight_layout(); save(fig, "fig6-9_bdt_profile")


def fig_sensitivity():
    off = float(C["profile_offset"])
    fig, ax = plt.subplots(figsize=(7.2, 4.2))
    ax.plot(DELTA, fairness_util, marker="o", color=PALETTE["green"], lw=1.8,
            label="collective fairness $\\Phi$")
    ax.plot(DELTA, variance_conf, marker="s", color=PALETTE["amber"], lw=1.8,
            label="aggregate variance $V$")
    ax.plot(DELTA, weight_gini, marker="^", color=PALETTE["red"], lw=1.8,
            label="weight concentration (Gini)")
    ax.plot(DELTA, [g - off for g in glob_by_delta], marker="d", color=PALETTE["blue"], lw=1.8,
            label="global maturity $-\\,%g$" % off)
    ax.set_xlabel(C["x_label_sensitivity"])
    ax.set_ylabel(C["y_label_sensitivity"])
    ax.set_xlim(-0.02, 0.52)
    # headroom above the highest series (Phi) so the two-column legend clears
    # every curve instead of sitting on the maturity series
    lo = min(min(variance_conf), min(weight_gini), min(g - off for g in glob_by_delta))
    ax.set_ylim(min(0.0, lo - 0.05), max(fairness_util) + 0.25)
    ax.legend(loc="upper left", fontsize=8.5, ncol=2)
    fig.tight_layout(); save(fig, "fig6-10_bdt_sensitivity")


if __name__ == "__main__":
    fig_q1agg(); fig_profile(); fig_sensitivity()
    print("\n=== SOLVED SYNTHETIC NUMBERS (seed %d, data %s) ===" % (SEED, WB))
    print("groups: sizes", GSIZE, "N=", sum(GSIZE), " inter-group weights Omega=",
          [round(o, 3) for o in OMEGA])
    print("q1 group reps", [round(m, 3) for m in q1_group_mu],
          "-> inter-group collective mean %.3f" % q1_agg)
    print("dimension means/weights (delta = %.1f):" % op)
    for n, mu, w in zip(DIMS, dim_mu, dim_w):
        print("   %-30s mean %.3f  w %.3f" % (n, mu, w))
    print("criterion weights sum to %.3f; dimension-weighted global maturity = %.4f"
          % (sum(dim_w), glob_mu))
    print("global maturity mean = %.3f  (%s)" % (glob_mu, C["global_level"]))
    for d, g, f, v, gi in zip(DELTA, glob_by_delta, fairness_util, variance_conf, weight_gini):
        print("   delta=%.1f  maturity %.3f  Phi %.3f  V %.3f  Gini(w) %.3f" % (d, g, f, v, gi))
