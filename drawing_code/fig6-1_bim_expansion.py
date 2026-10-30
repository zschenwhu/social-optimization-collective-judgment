#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
fig6-1_bim_expansion.py -- two figures for Case Study I (Section 6.2, BIM maturity
assessment), plus the synthetic numbers quoted in the prose and tables.

All data are read from ../simulation_data/fig6-1_bim_case.xlsx, which is the
authoritative source for this case; nothing numeric is hard-coded here. The
builder GENERATES the twenty-expert panel under seed 202603, recovers its three
groups by k-means, weights them by TOPSIS (Omega), solves the intra-group
consensus-confidence program for every group and indicator, and aggregates the
indicator collectives to the dimensions and the global maturity under best-worst
weights; the 'groups' sheet holds the solved q4 group opinions with Omega and the
'dimensions' sheet the solved dimension distributions with their best-worst
weights. The q4 collective and the global maturity are recomputed here from those
rows by quantile averaging (the builder records the same values in the 'derived'
sheet). The values are ILLUSTRATIVE SYNTHETIC ones that reproduce the
*qualitative* lesson of the published application (organization the strongest
dimension, technology the weakest, global maturity near the Managed--Defined
boundary), NOT its empirical estimates.

Outputs (../figures_out/, then copy *.pdf into book/figures/):
  fig6-1_bim_groupagg.pdf   -- between-group aggregation, Real-time Data indicator
  fig6-2_bim_dimensions.pdf -- three dimension PDFs + global maturity PDF

Run: python fig6-1_bim_expansion.py
"""
from _common import PALETTE, gauss, save, load, config
import numpy as np
import matplotlib.pyplot as plt

WB = "fig6-1_bim_case.xlsx"
S = load(WB)
C = config(S)
SEED = int(C["seed"])   # the seed under which the builder generated and solved the case


def qa_normal(mus, sigmas, w):
    """Quantile-averaged aggregate of normals is normal:
    mean = sum w_k mu_k, sd = sum w_k sigma_k (Proposition 2.1(iii))."""
    w = np.asarray(w, float); w = w / w.sum()
    return float(np.dot(w, mus)), float(np.dot(w, sigmas))


# ---------------------------------------------------------------- group layer --
# The three expert groups recovered by k-means, their solved group-opinion normals
# for the Real-time Data indicator (q4, a Technology indicator), and the TOPSIS
# credibility weights Omega.
G = S["groups"]
GROUPS = list(G["legend"].astype(str))
OMEGA = list(G["Omega"].astype(float))
q4_mu = list(G["mu"].astype(float))
q4_sd = list(G["sigma"].astype(float))
q4_agg_mu, q4_agg_sd = qa_normal(q4_mu, q4_sd, OMEGA)

# ------------------------------------------------------------ dimension layer --
# Dimension-level aggregates (BWM-weighted QA of the solved indicator collectives
# in each dimension) and the BWM dimension weights; the global maturity PDF is
# their QA aggregate.
D = S["dimensions"]
DIMS = list(D["dimension"].astype(str))
dim_mu = list(D["mu"].astype(float))
dim_sd = list(D["sigma"].astype(float))
dim_w = list(D["weight"].astype(float))
dim_col = [PALETTE[c] for c in D["color"].astype(str)]
glob_mu, glob_sd = qa_normal(dim_mu, dim_sd, dim_w)

# ------------------------- one generated expert's elicited q4 probabilities
E = S["elicitation_q4"]
LEVEL_MID = E["level_mid"].to_numpy(dtype=float)
p_q4 = E["probability"].to_numpy(dtype=float)
p_q4 = p_q4 / p_q4.sum()
fit_mu = float(np.dot(p_q4, LEVEL_MID))
fit_sd = float(np.sqrt(np.dot(p_q4, (LEVEL_MID - fit_mu) ** 2)))


def fig_groupagg():
    x = np.linspace(C["scale_min"], C["scale_max"], int(C["grid_points"]))
    fig, ax = plt.subplots(figsize=(7.2, 4.2))
    cols = [PALETTE["blue"], PALETTE["amber"], PALETTE["teal"]]
    for mu, sd, c, name, om in zip(q4_mu, q4_sd, cols, GROUPS, OMEGA):
        ax.plot(x, gauss(x, mu, sd), color=c, lw=1.8, ls="--",
                label=f"{name}, $\\Omega={om:.2f}$, $\\mathcal{{N}}({mu:.2f},{sd:.2f}^2)$", zorder=2)
    yv = gauss(x, q4_agg_mu, q4_agg_sd)
    # the collective is the one dark object; the groups are light tints
    ax.fill_between(x, 0, yv, color=PALETTE["gray_l"], alpha=0.55, zorder=1)
    ax.plot(x, yv, color=PALETTE["ink"], lw=2.6, ls="-",
            label=f"inter-group QA aggregate\n$\\mathcal{{N}}({q4_agg_mu:.2f},\\,{q4_agg_sd:.2f}^2)$",
            zorder=3)
    ax.set_xlabel(C["x_label_groupagg"])
    ax.set_ylabel(C["y_label"])
    ax.set_xlim(C["scale_min"], C["scale_max"]); ax.set_ylim(0, None)
    ax.legend(loc="upper right", fontsize=8.5)
    fig.tight_layout(); save(fig, "fig6-1_bim_groupagg")


def fig_dimensions():
    x = np.linspace(C["scale_min"], C["scale_max"], int(C["grid_points"]))
    fig, ax = plt.subplots(figsize=(7.2, 4.2))
    for mu, sd, c, name, w in zip(dim_mu, dim_sd, dim_col, DIMS, dim_w):
        ax.plot(x, gauss(x, mu, sd), color=c, lw=1.9, ls="--",
                label=f"{name} ($w={w:.2f}$), $\\mathcal{{N}}({mu:.2f},{sd:.2f}^2)$",
                zorder=2)
    yv = gauss(x, glob_mu, glob_sd)
    ax.fill_between(x, 0, yv, color=PALETTE["gray_l"], alpha=0.55, zorder=1)
    ax.plot(x, yv, color=PALETTE["ink"], lw=2.7, ls="-",
            label=f"global maturity\n$\\mathcal{{N}}({glob_mu:.2f},\\,{glob_sd:.2f}^2)$",
            zorder=3)
    for b in (C["band_lo"], C["band_hi"]):   # Managed | Defined | Standard borders
        ax.axvline(b, color=PALETTE["faint"], lw=0.8, ls=":", zorder=0)
    # extra headroom so the band name sits clear of the tallest curve peak
    ax.set_xlim(C["scale_min"], C["scale_max"]); ax.set_ylim(0, None)
    ymax = ax.get_ylim()[1] * 1.07
    ax.set_ylim(0, ymax)
    ax.text((C["band_lo"] + C["band_hi"]) / 2, ymax * 0.975, C["band_name"],
            ha="center", va="top", fontsize=8, color=PALETTE["slate"])
    ax.set_xlabel(C["x_label_dimensions"])
    ax.set_ylabel(C["y_label"])
    ax.legend(loc="upper left", fontsize=8.3)
    fig.tight_layout(); save(fig, "fig6-2_bim_dimensions")


if __name__ == "__main__":
    fig_groupagg()
    fig_dimensions()
    print("\n=== NUMBERS READ FROM THE WORKBOOK (synthetic, seed %d, data %s) ===" % (SEED, WB))
    print("q4 group opinions:", list(zip(GROUPS, q4_mu, q4_sd)), "Omega=", OMEGA)
    print("q4 inter-group QA aggregate: N(%.3f, %.3f^2)" % (q4_agg_mu, q4_agg_sd))
    print("dimension aggregates:")
    for n, mu, sd, w in zip(DIMS, dim_mu, dim_sd, dim_w):
        print("   %-13s w=%.2f  N(%.2f, %.2f^2)" % (n, w, mu, sd))
    print("global maturity: N(%.3f, %.3f^2)  (managed L5 / defined L6 boundary)"
          % (glob_mu, glob_sd))
    print("q4 elicited prob vector -> fitted normal: N(%.3f, %.3f^2)" % (fit_mu, fit_sd))
    print("q4 prob vector:", [round(float(v), 3) for v in p_q4])
