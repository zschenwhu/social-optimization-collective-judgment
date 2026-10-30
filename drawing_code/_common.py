#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
_common.py -- shared style and helpers for the figure-reproduction scripts.

Every figure script imports this, reads its workbook from ../simulation_data/,
and writes PNG + PDF to ../figures_out/ (and an editable SVG to ../svg/). The
plots use the book-wide palette, shared with the TikZ conceptual diagrams so
figures and diagrams read as one visual language, and set everything -- text
and mathtext alike -- in Times New Roman. Restyle freely here without touching
any data.

Palette (since 2026-09-14): gray and white carry the structure. Every
aggregate, collective, reference line, axis, and piece of text is a gray
("ink" for the primary object, "slate" for reference lines, "faint" for
background curves). Colour is used only where categories must be told apart,
and then only as light tints: sky, rose, sage, sand, and lilac for up to five
categories, with a deeper dusty rose ("wine") reserved for warnings and
negative outcomes. Each hue has a pale fill companion ("<key>_l").

The palette dict keeps its historical role keys ("blue", "amber", ...) so the
figure scripts and the colour names stored in the case workbooks keep working
("blue" -> sky, "amber" -> rose, "teal" -> sage, "green" -> sand,
"purple" -> lilac, "red" -> wine, "slate" -> mid gray). Descriptive aliases
are provided for new code.
"""
import os
import sys as _env_sys

HERE = os.path.dirname(os.path.abspath(__file__))

# The printed figures were rendered on the software stack named in ../environment.py.  With
# another matplotlib the same data are drawn, but text is placed about a point differently,
# so the stack is printed once and a notice is given when it is not the reference one.  A
# package that is not installed is named in a plain message before anything is imported.
_env_sys.path.insert(0, os.path.normpath(os.path.join(HERE, "..")))
try:
    import environment as _environment
except ImportError:
    _environment = None
if _environment is not None:
    _environment.require_packages()
    _environment.report()

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")            # headless: write files, no display window
import matplotlib.pyplot as plt

DATA = os.path.normpath(os.path.join(HERE, "..", "simulation_data"))
OUT = os.path.normpath(os.path.join(HERE, "..", "figures_out"))
SVG = os.path.normpath(os.path.join(HERE, "..", "svg"))
os.makedirs(OUT, exist_ok=True)
os.makedirs(SVG, exist_ok=True)

# ------------------------------------------------------------------ palette ----
# One palette for the whole book, pinned hexes, matching the TikZ conceptual
# diagrams. Grays first; light tints only where categories must be separated.
PALETTE = {
    # grays: structure, aggregates, reference, text
    "ink":    "#3A3A3A",   # primary object: the aggregate, the collective, the front
    "slate":  "#8A8A8A",   # neutral / reference line / equality / totals
    "faint":  "#C2C2C2",   # background curves, separators, faint edges
    "gray_l": "#E4E4E4",   # gray fill under the primary object
    "band":   "#F2F2F2",   # background bands
    # light tints: categorical distinction only
    "blue":   "#8FA9C4",   # sky:   first category / human experts
    "amber":  "#CB9AA3",   # rose:  second category / emphasis
    "teal":   "#94B79E",   # sage:  third category / machine experts
    "green":  "#C4B47E",   # sand:  fourth category
    "purple": "#B09CC2",   # lilac: fifth category
    "red":    "#A86A77",   # wine:  negative / warning / unreachable
    # pale fill companions
    "blue_l":   "#DFE7EF",
    "amber_l":  "#F1E2E5",
    "teal_l":   "#E1ECE4",
    "green_l":  "#EEE9D4",
    "purple_l": "#E8E1EE",
    "red_l":    "#EAD5D9",
    "sand":     "#EEE9D4",  # historical name of the band fill
}
# descriptive aliases (same objects, for new code)
PALETTE.update({
    "sky": PALETTE["blue"], "rose": PALETTE["amber"], "sage": PALETTE["teal"],
    "lilac": PALETTE["purple"], "wine": PALETTE["red"],
    "sky_l": PALETTE["blue_l"], "rose_l": PALETTE["amber_l"],
    "sage_l": PALETTE["teal_l"], "sand_l": PALETTE["green_l"],
    "lilac_l": PALETTE["purple_l"], "wine_l": PALETTE["red_l"],
    "text": "#333333",
})
# ordered categorical cycle: sky, rose, sage, sand, lilac, mid gray
CYCLE = [PALETTE["blue"], PALETTE["amber"], PALETTE["teal"],
         PALETTE["green"], PALETTE["purple"], PALETTE["slate"]]

# Back-compatibility: older scripts import GRAYS.
GRAYS = {"black": PALETTE["ink"], "dark": PALETTE["slate"], "mid": PALETTE["faint"],
         "light": PALETTE["gray_l"], "faint": "#D5D8DC"}
LINESTYLES = {"solid": "-", "dashed": "--", "dotted": ":", "dashdot": "-."}

plt.rcParams.update({
    # everything in Times New Roman, mathtext included
    "font.family": "Times New Roman",
    "mathtext.fontset": "custom",
    "mathtext.rm": "Times New Roman",
    "mathtext.it": "Times New Roman:italic",
    "mathtext.bf": "Times New Roman:bold",
    "mathtext.cal": "Times New Roman:italic",
    "mathtext.default": "it",
    "font.size": 11,
    "axes.linewidth": 0.9,
    "axes.edgecolor": "#666666",
    "axes.labelcolor": "#333333",
    "axes.titlecolor": "#333333",
    "text.color": "#333333",
    "xtick.color": "#666666",
    "ytick.color": "#666666",
    "axes.spines.top": False,
    "axes.spines.right": False,
    "legend.frameon": False,
    "figure.dpi": 120,
    "savefig.dpi": 300,
    "savefig.bbox": "tight",
    "savefig.facecolor": "white",
    "figure.facecolor": "white",
    # SVG exports draw glyphs as paths, so the files render and edit
    # identically on machines without the originating fonts
    "svg.fonttype": "path",
})


def load(name):
    """Return dict of {sheet_name: DataFrame} for ../simulation_data/<name>.xlsx."""
    path = os.path.join(DATA, name)
    return pd.read_excel(path, sheet_name=None)


def config(sheets):
    """Turn the 'config' sheet (key/value rows) into a plain dict."""
    c = sheets["config"]
    d = dict(zip(c["key"].astype(str), c["value"]))
    for k, v in list(d.items()):
        try:
            fv = float(v)
            d[k] = int(fv) if fv.is_integer() else fv
        except (TypeError, ValueError):
            pass
    return d


def gauss(x, mu, sigma):
    return np.exp(-(x - mu) ** 2 / (2.0 * sigma ** 2)) / (sigma * np.sqrt(2.0 * np.pi))


def save(fig, name):
    """name without extension; writes PNG and PDF to ../figures_out/ and an
    editable SVG (text drawn as paths, resolution-independent) to ../svg/.
    The SVG is a byproduct for hand editing; the shipped artifact remains the
    PDF, and adding the SVG leaves the PDF bytes untouched."""
    for ext in ("png", "pdf"):
        fig.savefig(os.path.join(OUT, name + "." + ext))
    fig.savefig(os.path.join(SVG, name + ".svg"))
    plt.close(fig)
    print("saved", name + ".png/.pdf (+ svg)")


def density_figure(workbook, outname):
    """Generic plotter for the Gaussian-density figures (2.1, 2.2, 6.3).

    Experts are faint dotted background curves; an optional highlighted expert
    and an optional linear pool are rose; the first (primary) single-Gaussian
    aggregate is drawn in ink with a light gray fill, and any further
    aggregates in the light category tints.
    """
    s = load(workbook)
    c = config(s)
    g = s["gaussians"]
    x = np.linspace(c["domain_min"], c["domain_max"], 600)
    fig, ax = plt.subplots(figsize=(7.2, 4.2))

    experts = g[g["role"].isin(["expert", "expert_highlight"])]
    aggs = g[g["role"] == "aggregate"]

    shown_expert_legend = False
    for _, r in experts.iterrows():
        y = gauss(x, r["mu"], r["sigma"])
        lab = str(r["legend"])
        if r["role"] == "expert_highlight":
            ax.plot(x, y, color=PALETTE["amber"], lw=1.9, ls="-",
                    label=(lab if lab != "_nolegend_" else "_nolegend_"), zorder=3)
        else:
            # faint dotted background expert; surface one generic legend entry
            if lab != "_nolegend_":
                showlab = lab if not shown_expert_legend else "_nolegend_"
                shown_expert_legend = True
            elif not shown_expert_legend:
                showlab = "individual experts"
                shown_expert_legend = True
            else:
                showlab = "_nolegend_"
            ax.plot(x, y, color=PALETTE["faint"], lw=1.0, ls=":", alpha=0.7,
                    label=showlab, zorder=1)

    # optional linear pool = normalized mixture of the expert components
    if int(c.get("draw_linear_pool", 0)) == 1 and len(experts) > 0:
        w = experts["weight"].to_numpy(dtype=float)
        w = w / w.sum()
        ypool = sum(wi * gauss(x, r["mu"], r["sigma"])
                    for wi, (_, r) in zip(w, experts.iterrows()))
        ax.plot(x, ypool, color=PALETTE["amber"], lw=2.3, ls="-",
                label="linear pool", zorder=4)

    # single-Gaussian aggregates: the primary one in ink, the rest in tints
    agg_colors = [PALETTE["ink"], PALETTE["teal"], PALETTE["purple"], PALETTE["green"]]
    for i, (_, r) in enumerate(aggs.iterrows()):
        ls = LINESTYLES.get(str(r.get("linestyle", "")), "-")
        col = agg_colors[i % len(agg_colors)]
        yv = gauss(x, r["mu"], r["sigma"])
        if i == 0:
            ax.fill_between(x, 0, yv, color=PALETTE["gray_l"], alpha=0.55, zorder=2)
        ax.plot(x, yv, color=col, lw=2.4, ls=ls, label=str(r["legend"]), zorder=5)

    if "annotations" in s:
        for _, a in s["annotations"].iterrows():
            ax.annotate(str(a["text"]), (a["x"], a["y"]), fontsize=8,
                        ha=str(a.get("ha", "center")), color=PALETTE["text"])

    ax.set_xlabel(c["x_label"]); ax.set_ylabel(c["y_label"])
    if "ymax" in c:
        ax.set_ylim(0, c["ymax"])
    ax.set_xlim(c["domain_min"], c["domain_max"])
    ax.legend(loc="best", fontsize=9)
    fig.tight_layout()
    save(fig, outname)
