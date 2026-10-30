#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Figure 5.2 (fig:calsharp): calibration-sharpness profile of the synthetic
human-machine panel. Data source: ../simulation_data/fig5-2_calibration_sharpness.xlsx
(derived from the ex:hmpanel study of fig5-3_loo.xlsx; RES_i = UNC_grid - POT_i).
The colour and marker of each group are read from the workbook's points sheet,
which gives them the convention of Figure 6.12 (config marker_rule). The callout
of the nearly coincident group, its text and its offsets from the group's mean
point are config entries (callout_group, callout_text, callout_*_dx/dy); the
builder checks at its tail that the named group does nearly coincide."""
from _common import load, config, save, PALETTE, plt

s = load("fig5-2_calibration_sharpness.xlsx")
c = config(s)
pts = s["points"]

# groups in workbook order; colour (a palette key) and marker from the workbook
GROUPS = list(dict.fromkeys(pts["group"].astype(str)))
SIZE = {"o": 6, "s": 6.5, "^": 8}          # optical sizes of the three marker shapes
CLUSTER = str(c["callout_group"])          # the nearly coincident group the callout names
assert CLUSTER in GROUPS, "callout_group %r is not a group of the points sheet" % CLUSTER

# annotation text is set in the neutral text gray; the marker colours alone
# carry the group coding, and the marker shapes repeat it for grayscale print
ANNOT_COL = {}

fig, ax = plt.subplots(figsize=(6.6, 4.4))
for grp in GROUPS:
    sub = pts[pts["group"].astype(str) == grp]
    mk = str(sub["marker"].iloc[0])
    col = PALETTE[str(sub["color"].iloc[0])]
    if grp == CLUSTER:
        # the members of the callout group nearly coincide; keep the data
        # positions EXACT but draw open markers of graded size so every
        # marker stays visible, and name the cluster with a leader line
        sizes = [11.0, 8.5, 6.0][:len(sub)]
        for k, (_, r) in enumerate(sub.iterrows()):
            ax.plot([r["RES"]], [r["REL"]], linestyle="none", marker=mk,
                    markersize=sizes[k], markerfacecolor="none",
                    markeredgecolor=col, markeredgewidth=1.4, alpha=0.95,
                    label=(grp if k == len(sub) - 1 else "_nolegend_"))   # legend shows the smallest
        cx = float(sub["RES"].mean())
        cy = float(sub["REL"].mean())
        ax.annotate(str(c["callout_text"]).replace("\\n", "\n"),
                    xy=(cx + c["callout_point_dx"], cy + c["callout_point_dy"]),
                    xytext=(cx + c["callout_text_dx"], cy + c["callout_text_dy"]),
                    fontsize=8, color=PALETTE["text"], ha="left", va="center",
                    arrowprops=dict(arrowstyle="->", color=PALETTE["slate"], lw=0.9,
                                    shrinkA=2, shrinkB=2))
    else:
        ax.plot(sub["RES"], sub["REL"], linestyle="none", marker=mk, color=col,
                markersize=SIZE.get(mk, 7), markeredgecolor="white", markeredgewidth=0.6,
                label=grp)

for _, a in s["annotations"].iterrows():
    txt = str(a["text"]).replace("\\n", "\n")
    key = " ".join(str(a["text"]).replace("\\n", " ").lower().split())
    col = next((v for k, v in ANNOT_COL.items() if k in key), PALETTE["text"])
    ax.text(a["x"], a["y"], txt, fontsize=8,
            color=col, ha=str(a.get("ha", "center")), va="center")

ax.set_xlim(c["xmin"], c["xmax"]); ax.set_ylim(c["ymin"], c["ymax"])
ax.set_xlabel(c["x_label"]); ax.set_ylabel(c["y_label"])
ax.legend(loc="upper right", fontsize=9)
fig.tight_layout()
save(fig, "fig5-2_calibration_sharpness")
