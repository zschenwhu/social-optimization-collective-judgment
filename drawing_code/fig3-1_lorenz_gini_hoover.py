#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Figure 3.1 (fig:lorenz): the Lorenz curve of a solved fairness-utility
profile, with its Gini area and Hoover gap.

Data source: ../simulation_data/fig3-1_lorenz_gini_hoover.xlsx, written by the
FIGURE 3.1 block of build_workbooks.py from the normalized fairness utilities
of the seven experts of the Case I fairness panel at the solved delta = 0
weights.  The Lorenz ordinates, the Gini coefficient and the Hoover index are
read from the workbook; this script only draws them, and it asserts that the
drawn geometry carries them (twice the shaded area is the Gini coefficient,
the arrow is the largest vertical gap) and that no label touches a line.
"""
from _common import load, config, save, PALETTE, plt
import numpy as np
from matplotlib.text import Text

s = load("fig3-1_lorenz_gini_hoover.xlsx")
c = config(s)
lz = s["lorenz"]
summ = dict(zip(s["summary"]["key"].astype(str), s["summary"]["value"]))
x = lz["x_share_of_experts"].to_numpy(float)
L = lz["L_share_of_utility"].to_numpy(float)
gini = float(summ["gini_pairwise"])
hoover = float(summ["hoover_max_gap"])
kh = int(summ["hoover_at_k"])
xh, yh = float(x[kh]), float(L[kh])

# the drawn geometry must carry the printed indices
area = float(np.sum(np.diff(x) * ((x[1:] - L[1:]) + (x[:-1] - L[:-1])) / 2.0))
assert abs(2.0 * area - gini) < 1e-12, "twice the shaded area is not the Gini coefficient"
assert abs((xh - yh) - hoover) < 1e-12 and kh == int(np.argmax(x - L)), "the arrow is not the largest gap"
# the values written on the figure are the ones the text prints
assert float("%.3f" % gini) == float(summ["gini_printed"]) and float("%.3f" % hoover) == float(summ["hoover_printed"])

fig, ax = plt.subplots(figsize=(5.2, 5.2))
ax.plot([0, 1], [0, 1], color=PALETTE["slate"], lw=1.6, label="line of equality", zorder=2)
ax.fill_between(x, L, x, color=PALETTE["gray_l"], alpha=0.75, lw=0, zorder=1)
ax.plot(x, L, color=PALETTE["ink"], lw=2.0, marker="o", ms=4.2, mfc="white", mew=1.1,
        label="Lorenz curve (%d experts)" % (len(x) - 1), zorder=3, clip_on=False)
arrow = ax.annotate("", xy=(xh, xh), xytext=(xh, yh),
                    arrowprops=dict(arrowstyle="<->", color=PALETTE["red"], lw=1.6,
                                    shrinkA=0, shrinkB=0), zorder=4)

# labels with leaders: the Hoover label sits in the empty region above the
# diagonal, the Gini label in the empty region below the curve
xm = 0.5
Lm = float(np.interp(xm, x, L))
gini_tip = (xm, 0.5 * (xm + Lm))
hoover_tip = (xh, xh)          # the leader meets the arrow's upper head on the diagonal
t_h = ax.annotate("Hoover index\n%.3f" % hoover, hoover_tip, xytext=(0.015, 0.27),
                  fontsize=10, color=PALETTE["text"], ha="left", va="center", zorder=5,
                  arrowprops=dict(arrowstyle="-", color=PALETTE["red"], lw=0.8, shrinkA=1, shrinkB=1))
t_g = ax.annotate("Gini coefficient %.3f\n(twice the shaded area)" % gini, gini_tip, xytext=(0.56, 0.20),
                  fontsize=10, color=PALETTE["text"], ha="left", va="center", zorder=5,
                  arrowprops=dict(arrowstyle="-", color=PALETTE["slate"], lw=0.8, shrinkA=1, shrinkB=1))

ax.set_xlim(0, 1); ax.set_ylim(0, 1)
ax.set_xticks([0, 0.5, 1]); ax.set_yticks([0, 0.5, 1])
ax.set_xlabel(c["x_label"]); ax.set_ylabel(c["y_label"])
ax.legend(loc="upper left", fontsize=10)
ax.set_aspect("equal")
fig.tight_layout()

# --- rendered check: every label box keeps >= 2 pt from every drawn line ----
fig.canvas.draw()
pt = fig.dpi / 72.0
lines = {"Lorenz curve": (x, L), "line of equality": (np.array([0.0, 1.0]), np.array([0.0, 1.0])),
         "Hoover arrow": (np.array([xh, xh]), np.array([yh, xh]))}


def _dense(px, py, n=400):
    t = np.linspace(0, 1, n)
    seg = [np.column_stack([px[i] + t * (px[i + 1] - px[i]), py[i] + t * (py[i + 1] - py[i])])
           for i in range(len(px) - 1)]
    return ax.transData.transform(np.vstack(seg))


for lab, t in (("Hoover label", t_h), ("Gini label", t_g)):
    bb = Text.get_window_extent(t)          # the text box alone, not the leader
    for name, (px, py) in lines.items():
        P = _dense(px, py)
        dx = np.maximum(np.maximum(bb.x0 - P[:, 0], P[:, 0] - bb.x1), 0.0)
        dy = np.maximum(np.maximum(bb.y0 - P[:, 1], P[:, 1] - bb.y1), 0.0)
        d = float(np.min(np.hypot(dx, dy))) / pt
        assert d >= 2.0 + 1.2, "%s is only %.1f pt from the %s" % (lab, d, name)
    lg = ax.get_legend().get_window_extent()
    assert not bb.overlaps(lg), "%s overlaps the legend" % lab
# the Gini leader ends strictly inside the shaded band, the Hoover leader on the arrow's upper head
assert np.interp(gini_tip[0], x, L) < gini_tip[1] < gini_tip[0]
assert hoover_tip == (xh, xh) and yh < xh

save(fig, "fig3-1_lorenz_gini_hoover")

if __name__ == "__main__":
    print("profile: %d experts, utilities %s" % (len(x) - 1, np.round(s["profile"]["utility_U"].to_numpy(float), 4)))
    print("Gini %.4f (twice the shaded area %.4f), Hoover %.4f at x = %d/%d"
          % (gini, 2 * area, hoover, kh, len(x) - 1))
