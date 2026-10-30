#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
build_workbooks.py
==================
Creates every Excel data file in this folder (simulation_data/) that backs the
reproducible figures and the synthetic-data tables of the monograph
"The Social Optimization of Collective Judgment" (Chen & Wang).

Each workbook is the *authoritative data source* for the figures and tables
it backs (simulation_data/INDEX.md lists them).  The plotting scripts in
drawing_code/ read only these Excel files -- never hard-coded numbers -- so the
figures are guaranteed to be generated from the tabulated data; the workbooks
named tab*.xlsx back tables of the book and are read by no plotting script.
The rule-12 self-checks at the end of this file re-read every written workbook
and abort the build on any violated invariant.

Run once:  python build_workbooks.py
Output:    ./*.xlsx (this folder; the drawing scripts read ../simulation_data/)
"""
import os
import math

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = HERE   # the workbooks are written beside this generator, in simulation_data/
os.makedirs(DATA, exist_ok=True)

# The shipped workbooks were built on the software stack named in ../environment.py (the pins
# of ../requirements.txt), and on that stack this file rebuilds them byte for byte.  On another
# stack the results agree to rounding, but non-unique weight vectors and solver diagnostics
# can differ, so the build stops before writing anything unless SOCJ_ALLOW_OTHER_STACK=1 is
# set; with it, environment.other_stack() runs this file with the self-checks reported instead
# of enforced and then ends the process, so nothing below runs a second time.
import sys as _env_sys
_env_sys.path.insert(0, os.path.normpath(os.path.join(HERE, "..")))
try:
    import environment as _environment
except ImportError:
    _environment = None
    print("   (environment.py not found beside simulation_data/; the software stack is not checked)")
if _environment is not None and not globals().get("_SOCJ_SOFT_RUN"):
    _environment.require_packages()    # a plain message, before any import, when a package is not installed
    if _environment.report():
        _environment.other_stack(os.path.abspath(__file__))

import pandas as pd


def cfg(d):
    return pd.DataFrame(list(d.items()), columns=["key", "value"])


# every workbook must carry a config/provenance sheet naming its generator,
# its seed (or the explicit statement that no draws are involved), and the
# synthetic flag; write() injects the missing keys so no workbook can ship
# without them (traceability rule 12).
_PROV_DEFAULTS = {
    "generator": "build_workbooks.py (the block that writes this workbook)",
    "seed": "none -- fixed illustrative parameters, no random draws",
    "synthetic": 1,
}


def write(name, sheets):
    """sheets: dict of sheet_name -> DataFrame (or list-of-dicts)."""
    sheets = dict(sheets)
    if "config" in sheets:
        c = sheets["config"]
        have = set(c["key"].astype(str))
        missing = {k: v for k, v in _PROV_DEFAULTS.items() if k not in have}
        if missing:
            sheets["config"] = pd.concat([c, cfg(missing)], ignore_index=True)
    elif "provenance" in sheets:
        # table-only workbooks that state their own provenance: inject only
        # the missing mandatory keys
        c = sheets["provenance"]
        have = set(c["key"].astype(str))
        missing = {k: v for k, v in dict(_PROV_DEFAULTS, workbook=name).items() if k not in have}
        if missing:
            sheets["provenance"] = pd.concat([c, cfg(missing)], ignore_index=True)
    else:
        # table-only workbooks: add a dedicated provenance sheet
        sheets["provenance"] = cfg(dict(_PROV_DEFAULTS, workbook=name))
    path = os.path.join(DATA, name)
    with pd.ExcelWriter(path, engine="openpyxl") as xl:
        for sheet, obj in sheets.items():
            df = obj if isinstance(obj, pd.DataFrame) else pd.DataFrame(obj)
            df.to_excel(xl, sheet_name=sheet, index=False)
        xl.book.properties.created = _FIXED_STAMP
    _fix_zip_stamps(path)
    print("wrote", os.path.relpath(path, HERE))


# openpyxl stamps the wall clock into docProps/core.xml (dcterms:modified is
# overwritten at save time) and into every zip entry header, so two builds of
# identical cell content are not byte-identical.  The post-pass below rewrites
# the archive with one fixed entry time and sets 'modified' equal to the fixed
# 'created' stamp, so that a rebuild from the same seed reproduces every
# workbook byte for byte (traceability rule 12; content is unchanged).
import datetime as _wb_datetime
import re as _wb_re
import zipfile as _wb_zipfile

_FIXED_STAMP = _wb_datetime.datetime(2026, 1, 1, 0, 0, 0)


def _fix_zip_stamps(path):
    with _wb_zipfile.ZipFile(path) as zin:
        entries = [(info, zin.read(info.filename)) for info in zin.infolist()]
    tmp = path + ".tmp"
    with _wb_zipfile.ZipFile(tmp, "w", compression=_wb_zipfile.ZIP_DEFLATED) as zout:
        for info, data in entries:
            if info.filename == "docProps/core.xml":
                data = _wb_re.sub(rb"(<dcterms:modified[^>]*>)[^<]*",
                               lambda m: m.group(1) + _FIXED_STAMP.strftime("%Y-%m-%dT%H:%M:%SZ").encode(),
                               data)
            new = _wb_zipfile.ZipInfo(info.filename, date_time=_FIXED_STAMP.timetuple()[:6])
            new.compress_type = _wb_zipfile.ZIP_DEFLATED
            new.external_attr = info.external_attr
            zout.writestr(new, data)
    os.replace(tmp, path)


# =====================================================================
# FIGURE 3.1 (fig:lorenz) and FIGURE 4.4 (fig:pareto) are SOLVED, not
# drawn from formulas: see the blocks "FIGURE 3.1" and "FIGURE 4.4" right
# after the block "CASE STUDY I, FAIRNESS PANEL" further down, which compute
# them on the seven-expert Case I panel with that block's functions.
# =====================================================================

# =====================================================================
# FIGURE 5.2  Calibration-sharpness profile of a mixed panel (fig:calsharp)
#   The workbook is written FURTHER DOWN, right after the fig5-3_loo
#   study: since 2026-08-28 its per-expert REL_i / RES_i values are
#   DERIVED from the same generated ex:hmpanel panel (the 'individual'
#   Hersbach split), so Figures 5.2 and 5.3 are two views of one
#   dataset and cannot drift apart.
# =====================================================================

# =====================================================================
# FIGURES 6.3 and 6.4 and TABLE 6.9 (Case I fairness panel) are SOLVED,
# not typed: see the block "CASE STUDY I, FAIRNESS PANEL" further down,
# after the shared numerical helpers, which writes fig6-3_bim_densities.xlsx,
# fig6-4_bim_tradeoff.xlsx, and tab6-9_bim_results.xlsx from the solve of
# eq:bim-triobjective.
# =====================================================================

# =====================================================================
# CASE STUDY I, FULL ASSESSMENT (Section 6.2, BIM maturity)
#   -> Figure fig6-1_bim_groupagg  (fig:bim-groupagg)
#   -> Figure fig6-2_bim_dimensions (fig:bim-dimensions)
#   -> Table  tab:bim-elicit (one generated expert's elicited q_4 profile and
#             its moment fit, as band probabilities; the fit parameters are
#             also quoted in the text of sec:bim-elicit)
#   -> Table  tab:bim-groups (the k-means groups, sizes, profiles, TOPSIS Omega)
#   -> Table  tab:bim-dimensions (the three dimensions and the global row)
# One workbook backs the whole case, because the two figures and the tables
# are views of a single synthetic dataset and must never drift apart.
#
# GENERATED AND SOLVED (since 2026-09-29, audit R24 U4a; before that the
# group opinions, Omega, the dimension distributions and the dimension
# weights were typed constants).  Every printed number of the full
# assessment is computed here from the stated inputs under seed 202603:
#   1. a stated 20 x 4 professional-attribute matrix (years of BIM
#      experience, BIM projects, project types, certifications), clustered
#      by k-means (Lloyd's algorithm, standardized columns, seeded restarts)
#      into three strata that must coincide with the designed 6/7/7 strata,
#      and TOPSIS on the raw cluster centroids (vector normalization, equal
#      criterion weights, benefit criteria, ideal and anti-ideal from the
#      centroids) giving the credibility weights Omega = closeness / sum;
#   2. a 20-expert x 24-indicator panel: each expert holds a latent normal
#      judgment N(mu_im, sd_im^2) with mu_im = dimension design level +
#      indicator offset + stratum offset + leniency_i + noise_im and sd_im
#      drawn uniformly on a stratum-specific range; the expert REPORTS the
#      ten band probabilities of that latent normal (tails assigned to the
#      end bands) rounded to hundredths by largest remainder, and the
#      reported profile is FITTED by a normal matched to its first two
#      moments at the band midpoints (sec:bim-elicit); the fits, untruncated,
#      are the inputs of the aggregation;
#   3. the base model of Algorithm alg:biobjective, max alpha*A~ - beta*V~
#      over the group simplex (A = cumulative overlap of def:consensus,
#      each overlap exact at the crossing points; V = (w.sigma)^2 of
#      def:confidence), min-max normalized by the four bounds A_min, A_max,
#      V_min, V_max solved over the simplex, at the stated operating point
#      alpha = 0.75, solved for every (group, indicator) by multi-start
#      SLSQP; the group opinion is the quantile average N(w.mu, (w.sigma)^2);
#   4. quantile averaging under Omega between groups, then under best-worst
#      criterion weights (linear model of Rezaei 2016, stated comparison
#      vectors) within each dimension and across the three dimensions, up
#      to the global maturity; the consistency of each comparison set is the
#      input-based ratio CR^I of Liang, Brunelli and Rezaei (2020), checked
#      against their tabulated thresholds (Table 3 of that paper).
# Random draws, in order: rng_panel (leniency 20, noise 20x24, sd uniforms
# 20x24); rng_kmeans (20 restarts x 3 initial centroids); rng_solve (per
# program, group-major then indicator: 24 Dirichlet(1) starts, shared by the
# four bound programs and the scalarized program of that group and indicator).
# =====================================================================
print("\n-- generating and solving Case I full assessment (fig6-1_bim_case.xlsx) ...")
import numpy as np
from scipy.optimize import minimize as _c1_minimize, linprog as _c1_linprog
from scipy.special import ndtr as _c1_ndtr
from decimal import Decimal as _c1_Dec, ROUND_HALF_UP as _c1_HALF_UP

_C1 = dict(
    seed=202603,
    N=20, K=3, sizes=(6, 7, 7),
    scale_min=0.0, scale_max=10.0, n_levels=10,
    dim_names=("Technology", "Organization", "Policy"),
    dim_colors=("blue", "green", "amber"),
    dim_sizes=(11, 7, 6),                          # tab:bim-instrument
    # generator: stated design inputs
    dim_design_level=(4.75, 5.65, 5.40),
    # indicator offsets inside each dimension (zero mean per dimension):
    # software, hardware, interoperability, information delivery, accuracy
    # and data richness below the technology level. The q1, q2, q3, q5 and q6
    # offsets follow the published application, which recommends improving
    # software application, hardware equipment, interoperability, and the
    # accuracy and delivery methods of information; the q7 (data richness)
    # offset is a design choice, as are the leadership, role and competency
    # (q15, q16, q17) offsets well above the organization level and the
    # small quality and safety (q14) offset marginally above it, which are
    # not a published finding; the rest spread around their level (the
    # design sentence of sec:bim-data is checked at the tail, (2f))
    indicator_offset=(-0.30, -0.20, -0.40, +0.10, -0.20, -0.25, -0.15, +0.45, +0.30, +0.35, +0.30,
                      -0.25, -0.20, +0.05, +0.30, +0.20, +0.15, -0.25,
                      +0.25, -0.35, +0.15, +0.10, -0.05, -0.10),
    stratum_offset=(-0.40, -0.05, +0.35),          # G1 junior, G2 mid-career, G3 senior
    leniency_sd=0.15,                              # per-expert leniency, N(0, sd)
    noise_sd=0.20,                                 # per-(expert, indicator) noise, N(0, sd)
    sd_range=((1.15, 1.55), (1.00, 1.35), (0.85, 1.20)),   # latent sd ~ U(a, b) per stratum
    report_resolution=0.01,                        # reported band probabilities, hundredths
    elicitation_expert=7,                          # tab:bim-elicit shows e7's q4 profile
    elicitation_indicator=4,
    # the intra-group base model
    alpha=0.75,                                    # beta = 1 - alpha
    n_random_starts=24,                            # Dirichlet(1) starts (+ the vertices + the centroid)
    slsqp_ftol=1e-10, slsqp_maxiter=300,
    tie_tol=1e-6,                                  # a start agrees when within this objective gap
    opinion_tol=1e-4,                              # agreeing starts must return the same group opinion to this
    member_tol=1e-5,                               # a group opinion coincides with a member when |dm| + |ds| < this
    # k-means and TOPSIS
    kmeans_restarts=20, kmeans_maxiter=100,
    topsis_weights=(0.25, 0.25, 0.25, 0.25),
)
_C1_STRATA = ("junior", "mid-career", "senior")
_C1_COLORS = ("blue", "amber", "teal")
_C1_ATTR_NAMES = ("years_bim_experience", "bim_projects", "project_types", "certifications")
# the stated 20 x 4 professional-attribute matrix (rows e1..e20; e1-e6 designed
# junior, e7-e13 mid-career, e14-e20 senior).  No stratum leads on every
# attribute: mid-career members have delivered the most projects but in the
# narrowest range of types and with the fewest certifications, and juniors
# hold more certifications than mid-career members.
_C1_ATTR = np.array([
    [2, 3, 2, 3], [3, 2, 3, 2], [1, 2, 2, 3], [4, 4, 3, 2], [2, 3, 2, 2], [3, 3, 3, 3],
    [6, 9, 1, 1], [8, 11, 2, 1], [7, 10, 2, 2], [9, 12, 1, 1], [5, 8, 2, 1], [7, 9, 1, 2], [8, 10, 2, 1],
    [12, 6, 4, 3], [15, 7, 5, 4], [11, 5, 4, 3], [14, 8, 4, 4], [13, 6, 5, 3], [16, 7, 4, 4], [12, 5, 3, 3],
], float)
_C1_IND_NAMES = ("Software Application", "Hardware Equipment", "Interoperability", "Real-time Data",
                 "Information Delivery Method", "Information Accuracy", "Data Richness & Reliability",
                 "Graphics", "Geospatial Capability", "Work Flow", "BIM + Technology",
                 "Change Management", "Risk Management", "Quality & Safety Management",
                 "Senior Leadership", "Role", "Competency Profile", "Specifications",
                 "Documentation & Modeling Standards", "Reward System", "Standard Operating Procedure",
                 "Strategic Planning", "Training Program", "Training Delivery Method")
# best-worst comparison vectors (linear model).  Within a dimension the
# criteria are its indicators in the order of tab:bim-instrument; 'best' and
# 'worst' are 1-based positions inside that list.
_C1_BWM = (
    dict(problem="Technology indicators", best=1, worst=9,
         best_to_others=(1, 3, 2, 3, 3, 2, 2, 4, 5, 2, 4),
         others_to_worst=(5, 2, 3, 2, 2, 3, 3, 1, 1, 3, 1)),
    dict(problem="Organization indicators", best=4, worst=7,
         best_to_others=(3, 2, 2, 1, 2, 2, 4),
         others_to_worst=(1, 2, 2, 4, 2, 2, 1)),
    dict(problem="Policy indicators", best=1, worst=2,
         best_to_others=(1, 4, 2, 2, 3, 3),
         others_to_worst=(4, 1, 2, 2, 1, 1)),
    dict(problem="Dimensions", best=1, worst=3,
         best_to_others=(1, 2, 4),
         others_to_worst=(4, 2, 1)),
)
# Thresholds for the input-based consistency ratio CR^I, Table 3 of Liang,
# Brunelli and Rezaei (2020), Omega 96:102175: rows are the scale a_BW
# (3..9), columns the number of criteria (3..9).  Stated reference data.
_C1_CRI_THRESH = {
    3: (0.1667, 0.1667, 0.1667, 0.1667, 0.1667, 0.1667, 0.1667),
    4: (0.1121, 0.1529, 0.1898, 0.2206, 0.2527, 0.2577, 0.2683),
    5: (0.1354, 0.1994, 0.2306, 0.2546, 0.2716, 0.2844, 0.2960),
    6: (0.1330, 0.1990, 0.2643, 0.3044, 0.3144, 0.3221, 0.3262),
    7: (0.1294, 0.2457, 0.2819, 0.3029, 0.3144, 0.3251, 0.3403),
    8: (0.1309, 0.2521, 0.2958, 0.3154, 0.3408, 0.3620, 0.3657),
    9: (0.1359, 0.2681, 0.3062, 0.3337, 0.3517, 0.3620, 0.3662),
}
_c1_N, _c1_K = _C1["N"], _C1["K"]
_c1_group_of = np.repeat(np.arange(_c1_K), _C1["sizes"])
_c1_dim_of = np.repeat(np.arange(3), _C1["dim_sizes"])
_c1_M = len(_c1_dim_of)
_c1_ind_off = np.array(_C1["indicator_offset"], float)
assert len(_c1_ind_off) == _c1_M == 24 and _C1_ATTR.shape == (_c1_N, 4)
for _c in range(3):
    assert abs(_c1_ind_off[_c1_dim_of == _c].sum()) < 1e-12, "indicator offsets must have zero mean per dimension"


def _c1_printed(mu, sd):
    """N(mu, sd^2) rounded half-up to two decimals, as the book prints."""
    # float(...) keeps the text of the number the same under NumPy 2, whose scalars print as np.float64(...)
    r = lambda v: _c1_Dec(repr(float(round(v, 9)))).quantize(_c1_Dec("0.01"), rounding=_c1_HALF_UP)
    return "N(%s, %s^2)" % (r(mu), r(sd))


# ---- 1. k-means on the attribute matrix, TOPSIS on the centroids ----------------
def _c1_kmeans(Xr, K, rng, restarts, maxiter):
    """Lloyd's algorithm on standardized columns; `restarts` seeded random
    initializations (K distinct rows), the lowest inertia kept; clusters
    numbered in order of first appearance along e1..eN."""
    Z = (Xr - Xr.mean(0)) / Xr.std(0)
    best = None
    for r in range(restarts):
        c = Z[rng.choice(len(Z), K, replace=False)]
        for _ in range(maxiter):
            lab = np.argmin(((Z[:, None, :] - c[None, :, :]) ** 2).sum(-1), axis=1)
            newc = np.array([Z[lab == k].mean(0) if np.any(lab == k) else c[k] for k in range(K)])
            if np.allclose(newc, c):
                break
            c = newc
        inertia = float(sum(((Z[lab == k] - c[k]) ** 2).sum() for k in range(K)))
        if best is None or inertia < best[0] - 1e-12:
            best = (inertia, lab.copy(), r)
    inertia, lab, r = best
    order = []
    for l in lab:
        if l not in order:
            order.append(l)
    return inertia, np.array([order.index(l) for l in lab]), r


_rng_c1_kmeans = np.random.default_rng(_C1["seed"])
_c1_inertia, _c1_lab, _c1_best_restart = _c1_kmeans(
    _C1_ATTR, _c1_K, _rng_c1_kmeans, _C1["kmeans_restarts"], _C1["kmeans_maxiter"])
_c1_sizes = np.bincount(_c1_lab, minlength=_c1_K)
assert tuple(int(s) for s in _c1_sizes) == tuple(_C1["sizes"]), \
    "k-means did not recover the designed 6/7/7 strata: %r" % (_c1_sizes,)
assert np.array_equal(_c1_lab, _c1_group_of), "k-means membership differs from the design"
_c1_cent = np.array([_C1_ATTR[_c1_lab == k].mean(0) for k in range(_c1_K)])
_c1_Xn = _c1_cent / np.sqrt((_c1_cent ** 2).sum(0))
_c1_Vt = _c1_Xn * np.array(_C1["topsis_weights"])
_c1_ideal, _c1_anti = _c1_Vt.max(0), _c1_Vt.min(0)
_c1_dpos = np.sqrt(((_c1_Vt - _c1_ideal) ** 2).sum(1))
_c1_dneg = np.sqrt(((_c1_Vt - _c1_anti) ** 2).sum(1))
_c1_close = _c1_dneg / (_c1_dpos + _c1_dneg)
_c1_Omega = _c1_close / _c1_close.sum()
print("   k-means strata %s (restart %d, inertia %.3f); TOPSIS Omega = %s"
      % (tuple(int(s) for s in _c1_sizes), _c1_best_restart, _c1_inertia, np.round(_c1_Omega, 4)))

# ---- 2. the panel: latent normals -> reported profiles -> moment fits -----------
_rng_c1_panel = np.random.default_rng(_C1["seed"])
_c1_leniency = _rng_c1_panel.normal(0.0, _C1["leniency_sd"], _c1_N)
_c1_noise = _rng_c1_panel.normal(0.0, _C1["noise_sd"], (_c1_N, _c1_M))
_c1_u = _rng_c1_panel.uniform(0.0, 1.0, (_c1_N, _c1_M))
_c1_sd_lo = np.array([_C1["sd_range"][k][0] for k in _c1_group_of])
_c1_sd_hi = np.array([_C1["sd_range"][k][1] for k in _c1_group_of])
_c1_ind_target = np.array(_C1["dim_design_level"])[_c1_dim_of] + _c1_ind_off
_c1_lat_mu = (_c1_ind_target[None, :] + np.array(_C1["stratum_offset"])[_c1_group_of][:, None]
              + _c1_leniency[:, None] + _c1_noise)
_c1_lat_sd = _c1_sd_lo[:, None] + (_c1_sd_hi - _c1_sd_lo)[:, None] * _c1_u
_c1_edges = np.arange(0.0, _C1["n_levels"] + 1.0)             # band borders 0, 1, ..., 10
_c1_mid = _c1_edges[:-1] + 0.5                                  # band midpoints l - 1/2
_c1_units = int(round(1.0 / _C1["report_resolution"]))


def _c1_report(m, s):
    """The ten band probabilities of N(m, s^2), the two tails assigned to L1
    and L10, rounded to hundredths by largest remainder (ties to the lower
    band); returned as integer hundredths summing to 100."""
    c = np.concatenate([[0.0], _c1_ndtr((_c1_edges[1:-1] - m) / s), [1.0]])
    q = np.diff(c) * _c1_units
    fl = np.floor(q)
    rem = int(round(_c1_units - fl.sum()))
    order = np.argsort(-(q - fl), kind="stable")
    fl[order[:rem]] += 1
    return fl.astype(int)


def _c1_fit(counts):
    """sec:bim-elicit: normal matched to the profile's mean and variance at the band midpoints."""
    p = counts / counts.sum()
    mu = float(p @ _c1_mid)
    return mu, math.sqrt(float(p @ (_c1_mid - mu) ** 2))


_c1_prof = np.zeros((_c1_N, _c1_M, _C1["n_levels"]), int)
_c1_fm = np.zeros((_c1_N, _c1_M))
_c1_fs = np.zeros((_c1_N, _c1_M))
for _i in range(_c1_N):
    for _m in range(_c1_M):
        _c1_prof[_i, _m] = _c1_report(_c1_lat_mu[_i, _m], _c1_lat_sd[_i, _m])
        _c1_fm[_i, _m], _c1_fs[_i, _m] = _c1_fit(_c1_prof[_i, _m])
assert (_c1_prof.sum(axis=2) == _c1_units).all()
_c1_tail = _c1_ndtr(-(_c1_fm - _C1["scale_min"]) / _c1_fs) + _c1_ndtr(-(_C1["scale_max"] - _c1_fm) / _c1_fs)
print("   panel: %d experts x %d indicators; fitted means %.2f-%.2f, sds %.3f-%.3f; max mass outside [0,10] %.4f"
      % (_c1_N, _c1_M, _c1_fm.min(), _c1_fm.max(), _c1_fs.min(), _c1_fs.max(), _c1_tail.max()))


# ---- 3. the intra-group base model -------------------------------------------------
_C1_SQRT2 = math.sqrt(2.0)


def _c1_Phi(x):
    """standard normal cdf through math.erfc (double precision, scalar)."""
    return 0.5 * math.erfc(-x / _C1_SQRT2)


_C1_INV_SQRT2PI = 1.0 / math.sqrt(2.0 * math.pi)


def _c1_ovl_grad(mu, sg, m, s):
    """Overlap int min{f_i, f} dx of the expert N(mu, sg^2) with the aggregate
    f = N(m, s^2), exact, and its partial derivatives in (m, s).  The two
    densities cross at the roots r1 < r2 of log f = log f_i (one root when
    s = sg); between the roots the narrower density is the larger one.  On
    the set R where f is the minimum, d/dm int_R f = -[f] and
    d/ds int_R f = -(1/s)[(x - m) f] evaluated at the ends of R (the terms
    from moving the crossing points cancel because f = f_i there)."""
    pdf = lambda x: _C1_INV_SQRT2PI / s * math.exp(-0.5 * ((x - m) / s) ** 2)
    if abs(s - sg) < 1e-12:
        c = 0.5 * (m + mu)
        o = 2.0 * _c1_Phi(-abs(m - mu) / (2.0 * s))
        if m == mu:
            return o, 0.0, 0.0
        sgn = 1.0 if m < mu else -1.0          # R = (c, inf) when m < mu, (-inf, c) otherwise
        return o, sgn * pdf(c), sgn * (c - m) * pdf(c) / s
    a = 0.5 / sg ** 2 - 0.5 / s ** 2
    b = m / s ** 2 - mu / sg ** 2
    c = 0.5 * mu ** 2 / sg ** 2 - 0.5 * m ** 2 / s ** 2 + math.log(sg / s)
    d = math.sqrt(b * b - 4.0 * a * c)
    r1, r2 = sorted(((-b - d) / (2.0 * a), (-b + d) / (2.0 * a)))
    pa = _c1_Phi((r2 - m) / s) - _c1_Phi((r1 - m) / s)       # aggregate mass between the roots
    pe = _c1_Phi((r2 - mu) / sg) - _c1_Phi((r1 - mu) / sg)   # expert mass between the roots
    f1, f2 = pdf(r1), pdf(r2)
    if s < sg:      # aggregate narrower: the expert is the minimum inside, the aggregate outside
        return pe + 1.0 - pa, f2 - f1, ((r2 - m) * f2 - (r1 - m) * f1) / s
    return pa + 1.0 - pe, f1 - f2, ((r1 - m) * f1 - (r2 - m) * f2) / s


def _c1_unit(w):
    w = np.clip(np.asarray(w, float), 0.0, None)
    s = w.sum()
    return w / s if s > 0 else np.full(len(w), 1.0 / len(w))


class _C1Program:
    """max_w alpha*A~(w) - (1-alpha)*V~(w) on the simplex of one group for one
    indicator (alg:biobjective with the group's fitted normals as the profile).
    Both objectives depend on w only through (m, s) = (w.mu, w.sigma), so their
    gradients are exact: dA/dw = A_m mu + A_s sigma and dV/dw = 2 s sigma."""

    def __init__(self, members, mu, sd):
        self.members, self.mu, self.sd, self.n = members, np.asarray(mu, float), np.asarray(sd, float), len(mu)
        self._pairs = [(float(a), float(b)) for a, b in zip(mu, sd)]

    def ms(self, w):
        return float(w @ self.mu), float(w @ self.sd)

    def A_grad(self, w):
        m, s = self.ms(w)
        o = dm = ds = 0.0
        for mi, si in self._pairs:
            a, b, c = _c1_ovl_grad(mi, si, m, s)
            o += a; dm += b; ds += c
        return o, dm * self.mu + ds * self.sd

    def A(self, w):
        return self.A_grad(w)[0]

    def V_grad(self, w):
        s = float(w @ self.sd)
        return s * s, 2.0 * s * self.sd

    def V(self, w):
        return self.V_grad(w)[0]

    def multistart(self, fg, starts):
        """minimize fg (value, gradient) over the simplex by SLSQP from every
        start; returns the best value, its w, all terminal values (at the
        renormalized terminal points), the terminal (m, s), and the number
        of runs SLSQP reports as converged."""
        cons = ({"type": "eq", "fun": lambda w: w.sum() - 1.0, "jac": lambda w: np.ones_like(w)},)
        vals, ws, ok = [], [], 0
        for w0 in starts:
            r = _c1_minimize(fg, w0, jac=True, method="SLSQP",
                             bounds=[(0.0, 1.0)] * self.n, constraints=cons,
                             options={"ftol": _C1["slsqp_ftol"], "maxiter": _C1["slsqp_maxiter"]})
            w = _c1_unit(r.x)
            vals.append(fg(w)[0]); ws.append(w); ok += int(r.success)
        vals = np.array(vals)
        k = int(np.argmin(vals))
        return float(vals[k]), ws[k], vals, np.array([self.ms(w) for w in ws]), ok

    def scalarized(self, w):
        """alpha*A~ - (1-alpha)*V~ at w and its gradient."""
        al = _C1["alpha"]
        a, ga = self.A_grad(w)
        v, gv = self.V_grad(w)
        dA, dV = self.Amax - self.Amin, self.Vmax - self.Vmin
        return (al * (a - self.Amin) / dA - (1.0 - al) * (v - self.Vmin) / dV,
                al * ga / dA - (1.0 - al) * gv / dV)

    def solve(self, rng):
        n = self.n
        starts = list(rng.dirichlet(np.ones(n), _C1["n_random_starts"])) + list(np.eye(n)) + [np.full(n, 1.0 / n)]
        self.n_starts = len(starts)
        tol = _C1["tie_tol"]
        neg = lambda fg: (lambda w: tuple(-v for v in fg(w)))
        self.Amin, _, va, _, _ = self.multistart(self.A_grad, starts)
        negAmax, _, vb, _, _ = self.multistart(neg(self.A_grad), starts)
        self.Amax = -negAmax
        self.Vmin, _, vc, _, _ = self.multistart(self.V_grad, starts)
        negVmax, _, vd, _, _ = self.multistart(neg(self.V_grad), starts)
        self.Vmax = -negVmax
        self.bound_agree = [int(np.sum(v - v.min() < tol)) for v in (va, vb, vc, vd)]
        assert self.Amax - self.Amin > 1e-6 and self.Vmax - self.Vmin > 1e-6, "degenerate objective range"
        negf, self.w, vals, msv, self.n_success = self.multistart(neg(self.scalarized), starts)
        self.objective = -negf
        at = vals - negf < tol
        self.agree = int(at.sum())
        self.m, self.s = self.ms(self.w)
        self.ms_spread = float(np.max(np.abs(msv[at] - np.array([self.m, self.s]))))
        others = vals[~at]
        self.runner_up_gap = float(others.min() - negf) if others.size else float("nan")
        self.at_member = int(np.min(np.abs(self.mu - self.m) + np.abs(self.sd - self.s)) < _C1["member_tol"])
        return self


_rng_c1_solve = np.random.default_rng(_C1["seed"])
import time as _c1_time
_c1_t0 = _c1_time.time()
_c1_progs = {}
for _k in range(_c1_K):
    _mem = np.where(_c1_group_of == _k)[0]
    for _m in range(_c1_M):
        _c1_progs[(_k, _m)] = _C1Program(_mem, _c1_fm[_mem, _m], _c1_fs[_mem, _m]).solve(_rng_c1_solve)
_c1_GM = np.array([[_c1_progs[(k, m)].m for m in range(_c1_M)] for k in range(_c1_K)])
_c1_GS = np.array([[_c1_progs[(k, m)].s for m in range(_c1_M)] for k in range(_c1_K)])
_c1_agree = [p.agree for p in _c1_progs.values()]
_c1_nstarts = [p.n_starts for p in _c1_progs.values()]
_c1_all_agree = all(p.agree == p.n_starts for p in _c1_progs.values())
_c1_bound_agree_min = min(min(p.bound_agree) for p in _c1_progs.values())
_c1_ms_spread = max(p.ms_spread for p in _c1_progs.values())
_c1_n_member = sum(p.at_member for p in _c1_progs.values())
assert _c1_ms_spread < _C1["opinion_tol"], "agreeing starts return different group opinions (%.2e)" % _c1_ms_spread
# Where a group opinion coincides with one member's fit, record whether that
# member's (mean, sd) point lies strictly inside the convex hull of the
# group's points (so the coincidence is not a vertex of the feasible (m, s)
# set but the kink of that member's overlap, which equals one there) and
# whether the recorded weight vector is concentrated on the member.
from scipy.spatial import ConvexHull as _c1_ConvexHull
_c1_member_interior = 0
_c1_member_wmax_one = 0
for _p in _c1_progs.values():
    if not _p.at_member:
        continue
    _P2 = np.column_stack([_p.mu, _p.sd])
    _j = int(np.argmin(np.abs(_p.mu - _p.m) + np.abs(_p.sd - _p.s)))
    _eqs = _c1_ConvexHull(_P2).equations
    _c1_member_interior += int(np.all(_eqs[:, :2] @ _P2[_j] + _eqs[:, 2] < -1e-9))
    _c1_member_wmax_one += int(float(np.max(_p.w)) > 1.0 - _C1["member_tol"])
print("   intra-group base model (alpha %.2f): %d programs in %.0fs; starts at the optimum %d-%d of %d-%d; "
      "(m, s) spread %.1e; %d group opinions coincide with one member (%d strictly inside the group hull, "
      "%d with recorded max weight 1)"
      % (_C1["alpha"], len(_c1_progs), _c1_time.time() - _c1_t0, min(_c1_agree), max(_c1_agree),
         min(_c1_nstarts), max(_c1_nstarts), _c1_ms_spread, _c1_n_member, _c1_member_interior, _c1_member_wmax_one))

# ---- 4. inter-group QA under Omega; best-worst weights; the hierarchy --------------
_c1_IM = _c1_Omega @ _c1_GM                                    # indicator collectives (normal)
_c1_IS = _c1_Omega @ _c1_GS


def _c1_bwm(aB, aW, best, worst):
    """Linear best-worst model (Rezaei 2016): min xi s.t. |w_B - a_Bj w_j| <= xi,
    |w_j - a_jW w_W| <= xi, sum w = 1, w >= 0; returns (w, xi)."""
    n = len(aB)
    c = np.zeros(n + 1); c[-1] = 1.0
    A_ub, b_ub = [], []
    for j in range(n):
        if j != best:
            for sgn in (1.0, -1.0):
                row = np.zeros(n + 1); row[best] = sgn; row[j] = -sgn * aB[j]; row[-1] = -1.0
                A_ub.append(row); b_ub.append(0.0)
        if j != worst:
            for sgn in (1.0, -1.0):
                row = np.zeros(n + 1); row[j] = sgn; row[worst] = -sgn * aW[j]; row[-1] = -1.0
                A_ub.append(row); b_ub.append(0.0)
    A_eq = np.r_[np.ones(n), 0.0][None, :]
    res = _c1_linprog(c, A_ub=np.array(A_ub), b_ub=np.array(b_ub), A_eq=A_eq, b_eq=[1.0],
                      bounds=[(0.0, None)] * (n + 1), method="highs")
    assert res.success, "BWM linear program failed: %s" % res.message
    # the optimal weights must be unique: min and max of each w_j over the optimal face
    xi = float(res.x[-1])
    spread = 0.0
    for j in range(n):
        for sgn in (1.0, -1.0):
            cj = np.zeros(n + 1); cj[j] = sgn
            r2 = _c1_linprog(cj, A_ub=np.vstack([np.array(A_ub), np.r_[np.zeros(n), 1.0][None, :]]),
                             b_ub=np.r_[np.array(b_ub), xi + 1e-12], A_eq=A_eq, b_eq=[1.0],
                             bounds=[(0.0, None)] * (n + 1), method="highs")
            spread = max(spread, abs(float(r2.x[j]) - float(res.x[j])))
    return res.x[:n] / res.x[:n].sum(), xi, spread


def _c1_cri(aB, aW, worst):
    """Input-based consistency ratio of Liang, Brunelli and Rezaei (2020), eqs. (6)-(7):
    CR^I = max_j |a_Bj a_jW - a_BW| / (a_BW^2 - a_BW), and 0 when a_BW = 1."""
    aBW = aB[worst]
    if aBW == 1:
        return 0.0
    return float(max(abs(aB[j] * aW[j] - aBW) for j in range(len(aB))) / (aBW * aBW - aBW))


_c1_bwm_rows, _c1_bwm_summary = [], []
_c1_wind = np.zeros(_c1_M)
_c1_wdim = None
for _b, _P in enumerate(_C1_BWM):
    _aB = np.array(_P["best_to_others"], float)
    _aW = np.array(_P["others_to_worst"], float)
    _bi, _wi = _P["best"] - 1, _P["worst"] - 1
    assert _aB[_bi] == 1 and _aW[_wi] == 1 and _aB[_wi] == _aW[_bi], "BWM vectors of %s malformed" % _P["problem"]
    _w, _xi, _spr = _c1_bwm(_aB, _aW, _bi, _wi)
    assert _spr < 1e-7, "BWM weights of %s are not unique (spread %.2e)" % (_P["problem"], _spr)
    _cr = _c1_cri(_aB, _aW, _wi)
    _aBW = int(_aB[_wi])
    _n = len(_aB)
    _thr_row = _C1_CRI_THRESH.get(_aBW)
    _thr = _thr_row[_n - 3] if (_thr_row is not None and 3 <= _n <= 9) else float("nan")
    _thr_min = min(_thr_row) if _thr_row is not None else float("nan")
    if _b < 3:
        _ids = np.where(_c1_dim_of == _b)[0]
        _c1_wind[_ids] = _w
        _names = ["q%d %s" % (j + 1, _C1_IND_NAMES[j]) for j in _ids]
    else:
        _c1_wdim = _w
        _names = list(_C1["dim_names"])
    for _j in range(_n):
        _c1_bwm_rows.append({"problem": _P["problem"], "criterion": _names[_j],
                             "best_to_others": int(_aB[_j]), "others_to_worst": int(_aW[_j]),
                             "weight": float(_w[_j]),
                             "CR_I_j": (abs(_aB[_j] * _aW[_j] - _aBW) / (_aBW * _aBW - _aBW)) if _aBW > 1 else 0.0})
    _c1_bwm_summary.append({"problem": _P["problem"], "n_criteria": _n,
                            "best": _names[_bi], "worst": _names[_wi], "a_BW": _aBW,
                            "xi_L": _xi, "CR_I": _cr,
                            "threshold_liang2020": _thr,
                            "min_threshold_for_scale": _thr_min,
                            "weight_uniqueness_spread": _spr})
    print("   BWM %-24s a_BW %d  n %2d  xi^L %.4f  CR^I %.4f  (threshold %s; smallest for the scale %.4f)"
          % (_P["problem"], _aBW, _n, _xi, _cr, ("%.4f" % _thr) if _thr == _thr else "n/a", _thr_min))
_c1_DM = np.array([_c1_wind[_c1_dim_of == c] @ _c1_IM[_c1_dim_of == c] for c in range(3)])
_c1_DS = np.array([_c1_wind[_c1_dim_of == c] @ _c1_IS[_c1_dim_of == c] for c in range(3)])
_c1_glob_mu, _c1_glob_sd = float(_c1_wdim @ _c1_DM), float(_c1_wdim @ _c1_DS)
# the published application gave the three dimensions equal importance in its global
# aggregate (best-worst weights only for the indicators); the case departs from that,
# and records the equal-weight global beside its own for comparison
_c1_glob_eq_mu, _c1_glob_eq_sd = float(_c1_DM.mean()), float(_c1_DS.mean())
_c1_q = _C1["elicitation_indicator"] - 1
_c1_q4_mu, _c1_q4_sd = float(_c1_IM[_c1_q]), float(_c1_IS[_c1_q])
# a linear pool of the same three q4 group opinions, for comparison with the QA spread
_c1_q4_lp_sd = math.sqrt(float(_c1_Omega @ (_c1_GS[:, _c1_q] ** 2 + (_c1_GM[:, _c1_q] - _c1_q4_mu) ** 2)))
_c1_e = _C1["elicitation_expert"] - 1
_c1_fit_mu, _c1_fit_sd = float(_c1_fm[_c1_e, _c1_q]), float(_c1_fs[_c1_e, _c1_q])
_c1_level_names = ("Not present", "Initial", "Experiential", "Adjusted", "Managed", "Defined",
                   "Standard", "Quantitatively managed", "Optimized", "Full-fledged")
_c1_band = lambda x: _c1_level_names[int(min(max(math.floor(x), 0), 9))]
print("   q4 group opinions %s -> collective N(%.4f, %.4f^2) (linear pool sd %.4f)"
      % ([_c1_printed(a, b) for a, b in zip(_c1_GM[:, _c1_q], _c1_GS[:, _c1_q])], _c1_q4_mu, _c1_q4_sd, _c1_q4_lp_sd))
print("   dimensions %s; global N(%.4f, %.4f^2) (%s band)"
      % ([_c1_printed(a, b) for a, b in zip(_c1_DM, _c1_DS)], _c1_glob_mu, _c1_glob_sd, _c1_band(_c1_glob_mu)))
print("   tab:bim-elicit: e%d profile %s -> fit N(%.4f, %.4f^2)"
      % (_c1_e + 1, (_c1_prof[_c1_e, _c1_q] / _c1_units).tolist(), _c1_fit_mu, _c1_fit_sd))


def _c1_range(v):
    lo, hi = int(v.min()), int(v.max())
    return "%d" % lo if lo == hi else "%d-%d" % (lo, hi)


_c1_groups_rows = []
for _k in range(_c1_K):
    _A_k = _C1_ATTR[_c1_lab == _k]
    _row = {"group": "G%d (%s)" % (_k + 1, _C1_STRATA[_k]), "legend": "G%d (%s)" % (_k + 1, _C1_STRATA[_k]),
            "stratum": _C1_STRATA[_k], "size": int(_c1_sizes[_k]),
            "mu": float(_c1_GM[_k, _c1_q]), "sigma": float(_c1_GS[_k, _c1_q]),
            "Omega": float(_c1_Omega[_k]), "color": _C1_COLORS[_k]}
    for _a, _nm in enumerate(_C1_ATTR_NAMES):
        _row["range_" + _nm] = _c1_range(_A_k[:, _a])
    for _a, _nm in enumerate(_C1_ATTR_NAMES):
        _row["centroid_" + _nm] = float(_c1_cent[_k, _a])
    for _a, _nm in enumerate(_C1_ATTR_NAMES):
        _row["norm_" + _nm] = float(_c1_Xn[_k, _a])
    _row.update({"d_plus": float(_c1_dpos[_k]), "d_minus": float(_c1_dneg[_k]), "closeness": float(_c1_close[_k])})
    _c1_groups_rows.append(_row)
_c1_dims_rows = []
for _c in range(3):
    _ids = np.where(_c1_dim_of == _c)[0]
    _lo_i, _hi_i = _ids[int(np.argmin(_c1_IM[_ids]))], _ids[int(np.argmax(_c1_IM[_ids]))]
    _c1_dims_rows.append({"dimension": _C1["dim_names"][_c], "mu": float(_c1_DM[_c]), "sigma": float(_c1_DS[_c]),
                          "weight": float(_c1_wdim[_c]), "color": _C1["dim_colors"][_c],
                          "n_indicators": int(_C1["dim_sizes"][_c]),
                          "design_level": float(_C1["dim_design_level"][_c]),
                          "band": _c1_band(float(_c1_DM[_c])),
                          "printed": _c1_printed(float(_c1_DM[_c]), float(_c1_DS[_c])),
                          "lowest_indicator": "q%d %s" % (_lo_i + 1, _C1_IND_NAMES[_lo_i]),
                          "highest_indicator": "q%d %s" % (_hi_i + 1, _C1_IND_NAMES[_hi_i])})
_c1_ind_rows = []
for _m in range(_c1_M):
    _r = {"indicator": "q%d" % (_m + 1), "name": _C1_IND_NAMES[_m], "dimension": _C1["dim_names"][_c1_dim_of[_m]],
          "design_target": float(_c1_ind_target[_m])}
    for _k in range(_c1_K):
        _r["G%d_mu" % (_k + 1)] = float(_c1_GM[_k, _m])
        _r["G%d_sigma" % (_k + 1)] = float(_c1_GS[_k, _m])
    _r.update({"mu": float(_c1_IM[_m]), "sigma": float(_c1_IS[_m]), "bwm_weight": float(_c1_wind[_m]),
               "rank_in_dimension": int(1 + np.sum(_c1_IM[_c1_dim_of == _c1_dim_of[_m]] < _c1_IM[_m]))})
    _c1_ind_rows.append(_r)
_c1_expert_rows = [dict({"expert": "e%d" % (_i + 1)},
                        **{_nm: float(_C1_ATTR[_i, _a]) for _a, _nm in enumerate(_C1_ATTR_NAMES)},
                        **{"group_design": "G%d" % (_c1_group_of[_i] + 1), "group_kmeans": "G%d" % (_c1_lab[_i] + 1),
                           "leniency": float(_c1_leniency[_i])}) for _i in range(_c1_N)]
_c1_panel_rows = []
for _i in range(_c1_N):
    for _m in range(_c1_M):
        _r = {"expert": "e%d" % (_i + 1), "group": "G%d" % (_c1_group_of[_i] + 1), "indicator": "q%d" % (_m + 1),
              "dimension": _C1["dim_names"][_c1_dim_of[_m]], "design_target": float(_c1_ind_target[_m]),
              "latent_mu": float(_c1_lat_mu[_i, _m]), "latent_sd": float(_c1_lat_sd[_i, _m])}
        for _l in range(_C1["n_levels"]):
            _r["p_L%d" % (_l + 1)] = _c1_prof[_i, _m, _l] / _c1_units
        _r.update({"fit_mu": float(_c1_fm[_i, _m]), "fit_sigma": float(_c1_fs[_i, _m]),
                   "mass_outside_0_10": float(_c1_tail[_i, _m])})
        _c1_panel_rows.append(_r)
_c1_sol_rows = []
for (_k, _m), _p in sorted(_c1_progs.items()):
    _r = {"group": "G%d" % (_k + 1), "indicator": "q%d" % (_m + 1), "dimension": _C1["dim_names"][_c1_dim_of[_m]]}
    for _i in range(_c1_N):
        _r["w_e%d" % (_i + 1)] = float(_p.w[list(_p.members).index(_i)]) if _i in _p.members else np.nan
    _r.update({"m": _p.m, "s": _p.s, "A": _p.A(_p.w), "V": _p.V(_p.w),
               "A_min": _p.Amin, "A_max": _p.Amax, "V_min": _p.Vmin, "V_max": _p.Vmax,
               "objective": _p.objective, "n_starts": _p.n_starts, "starts_at_optimum": _p.agree,
               "bound_starts_at_optimum": ",".join(str(v) for v in _p.bound_agree),
               "ms_spread_at_optimum": _p.ms_spread, "runner_up_gap": _p.runner_up_gap,
               "slsqp_success": _p.n_success, "coincides_with_member": _p.at_member})
    _c1_sol_rows.append(_r)
# the band probabilities of the fitted normal (tails to the end bands), shown
# beside the reported profile in tab:bim-elicit
_c1_fit_band = np.diff(np.concatenate([[0.0], _c1_ndtr((_c1_edges[1:-1] - _c1_fit_mu) / _c1_fit_sd), [1.0]]))
_c1_elicit_rows = [{"level": _l + 1, "level_mid": float(_c1_mid[_l]),
                    "probability": _c1_prof[_c1_e, _c1_q, _l] / _c1_units,
                    "fit_probability": float(_c1_fit_band[_l])}
                   for _l in range(_C1["n_levels"])]

write("fig6-1_bim_case.xlsx", {
    "config": cfg({
        "case": "I -- BIM maturity assessment (Section 6.2)",
        "figures": "fig6-1_bim_groupagg, fig6-2_bim_dimensions",
        "labels": "fig:bim-groupagg, fig:bim-dimensions",
        "tables": ("tab:bim-elicit (elicitation_q4: the profile and the band probabilities of its fit; "
                   "fit parameters in derived), tab:bim-groups (groups), "
                   "tab:bim-dimensions (dimensions + derived global row)"),
        "source_paper": "chen2023bimdc",
        "synthetic": 1,
        "seed": _C1["seed"],
        "generator": ("build_workbooks.py Case I full-assessment block: generated panel, elicited profiles and "
                      "moment fits, k-means + TOPSIS, multi-start SLSQP intra-group base model, "
                      "best-worst weights, quantile-averaged hierarchy"),
        "rng_order": ("rng_panel(seed): leniency[20], noise[20x24], sd uniforms[20x24]; "
                      "rng_kmeans(seed): 20 restarts x choice of 3 rows; "
                      "rng_solve(seed): per program (group-major, then indicator) %d Dirichlet(1) starts" % _C1["n_random_starts"]),
        "n_experts": _c1_N, "n_indicators": _c1_M, "n_dimensions": 3, "n_groups": _c1_K,
        "group_sizes_design": ",".join(str(s) for s in _C1["sizes"]),
        "scale_min": _C1["scale_min"], "scale_max": _C1["scale_max"], "n_levels": _C1["n_levels"],
        "dim_design_level": ",".join("%.2f" % v for v in _C1["dim_design_level"]),
        "indicator_offset": ",".join("%+.2f" % v for v in _C1["indicator_offset"]),
        "stratum_offset_G1_G2_G3": ",".join("%+.2f" % v for v in _C1["stratum_offset"]),
        "leniency_sd": _C1["leniency_sd"], "noise_sd": _C1["noise_sd"],
        "latent_sd_range_G1_G2_G3": "; ".join("U(%.2f, %.2f)" % r for r in _C1["sd_range"]),
        "latent_rule": "mu_im = dimension design level + indicator offset + stratum offset + leniency_i + noise_im",
        "report_rule": ("band probabilities of the latent N(mu_im, sd_im^2) over [l-1, l), the tails assigned to "
                        "L1 and L10, rounded to hundredths by largest remainder (largest-remainder ties assigned "
                        "to the lower band, by a stable sort of the remainders)"),
        "report_resolution": _C1["report_resolution"],
        "fit_rule": "moment matching at band midpoints: mu = sum p_l m_l, sigma^2 = sum p_l (m_l - mu)^2 with m_l = l - 0.5",
        "density_family": "untruncated normal (the fits enter the aggregation as N(fit_mu, fit_sigma^2))",
        "fitted_mean_min": float(_c1_fm.min()), "fitted_mean_max": float(_c1_fm.max()),
        "fitted_sd_min": float(_c1_fs.min()), "fitted_sd_max": float(_c1_fs.max()),
        "max_mass_outside_0_10_fits": float(_c1_tail.max()),
        "max_mass_outside_0_10_aggregates": float(max(
            np.max(_c1_ndtr(-_c1_IM / _c1_IS) + _c1_ndtr(-(10.0 - _c1_IM) / _c1_IS)),
            np.max(_c1_ndtr(-_c1_GM / _c1_GS) + _c1_ndtr(-(10.0 - _c1_GM) / _c1_GS)),
            np.max(_c1_ndtr(-_c1_DM / _c1_DS) + _c1_ndtr(-(10.0 - _c1_DM) / _c1_DS)))),
        "elicitation_expert": "e%d" % _C1["elicitation_expert"],
        "elicitation_expert_group": "G%d" % (_c1_group_of[_c1_e] + 1),
        "elicitation_latent": "N(%.4f, %.4f^2)" % (_c1_lat_mu[_c1_e, _c1_q], _c1_lat_sd[_c1_e, _c1_q]),
        "elicitation_mass_managed_defined": float((_c1_prof[_c1_e, _c1_q, 4] + _c1_prof[_c1_e, _c1_q, 5]) / _c1_units),
        "program": ("alg:biobjective per (group, indicator): max_w alpha*A~(w) - (1-alpha)*V~(w) over the group "
                    "simplex, A = sum of exact overlaps of the members' fits with N(w.mu, (w.sigma)^2), V = (w.sigma)^2"),
        "alpha": _C1["alpha"], "beta": 1.0 - _C1["alpha"],
        "normalization": "min-max with A_min, A_max, V_min, V_max solved over the simplex by the same multi-start",
        "solver": "scipy SLSQP, ftol %g, maxiter %d, exact gradients (dA/dw = A_m mu + A_s sigma, dV/dw = 2 s sigma)" % (
            _C1["slsqp_ftol"], _C1["slsqp_maxiter"]),
        "starts": ("%d Dirichlet(1) starts + the vertices + the centroid (31 starts for a group of 6, 32 for 7), "
                   "the same start set for the four bound programs and the scalarized program" % _C1["n_random_starts"]),
        "n_random_starts": _C1["n_random_starts"], "slsqp_ftol": _C1["slsqp_ftol"],
        "kmeans_restarts": _C1["kmeans_restarts"],
        "max_abs_indicator_offset": float(np.max(np.abs(_c1_ind_off))),
        "agreement_rule": "a start agrees when its terminal objective lies within %g of the retained value" % _C1["tie_tol"],
        "n_programs": len(_c1_progs),
        "all_starts_agree_in_every_program": int(_c1_all_agree),
        "min_starts_at_optimum": int(min(_c1_agree)),
        "min_bound_starts_at_optimum": int(_c1_bound_agree_min),
        "max_group_opinion_spread_among_agreeing_starts": _c1_ms_spread,
        "group_opinion_agreement_bound": _C1["opinion_tol"],
        "n_scalarized_runs_not_converged": int(sum(p.n_starts - p.n_success for p in _c1_progs.values())),
        "n_group_opinions_equal_to_one_member": int(_c1_n_member),
        "member_coincidence_rule": ("a group opinion coincides with a member when |m - mu_i| + |s - sigma_i| < %g"
                                    % _C1["member_tol"]),
        "member_coincidence_tol": _C1["member_tol"],
        "n_coinciding_members_strictly_inside_group_hull": int(_c1_member_interior),
        "n_coinciding_opinions_with_recorded_max_weight_one": int(_c1_member_wmax_one),
        "member_coincidence_note": ("where the group opinion coincides with a member, that member's overlap equals "
                                    "one and has a kink there; the member's (mu, sigma) point is strictly inside the "
                                    "convex hull of the group's points in all but the recorded exceptions, so the "
                                    "coincidence is not a vertex of the feasible (m, s) set, and the recorded w is "
                                    "one of many optimal weight vectors (max weight within member_coincidence_tol "
                                    "of one in the recorded count)"),
        "weight_uniqueness_note": ("the program depends on w only through (w.mu, w.sigma), so the optimal weight "
                                   "vector need not be unique; in every program all agreeing starts return the "
                                   "same group opinion (spread recorded in "
                                   "max_group_opinion_spread_among_agreeing_starts), and the sheet records the "
                                   "best start's w"),
        "kmeans": ("Lloyd's algorithm on standardized attribute columns (population sd), %d seeded restarts, each "
                   "initialized at %d distinct rows and run for at most %d Lloyd iterations, stopping when the "
                   "centroids are unchanged (numpy allclose, default tolerances rtol 1e-5, atol 1e-8), an empty "
                   "cluster keeping its previous centroid; the restart of minimum inertia is kept, a later restart "
                   "replacing an earlier one only when it lowers the inertia by more than 1e-12") % (
            _C1["kmeans_restarts"], _c1_K, _C1["kmeans_maxiter"]),
        "kmeans_maxiter": _C1["kmeans_maxiter"],
        "kmeans_convergence": "numpy.allclose(new centroids, old centroids) with default tolerances",
        "attribute_matrix": ("stated design input, not drawn (sheet expert_attributes, rows e1..e20); the synthetic "
                             "case clusters the four raw attributes directly rather than through the published "
                             "frequency ratio (projects per year of BIM experience)"),
        "kmeans_best_restart": _c1_best_restart, "kmeans_inertia": _c1_inertia,
        "kmeans_labeling": "clusters numbered in order of first appearance along e1..e20",
        "topsis": ("vector normalization over the three centroids, criterion weights %s, all benefit criteria, "
                   "ideal and anti-ideal = best and worst centroid value per criterion, Omega = closeness / sum") % (
            ",".join("%.2f" % v for v in _C1["topsis_weights"])),
        "attribute_columns": ",".join(_C1_ATTR_NAMES),
        "inter_group": "quantile averaging under Omega (normal closed form, Prop. qa-properties(iii))",
        "indicator_weights": "best-worst method, linear model (Rezaei 2016), one problem per dimension (sheet bwm)",
        "dimension_weights": "best-worst method, linear model (Rezaei 2016), best Technology, worst Policy (sheet bwm)",
        "consistency": ("input-based consistency ratio CR^I of Liang, Brunelli and Rezaei (2020), eqs. (6)-(7), "
                        "against their Table 3 thresholds (sheets bwm_summary, bwm_cri_thresholds)"),
        "max_CR_I": float(max(r["CR_I"] for r in _c1_bwm_summary)),
        "global_mu": _c1_glob_mu, "global_sigma": _c1_glob_sd, "global_band": _c1_band(_c1_glob_mu),
        # for comparison only: the global under the published application's equal importance
        # of the three dimensions (the case itself uses the best-worst dimension weights)
        "global_mu_equal_dimension_weights": _c1_glob_eq_mu,
        "global_sigma_equal_dimension_weights": _c1_glob_eq_sd,
        "global_equal_dimension_weights_printed": _c1_printed(_c1_glob_eq_mu, _c1_glob_eq_sd),
        "global_equal_dimension_weights_note": ("QA of the three dimension collectives with weights 1/3 each, "
                                                "the published application's rule; not used by the case"),
        "q4_linear_pool_sd": _c1_q4_lp_sd,
        # the margins by which the dimension order holds (sec:bim-results, robustness)
        "margin_organization_over_policy": float(_c1_DM[1] - _c1_DM[2]),
        "margin_policy_over_technology": float(_c1_DM[2] - _c1_DM[0]),
        "indicator": "q4 -- Real-time Data (a technology indicator)",
        "x_label_groupagg": "maturity level, Real-time Data indicator ($q_4$)",
        "x_label_dimensions": "BIM maturity level",
        "y_label": "probability density",
        # the two level borders drawn on the dimension figure
        "band_lo": 5.0, "band_hi": 6.0, "band_name": "Defined",
        "grid_points": 700,
    }),
    # COMPUTED: the k-means groups, their sizes and attribute ranges, the
    # TOPSIS credibility weights, and each group's solved opinion on q4
    "groups": _c1_groups_rows,
    # COMPUTED: dimension collectives (BWM-weighted QA of the indicator
    # collectives) and the BWM dimension weights
    "dimensions": _c1_dims_rows,
    # GENERATED: expert e7's reported profile over the ten levels for q4 (tab:bim-elicit)
    "elicitation_q4": _c1_elicit_rows,
    # COMPUTED: the two quantile averages the text reports and the tab:bim-elicit fit.
    "derived": [
        {"quantity": "q4 collective (inter-group QA under Omega)", "mu": _c1_q4_mu, "sigma": _c1_q4_sd,
         "printed": _c1_printed(_c1_q4_mu, _c1_q4_sd),
         "rule": "mu = sum Omega_k mu_k, sigma = sum Omega_k sigma_k",
         "shown_in": "Figure 6.1 legend, text sec:bim-hierarchy"},
        {"quantity": "global maturity (QA of dimensions under BWM weights)", "mu": _c1_glob_mu, "sigma": _c1_glob_sd,
         "printed": _c1_printed(_c1_glob_mu, _c1_glob_sd),
         "rule": "mu = sum w_d mu_d, sigma = sum w_d sigma_d",
         "shown_in": "Figure 6.2 legend, Table 6.8 global row, text sec:bim-results"},
        {"quantity": "equal-dimension-weight global (the published rule, for comparison)",
         "mu": _c1_glob_eq_mu, "sigma": _c1_glob_eq_sd,
         "printed": _c1_printed(_c1_glob_eq_mu, _c1_glob_eq_sd),
         "rule": "mu = mean of the dimension mu_d, sigma = mean of the dimension sigma_d (weights 1/3)",
         "shown_in": "comparison only (sec:bim-results empirical box, if printed)"},
        {"quantity": "tab:bim-elicit normal fit of the q4 profile", "mu": _c1_fit_mu, "sigma": _c1_fit_sd,
         "printed": _c1_printed(_c1_fit_mu, _c1_fit_sd),
         "rule": "moment matching at band midpoints",
         "shown_in": "Table 6.6 fit row (label and band probabilities), text sec:bim-elicit"},
    ],
    "expert_attributes": _c1_expert_rows,
    "panel": _c1_panel_rows,
    "intra_group_solutions": _c1_sol_rows,
    "indicator_collectives": _c1_ind_rows,
    "bwm": _c1_bwm_rows,
    "bwm_summary": _c1_bwm_summary,
    "bwm_cri_thresholds": [dict({"scale_a_BW": s}, **{"n=%d" % (n + 3): v for n, v in enumerate(row)})
                           for s, row in sorted(_C1_CRI_THRESH.items())],
})

# =====================================================================
# CASE STUDY II (Section 6.3, digital-transformation barriers in SMCEs)
#   -> Figure fig6-5_dt_ranking14 (fig:dt-ranking14)
#   -> Figure fig6-6_dt_crossswf  (fig:dt-crossswf)
#   -> Figure fig6-7_dt_triad     (fig:dt-triad)
#   -> Table  tab:dt-swf-results, tab:dt-elicit
# GENERATED AND SOLVED HERE (since 2026-09-14).  A 30 x 14 panel of
# truncated-Weibull experts is generated under seed 202603 by the design
# chapter6.tex states (target importances calibrated to stated design
# means, leniency, background bias, noise, five-level profile, grouped-MLE
# Weibull fit), and the welfare-confidence program eq:dt-welfare-conf is
# solved 70 times (14 barriers x 5 equity indices) by multi-start SLSQP:
#     max_w  W~_s(w) - V~(w)   on the simplex,
#     W_s(w) = Ubar(w) (1 - I_s(U(w)))   (Sen's abbreviated welfare form on
#     the exp-normalized Fehr-Schmidt utilities of the overlaps, eq:exp-norm),
#     V(w) = variance of the quantile-averaged aggregate,
#     ~ = min-max normalization over the simplex (bounds solved).
# Since 2026-09-29 (R24) the welfare bounds and the program are searched
# over every vertex, a grid on every two-expert edge refined by zooming on
# the best edges, a random sample of points on the faces of three to six
# experts, and multi-start SLSQP (interior starts plus the best edge points);
# the welfare bounds are refined until no evaluated solution of the program
# lies outside them, and every retained solution is checked against all of
# these candidates and against every pairwise weight transfer.  Every start
# is recorded (sheet 'starts'); the programs behind the adjacent pairs that
# tie at two decimals and the flat programs are solved again on a doubled
# discretization (sheets 'fine_check', 'fine_weights'); and every number the
# text of Section 6.3 prints is written to the sheet 'text_numbers'.
# Every printed number of the case (tab:dt-swf-results, the ranking, the
# triad, the spreads, tab:dt-elicit's profile and fit) is read from this workbook.
# =====================================================================
import math
import time
import numpy as np
from scipy.optimize import minimize as _sp_minimize

print("\n-- generating and solving Case II (fig6-5_dt_case.xlsx) ...")
_dt_t0 = time.time()

DT_SEED = 202603
DT_N, DT_M = 30, 14
DT_HI = 5.0                     # importance domain [0, 5], five unit bands
DT_RHO, DT_THETA = 0.4, 0.2     # Fehr-Schmidt jealousy / sympathy (Case I values)
DT_LENIENCY_SD = 0.15           # a_i ~ N(0, 0.15^2)
DT_NOISE_SD = 0.25              # eps_ik ~ N(0, 0.25^2)
DT_BIAS = 0.30                  # background bias on the expert's own TOE domain
DT_PROFILE_SD = 0.9             # Gaussian kernel sd of the five-level profile
DT_CALIB_TOL = 2e-4             # |equal-weight QA mean - design mean| target
DT_CALIB_MAXIT = 10
DT_PGRID = 2000                 # midpoint probability grid for quantile averaging
DT_XBINS = 600                  # x bins on [0, 5] for the overlaps
DT_STARTS_BOUNDS = 6            # Dirichlet(1) starts (plus barycentre) for W_min, W_max
DT_STARTS_PROGRAM = 12          # Dirichlet(1) starts (plus barycentre) for the program
DT_FTOL, DT_MAXITER, DT_FDH = 1e-10, 400, 1e-7
DT_ACTIVE_TOL = 1e-4            # w_i > tol counts as an active expert
DT_BOUND_REFINE_MAX = 6         # max refinements of a welfare bound from the program's solutions
DT_BOUND_TOL = 1e-9             # tolerance for W(w) within [W_min, W_max] at every retained candidate
# candidate search and verification (R24): every two-expert edge on a grid
# (its endpoints are the 30 vertices), a zoom on the best edges, a random
# face sample, SLSQP from the best edge points, and a pairwise-transfer check
DT_EDGE_GRID = 101              # grid points per edge, t = 0, 0.01, ..., 1
DT_EDGE_TOP = 8                 # best edges refined by zooming, per search
DT_EDGE_STARTS = 3              # best zoomed edge points used as SLSQP starts
DT_ZOOM_PTS = 21                # points per zoom round on an edge
DT_ZOOM_WIDTH = 1e-10           # zooming stops at this bracket width
DT_FACE_SAMPLE = 3000           # random points on faces of 3..6 experts per barrier
DT_FACE_SIZES = (3, 4, 5, 6)    # face sizes, drawn uniformly
DT_TRANSFER_STEPS = (1e-3, 1e-4, 1e-5)   # pairwise transfers checked at the solution
DT_TRANSFER_TOL = 1e-9          # a transfer gaining more than this triggers a polish
DT_POLISH_MAX = 50              # max transfer-polish steps
DT_AGREE_TOL = 1e-6             # a start reaches the retained value within this
DT_NEAR_TOL = 1e-3              # a start is near-optimal within this (flat-landscape diagnostic)
DT_FLAT_TOL = 0.005             # a program is flat if a near-optimal start moves its mean by more
DT_FINE_FACTOR = 2              # doubled-discretization check: 2 x DT_PGRID probabilities, 2 x DT_XBINS bins
DT_ELICIT_BARRIER = "B1"        # tab:dt-elicit: the barrier illustrated
DT_TRIAD_EQ = "EQ4"             # objective whose collective densities Figure 6.7 shows
_dt_pg = (np.arange(DT_PGRID) + 0.5) / DT_PGRID
_dt_xg = np.linspace(0.0, DT_HI, DT_XBINS + 1)

# the 14 barriers (tab:dt-barriers) and the design mean of each: the value
# the equal-weight quantile-averaged mean of the generated panel is
# calibrated to, set in the order of the complete final ranking of
# Chen et al. (2024c), Table 11 (B14, B6, B1 at the head, B11, B13, B9 at the foot)
_DT_BARRIERS = [
    ("B1",  "Inadequate infrastructure",        "T", 3.86),
    ("B2",  "Poor system integration",          "T", 3.02),
    ("B3",  "Low level of digital technology",  "T", 3.13),
    ("B4",  "Data security risks",              "T", 3.56),
    ("B5",  "Insufficient funding",             "O", 3.67),
    ("B6",  "Lack of strategic planning",       "O", 4.00),
    ("B7",  "Lack of leadership support",       "O", 3.73),
    ("B8",  "Lack of skilled employees",        "O", 3.43),
    ("B9",  "Difficulty in technical training", "O", 2.29),
    ("B10", "Employee resistance",              "O", 3.79),
    ("B11", "Industry resistance",              "E", 2.83),
    ("B12", "Customer diversity",               "E", 3.29),
    ("B13", "Limited financial support",        "E", 2.56),
    ("B14", "Lack of standards & regulations",  "E", 4.10),
]
_DT_IDS = [b[0] for b in _DT_BARRIERS]
_DT_DIM = [b[2] for b in _DT_BARRIERS]
_DT_DESIGN = np.array([b[3] for b in _DT_BARRIERS])
_DT_EQ = ["EQ1", "EQ2", "EQ3", "EQ4", "EQ5"]
# experts 1-15 technologists (bias on the technological barriers B1-B4),
# 16-30 managers (bias on the organizational barriers B5-B10)
_DT_BACKGROUND = ["technologist"] * 15 + ["manager"] * 15
_DT_BG_DOMAIN = {"technologist": "T", "manager": "O"}
# tab:dt-elicit shows a member of the generated panel (selected below by a
# stated rule), not a typed profile


# ------------------------------------------------ truncated Weibull (eq:dt-weibull-trunc)
def _dt_wb_H(x, k, lam):
    return 1.0 - np.exp(-(np.clip(np.asarray(x, float), 0.0, None) / lam) ** k)


def _dt_trunc_quantile(p, k, lam):
    """Inverse of H(x)/H(5) on (0,1): x = lam (-ln(1 - p H(5)))^(1/k)."""
    return lam * (-np.log(1.0 - p * _dt_wb_H(DT_HI, k, lam))) ** (1.0 / k)


def _dt_trunc_cdf(x, k, lam):
    return np.clip(_dt_wb_H(x, k, lam) / _dt_wb_H(DT_HI, k, lam), 0.0, 1.0)


def _dt_fit_trunc_weibull(prob):
    """Grouped maximum likelihood of the truncated Weibull on the five unit
    bands l_t = [t-1, t):  max_{k,lam} sum_t p_t log([H(t)-H(t-1)] / H(5)).
    Nine starts, L-BFGS-B in log-parameters, best kept.
    Returns (kappa, lambda, mean of the fitted truncated Weibull, H(5))."""
    prob = np.asarray(prob, float)
    edges = np.arange(0.0, DT_HI + 1.0)

    def nll(th):
        k, lam = np.exp(th)
        H = _dt_wb_H(edges, k, lam)
        with np.errstate(divide="ignore", invalid="ignore"):
            q = np.clip(np.diff(H) / H[-1], 1e-300, None)
        return -float(np.dot(prob, np.nan_to_num(np.log(q), nan=-700.0)))

    best = None
    for k0 in (1.5, 3.0, 6.0):
        for l0 in (2.0, 3.5, 4.5):
            r = _sp_minimize(nll, np.log([k0, l0]), method="L-BFGS-B",
                             bounds=[(np.log(0.3), np.log(60.0)), (np.log(0.2), np.log(40.0))])
            if best is None or r.fun < best.fun:
                best = r
    k, lam = np.exp(best.x)
    mean = float(np.mean(_dt_trunc_quantile(_dt_pg, k, lam)))
    return float(k), float(lam), mean, float(_dt_wb_H(DT_HI, k, lam))


def _dt_profile(rating):
    """Five-level probability vector concentrated around the rating x_ik
    (the book's symbol; r_ik is the background bias):
    p_t proportional to exp(-(m_t - x_ik)^2 / (2 sd^2)), m_t = t - 1/2."""
    mids = np.arange(5) + 0.5
    p = np.exp(-(mids[None, :] - np.asarray(rating, float)[:, None]) ** 2 / (2.0 * DT_PROFILE_SD ** 2))
    return p / p.sum(axis=1, keepdims=True)


# ------------------------------------------------------------ the panel of one barrier
class _DtPanel:
    """N truncated-Weibull experts; quantile averaging on the p grid
    (def:qa), overlaps as shared bin mass on the x grid (def:consensus),
    aggregate variance as a quadratic form (def:confidence)."""

    def __init__(self, ks, lams, pgrid=DT_PGRID, xbins=DT_XBINS):
        # pgrid, xbins: the discretization (the case's DT_PGRID, DT_XBINS by
        # default; the doubled-discretization check passes twice as many)
        self.k = np.asarray(ks, float)
        self.lam = np.asarray(lams, float)
        self.N = self.k.size
        self.pg = _dt_pg if pgrid == DT_PGRID else (np.arange(pgrid) + 0.5) / pgrid
        self.xg = _dt_xg if xbins == DT_XBINS else np.linspace(0.0, DT_HI, xbins + 1)
        self.Qi = np.vstack([_dt_trunc_quantile(self.pg, k, l) for k, l in zip(self.k, self.lam)])
        self.mu_i = self.Qi.mean(axis=1)
        Fi = np.vstack([_dt_trunc_cdf(self.xg, k, l) for k, l in zip(self.k, self.lam)])
        self.dFi = np.diff(Fi, axis=1)                       # N x X bin masses
        Qc = self.Qi - self.mu_i[:, None]
        self.Sig = Qc @ Qc.T / pgrid                         # V(w) = w' Sig w
        self.Vi = np.diag(self.Sig).copy()

    def quantile(self, w):
        return np.asarray(w, float) @ self.Qi

    def overlaps(self, W):
        """Batched overlaps A (B x N): F_QA is the inverse of the aggregate
        quantile table; A_i = sum over bins of min{dF_QA, dF_i}."""
        W = np.atleast_2d(W)
        Q = W @ self.Qi
        F = np.empty((W.shape[0], self.xg.size))
        for b in range(W.shape[0]):
            F[b] = np.interp(self.xg, Q[b], self.pg, left=0.0, right=1.0)
        F[:, 0] = 0.0
        F[:, -1] = 1.0
        dF = np.diff(F, axis=1)
        return np.minimum(dF[:, None, :], self.dFi[None, :, :]).sum(axis=2)

    def var(self, W):
        W = np.atleast_2d(W)
        return np.einsum("bi,ij,bj->b", W, self.Sig, W)

    def mean(self, w):
        return float(np.asarray(w, float) @ self.mu_i)

    def density(self, w, x):
        """f_QA on the points x by the inverse-function rule: along the
        aggregate quantile function, f_QA(Q(p)) = 1 / Q'(p), with Q' the
        central difference on the p grid; zero below Q(p_1), and the value at
        Q(p_P) carried to the domain boundary 5 (the truncated density is
        positive there)."""
        Q = self.quantile(w)
        f_on_Q = 1.0 / np.gradient(Q, self.pg)
        return np.where((x >= Q[0]) & (x <= DT_HI), np.interp(x, Q, f_on_Q, right=f_on_Q[-1]), 0.0)


# ------------------------------------------------ behavioral layer and welfare
def _dt_fs_expnorm(A, rho, theta):
    """Batched Fehr-Schmidt utilities of the overlaps (def:ifcu) mapped to
    [0,1] by the exponential normalization eq:exp-norm.  A is B x N."""
    A = np.atleast_2d(A)
    n = A.shape[1]
    M = A[:, None, :] - A[:, :, None]                        # M[b,i,j] = A_j - A_i
    D = np.clip(M, 0.0, None).sum(axis=2) / (n - 1)
    S = np.clip(-M, 0.0, None).sum(axis=2) / (n - 1)
    Ufs = A - rho * D - theta * S
    pos = A > 0
    with np.errstate(divide="ignore", invalid="ignore"):
        r = np.where(pos, Ufs / np.where(pos, A, 1.0), 0.0)
    U = np.where(pos & (Ufs >= 0) & (Ufs <= A),
                 (np.exp(np.clip(r, 0.0, 1.0)) - 1.0) / (math.e - 1.0), 0.0)
    return np.where(pos & (Ufs > A), 1.0, U)


def _dt_ineq(U, s):
    """Batched inequality index I_s of tab:swf on the rows of U (relative
    range, relative mean deviation, coefficient of variation, Gini, Hoover);
    zero on an all-zero row by the book's convention."""
    U = np.atleast_2d(U)
    n = U.shape[1]
    m = U.mean(axis=1)
    safe = np.where(m > 0, m, 1.0)
    if s == 1:
        val = (U.max(axis=1) - U.min(axis=1)) / safe
    elif s == 2:
        val = np.abs(U - m[:, None]).sum(axis=1) / (n * safe)
    elif s == 3:
        val = np.sqrt(((U - m[:, None]) ** 2).mean(axis=1)) / safe
    elif s == 4:
        val = np.abs(U[:, :, None] - U[:, None, :]).sum(axis=(1, 2)) / (2.0 * n * n * safe)
    elif s == 5:
        val = np.abs(U - m[:, None]).sum(axis=1) / (2.0 * n * safe)
    else:
        raise ValueError(s)
    return np.where(m > 0, val, 0.0)


def _dt_welfare(panel, W, s):
    """Complete welfare in Sen's abbreviated form, W_s(w) = Ubar(w) (1 - I_s(U(w)))."""
    U = _dt_fs_expnorm(panel.overlaps(W), DT_RHO, DT_THETA)
    return U.mean(axis=1) * (1.0 - _dt_ineq(U, s))


def _dt_run(fun, n, w0, maximize=True):
    """One SLSQP run on the simplex (sum w = 1, 0 <= w_i <= 1) from w0 with a
    batched forward-difference gradient (step DT_FDH).  fun maps a (B x n)
    batch to (B,).  Returns (w, fun(w), SLSQP exit status), w projected back
    onto the simplex (clip at zero, renormalize) before it is evaluated."""
    sign = -1.0 if maximize else 1.0
    eye = np.eye(n)

    def f(w):
        return sign * float(fun(w[None, :])[0])

    def jac(w):
        v = fun(np.vstack([w[None, :], w[None, :] + DT_FDH * eye]))
        return sign * (v[1:] - v[0]) / DT_FDH

    cons = ({"type": "eq", "fun": lambda w: np.sum(w) - 1.0, "jac": lambda w: np.ones_like(w)},)
    r = _sp_minimize(f, w0, jac=jac, method="SLSQP", bounds=[(0.0, 1.0)] * n, constraints=cons,
                     options={"ftol": DT_FTOL, "maxiter": DT_MAXITER})
    w = np.clip(r.x, 0.0, None)
    w = w / w.sum()
    return w, float(fun(w[None, :])[0]), int(r.status)


def _dt_slsqp(fun, n, starts, maximize=True):
    """Multi-start SLSQP: the best local optimum over the starts is retained."""
    best = None
    for w0 in starts:
        w, v, _ = _dt_run(fun, n, w0, maximize)
        if best is None or (v > best[1] if maximize else v < best[1]):
            best = (w, v)
    return best


# ---------------------------------------- candidate search on the simplex (R24)
# Every solution of the program found so far puts its weight on two experts,
# and the welfare maximum is found on the same edges, so the search covers
# the edges of the simplex explicitly: a grid on each of the 435 two-expert
# edges (its endpoints are the 30 vertices), a zoom on the best edges, and
# SLSQP from the best zoomed points, next to the interior starts (the
# barycentre and the Dirichlet(1) draws).  A random
# sample on the faces of three to six experts checks the rest of the simplex.
_DT_PAIRS = [(i, j) for i in range(DT_N) for j in range(i + 1, DT_N)]
_dt_tgrid = np.linspace(0.0, 1.0, DT_EDGE_GRID)


def _dt_edge_points(i, j, t):
    """Points t e_i + (1 - t) e_j of the edge (i, j), one row per t."""
    t = np.atleast_1d(np.asarray(t, float))
    W = np.zeros((t.size, DT_N))
    W[:, i] += t
    W[:, j] += 1.0 - t
    return W


def _dt_edge_grid():
    W = np.zeros((len(_DT_PAIRS) * DT_EDGE_GRID, DT_N))
    for q, (i, j) in enumerate(_DT_PAIRS):
        W[q * DT_EDGE_GRID:(q + 1) * DT_EDGE_GRID] = _dt_edge_points(i, j, _dt_tgrid)
    return W


def _dt_face_sample(rng):
    """DT_FACE_SAMPLE points, each on a face of m experts (m uniform on
    DT_FACE_SIZES, the experts uniform without replacement, the weights
    Dirichlet(1) on them); draws in that order, point by point."""
    F = np.zeros((DT_FACE_SAMPLE, DT_N))
    for r in range(DT_FACE_SAMPLE):
        m = int(rng.choice(DT_FACE_SIZES))
        idx = rng.choice(DT_N, size=m, replace=False)
        F[r, idx] = rng.dirichlet(np.ones(m))
    return F


def _dt_utilities(panel, W, chunk=2000):
    """Normalized Fehr-Schmidt utilities U (B x N) of a batch, in chunks."""
    return np.vstack([_dt_fs_expnorm(panel.overlaps(W[c:c + chunk]), DT_RHO, DT_THETA)
                      for c in range(0, len(W), chunk)])


def _dt_welfare_from_U(U, s, chunk=4000):
    out = np.empty(len(U))
    for c in range(0, len(U), chunk):
        u = U[c:c + chunk]
        out[c:c + chunk] = u.mean(axis=1) * (1.0 - _dt_ineq(u, s))
    return out


def _dt_top_edges(grid_vals, maximize, n):
    """The n edges whose best grid value is best, each with its best t."""
    v = grid_vals.reshape(len(_DT_PAIRS), DT_EDGE_GRID)
    if not maximize:
        v = -v
    g = v.argmax(axis=1)
    order = np.argsort(-v[np.arange(len(_DT_PAIRS)), g], kind="stable")[:n]
    return [(_DT_PAIRS[q], float(_dt_tgrid[g[q]])) for q in order]


def _dt_edge_zoom(fun, i, j, t0, maximize=True):
    """Refine fun along the edge (i, j) from the grid point t0: DT_ZOOM_PTS
    points on the bracket of the neighbouring grid points, then on the
    bracket of the best point's neighbours, until the bracket is narrower
    than DT_ZOOM_WIDTH.  Returns (t, fun at t)."""
    sign = 1.0 if maximize else -1.0
    h = 1.0 / (DT_EDGE_GRID - 1)
    lo, hi = max(0.0, t0 - h), min(1.0, t0 + h)
    bt, bv = t0, sign * float(fun(_dt_edge_points(i, j, t0))[0])
    while hi - lo > DT_ZOOM_WIDTH:
        ts = np.linspace(lo, hi, DT_ZOOM_PTS)
        v = sign * fun(_dt_edge_points(i, j, ts))
        g = int(np.argmax(v))
        if v[g] > bv:
            bt, bv = float(ts[g]), float(v[g])
        step = ts[1] - ts[0]
        lo, hi = max(0.0, bt - step), min(1.0, bt + step)
    return bt, sign * bv


def _dt_transfers(w):
    """Every pairwise transfer of DT_TRANSFER_STEPS (capped at the giver's
    weight) from an active expert i to another expert j, one row each."""
    rows = []
    for i in np.where(w > 0.0)[0]:
        for j in range(DT_N):
            if j != i:
                for e in DT_TRANSFER_STEPS:
                    d = w.copy()
                    ee = min(e, d[i])
                    d[i] -= ee
                    d[j] += ee
                    rows.append(d)
    return np.vstack(rows)


# ------------------------------------------------------------ generate the panel
# Draw order under default_rng(DT_SEED): (1) the 30 leniencies a_i, (2) the
# 30 x 14 noise matrix eps_ik (expert-major).  Nothing else is drawn for the
# panel; the target importances tau_k are then calibrated deterministically.
_dt_rng = np.random.default_rng(DT_SEED)
_dt_len = _dt_rng.normal(0.0, DT_LENIENCY_SD, DT_N)
_dt_eps = _dt_rng.normal(0.0, DT_NOISE_SD, (DT_N, DT_M))
_dt_bias = np.array([[DT_BIAS if _DT_BG_DOMAIN[_DT_BACKGROUND[i]] == _DT_DIM[k] else 0.0
                      for k in range(DT_M)] for i in range(DT_N)])


def _dt_fit_column(tau_k, k):
    # the book's x_ik = clip(tau_k + a_i + r_ik + eps_ik, 0, 5), r_ik the background bias
    x = np.clip(tau_k + _dt_len + _dt_bias[:, k] + _dt_eps[:, k], 0.0, DT_HI)
    prof = _dt_profile(x)
    fits = np.array([_dt_fit_trunc_weibull(p) for p in prof])
    return x, prof, fits[:, 0], fits[:, 1], fits[:, 2], fits[:, 3]


_dt_tau = np.empty(DT_M)
_dt_calib_it = np.zeros(DT_M, int)
_dt_rating = np.empty((DT_N, DT_M))
_dt_prof = np.empty((DT_N, DT_M, 5))
_dt_kap = np.empty((DT_N, DT_M))
_dt_lam = np.empty((DT_N, DT_M))
_dt_wmean = np.empty((DT_N, DT_M))
_dt_H5 = np.empty((DT_N, DT_M))
for _k in range(DT_M):
    # secant iteration on tau_k: the equal-weight quantile-averaged mean of
    # the fitted panel (the average of the fitted expert means, QA means
    # being linear) is driven to the design mean
    t_prev, m_prev = None, None
    t_cur = _DT_DESIGN[_k]
    for _it in range(DT_CALIB_MAXIT):
        r, prof, kk, ll, wm, h5 = _dt_fit_column(t_cur, _k)
        m_cur = float(wm.mean())
        _dt_calib_it[_k] = _it + 1
        if abs(m_cur - _DT_DESIGN[_k]) < DT_CALIB_TOL:
            break
        if t_prev is None or abs(m_cur - m_prev) < 1e-9:
            t_new = t_cur + (_DT_DESIGN[_k] - m_cur)
        else:
            slope = (m_cur - m_prev) / (t_cur - t_prev)
            t_new = t_cur + (_DT_DESIGN[_k] - m_cur) / max(slope, 0.3)
        t_prev, m_prev, t_cur = t_cur, m_cur, t_new
    assert abs(m_cur - _DT_DESIGN[_k]) < 5 * DT_CALIB_TOL, \
        "Case II calibration of %s did not converge (%.5f vs %.2f)" % (_DT_IDS[_k], m_cur, _DT_DESIGN[_k])
    _dt_tau[_k] = t_cur
    _dt_rating[:, _k], _dt_prof[:, _k], _dt_kap[:, _k], _dt_lam[:, _k], _dt_wmean[:, _k], _dt_H5[:, _k] = \
        r, prof, kk, ll, wm, h5
print("   panel generated and fitted: shapes %.2f..%.2f, scales %.2f..%.2f, %.0f s"
      % (_dt_kap.min(), _dt_kap.max(), _dt_lam.min(), _dt_lam.max(), time.time() - _dt_t0))

# tab:dt-elicit shows one member of the generated panel, chosen by a stated
# rule: on barrier DT_ELICIT_BARRIER, the expert whose fitted mean lies
# closest to the barrier's equal-weight collective mean (the average of the
# fitted expert means); its profile and fit are the panel's own values.
_dt_el_k = _DT_IDS.index(DT_ELICIT_BARRIER)
_dt_el_i = int(np.argmin(np.abs(_dt_wmean[:, _dt_el_k] - _dt_wmean[:, _dt_el_k].mean())))
_dt_el_prof = _dt_prof[_dt_el_i, _dt_el_k].copy()
_dt_el_fit = (float(_dt_kap[_dt_el_i, _dt_el_k]), float(_dt_lam[_dt_el_i, _dt_el_k]),
              float(_dt_wmean[_dt_el_i, _dt_el_k]), float(_dt_H5[_dt_el_i, _dt_el_k]))

# ------------------------------------------------------------------ solve
# Start order under a fresh default_rng(DT_SEED): for each barrier in order
# B1..B14, DT_STARTS_PROGRAM Dirichlet(1) draws (the same starts as before
# R24).  The face sample (_dt_face_sample) is drawn from a second fresh
# default_rng(DT_SEED), barrier by barrier in the same order, so that it
# leaves the starts unchanged.  The bound searches use the barycentre and
# the first DT_STARTS_BOUNDS Dirichlet starts, the scalarized program the
# barycentre and all of them; the vertices, the edge grid, and the zoomed
# edge points involve no draws.
_dt_rng_starts = np.random.default_rng(DT_SEED)
_dt_rng_face = np.random.default_rng(DT_SEED)
_dt_solves = []
_dt_starts_rec = []
_dt_W = {}
_dt_A = {}
_dt_U = {}
_dt_panels = {}
_dt_Aeq = {}
_dt_faces = {}
_dt_bary = np.full(DT_N, 1.0 / DT_N)
_dt_Wedge = _dt_edge_grid()
_dt_eye = np.eye(DT_N)


def _dt_solve_cell(panel, s, cl, rs, Vmin, Vmax):
    """The welfare bounds and the program eq:dt-welfare-conf of one barrier
    under the inequality index I_s.

    W_min: the smallest welfare over the 30 vertices, the edge grid, and the
      face sample, and of SLSQP minimizations from the best vertex, the best
      face point, the barycentre, and the first DT_STARTS_BOUNDS Dirichlet starts.
    W_max: the largest welfare over the edge grid, the face sample, a zoom on
      the DT_EDGE_TOP best edges, and SLSQP maximizations from the
      DT_EDGE_STARTS best zoomed points, the best face point, the barycentre,
      and the first DT_STARTS_BOUNDS Dirichlet starts.
    Program: SLSQP from the barycentre and every Dirichlet start (the interior
      starts), the DT_EDGE_STARTS best zoomed edge points of the scalarized
      objective and SLSQP from each, and SLSQP from the best face point; the
      best candidate is retained, then
      checked against every pairwise transfer and polished while one gains.
    Bound refinement: if the welfare at any candidate of the program leaves
      [W_min, W_max], the bound is moved to it (and further by a bound solve
      started there) and the program is solved again, until none does."""

    def Wfun(Wb):
        return _dt_welfare(panel, Wb, s)

    We = _dt_welfare_from_U(cl["U_edge"], s)
    Wf = _dt_welfare_from_U(cl["U_face"], s)
    Wv = _dt_welfare_from_U(cl["U_vert"], s)
    cand = [float(Wv.min()), float(We.min()), float(Wf.min())]
    for w0 in [_dt_eye[int(np.argmin(Wv))], cl["W_face"][int(np.argmin(Wf))], _dt_bary] + rs[:DT_STARTS_BOUNDS]:
        cand.append(_dt_run(Wfun, DT_N, w0, maximize=False)[1])
    Wmin = min(cand)
    cand = [float(We.max()), float(Wf.max())]
    zs = []
    for (i, j), t in _dt_top_edges(We, True, DT_EDGE_TOP):
        tz, vz = _dt_edge_zoom(Wfun, i, j, t, True)
        zs.append((vz, _dt_edge_points(i, j, tz)[0]))
        cand.append(vz)
    zs.sort(key=lambda z: -z[0])
    for w0 in ([z[1] for z in zs[:DT_EDGE_STARTS]] + [cl["W_face"][int(np.argmax(Wf))], _dt_bary]
               + rs[:DT_STARTS_BOUNDS]):
        cand.append(_dt_run(Wfun, DT_N, w0, maximize=True)[1])
    Wmax = max(cand)
    out = {"W_min_vertex": float(Wv.min()), "W_min_search": Wmin, "W_max_search": Wmax,
           "W_max_cloud": float(max(We.max(), Wf.max())), "W_min_cloud": float(min(We.min(), Wf.min(), Wv.min()))}
    dV = Vmax - Vmin
    for ref in range(DT_BOUND_REFINE_MAX + 1):
        dW = Wmax - Wmin
        assert dW > 1e-9 and dV > 1e-9, "degenerate min-max range"

        def scal(Wb, Wmin=Wmin, dW=dW):
            return (Wfun(Wb) - Wmin) / dW - (panel.var(Wb) - Vmin) / dV

        Se = (We - Wmin) / dW - (cl["V_edge"] - Vmin) / dV
        Sf = (Wf - Wmin) / dW - (cl["V_face"] - Vmin) / dV
        runs = []
        for n, w0 in enumerate([_dt_bary] + rs):
            w, v, st = _dt_run(scal, DT_N, w0, True)
            runs.append({"start": "barycentre" if n == 0 else "dirichlet_%d" % n, "kind": "interior",
                         "w": w, "objective": v, "status": st})
        ez = []
        for (i, j), t in _dt_top_edges(Se, True, DT_EDGE_TOP):
            tz, vz = _dt_edge_zoom(scal, i, j, t, True)
            ez.append((vz, _dt_edge_points(i, j, tz)[0]))
        ez.sort(key=lambda z: -z[0])
        for n, (vz, wz) in enumerate(ez[:DT_EDGE_STARTS], 1):
            runs.append({"start": "edge_point_%d" % n, "kind": "edge point", "w": wz,
                         "objective": vz, "status": -1})
            w, v, st = _dt_run(scal, DT_N, wz, True)
            runs.append({"start": "edge_%d" % n, "kind": "edge start", "w": w, "objective": v, "status": st})
        w, v, st = _dt_run(scal, DT_N, cl["W_face"][int(np.argmax(Sf))], True)
        runs.append({"start": "face_1", "kind": "face start", "w": w, "objective": v, "status": st})
        best = max(runs, key=lambda r: r["objective"])
        n_polish = 0
        while True:
            D = _dt_transfers(best["w"])
            g = scal(D) - best["objective"]
            gain = float(g.max())
            if gain <= DT_TRANSFER_TOL:
                break
            assert n_polish < DT_POLISH_MAX, "transfer polish did not settle"
            n_polish += 1
            wt, vt = D[int(np.argmax(g))], best["objective"] + gain
            w, v, st = _dt_run(scal, DT_N, wt, True)
            if v <= vt:
                w, v, st = wt, vt, -1
            best = {"start": "polish_%d" % n_polish, "kind": "polish", "w": w, "objective": v, "status": st}
            runs.append(best)
        Wall = Wfun(np.vstack([r["w"] for r in runs]))
        hi, lo = float(Wall.max()), float(Wall.min())
        if hi <= Wmax + DT_BOUND_TOL and lo >= Wmin - DT_BOUND_TOL:
            break
        assert ref < DT_BOUND_REFINE_MAX, "welfare bound not settled after %d refinements" % ref
        if hi > Wmax + DT_BOUND_TOL:
            Wmax = max(hi, _dt_run(Wfun, DT_N, runs[int(np.argmax(Wall))]["w"], True)[1])
        if lo < Wmin - DT_BOUND_TOL:
            Wmin = min(lo, _dt_run(Wfun, DT_N, runs[int(np.argmin(Wall))]["w"], False)[1])
    Vall = panel.var(np.vstack([r["w"] for r in runs]))
    for r, Wr, Vr in zip(runs, Wall, Vall):
        r["W"], r["V"] = float(Wr), float(Vr)
    out.update({"w": best["w"], "objective": best["objective"], "source": best["start"],
                "W_min": Wmin, "W_max": Wmax, "dW": dW, "dV": dV, "refinements": ref,
                "objective_barycentre": float(scal(_dt_bary[None, :])[0]),
                "objective_edge_grid_max": float(Se.max()), "objective_face_max": float(Sf.max()),
                "objective_edge_zoom_best": float(ez[0][0]),
                "objective_interior_best": max(r["objective"] for r in runs if r["kind"] == "interior"),
                "transfer_gain": gain, "polish_steps": n_polish, "runs": runs})
    return out


for _k in range(DT_M):
    _bid = _DT_IDS[_k]
    panel = _DtPanel(_dt_kap[:, _k], _dt_lam[:, _k])
    _dt_panels[_bid] = panel
    _rs = [_dt_rng_starts.dirichlet(np.ones(DT_N)) for _ in range(DT_STARTS_PROGRAM)]
    _face = _dt_face_sample(_dt_rng_face)
    _dt_faces[_bid] = _face
    _cl = {"U_edge": _dt_utilities(panel, _dt_Wedge), "V_edge": panel.var(_dt_Wedge),
           "U_face": _dt_utilities(panel, _face), "V_face": panel.var(_face),
           "U_vert": _dt_utilities(panel, _dt_eye), "W_face": _face}
    # confidence bounds: V is a convex quadratic on the simplex, so its
    # minimum is found from the barycentre and its maximum sits at a vertex
    _wVmin, Vmin = _dt_slsqp(panel.var, DT_N, [_dt_bary], maximize=False)
    Vmax = float(panel.Vi.max())
    # equal weights: the comparator of the leveling-down diagnostics
    _Aeq = panel.overlaps(_dt_bary[None, :])[0]
    _Ueq = _dt_fs_expnorm(_Aeq[None, :], DT_RHO, DT_THETA)[0]
    _Veq = float(panel.var(_dt_bary[None, :])[0])
    _dt_Aeq[_bid] = _Aeq
    _center = panel.mean(_dt_bary)
    _hi_i, _lo_i = int(np.argmax(panel.mu_i)), int(np.argmin(panel.mu_i))
    _sharp_i, _flat_i = int(np.argmin(panel.Vi)), int(np.argmax(panel.Vi))
    for _s in range(1, 6):
        _eq = "EQ%d" % _s
        _c = _dt_solve_cell(panel, _s, _cl, _rs, Vmin, Vmax)
        w = _c["w"]
        A = panel.overlaps(w[None, :])[0]
        U = _dt_fs_expnorm(A[None, :], DT_RHO, DT_THETA)[0]
        I = float(_dt_ineq(U[None, :], _s)[0])
        Ws = float(U.mean() * (1.0 - I))
        V = float(panel.var(w[None, :])[0])
        mu = panel.mean(w)
        _act = np.where(w > DT_ACTIVE_TOL)[0]
        _dt_W[(_bid, _eq)] = w
        _dt_A[(_bid, _eq)] = A
        _dt_U[(_bid, _eq)] = U
        _rr = _c["runs"]
        _reach = [r for r in _rr if r["objective"] >= _c["objective"] - DT_AGREE_TOL]
        _near = [r for r in _rr if r["objective"] >= _c["objective"] - DT_NEAR_TOL]
        for r in _rr:
            _dt_starts_rec.append({
                "barrier": _bid, "eq": _eq, "start": r["start"], "kind": r["kind"],
                "objective": r["objective"], "gap": _c["objective"] - r["objective"],
                "reaches": bool(r["objective"] >= _c["objective"] - DT_AGREE_TOL),
                "mu": panel.mean(r["w"]), "n_active": int((r["w"] > DT_ACTIVE_TOL).sum()),
                "W": r["W"], "V": r["V"], "slsqp_status": r["status"], "retained": r["start"] == _c["source"],
            })
        _dt_solves.append({
            "barrier": _bid, "eq": _eq, "mu": mu, "V": V,
            "Ubar": float(U.mean()), "Umin": float(U.min()), "I": I, "W": Ws,
            "W_min": _c["W_min"], "W_max": _c["W_max"], "V_min": Vmin, "V_max": Vmax,
            "W_tilde": (Ws - _c["W_min"]) / _c["dW"], "V_tilde": (V - Vmin) / _c["dV"],
            "objective": _c["objective"], "source": _c["source"],
            "objective_barycentre": _c["objective_barycentre"],
            "objective_interior_best": _c["objective_interior_best"],
            "objective_edge_zoom_best": _c["objective_edge_zoom_best"],
            "objective_edge_grid_max": _c["objective_edge_grid_max"],
            "objective_face_max": _c["objective_face_max"],
            "transfer_gain": _c["transfer_gain"], "polish_steps": _c["polish_steps"],
            "W_min_vertex": _c["W_min_vertex"], "W_min_cloud": _c["W_min_cloud"],
            "W_max_cloud": _c["W_max_cloud"], "W_min_search": _c["W_min_search"],
            "W_max_search": _c["W_max_search"],
            "W_max_starts": max(r["W"] for r in _rr), "W_min_starts": min(r["W"] for r in _rr),
            "bound_refinements": _c["refinements"],
            "n_interior_starts": sum(r["kind"] == "interior" for r in _rr),
            "n_interior_reach": sum(r["kind"] == "interior" for r in _reach),
            "n_reach": len(_reach),
            "mu_spread_reach": max(panel.mean(r["w"]) for r in _reach) - min(panel.mean(r["w"]) for r in _reach),
            "w_l1_dev_reach": max(float(np.abs(r["w"] - w).sum()) for r in _reach),
            "mu_dev_near": max(abs(panel.mean(r["w"]) - mu) for r in _near),
            "Abar": float(A.mean()), "Amin": float(A.min()),
            "Abar_equal_weights": float(_Aeq.mean()),
            "dAbar": float(A.mean() - _Aeq.mean()),
            "n_overlap_lower": int((A < _Aeq).sum()),
            "I_overlap": float(_dt_ineq(A[None, :], _s)[0]),
            "I_overlap_equal_weights": float(_dt_ineq(_Aeq[None, :], _s)[0]),
            "I_equal_weights": float(_dt_ineq(_Ueq[None, :], _s)[0]),
            "Ubar_equal_weights": float(_Ueq.mean()),
            "V_equal_weights": _Veq,
            "leveling_down": bool(A.mean() < _Aeq.mean()
                                  and _dt_ineq(A[None, :], _s)[0] < _dt_ineq(_Aeq[None, :], _s)[0]),
            "n_active": int(_act.size), "w_max": float(w.max()),
            "active_experts": ",".join(str(i + 1) for i in _act),
            "active_is_mean_extremes": bool(set(_act.tolist()) == {_hi_i, _lo_i}),
            "active_has_sharpest": bool(_sharp_i in _act), "active_has_flattest": bool(_flat_i in _act),
            "active_straddle_center": bool(panel.mu_i[_act].max() > _center > panel.mu_i[_act].min()),
            "weight_gini": float(np.abs(w[:, None] - w[None, :]).sum() / (2.0 * DT_N * DT_N * w.mean())),
        })
    print("   %-4s mu by EQ1..EQ5: %s  active %s  interior starts reaching %s  (%.0f s)" % (
        _bid, " ".join("%.3f" % d["mu"] for d in _dt_solves[-5:]),
        [d["n_active"] for d in _dt_solves[-5:]], [d["n_interior_reach"] for d in _dt_solves[-5:]],
        time.time() - _dt_t0))

# the 14 x 5 matrix (eq:dt-expectation), the average (eq:dt-average), the rank
_dt_Mx = np.array([[next(d["mu"] for d in _dt_solves if d["barrier"] == b and d["eq"] == e)
                    for e in _DT_EQ] for b in _DT_IDS])
_dt_avg = _dt_Mx.mean(axis=1)
_dt_order = sorted(range(DT_M), key=lambda r: -_dt_avg[r])
_dt_rank = np.empty(DT_M, int)
for _p, _r in enumerate(_dt_order, 1):
    _dt_rank[_r] = _p
_dt_triad = [_DT_IDS[r] for r in _dt_order[:3]]
_dt_ew_mean = np.array([_dt_panels[b].mean(_dt_bary) for b in _DT_IDS])
_dt_eq_orders = {e: ">".join(_DT_IDS[r] for r in np.argsort(-_dt_Mx[:, j], kind="stable"))
                 for j, e in enumerate(_DT_EQ)}
print("   ranking: " + ">".join(_DT_IDS[r] for r in _dt_order))
print("   triad %s; bottom three %s; active experts %d..%d; EQ2-EQ5 max |diff| %.2e"
      % (_dt_triad, [_DT_IDS[r] for r in _dt_order[-3:]],
         min(d["n_active"] for d in _dt_solves), max(d["n_active"] for d in _dt_solves),
         float(np.abs(_dt_Mx[:, 1] - _dt_Mx[:, 4]).max())))

# collective densities of the leading triad under DT_TRIAD_EQ (Figure 6.7)
_dt_dens = {"x": _dt_xg}
for _b in _dt_triad:
    _dt_dens["f_" + _b] = _dt_panels[_b].density(_dt_W[(_b, DT_TRIAD_EQ)], _dt_xg)

# ---- the numbers the text of Section 6.3 (and its echoes in tab:robustness
# and Chapter 7) prints, computed here from the solved outputs; the tail
# self-checks recompute each one from the other sheets.  Comparisons of the
# collective means use the tolerance DT_TIE_TOL, so that differences below
# solver precision are not counted as wins.
DT_TIE_TOL = 1e-6
_dt_st = pd.DataFrame(_dt_starts_rec)
for _d in _dt_solves:
    _g = _dt_st[(_dt_st["barrier"] == _d["barrier"]) & (_dt_st["eq"] == _d["eq"])]
    _gr = _g[_g["objective"] >= _d["objective"] - DT_AGREE_TOL]
    _d["mu_reach_min"], _d["mu_reach_max"] = float(_gr["mu"].min()), float(_gr["mu"].max())
    _g = _g[_g["objective"] >= _d["objective"] - DT_NEAR_TOL]
    _d["mu_near_min"], _d["mu_near_max"] = float(_g["mu"].min()), float(_g["mu"].max())
_dts = pd.DataFrame(_dt_solves)
_dt_nr = _dts.set_index(["barrier", "eq"])
_dt_spread = _dt_Mx.max(axis=1) - _dt_Mx.min(axis=1)
_dt_eq1_low = [_DT_IDS[k] for k in range(DT_M) if _dt_Mx[k, 0] < _dt_Mx[k, 1:].min() - DT_TIE_TOL]
_dt_eq1_tie = [_DT_IDS[k] for k in range(DT_M) if abs(_dt_Mx[k, 0] - _dt_Mx[k, 1:].min()) <= DT_TIE_TOL]
_dt_eq1_high = [_DT_IDS[k] for k in range(DT_M) if _dt_Mx[k, 0] > _dt_Mx[k, 1:].max() + DT_TIE_TOL]
_dt_eq1_dep = np.abs(_dt_Mx[:, 0] - _dt_Mx[:, 1:].mean(axis=1))
_dt_span4 = _dt_Mx[:, 1:].max(axis=1) - _dt_Mx[:, 1:].min(axis=1)
_dt_d25 = np.abs(_dt_Mx[:, 1] - _dt_Mx[:, 4])
# adjacent pairs of the common order that tie at the table's two decimals
_dt_ties = []
for _j, _e in enumerate(_DT_EQ):
    for _p in range(DT_M - 1):
        _a, _b = _dt_order[_p], _dt_order[_p + 1]
        if "%.2f" % _dt_Mx[_a, _j] == "%.2f" % _dt_Mx[_b, _j]:
            _dt_ties.append({"eq": _e, "upper": _DT_IDS[_a], "lower": _DT_IDS[_b],
                             "upper_mu": _dt_Mx[_a, _j], "lower_mu": _dt_Mx[_b, _j],
                             "margin": _dt_Mx[_a, _j] - _dt_Mx[_b, _j]})
# adjacent pairs of the common order, under each index, whose near-optimal
# mean ranges (starts within DT_NEAR_TOL of the retained value) overlap
_dt_near_overlap = []
for _j, _e in enumerate(_DT_EQ):
    for _p in range(DT_M - 1):
        _a, _b = _DT_IDS[_dt_order[_p]], _DT_IDS[_dt_order[_p + 1]]
        if _dt_nr.loc[(_b, _e), "mu_near_max"] >= _dt_nr.loc[(_a, _e), "mu_near_min"]:
            _dt_near_overlap.append((_e, _a, _b))
for _t in _dt_ties:
    _t["near_optimal_overlap"] = (_t["eq"], _t["upper"], _t["lower"]) in _dt_near_overlap
_dt_two = _dts[_dts["n_active"] == 2]
_dt_Vrel = (_dts["V"] / _dts["V_equal_weights"] - 1.0).abs()
_dt_emsd = np.array([_dt_wmean[:, k].std() for k in range(DT_M)])
_dt_near_big = _dts[_dts["mu_dev_near"] > DT_FLAT_TOL]

# ---- doubled-discretization check (R24).  The programs behind every
# adjacent pair of the common order that ties at the table's two decimals
# (sheet near_ties), and every flat program, whose near-optimal starts move
# the collective mean by more than DT_FLAT_TOL, are solved again on a grid
# of DT_FINE_FACTOR times as many probabilities and bins by the same search
# without the Dirichlet starts (the vertices, the edge grid, the zoom on the
# best edges and SLSQP from the best zoomed points, SLSQP from the
# barycentre and from the best point of the barrier's face sample, the
# transfer check and polish, and the bound refinement).  Sheet fine_check.
_dt_fine_cells = []
for _t in _dt_ties:
    for _b in (_t["upper"], _t["lower"]):
        if (_b, _t["eq"]) not in _dt_fine_cells:
            _dt_fine_cells.append((_b, _t["eq"]))
for _r in _dt_near_big.itertuples():
    if (_r.barrier, _r.eq) not in _dt_fine_cells:
        _dt_fine_cells.append((_r.barrier, _r.eq))
_dt_fine = []
_dt_fine_W = {}
_dt_fine_panels = {}
for _b, _e in _dt_fine_cells:
    _k = _DT_IDS.index(_b)
    if _b not in _dt_fine_panels:
        _fp = _DtPanel(_dt_kap[:, _k], _dt_lam[:, _k], pgrid=DT_FINE_FACTOR * DT_PGRID,
                       xbins=DT_FINE_FACTOR * DT_XBINS)
        _fcl = {"U_edge": _dt_utilities(_fp, _dt_Wedge, chunk=1000), "V_edge": _fp.var(_dt_Wedge),
                "U_face": _dt_utilities(_fp, _dt_faces[_b], chunk=1000), "V_face": _fp.var(_dt_faces[_b]),
                "U_vert": _dt_utilities(_fp, _dt_eye), "W_face": _dt_faces[_b]}
        _dt_fine_panels[_b] = (_fp, _fcl, _dt_slsqp(_fp.var, DT_N, [_dt_bary], maximize=False)[1],
                               float(_fp.Vi.max()))
    _fp, _fcl, _fVmin, _fVmax = _dt_fine_panels[_b]
    _c = _dt_solve_cell(_fp, int(_e[2:]), _fcl, [], _fVmin, _fVmax)
    _wf, _wc = _c["w"], _dt_W[(_b, _e)]
    _dt_fine_W[(_b, _e)] = _wf
    _dt_fine.append({
        "barrier": _b, "eq": _e,
        "reason": ", ".join(_x for _x, _on in (
            ("near tie", any(_b in (_t["upper"], _t["lower"]) and _e == _t["eq"] for _t in _dt_ties)),
            ("flat", (_b, _e) in set(zip(_dt_near_big["barrier"], _dt_near_big["eq"])))) if _on),
        "mu": float(_dt_nr.loc[(_b, _e), "mu"]), "mu_fine": _fp.mean(_wf),
        "mu_fine_minus_mu": _fp.mean(_wf) - float(_dt_nr.loc[(_b, _e), "mu"]),
        "active_experts": str(_dt_nr.loc[(_b, _e), "active_experts"]),
        "active_experts_fine": ",".join(str(i + 1) for i in np.where(_wf > DT_ACTIVE_TOL)[0]),
        "w_l1_fine_vs_coarse": float(np.abs(_wf - _wc).sum()),
        "objective_fine": _c["objective"], "source_fine": _c["source"],
        "W_min_fine": _c["W_min"], "W_max_fine": _c["W_max"], "V_min_fine": _fVmin, "V_max_fine": _fVmax,
        "bound_refinements_fine": _c["refinements"], "transfer_gain_fine": _c["transfer_gain"],
        "objective_edge_grid_max_fine": _c["objective_edge_grid_max"],
        "objective_face_max_fine": _c["objective_face_max"],
    })
    print("   doubled discretization %-4s %s: mean %.5f (grid %d/%d: %.5f), active %s -> %s"
          % (_b, _e, _fp.mean(_wf), DT_PGRID, DT_XBINS, float(_dt_nr.loc[(_b, _e), "mu"]),
             _dt_fine[-1]["active_experts"], _dt_fine[-1]["active_experts_fine"]))
_dt_fnr = {(_f["barrier"], _f["eq"]): _f for _f in _dt_fine}
for _t in _dt_ties:
    _t["upper_mu_fine"] = _dt_fnr[(_t["upper"], _t["eq"])]["mu_fine"]
    _t["lower_mu_fine"] = _dt_fnr[(_t["lower"], _t["eq"])]["mu_fine"]
    _t["margin_fine"] = _t["upper_mu_fine"] - _t["lower_mu_fine"]
_dt_tn = []


def _dt_num(key, value, printed, where):
    _dt_tn.append({"key": key, "value": value, "printed": printed, "where": where})


def _f2(x):
    return "%.2f" % x


_dt_num("elicit_expert", _dt_el_i + 1, "%d" % (_dt_el_i + 1), "sec:dt-elicit, tab:dt-elicit")
_dt_num("elicit_p1_max", float(_dt_el_prof[0]), "0.001", "sec:dt-elicit (the first level receives less than this)")
_dt_num("elicit_kappa", _dt_el_fit[0], _f2(_dt_el_fit[0]), "sec:dt-elicit, tab:dt-elicit")
_dt_num("elicit_lambda", _dt_el_fit[1], _f2(_dt_el_fit[1]), "sec:dt-elicit, tab:dt-elicit")
_dt_num("elicit_mean", _dt_el_fit[2], _f2(_dt_el_fit[2]), "sec:dt-elicit, tab:dt-elicit")
for _p in (0, 1, 2, 3):
    _r = _dt_order[_p]
    _dt_num("avg_rank%d_%s" % (_p + 1, _DT_IDS[_r]), _dt_avg[_r], _f2(_dt_avg[_r]), "sec:dt-results")
for _p in (DT_M - 1, DT_M - 2, DT_M - 3):
    _r = _dt_order[_p]
    _dt_num("avg_rank%d_%s" % (_p + 1, _DT_IDS[_r]), _dt_avg[_r], _f2(_dt_avg[_r]), "sec:dt-results")
_dt_num("n_orders_equal_to_average", sum(o == ">".join(_DT_IDS[r] for r in _dt_order) for o in _dt_eq_orders.values()),
        "five", "sec:dt-results (the order is the same under every index)")
_dt_num("n_ties_at_2dp", len(_dt_ties), str(len(_dt_ties)), "sec:dt-results")
for _t in _dt_ties:
    _dt_num("tie_%s_%s_%s" % (_t["eq"], _t["upper"], _t["lower"]), _t["margin"], "%.4f" % _t["margin"],
            "sec:dt-results (margin under %s)" % _t["eq"])
for _k in np.argsort(_dt_spread, kind="stable")[:2].tolist() + np.argsort(-_dt_spread, kind="stable")[:2].tolist():
    _dt_num("spread_%s" % _DT_IDS[_k], _dt_spread[_k], "%.3f" % _dt_spread[_k], "sec:dt-results")
for _b in _dt_triad[1:]:
    _k = _DT_IDS.index(_b)
    _dt_num("spread_%s" % _b, _dt_spread[_k], "%.3f" % _dt_spread[_k], "sec:dt-results (within the triad)")
_dt_num("spread_max", float(_dt_spread.max()), _f2(_dt_spread.max()), "tab:robustness, chapter7 (by up to)")
_dt_num("n_eq1_lowest", len(_dt_eq1_low), str(len(_dt_eq1_low)), "sec:dt-results: " + ",".join(_dt_eq1_low))
_dt_num("n_eq1_tie_lowest", len(_dt_eq1_tie), str(len(_dt_eq1_tie)), "sec:dt-results: " + ",".join(_dt_eq1_tie))
for _b in _dt_eq1_tie:
    _k = _DT_IDS.index(_b)
    _jt = 1 + int(np.argmin(_dt_Mx[_k, 1:]))
    _dt_num("eq1_tie_%s_partner" % _b, _jt + 1, _DT_EQ[_jt], "sec:dt-results (the index EQ1 ties for the lowest)")
    _dt_num("eq1_tie_%s_weight_diff" % _b, float(np.abs(_dt_W[(_b, "EQ1")] - _dt_W[(_b, _DT_EQ[_jt])]).max()),
            "same solution", "sec:dt-results (largest weight difference of the two solutions)")
_dt_num("n_eq1_highest", len(_dt_eq1_high), str(len(_dt_eq1_high)), "sec:dt-results: " + ",".join(_dt_eq1_high))
for _k in np.argsort(-_dt_eq1_dep, kind="stable")[:2]:
    _dt_num("eq1_departure_%s" % _DT_IDS[_k], _dt_eq1_dep[_k], _f2(_dt_eq1_dep[_k]),
            "sec:dt-results (largest departures of EQ1 from the mean of the other four)")
_k2 = int(np.argmax(_dt_span4))
_dt_num("span4_max_%s" % _DT_IDS[_k2], _dt_span4[_k2], _f2(_dt_span4[_k2]), "sec:dt-results")
_dt_num("span4_max_other", float(np.delete(_dt_span4, _k2).max()), _f2(np.delete(_dt_span4, _k2).max()),
        "sec:dt-results (the other four indices within this on every other barrier)")
_k3 = int(np.argmax(_dt_d25))
_dt_num("eq2_eq5_max_%s" % _DT_IDS[_k3], _dt_d25[_k3], "%.3f" % _dt_d25[_k3], "sec:dt-results")
_dt_num("n_two_expert", len(_dt_two), str(len(_dt_two)), "sec:dt-results")
_dt_num("n_active_max", int(_dts["n_active"].max()), str(int(_dts["n_active"].max())), "sec:dt-results")
_dt_num("n_pair_mean_extremes", int(_dt_two["active_is_mean_extremes"].sum()),
        str(int(_dt_two["active_is_mean_extremes"].sum())), "sec:dt-results")
_dt_num("n_pair_straddle", int(_dt_two["active_straddle_center"].sum()),
        str(int(_dt_two["active_straddle_center"].sum())), "sec:dt-results")
_dt_num("n_sharpest_active", int(_dts["active_has_sharpest"].sum()), str(int(_dts["active_has_sharpest"].sum())),
        "sec:dt-results")
_dt_num("n_flattest_active", int(_dts["active_has_flattest"].sum()), str(int(_dts["active_has_flattest"].sum())),
        "sec:dt-results")
_dt_fl = sorted({str(b) for b in _dts.loc[~_dts["active_has_flattest"], "barrier"]}, key=_DT_IDS.index)
_dt_num("barriers_flattest_not_active", len(_dt_fl), ",".join(_dt_fl), "sec:dt-results")
_dt_fla = _dts.groupby("barrier")["active_has_flattest"].agg(["all", "any"])
_dt_num("n_barriers_flattest_active", int(_dt_fla["all"].sum()), str(int(_dt_fla["all"].sum())),
        "sec:dt-results (barriers whose flattest expert is active under every index)")
_dt_num("n_barriers_flattest_mixed", int((_dt_fla["any"] & ~_dt_fla["all"]).sum()),
        str(int((_dt_fla["any"] & ~_dt_fla["all"]).sum())), "diagnostic (must be 0 for the text's reading)")
_dt_num("Abar_min", float(_dts["Abar"].min()), _f2(_dts["Abar"].min()), "sec:dt-results")
_dt_num("Abar_max", float(_dts["Abar"].max()), _f2(_dts["Abar"].max()), "sec:dt-results")
_dt_num("Abar_eq_min", float(_dts["Abar_equal_weights"].min()), _f2(_dts["Abar_equal_weights"].min()), "sec:dt-results")
_dt_num("Abar_eq_max", float(_dts["Abar_equal_weights"].max()), _f2(_dts["Abar_equal_weights"].max()), "sec:dt-results")
_dt_num("n_Abar_lower", int((_dts["dAbar"] < 0).sum()), str(int((_dts["dAbar"] < 0).sum())), "sec:dt-results")
_dt_num("Abar_drop_min", float(-_dts["dAbar"].max()), "%.3f" % (-_dts["dAbar"].max()), "sec:dt-results")
_dt_num("Abar_drop_max", float(-_dts["dAbar"].min()), "%.3f" % (-_dts["dAbar"].min()), "sec:dt-results")
_dt_num("Abar_drop_max_pct", float((-_dts["dAbar"] / _dts["Abar_equal_weights"]).max() * 100),
        "%.1f" % ((-_dts["dAbar"] / _dts["Abar_equal_weights"]).max() * 100), "sec:dt-results (percent)")
_dt_num("n_all_overlaps_lower", int((_dts["n_overlap_lower"] == DT_N).sum()),
        str(int((_dts["n_overlap_lower"] == DT_N).sum())), "sec:dt-results")
_dt_num("n_barriers_leveling_down", int(_dts.groupby("barrier")["leveling_down"].any().sum()),
        str(int(_dts.groupby("barrier")["leveling_down"].any().sum())), "sec:dt-results (under at least one index)")
_dt_num("n_Ubar_higher", int((_dts["Ubar"] > _dts["Ubar_equal_weights"]).sum()),
        str(int((_dts["Ubar"] > _dts["Ubar_equal_weights"]).sum())), "sec:dt-results (Ubar above its equal-weight value)")
_dt_num("n_leveling_down", int(_dts["leveling_down"].sum()), str(int(_dts["leveling_down"].sum())),
        "sec:dt-results (Prop. consensus-fairness (iii) on the overlap profile)")
_dt_num("Ubar_min", float(_dts["Ubar"].min()), _f2(_dts["Ubar"].min()), "sec:dt-results")
_dt_num("Ubar_max", float(_dts["Ubar"].max()), _f2(_dts["Ubar"].max()), "sec:dt-results")
_dt_num("V_rel_max_pct", float(_dt_Vrel.max() * 100), "%.0f" % np.ceil(_dt_Vrel.max() * 100),
        "sec:dt-results (aggregate variance within this percent of its equal-weight value)")
_dt_num("W_tilde_min", float(_dts["W_tilde"].min()), _f2(_dts["W_tilde"].min()), "sec:dt-results")
_dt_num("W_tilde_max", float(_dts["W_tilde"].max()), _f2(_dts["W_tilde"].max()), "sec:dt-results")
_dt_num("expert_mean_sd_min", float(_dt_emsd.min()), _f2(_dt_emsd.min()), "sec:dt-results")
_dt_num("expert_mean_sd_max", float(_dt_emsd.max()), _f2(_dt_emsd.max()), "sec:dt-results")
for _b in _dt_triad:
    _k = _DT_IDS.index(_b)
    _dt_num("triad_%s_%s" % (_b, DT_TRIAD_EQ), _dt_Mx[_k, _DT_EQ.index(DT_TRIAD_EQ)],
            _f2(_dt_Mx[_k, _DT_EQ.index(DT_TRIAD_EQ)]), "sec:dt-results (Figure 6.7)")
def _ceil_to(x, d):
    """The smallest multiple of 10^-d at or above x, printed with d decimals
    (the bound a 'less than' or 'within' claim of the text prints)."""
    return "%.*f" % (d, math.ceil(x * 10 ** d - 1e-9) / 10 ** d)


# restart stability (sec:dt-data): the interior starts (the barycentre and the
# Dirichlet(1) draws) reaching the retained value, the program none reaches, and
# the mean agreement among the starts
_dt_num("interior_reach_median", float(_dts["n_interior_reach"].median()), "%g" % _dts["n_interior_reach"].median(),
        "sec:dt-data (of the %d interior starts)" % (DT_STARTS_PROGRAM + 1))
_dt_rz = _dts[_dts["n_interior_reach"] == 0]
_dt_num("n_interior_reach_zero", len(_dt_rz), str(len(_dt_rz)),
        "sec:dt-data: " + ",".join("%s/%s" % (r.barrier, r.eq) for r in _dt_rz.itertuples()))
_dt_num("interior_reach_min_nonzero", int(_dts.loc[_dts["n_interior_reach"] > 0, "n_interior_reach"].min()),
        str(int(_dts.loc[_dts["n_interior_reach"] > 0, "n_interior_reach"].min())),
        "sec:dt-data (at least this many in every other program)")
_dt_z = _dt_st[_dt_st["kind"] == "interior"].merge(_dt_rz[["barrier", "eq", "mu"]], on=["barrier", "eq"],
                                                 suffixes=("", "_cell"))
_dt_zc = _dt_z[_dt_z["n_active"] == 1]
_dt_zs = _dt_z[_dt_z["n_active"] > 1]
_dt_num("zero_reach_n_collapsed", len(_dt_zc), str(len(_dt_zc)),
        "sec:dt-data (interior starts of that program ending on a single expert)")
_dt_num("zero_reach_n_spread", len(_dt_zs), str(len(_dt_zs)), "sec:dt-data (the others)")
_dt_num("zero_reach_nact_min", int(_dt_zs["n_active"].min()), str(int(_dt_zs["n_active"].min())),
        "sec:dt-data (the others end on this many experts, at least)")
_dt_num("zero_reach_nact_max", int(_dt_zs["n_active"].max()), str(int(_dt_zs["n_active"].max())),
        "sec:dt-data (and at most)")
_dt_num("zero_reach_gap_min", float(_dt_zs["gap"].min()), "%.4f" % _dt_zs["gap"].min(),
        "sec:dt-data (the others stop this far below the retained value, at least)")
_dt_num("zero_reach_gap_max", float(_dt_zs["gap"].max()), "%.4f" % _dt_zs["gap"].max(),
        "sec:dt-data (and at most)")
_dt_num("zero_reach_mu_dev_max", float((_dt_zs["mu"] - _dt_zs["mu_cell"]).abs().max()),
        _ceil_to(float((_dt_zs["mu"] - _dt_zs["mu_cell"]).abs().max()), 5),
        "sec:dt-data (their collective means within this of the retained one)")
_dt_msr = (_dts["mu_reach_max"] - _dts["mu_reach_min"])
_dt_kx = int(np.argmax(_dt_msr.to_numpy()))
_dt_num("mu_reach_split_cell", _dt_kx, "%s/%s" % (_dts["barrier"].iloc[_dt_kx], _dts["eq"].iloc[_dt_kx]),
        "sec:dt-data (the program whose reaching starts return two means)")
_dt_num("mu_reach_split_lo", float(_dts["mu_reach_min"].iloc[_dt_kx]), "%.4f" % _dts["mu_reach_min"].iloc[_dt_kx],
        "sec:dt-data")
_dt_num("mu_reach_split_hi", float(_dts["mu_reach_max"].iloc[_dt_kx]), "%.4f" % _dts["mu_reach_max"].iloc[_dt_kx],
        "sec:dt-data (both print as %s)" % _f2(_dts["mu_reach_max"].iloc[_dt_kx]))
_dt_num("mu_reach_spread_max_other", float(_dt_msr.drop(_dt_msr.index[_dt_kx]).max()),
        _ceil_to(float(_dt_msr.drop(_dt_msr.index[_dt_kx]).max()), 4),
        "sec:dt-data (every other program, within this)")
# near-optimal starts (within DT_NEAR_TOL): the flat programs
_dt_num("n_flat", len(_dt_near_big), str(len(_dt_near_big)),
        "sec:dt-data: " + ",".join("%s/%s %.3f" % (r.barrier, r.eq, r.mu_dev_near) for r in _dt_near_big.itertuples()))
for _r in _dt_near_big.itertuples():
    _dt_num("mu_dev_near_%s_%s" % (_r.barrier, _r.eq), float(_r.mu_dev_near), "%.3f" % _r.mu_dev_near, "sec:dt-data")
_dt_num("mu_dev_near_max_other", float(_dts.loc[_dts["mu_dev_near"] <= DT_FLAT_TOL, "mu_dev_near"].max()),
        _ceil_to(float(_dts.loc[_dts["mu_dev_near"] <= DT_FLAT_TOL, "mu_dev_near"].max()), 3),
        "sec:dt-data (every other program, less than this)")
# the doubled-discretization check (sec:dt-data, sec:dt-results)
_dt_num("fine_pgrid", DT_FINE_FACTOR * DT_PGRID, "{:,}".format(DT_FINE_FACTOR * DT_PGRID), "sec:dt-data")
_dt_num("fine_xbins", DT_FINE_FACTOR * DT_XBINS, "{:,}".format(DT_FINE_FACTOR * DT_XBINS), "sec:dt-data")
_dt_nft = sum("near tie" in _f["reason"] for _f in _dt_fine)
_dt_num("n_fine_tie_cells", _dt_nft, str(_dt_nft), "sec:dt-data (the programs behind the near ties)")
_dt_num("n_fine_same_pair", sum(_f["active_experts"] == _f["active_experts_fine"] for _f in _dt_fine),
        str(sum(_f["active_experts"] == _f["active_experts_fine"] for _f in _dt_fine)),
        "sec:dt-data, tab:robustness, ch6 closing synthesis (of the %d programs solved again, the same "
        "two experts)" % len(_dt_fine))
_dt_fsh = max(abs(_f["mu_fine_minus_mu"]) for _f in _dt_fine)
_dt_num("fine_mu_shift_max", _dt_fsh, _ceil_to(_dt_fsh, 4),
        "sec:dt-data, tab:robustness, ch6 closing synthesis (collective means move by less than this)")
for _t in _dt_ties:
    _dt_num("tie_fine_%s_%s_%s" % (_t["eq"], _t["upper"], _t["lower"]), _t["margin_fine"], "%.4f" % _t["margin_fine"],
            "sec:dt-results (margin on the doubled discretization)")
_dt_num("n_near_optimal_overlaps", len(_dt_near_overlap), str(len(_dt_near_overlap)),
        "sec:dt-results: " + ",".join("%s %s/%s" % t for t in _dt_near_overlap))
for _e, _a, _b in _dt_near_overlap:
    _dt_num("near_overlap_%s_%s_%s_lower_max" % (_e, _a, _b), float(_dt_nr.loc[(_b, _e), "mu_near_max"]),
            "%.4f" % _dt_nr.loc[(_b, _e), "mu_near_max"],
            "sec:dt-results (a near-optimal solution of the lower barrier's program, above the upper's %.4f)"
            % _dt_nr.loc[(_a, _e), "mu"])
_dt_num("n_bound_refined", int((_dts["bound_refinements"] > 0).sum()), str(int((_dts["bound_refinements"] > 0).sum())),
        "sec:dt-data")
_dt_num("n_W_min_at_vertex", int((np.abs(_dts["W_min"] - _dts["W_min_vertex"]) <= 1e-12).sum()),
        str(int((np.abs(_dts["W_min"] - _dts["W_min_vertex"]) <= 1e-12).sum())), "sec:dt-data")
_dt_nrr = int(_dts["source"].isin(["barycentre"] + ["dirichlet_%d" % n for n in range(1, DT_STARTS_PROGRAM + 1)]).sum())
_dt_num("n_retained_from_interior", _dt_nrr, str(_dt_nrr), "diagnostic")
_dt_num("n_retained_from_edges", int(_dts["source"].str.startswith("edge").sum()),
        str(int(_dts["source"].str.startswith("edge").sum())), "diagnostic")
for _r in _dt_tn:
    print("   text number %-34s %-12s %s" % (_r["key"], _r["printed"], _r["where"]))

_dt_cols = ["%s_%s" % (b, e) for b in _DT_IDS for e in _DT_EQ]
_dt_elapsed = time.time() - _dt_t0
write("fig6-5_dt_case.xlsx", {
    "config": cfg({
        "case": "II -- barriers to digital transformation in SMCEs (Section 6.3)",
        "figures": "fig6-5_dt_ranking14, fig6-6_dt_crossswf, fig6-7_dt_triad",
        "labels": "fig:dt-ranking14, fig:dt-crossswf, fig:dt-triad",
        "source_paper": "chen2024socialwelfare",
        "synthetic": 1,
        "generator": "build_workbooks.py Case II block (generated and solved; nothing typed)",
        "seed": DT_SEED,
        "seed_use": ("default_rng(seed): (1) 30 leniencies a_i, (2) 30x14 noise eps_ik; "
                     "a fresh default_rng(seed) then draws the SLSQP starts, %d Dirichlet(1) "
                     "vectors per barrier in order B1..B14; a second fresh default_rng(seed) draws, "
                     "barrier by barrier in the same order, the %d points of the face sample "
                     "(per point: face size, the experts of the face, their Dirichlet(1) weights)"
                     % (DT_STARTS_PROGRAM, DT_FACE_SAMPLE)),
        "N_experts": DT_N, "M_barriers": DT_M,
        "scale_min": 0.0, "scale_max": DT_HI, "n_levels": 5,
        "backgrounds": "experts 1-15 technologists (bias on B1-B4), 16-30 managers (bias on B5-B10)",
        "design_means": ", ".join("%s %.2f" % (b, d) for b, d in zip(_DT_IDS, _DT_DESIGN)),
        "target_calibration": ("tau_k found by secant iteration so that the equal-weight quantile-"
                               "averaged mean of the fitted panel equals the design mean within %g "
                               "(at most %d refits)" % (DT_CALIB_TOL, DT_CALIB_MAXIT)),
        "rating_rule": "x_ik = clip(tau_k + a_i + r_ik + eps_ik, 0, 5)  (symbols of chapter6.tex)",
        "background_bias_rule": ("r_ik = %.2f on the barriers of the expert's own dimension "
                                 "(B1-B4 technologist, B5-B10 manager), 0 elsewhere" % DT_BIAS),
        "leniency_sd": DT_LENIENCY_SD, "noise_sd": DT_NOISE_SD, "background_bias": DT_BIAS,
        "profile_rule": ("p_t proportional to exp(-(m_t - x_ik)^2 / (2 sd^2)), m_t = t - 1/2, "
                         "sd = %.1f" % DT_PROFILE_SD),
        "profile_kernel_sd": DT_PROFILE_SD,
        "weibull_estimator": ("grouped maximum likelihood of the truncated Weibull on the five "
                              "bands, nine L-BFGS-B starts in log-parameters"),
        "elicit_barrier": DT_ELICIT_BARRIER, "elicit_expert": _dt_el_i + 1,
        "elicit_rule": ("tab:dt-elicit shows a member of the generated panel: on the barrier "
                        "elicit_barrier, the expert whose fitted mean is closest to the barrier's "
                        "equal-weight collective mean (sheet elicitation_b1 copies that panel row)"),
        "elicit_profile": ", ".join("%.3f" % p for p in _dt_el_prof),
        "elicit_fit_kappa": _dt_el_fit[0], "elicit_fit_lambda": _dt_el_fit[1],
        "elicit_fit_mean": _dt_el_fit[2], "elicit_fit_H5": _dt_el_fit[3],
        "rho": DT_RHO, "theta": DT_THETA,
        "utility_normalization": "exponential unit-interval normalization eq:exp-norm",
        "welfare_form": "W_s(w) = Ubar(w) (1 - I_s(U(w))), Sen's abbreviated form (eq:complete-welfare, multiplicative)",
        "program": "max_w  W~_s(w) - V~(w) on the simplex, equal weights, both terms min-max normalized",
        "minmax_W": ("W_min = the smallest welfare over the 30 vertices, the edge grid, the face sample, and "
                     "SLSQP minimizations from the best vertex, the best face point, the barycentre and the "
                     "first %d Dirichlet(1) starts; W_max = the largest welfare over the edge grid, the face "
                     "sample, a zoom on the %d best edges, and SLSQP maximizations from the %d best zoomed "
                     "points, the best face point, the barycentre and the first %d Dirichlet(1) starts; a bound "
                     "that the welfare of any candidate of the program exceeds is moved to it (and further by "
                     "a bound solve started there) and the program is re-solved, until no candidate exceeds "
                     "it (at most %d times), so W~ lies in [0, 1] at every recorded start and solution"
                     % (DT_STARTS_BOUNDS, DT_EDGE_TOP, DT_EDGE_STARTS, DT_STARTS_BOUNDS, DT_BOUND_REFINE_MAX)),
        "bound_refine_max": DT_BOUND_REFINE_MAX, "bound_tol": DT_BOUND_TOL,
        "edge_grid": ("each of the %d two-expert edges t e_i + (1 - t) e_j on t = 0, 1/%d, ..., 1 "
                      "(its endpoints are the 30 vertices)" % (len(_DT_PAIRS), DT_EDGE_GRID - 1)),
        "edge_grid_points": DT_EDGE_GRID, "edge_top": DT_EDGE_TOP, "edge_starts": DT_EDGE_STARTS,
        "edge_zoom": ("from the best grid point of an edge, %d points on the bracket of its neighbours, "
                      "repeated on the bracket of the best point's neighbours until the bracket is "
                      "narrower than %g" % (DT_ZOOM_PTS, DT_ZOOM_WIDTH)),
        "face_sample": ("%d points per barrier, each on a face of m experts, m uniform on %s, the experts "
                        "uniform without replacement, Dirichlet(1) weights on them"
                        % (DT_FACE_SAMPLE, list(DT_FACE_SIZES))),
        "program_search": ("SLSQP from the barycentre and the %d Dirichlet(1) starts (the interior starts), "
                           "the %d best zoomed points of the %d best edges of the scalarized objective and "
                           "SLSQP from each, and SLSQP from the best face point; the best candidate is retained"
                           % (DT_STARTS_PROGRAM, DT_EDGE_STARTS, DT_EDGE_TOP)),
        "verification": ("the retained solution is at least as good as every start, every zoomed edge point, "
                         "the edge grid and the face sample (sheet solves), and no pairwise transfer of "
                         "%s of weight from an active expert to another gains more than %g; a transfer that "
                         "gains is taken and SLSQP restarted from it (polish, at most %d steps)"
                         % (list(DT_TRANSFER_STEPS), DT_TRANSFER_TOL, DT_POLISH_MAX)),
        "agreement_tol": DT_AGREE_TOL, "near_optimal_tol": DT_NEAR_TOL, "tie_tol": DT_TIE_TOL,
        "flat_tol": DT_FLAT_TOL, "fine_factor": DT_FINE_FACTOR,
        "fine_check": ("doubled-discretization check (sheets fine_check, fine_weights, near_ties *_fine): the "
                       "programs of every adjacent pair of the common order that ties at two decimals and every "
                       "program whose near-optimal starts (within near_optimal_tol) move the collective mean by "
                       "more than flat_tol are solved again with fine_factor x p_grid probabilities and "
                       "fine_factor x x_bins bins, by the same search without the Dirichlet starts (vertices, "
                       "edge grid, zoom, SLSQP from the best zoomed points, the barycentre and the best face "
                       "point, transfer check and polish, bound refinement)"),
        "transfer_tol": DT_TRANSFER_TOL, "transfer_steps": ", ".join("%g" % e for e in DT_TRANSFER_STEPS),
        "polish_max": DT_POLISH_MAX, "dirichlet_starts_program": DT_STARTS_PROGRAM,
        "dirichlet_starts_bounds": DT_STARTS_BOUNDS,
        "zoom_points": DT_ZOOM_PTS, "zoom_width": DT_ZOOM_WIDTH, "face_sample_size": DT_FACE_SAMPLE,
        "face_sizes": ", ".join(str(m) for m in DT_FACE_SIZES),
        "restart_stability": ("sheet starts records every start of the final program solve; n_interior_reach "
                              "counts the interior starts (the barycentre and the Dirichlet(1) draws, "
                              "kind 'interior' in sheet starts) within agreement_tol of the retained value, "
                              "mu_spread_reach (= mu_reach_max - mu_reach_min) the range of the collective "
                              "mean over the starts within it, mu_dev_near the largest deviation of the "
                              "collective mean over the starts within near_optimal_tol (mu_near_min, "
                              "mu_near_max their range); a program with mu_dev_near > flat_tol is flat"),
        "minmax_V": "V_min by SLSQP from the barycentre (V convex); V_max = max_i V_i (vertex of the simplex)",
        "solver": ("scipy SLSQP, equality sum w = 1, bounds [0,1], batched forward-difference gradient "
                   "(step %g), ftol %g, maxiter %d" % (DT_FDH, DT_FTOL, DT_MAXITER)),
        "starts_program": ("barycentre + %d Dirichlet(1) + %d zoomed edge points (and SLSQP from each) "
                           "+ the best face point" % (DT_STARTS_PROGRAM, DT_EDGE_STARTS)),
        "starts_bounds": ("barycentre + %d Dirichlet(1) + best face point + best vertex (W_min) or %d "
                          "zoomed edge points (W_max)" % (DT_STARTS_BOUNDS, DT_EDGE_STARTS)),
        "start_law": "Dirichlet(1,...,1), uniform on the simplex",
        "p_grid": DT_PGRID, "p_rule": "midpoint rule, p_j = (j - 1/2)/P",
        "x_bins": DT_XBINS, "overlap_rule": "sum over bins of min{dF_QA, dF_i} on [0,5]",
        "variance_rule": "V(w) = w' Sigma w, Sigma the covariance of the expert quantile functions on the p grid",
        "active_threshold": DT_ACTIVE_TOL,
        "triad": ",".join(_dt_triad), "triad_objective": DT_TRIAD_EQ,
        "band_lo": 3.0, "band_hi": 4.0,
        "x_label_ranking": "mean importance of the collective distribution, averaged over the five equity indices",
        "x_label_triad": "importance level",
        "y_label_triad": "collective probability density",
    }),
    "barriers": pd.DataFrame([{
        "id": b[0], "label": b[1], "dimension": b[2], "design_mean": b[3],
        "tau_calibrated": _dt_tau[k], "calibration_refits": int(_dt_calib_it[k]),
        "equal_weight_mean": _dt_ew_mean[k],
        "expert_mean_min": _dt_wmean[:, k].min(), "expert_mean_max": _dt_wmean[:, k].max(),
        "expert_mean_sd": _dt_wmean[:, k].std(),
    } for k, b in enumerate(_DT_BARRIERS)]),
    "dimensions": [
        {"code": "T", "name": "Technology",   "color": "blue"},
        {"code": "O", "name": "Organization", "color": "green"},
        {"code": "E", "name": "Environment",  "color": "amber"},
    ],
    "equity_objectives": [
        {"eq": "EQ1", "name": "EQ1 (rel. range)",     "index": "I1 relative range",          "color": "red"},
        {"eq": "EQ2", "name": "EQ2 (rel. mean dev.)", "index": "I2 relative mean deviation", "color": "blue"},
        {"eq": "EQ3", "name": "EQ3 (coef. var.)",     "index": "I3 coefficient of variation","color": "teal"},
        {"eq": "EQ4", "name": "EQ4 (Gini)",           "index": "I4 Gini",                    "color": "green"},
        {"eq": "EQ5", "name": "EQ5 (Hoover)",         "index": "I5 Hoover",                  "color": "purple"},
    ],
    "panel_experts": pd.DataFrame({"expert": np.arange(1, DT_N + 1), "background": _DT_BACKGROUND,
                                   "leniency": _dt_len}),
    "panel": pd.DataFrame([{
        "expert": i + 1, "background": _DT_BACKGROUND[i], "barrier": _DT_IDS[k],
        "tau": _dt_tau[k], "leniency": _dt_len[i], "bias": _dt_bias[i, k], "noise": _dt_eps[i, k],
        "rating": _dt_rating[i, k],
        **{"p%d" % (t + 1): _dt_prof[i, k, t] for t in range(5)},
        "kappa": _dt_kap[i, k], "lambda": _dt_lam[i, k], "H5": _dt_H5[i, k],
        "fitted_mean": _dt_wmean[i, k],
    } for i in range(DT_N) for k in range(DT_M)]),
    "elicitation_b1": pd.DataFrame([{"barrier": DT_ELICIT_BARRIER, "expert": _dt_el_i + 1, "level": t + 1,
                                     "level_mid": t + 0.5, "probability": p}
                                    for t, p in enumerate(_dt_el_prof)]),
    "solves": pd.DataFrame(_dt_solves),
    "starts": pd.DataFrame(_dt_starts_rec),
    "overlaps_equal_weights": pd.DataFrame({b: _dt_Aeq[b] for b in _DT_IDS},
                                           index=pd.Index(np.arange(1, DT_N + 1), name="expert")).reset_index(),
    "near_ties": pd.DataFrame(_dt_ties, columns=["eq", "upper", "lower", "upper_mu", "lower_mu", "margin",
                                                 "near_optimal_overlap", "upper_mu_fine", "lower_mu_fine",
                                                 "margin_fine"]),
    "fine_check": pd.DataFrame(_dt_fine),
    "fine_weights": pd.DataFrame({"%s_%s" % c: _dt_fine_W[c] for c in _dt_fine_cells},
                                 index=pd.Index(np.arange(1, DT_N + 1), name="expert")).reset_index(),
    "text_numbers": pd.DataFrame(_dt_tn),
    "weights": pd.DataFrame({c: _dt_W[(c.split("_")[0], c.split("_")[1])] for c in _dt_cols},
                            index=pd.Index(np.arange(1, DT_N + 1), name="expert")).reset_index(),
    "overlaps": pd.DataFrame({c: _dt_A[(c.split("_")[0], c.split("_")[1])] for c in _dt_cols},
                             index=pd.Index(np.arange(1, DT_N + 1), name="expert")).reset_index(),
    "utilities": pd.DataFrame({c: _dt_U[(c.split("_")[0], c.split("_")[1])] for c in _dt_cols},
                              index=pd.Index(np.arange(1, DT_N + 1), name="expert")).reset_index(),
    "matrix": pd.DataFrame([{
        "id": _DT_IDS[k], "dimension": _DT_DIM[k],
        **{e: _dt_Mx[k, j] for j, e in enumerate(_DT_EQ)},
        "avg": _dt_avg[k], "rank": int(_dt_rank[k]),
        "spread": float(_dt_Mx[k].max() - _dt_Mx[k].min()),
        "eq1_departure": float(_dt_eq1_dep[k]), "span_eq2_eq5": float(_dt_span4[k]),
        "abs_eq2_minus_eq5": float(_dt_d25[k]),
        "design_mean": _DT_DESIGN[k], "equal_weight_mean": _dt_ew_mean[k],
        "n_active_min": min(d["n_active"] for d in _dt_solves if d["barrier"] == _DT_IDS[k]),
        "n_active_max": max(d["n_active"] for d in _dt_solves if d["barrier"] == _DT_IDS[k]),
    } for k in range(DT_M)]),
    "ranking": pd.DataFrame([{"rank": p, "id": _DT_IDS[r], "avg": _dt_avg[r]}
                             for p, r in enumerate(_dt_order, 1)]),
    "per_objective_order": pd.DataFrame([{"eq": e, "order": o} for e, o in _dt_eq_orders.items()]),
    "triad_density": pd.DataFrame(_dt_dens),
})
print("   Case II solved in %.0f s" % _dt_elapsed)

# =====================================================================
# SYNTHETIC-DATA TABLES  (data only -- no plotting script)
# =====================================================================
# The four table-only workbooks are written from solved values by their own
# blocks below: tab2-1 and tab2-5 by the Chapter 2 block, tab6-9 by the Case I
# fairness-panel block, tab5-3 by the ex:hmpanel block.  Nothing is typed here.


# =====================================================================
# =====================================================================
#  ADDITIONS 2026-08-28 -- the six figures promised in the proposal
#  review response, plus the workbook of the new Chapter 6 case study
#  (Case IV, rail-transit station accessibility with stakeholder
#  equity).  All heavy computation lives HERE so that every number a
#  figure plots or the text prints is carried by a workbook; the figure
#  scripts read and plot only.  Everything below is deterministic: the
#  generator seed is 202603 and no other entropy source is used.
# =====================================================================
# =====================================================================
import math
import numpy as np

NEW_SEED = 202603
_trapz = getattr(np, "trapezoid", np.trapz)


# =====================================================================
# WORKED EXAMPLES OF THE THEORY CHAPTERS (Chapters 2 to 5)
#   The short examples that Chapters 2 to 5 work by hand, computed here
#   in closed form from their stated inputs by the book's own definitions
#   and written to a sheet 'worked_examples' of one workbook of each chapter:
#     Chapter 2 -> fig2-2_three_pooling_operators.xlsx
#       * the correlated two-expert instance of the generalized least-squares
#         weights of sec:foundations-suprabayesian (sigma_1 = 1, sigma_2 = 2,
#         rho = 0.8, reports 5 and 6)
#       * the three-expert DeGroot trust matrix of sec:foundations-related and
#         its stationary distribution
#     Chapter 3 -> fig3-1_lorenz_gini_hoover.xlsx
#       * Example ex:bonferroni (Phi_{1,1} of two four-expert profiles)
#       * Example ex:skewed-panel (def:ifcu with heterogeneous dispositions,
#         eq:parabola-norm at kappa = 2, eq:bonferroni-closed)
#       * the three-expert welfare illustration of sec:conscience-welfare
#         (def:ifcu, eq:exp-norm, the five indices of tab:swf, the complete
#         welfare W_{1,4} of eq:complete-welfare) and the Atkinson index at
#         unit aversion of the same profile
#       * the leveling-down illustration after prop:consensus-fairness
#       * the ultimatum-game identification alpha = s / (1 - 2 s) of
#         sec:conscience-elicitation
#     Chapter 4 -> fig4-3_scalability.xlsx
#       * Example ex:unbalanced-clustering (the intra-group interaction
#         counts sum_k n_k (n_k - 1) of a balanced and a skewed partition of
#         N = 90 experts into K = 3 subgroups, against the flat N (N - 1))
#       * the fuzzy C-means memberships of sec:scalability-hierarchical
#         (fuzzifier m = 2, an expert at distances 1 and 2 from the two
#         centroids, and an expert equidistant from both)
#       (the illustrative epigraph-size count of sec:scalability-convex is
#       written by the Figure 4.3 block itself, sheet lp_reformulation_size)
#     Chapter 5 -> fig5-4_depnet.xlsx
#       * the equicorrelated ceiling 1/r at r = 0.7 and Example ex:neff
#       * the lineage count 5 / [1 + 4 (0.85)] of Example ex:corr-lineage
#   The stated inputs are the only typed numbers of this block.  Every
#   computed row carries a second, independent route to the same value
#   (column cross_check: the closed form of eq:bonferroni-closed against the
#   double sum of def:collective-fairness, matrix algebra against the scalar
#   formulas, order statistics and the Lorenz curve against the pairwise
#   sums, the ordered pairs enumerated on the label matrix against the
#   interaction-count formula, numpy's vectorized membership update against
#   the exact fraction, and so on), asserted equal here; the decimals at which the text
#   prints it; and the printed string (half-up rounding, or the exact
#   fraction where the text prints a fraction).  No random draws.
# =====================================================================
from fractions import Fraction as _WxF
from decimal import Decimal as _WxD, ROUND_HALF_UP as _WX_HALF_UP
from scipy.optimize import brentq as _wx_brentq

_WX_ROWS = []
_WX_CHAPTER_WORKBOOK = {2: "fig2-2_three_pooling_operators.xlsx", 3: "fig3-1_lorenz_gini_hoover.xlsx",
                        4: "fig4-3_scalability.xlsx", 5: "fig5-4_depnet.xlsx"}


def _wx_fmt(v, d):
    """v rounded half-up to d decimals, as the text prints it (no negative zero)."""
    q = _WxD(repr(round(float(v), 12))).quantize(_WxD(1).scaleb(-int(d)), rounding=_WX_HALF_UP)
    return str(abs(q) if q == 0 else q)


def _wx_add(ch, example, section, quantity, value, d=None, role="computed", cross=None, rule="", tol=1e-12):
    """One row of a worked_examples sheet.  d = the decimals the text prints, 'fraction'
    where it prints the exact fraction (value then a Fraction), None where the value is
    not printed as a numeral.  A computed row must agree with its cross-check."""
    frac = value if isinstance(value, _WxF) else None
    v = float(value)
    if d is None:
        dec, printed = "", ""
    elif d == "fraction":
        assert frac is not None, "%s / %s: a fraction is printed but the value is not exact" % (example, quantity)
        dec, printed = "fraction", str(frac)
    else:
        dec, printed = str(int(d)), _wx_fmt(v, d)
    if role == "computed":
        assert cross is not None, "%s / %s: a computed row needs a cross-check" % (example, quantity)
        assert abs(v - float(cross)) <= tol, "%s / %s: %.15g, cross-check %.15g" % (example, quantity, v, float(cross))
    else:
        assert role == "stated input" and cross is None, (example, quantity, role)
    _WX_ROWS.append({"chapter": ch, "example": example, "section": section, "quantity": quantity,
                     "role": role, "value": v, "cross_check": float(cross) if cross is not None else float("nan"),
                     "decimals": dec, "printed": printed, "rule": rule})
    return value


def _wx_input(ch, example, section, quantity, value, d, rule="typed, stated in the text"):
    return _wx_add(ch, example, section, quantity, value, d, "stated input", rule=rule)


def _wx_sheet(ch):
    rows = [r for r in _WX_ROWS if r["chapter"] == ch]
    assert rows, "no worked example of Chapter %d" % ch
    return pd.DataFrame(rows, columns=["chapter", "example", "section", "quantity", "role", "value",
                                       "cross_check", "decimals", "printed", "rule"])


def _wx_solve_exact(M, b):
    """Gauss-Jordan elimination in exact rational arithmetic (M square and nonsingular)."""
    n = len(M)
    T = [list(M[i]) + [b[i]] for i in range(n)]
    for c in range(n):
        p = next(r for r in range(c, n) if T[r][c] != 0)
        T[c], T[p] = T[p], T[c]
        T[c] = [x / T[c][c] for x in T[c]]
        for r in range(n):
            if r != c and T[r][c] != 0:
                f = T[r][c]
                T[r] = [x - f * y for x, y in zip(T[r], T[c])]
    return [T[i][n] for i in range(n)]


def _wx_ifcu(A, rho, theta):
    """def:ifcu as displayed: the two comparison sums over j != i, divisor N - 1."""
    N = len(A)
    out = []
    for i in range(N):
        dis = sum(max(A[j] - A[i], 0.0) for j in range(N) if j != i)
        adv = sum(max(A[i] - A[j], 0.0) for j in range(N) if j != i)
        out.append(A[i] - rho[i] / (N - 1) * dis - theta[i] / (N - 1) * adv)
    return out


def _wx_ifcu_sorted(A, rho, theta):
    """Cross-check of def:ifcu through the order statistics: the disadvantage sum of expert
    i is (sum of the larger overlaps) - (their number) A_i, the advantage sum likewise."""
    A = np.asarray(A, float)
    n = A.size
    out = np.empty(n)
    for i in range(n):
        hi, lo = A[A > A[i]], A[A < A[i]]
        out[i] = (A[i] - rho[i] * (hi.sum() - hi.size * A[i]) / (n - 1)
                  - theta[i] * (lo.size * A[i] - lo.sum()) / (n - 1))
    return out


def _wx_power_norm(U, A, kappa):
    """eq:parabola-norm, branch by branch."""
    if A == 0 or U < 0:
        return 0.0
    return 1.0 if U > A else (U / A) ** kappa


def _wx_exp_norm(U, A):
    """eq:exp-norm, branch by branch."""
    if A == 0 or U < 0:
        return 0.0
    return 1.0 if U > A else (math.exp(U / A) - 1.0) / (math.e - 1.0)


def _wx_bonf_double(U):
    """def:collective-fairness at a = b = 1: the double sum over i != j."""
    N = len(U)
    return math.sqrt(sum(U[i] * U[j] for i in range(N) for j in range(N) if i != j) / (N * (N - 1)))


def _wx_bonf_closed(U):
    """eq:bonferroni-closed: the two running sums S_1 and S_2."""
    U = np.asarray(U, float)
    N = U.size
    return math.sqrt(max((U.sum() ** 2 - (U ** 2).sum()) / (N * (N - 1)), 0.0))


def _wx_indices(U):
    """The five inequality measures of tab:swf, as displayed."""
    U = np.asarray(U, float)
    N, m = U.size, U.mean()
    return {1: (U.max() - U.min()) / m, 2: np.abs(U - m).sum() / (N * m),
            3: math.sqrt(((U - m) ** 2).sum() / N) / m,
            4: np.abs(U[:, None] - U[None, :]).sum() / (2.0 * N * N * m),
            5: np.abs(U - m).sum() / (2.0 * N * m)}


def _wx_indices_cross(U):
    """Independent routes to the five measures: order statistics and the Lorenz curve
    (the Hoover index is its largest gap, the relative mean deviation twice that, the
    Gini index sum_i (2i - N - 1) x_(i) / (N^2 mean)), numpy's population sd."""
    x = np.sort(np.asarray(U, float))
    N, m = x.size, x.mean()
    L = np.concatenate([[0.0], np.cumsum(x) / x.sum()])
    gap = float(np.max(np.arange(N + 1) / N - L))
    return {1: float(np.ptp(x)) / m, 2: 2.0 * gap, 3: float(np.std(x)) / m,
            4: float(np.sum((2.0 * np.arange(1, N + 1) - N - 1) * x)) / (N * N * m), 5: gap}


def _wx_neff(R, w):
    """def:neff: N_eff = 1 / (w' R w)."""
    w = np.asarray(w, float)
    return 1.0 / float(w @ R @ w)


# ---- Chapter 2: the generalized least-squares weights of two correlated experts
_wx_e, _wx_s = "GLS weights of two correlated experts", "sec:foundations-suprabayesian"
_wx_s1 = _wx_input(2, _wx_e, _wx_s, "sigma_1", _WxF(1), 0)
_wx_s2 = _wx_input(2, _wx_e, _wx_s, "sigma_2", _WxF(2), 0)
_wx_rho = _wx_input(2, _wx_e, _wx_s, "rho", _WxF("0.8"), 1)
_wx_r1 = _wx_input(2, _wx_e, _wx_s, "r_1", _WxF(5), 0)
_wx_r2 = _wx_input(2, _wx_e, _wx_s, "r_2", _WxF(6), 0)
_wx_Sig = np.array([[float(_wx_s1 ** 2), float(_wx_rho * _wx_s1 * _wx_s2)],
                    [float(_wx_rho * _wx_s1 * _wx_s2), float(_wx_s2 ** 2)]])
_wx_Si = np.linalg.inv(_wx_Sig)
_wx_eSe = float(np.ones(2) @ _wx_Si @ np.ones(2))
_wx_det = float(np.linalg.det(_wx_Sig))
_wx_wg = _wx_Si @ np.ones(2) / _wx_eSe
_wx_add(2, _wx_e, _wx_s, "sigma_1^2", _wx_s1 ** 2, 0, cross=_wx_Sig[0, 0], rule="first term of the denominator")
_wx_add(2, _wx_e, _wx_s, "sigma_2^2", _wx_s2 ** 2, 0, cross=_wx_Sig[1, 1], rule="second term of the denominator")
_wx_add(2, _wx_e, _wx_s, "2 rho sigma_1 sigma_2", 2 * _wx_rho * _wx_s1 * _wx_s2, 1, cross=2.0 * _wx_Sig[0, 1],
        rule="twice the error covariance")
_wx_D = _wx_add(2, _wx_e, _wx_s, "denominator", _wx_s1 ** 2 + _wx_s2 ** 2 - 2 * _wx_rho * _wx_s1 * _wx_s2, 1,
                cross=float(np.array([1.0, -1.0]) @ _wx_Sig @ np.array([1.0, -1.0])),
                rule="sigma_1^2 + sigma_2^2 - 2 rho sigma_1 sigma_2; cross-check: Var(eps_1 - eps_2) = d' Sigma d, d = (1, -1)")
_wx_n1 = _wx_add(2, _wx_e, _wx_s, "numerator of w_1", _wx_s2 ** 2 - _wx_rho * _wx_s1 * _wx_s2, 1,
                 cross=_wx_det * (_wx_Si @ np.ones(2))[0], rule="sigma_2^2 - rho sigma_1 sigma_2; cross-check: det(Sigma) (Sigma^-1 e)_1")
_wx_n2 = _wx_add(2, _wx_e, _wx_s, "numerator of w_2", _wx_s1 ** 2 - _wx_rho * _wx_s1 * _wx_s2, 1,
                 cross=_wx_det * (_wx_Si @ np.ones(2))[1], rule="sigma_1^2 - rho sigma_1 sigma_2; cross-check: det(Sigma) (Sigma^-1 e)_2")
_wx_w1 = _wx_add(2, _wx_e, _wx_s, "w_1", _wx_n1 / _wx_D, "fraction", cross=_wx_wg[0],
                 rule="exact; cross-check: Sigma^-1 e / (e' Sigma^-1 e) by numpy.linalg.inv")
_wx_w2 = _wx_add(2, _wx_e, _wx_s, "w_2", _wx_n2 / _wx_D, "fraction", cross=_wx_wg[1],
                 rule="exact; cross-check: Sigma^-1 e / (e' Sigma^-1 e) by numpy.linalg.inv")
_wx_add(2, _wx_e, _wx_s, "sigma_1^2 sigma_2^2 (1 - rho^2)", _wx_s1 ** 2 * _wx_s2 ** 2 * (1 - _wx_rho ** 2), 2,
        cross=_wx_det, rule="cross-check: det(Sigma)")
_wx_add(2, _wx_e, _wx_s, "diffuse-prior posterior variance", _wx_s1 ** 2 * _wx_s2 ** 2 * (1 - _wx_rho ** 2) / _wx_D, 1,
        cross=1.0 / _wx_eSe, rule="sigma_1^2 sigma_2^2 (1 - rho^2) / denominator; cross-check: 1 / (e' Sigma^-1 e)")
_wx_est = _wx_w1 * _wx_r1 + _wx_w2 * _wx_r2
_wx_add(2, _wx_e, _wx_s, "combined estimate (fraction)", _wx_est, "fraction",
        cross=float(_wx_wg @ np.array([float(_wx_r1), float(_wx_r2)])), rule="w_1 r_1 + w_2 r_2, exact")
_wx_add(2, _wx_e, _wx_s, "combined estimate", _wx_est, 2,
        cross=float(_wx_wg @ np.array([float(_wx_r1), float(_wx_r2)])), rule="w_1 r_1 + w_2 r_2, two decimals")
_wx_add(2, _wx_e, _wx_s, "correlation at which w_2 turns negative", _wx_s1 / _wx_s2, None,
        cross=float(_wx_s1 ** 2) / float(_wx_s1 * _wx_s2), rule="sigma_1 / sigma_2 (the text states it symbolically)")

# ---- Chapter 2: the three-expert DeGroot trust matrix and its stationary distribution
_wx_e, _wx_s = "DeGroot trust matrix", "sec:foundations-related"
_wx_W = [[_WxF(1, 2), _WxF(1, 2), _WxF(0)], [_WxF(1, 4), _WxF(1, 2), _WxF(1, 4)], [_WxF(0), _WxF(1, 2), _WxF(1, 2)]]
for _i in range(3):
    for _j in range(3):
        _wx_input(2, _wx_e, _wx_s, "W[%d,%d]" % (_i + 1, _j + 1), _wx_W[_i][_j], "fraction")
_wx_M = [[_wx_W[_j][_i] - (1 if _i == _j else 0) for _j in range(3)] for _i in range(2)] + [[_WxF(1)] * 3]
_wx_pi = _wx_solve_exact(_wx_M, [_WxF(0), _WxF(0), _WxF(1)])
_wx_Wp = np.linalg.matrix_power(np.array([[float(x) for x in r] for r in _wx_W]), 256)
for _i in range(3):
    _wx_add(2, _wx_e, _wx_s, "pi_%d" % (_i + 1), _wx_pi[_i], "fraction", cross=_wx_Wp[0, _i],
            rule="pi' W = pi', sum pi = 1, exact; cross-check: first row of W^256")
_wx_add(2, _wx_e, _wx_s, "pi_2 / pi_1", _wx_pi[1] / _wx_pi[0], None, cross=_wx_Wp[0, 1] / _wx_Wp[0, 0],
        rule="the text: the second expert weighted twice as heavily as either neighbor")

# ---- Chapter 3: Example ex:bonferroni
_wx_e, _wx_s = "ex:bonferroni", "sec:conscience-fairness"
_wx_UA = [_wx_input(3, _wx_e, _wx_s, "U_A[%d]" % (_i + 1), _v, 1) for _i, _v in enumerate((0.8, 0.8, 0.8, 0.0))]
_wx_UB = [_wx_input(3, _wx_e, _wx_s, "U_B[%d]" % (_i + 1), _v, 1) for _i, _v in enumerate((0.6, 0.6, 0.6, 0.6))]
_wx_add(3, _wx_e, _wx_s, "arithmetic mean of A (and of B)", float(np.mean(_wx_UA)), 1, cross=float(np.mean(_wx_UB)),
        rule="the two profiles share the mean; cross-check: the mean of B")
_wx_add(3, _wx_e, _wx_s, "Phi_11(A)", _wx_bonf_double(_wx_UA), 3, cross=_wx_bonf_closed(_wx_UA),
        rule="def:collective-fairness, a = b = 1 (double sum); cross-check: eq:bonferroni-closed")
_wx_add(3, _wx_e, _wx_s, "Phi_11(B)", _wx_bonf_double(_wx_UB), 3, cross=_wx_bonf_closed(_wx_UB),
        rule="def:collective-fairness, a = b = 1 (double sum); cross-check: eq:bonferroni-closed")

# ---- Chapter 3: Example ex:skewed-panel
_wx_e = "ex:skewed-panel"
_wx_A5 = [_wx_input(3, _wx_e, _wx_s, "A[%d]" % (_i + 1), _v, 2) for _i, _v in enumerate((0.95, 0.30, 0.22, 0.18, 0.15))]
_wx_r5 = [_wx_input(3, _wx_e, _wx_s, "rho[%d]" % (_i + 1), _v, 1) for _i, _v in enumerate((0.2, 0.5, 0.6, 0.7, 0.8))]
_wx_t5 = [_wx_input(3, _wx_e, _wx_s, "vartheta[%d]" % (_i + 1), _v, 1) for _i, _v in enumerate((-0.3, 0.1, 0.2, 0.3, 0.4))]
_wx_k = _wx_input(3, _wx_e, _wx_s, "kappa", 2, 0)
_wx_U5 = _wx_ifcu(_wx_A5, _wx_r5, _wx_t5)
_wx_U5x = _wx_ifcu_sorted(_wx_A5, _wx_r5, _wx_t5)
for _i in range(5):
    _wx_add(3, _wx_e, _wx_s, "U_FS[%d]" % (_i + 1), _wx_U5[_i], 3, cross=_wx_U5x[_i],
            rule="def:ifcu; cross-check: the comparison sums through the order statistics")
_wx_N5 = [_wx_power_norm(_wx_U5[_i], _wx_A5[_i], _wx_k) for _i in range(5)]
_wx_N5x = np.clip(np.asarray(_wx_U5x) / np.asarray(_wx_A5), 0.0, 1.0) ** _wx_k
for _i in range(5):
    _wx_add(3, _wx_e, _wx_s, "normalized U_FS[%d]" % (_i + 1), _wx_N5[_i], 3, cross=_wx_N5x[_i],
            rule="eq:parabola-norm at kappa = 2, branch by branch; cross-check: clip(U/A, 0, 1)^kappa")
_wx_add(3, _wx_e, _wx_s, "arithmetic mean of the normalized profile", float(np.mean(_wx_N5)), 3,
        cross=float(np.sum(_wx_N5x)) / 5.0, rule="mean of the normalized profile")
_wx_add(3, _wx_e, _wx_s, "Phi_11", _wx_bonf_closed(_wx_N5), 3, cross=_wx_bonf_double(_wx_N5),
        rule="eq:bonferroni-closed (as the text computes it); cross-check: the double sum of def:collective-fairness")

# ---- Chapter 3: the three-expert welfare illustration and its Atkinson index
_wx_e, _wx_s = "three-expert welfare profile", "sec:conscience-welfare"
_wx_A3 = [_wx_input(3, _wx_e, _wx_s, "A[%d]" % (_i + 1), _v, 1) for _i, _v in enumerate((0.9, 0.6, 0.3))]
_wx_rh3 = _wx_input(3, _wx_e, _wx_s, "rho (every expert)", 0.4, 1)
_wx_th3 = _wx_input(3, _wx_e, _wx_s, "vartheta (every expert)", 0.2, 1)
_wx_U3 = _wx_ifcu(_wx_A3, [_wx_rh3] * 3, [_wx_th3] * 3)
_wx_U3x = _wx_ifcu_sorted(_wx_A3, [_wx_rh3] * 3, [_wx_th3] * 3)
for _i in range(3):
    _wx_add(3, _wx_e, _wx_s, "U_FS[%d]" % (_i + 1), _wx_U3[_i], 2, cross=_wx_U3x[_i],
            rule="def:ifcu; cross-check: the comparison sums through the order statistics")
for _i in range(3):
    _wx_add(3, _wx_e, _wx_s, "U_FS[%d] / A[%d]" % (_i + 1, _i + 1), _wx_U3[_i] / _wx_A3[_i], 2,
            cross=_wx_U3x[_i] / _wx_A3[_i], rule="the ratio the text compares across the experts")
_wx_E3 = [_wx_exp_norm(_wx_U3[_i], _wx_A3[_i]) for _i in range(3)]
_wx_E3x = (np.exp(np.clip(_wx_U3x / np.asarray(_wx_A3), 0.0, 1.0)) - 1.0) / (math.e - 1.0)
for _i in range(3):
    _wx_add(3, _wx_e, _wx_s, "U[%d]" % (_i + 1), _wx_E3[_i], 3, cross=_wx_E3x[_i],
            rule="eq:exp-norm, branch by branch; cross-check: (exp(clip(U/A, 0, 1)) - 1) / (e - 1)")
_wx_m3 = float(np.mean(_wx_E3))
_wx_add(3, _wx_e, _wx_s, "mean U", _wx_m3, 3, cross=float(np.sum(_wx_E3x)) / 3.0,
        rule="the utilitarian mean of the normalized profile")
_wx_add(3, _wx_e, _wx_s, "minimum U", min(_wx_E3), 3, cross=float(np.sort(_wx_E3x)[0]),
        rule="the egalitarian minimum")
_wx_I3, _wx_I3x = _wx_indices(_wx_E3), _wx_indices_cross(_wx_E3x)
for _k in range(1, 6):
    _wx_add(3, _wx_e, _wx_s, "EQ_%d" % _k, -_wx_I3[_k], 3, cross=-_wx_I3x[_k],
            rule="EQ_k = -I_k of tab:swf as displayed; cross-check: order statistics and the Lorenz curve")
_wx_add(3, _wx_e, _wx_s, "I_4", _wx_I3[4], 3, cross=_wx_I3x[4], rule="the Gini index of the profile")
_wx_lam = _wx_input(3, _wx_e, _wx_s, "lambda", 1, 0)
_wx_add(3, _wx_e, _wx_s, "W_{1,4}", _wx_m3 - _wx_lam * _wx_I3[4], 3, cross=float(np.mean(_wx_E3x)) - _wx_lam * _wx_I3x[4],
        rule="eq:complete-welfare, W_{lambda,k} = Ubar - lambda I_k at lambda = 1, k = 4")
_wx_e = "Atkinson index of the three-expert profile"
_wx_eps = _wx_input(3, _wx_e, _wx_s, "varepsilon_Atk", 1, 0)
_wx_gm = float(np.prod(_wx_E3)) ** (1.0 / 3.0)
_wx_add(3, _wx_e, _wx_s, "geometric mean", _wx_gm, 3, cross=math.exp(float(np.mean(np.log(_wx_E3x)))),
        rule="the equally-distributed-equivalent level at unit aversion; cross-check: exp(mean log U)")
_wx_gen = math.exp(math.log(float(np.mean(np.asarray(_wx_E3x) ** 1e-7))) / 1e-7)
_wx_add(3, _wx_e, _wx_s, "I_Atk(1)", 1.0 - _wx_gm / _wx_m3, 3, cross=1.0 - _wx_gen / float(np.mean(_wx_E3x)),
        rule="1 - geometric mean / mean; cross-check: the generalized mean of order 1 - varepsilon at varepsilon = 1 - 1e-7",
        tol=1e-6)

# ---- Chapter 3: the leveling-down illustration after prop:consensus-fairness
_wx_e = "leveling-down four-expert panel"
_wx_Al = [_wx_input(3, _wx_e, _wx_s, "A[%d]" % (_i + 1), _WxF(_v), 1) for _i, _v in enumerate(("0.9", "0.9", "0.9", "0.1"))]
_wx_Ap = [_wx_input(3, _wx_e, _wx_s, "A'[%d]" % (_i + 1), _WxF(_v), 1) for _i, _v in enumerate(("0.4", "0.4", "0.4", "0.4"))]


def _wx_gini_exact(a):
    n, m = len(a), sum(a) / len(a)
    return sum(abs(x - y) for x in a for y in a) / (2 * n * n * m)


_wx_mA, _wx_mAp = sum(_wx_Al) / 4, sum(_wx_Ap) / 4
_wx_GA, _wx_GAp = _wx_gini_exact(_wx_Al), _wx_gini_exact(_wx_Ap)
_wx_Alx, _wx_Apx = np.array([float(x) for x in _wx_Al]), np.array([float(x) for x in _wx_Ap])
_wx_add(3, _wx_e, _wx_s, "mean A", _wx_mA, 2, cross=float(_wx_Alx.mean()), rule="exact")
_wx_add(3, _wx_e, _wx_s, "I_4(A)", _wx_GA, 3, cross=_wx_indices_cross(_wx_Alx)[4], rule="exact pairwise sum; cross-check: order statistics")
_wx_add(3, _wx_e, _wx_s, "-I_4(A)", -_wx_GA, 3, cross=-_wx_indices_cross(_wx_Alx)[4], rule="the bare equity objective")
_wx_add(3, _wx_e, _wx_s, "mean A'", _wx_mAp, 2, cross=float(_wx_Apx.mean()), rule="exact")
_wx_add(3, _wx_e, _wx_s, "I_4(A')", _wx_GAp, 0, cross=_wx_indices_cross(_wx_Apx)[4], rule="exact (perfect equity)")
_wx_add(3, _wx_e, _wx_s, "drop in mean overlap", _wx_mA - _wx_mAp, 2, cross=float(_wx_Alx.mean() - _wx_Apx.mean()), rule="exact")
_wx_add(3, _wx_e, _wx_s, "threshold lambda", (_wx_mA - _wx_mAp) / (_wx_GA - _wx_GAp), 2,
        cross=_wx_brentq(lambda l: (float(_wx_Alx.mean()) - l * _wx_indices_cross(_wx_Alx)[4])
                         - (float(_wx_Apx.mean()) - l * _wx_indices_cross(_wx_Apx)[4]), 0.0, 100.0, xtol=1e-15),
        rule="drop / gain in I_4, exact; cross-check: the lambda at which Abar - lambda I_4 ties on the two profiles (brentq)",
        tol=1e-12)

# ---- Chapter 3: the ultimatum-game identification of the jealousy coefficient
_wx_e, _wx_s = "ultimatum-game identification", "sec:conscience-elicitation"
for _tag, _sv, _ds, _da in (("thirty percent", "0.3", 1, 2), ("a quarter", "0.25", 2, 1)):
    _shat = _wx_input(3, _wx_e, _wx_s, "s_hat (%s)" % _tag, _WxF(_sv), _ds)
    _den = _wx_add(3, _wx_e, _wx_s, "1 - 2 s_hat (%s)" % _tag, 1 - 2 * _shat, 1 if _tag == "thirty percent" else None,
                   cross=1.0 - 2.0 * float(_sv), rule="the denominator of alpha = s / (1 - 2 s)")
    _wx_add(3, _wx_e, _wx_s, "alpha (%s)" % _tag, _shat / _den, _da,
            cross=_wx_brentq(lambda a: float(_sv) - a * (1.0 - 2.0 * float(_sv)), 0.0, 100.0, xtol=1e-15),
            rule="alpha = s / (1 - 2 s), exact; cross-check: the alpha at which accepting s leaves the responder indifferent (brentq)")

# ---- Chapter 4: Example ex:unbalanced-clustering (the intra-group interaction counts)
_wx_e, _wx_s = "ex:unbalanced-clustering", "sec:scalability-hierarchical"
_wx4_N = _wx_input(4, _wx_e, _wx_s, "N", 90, 0)
_wx4_K = _wx_input(4, _wx_e, _wx_s, "K", 3, 0)
_wx4_sk = [_wx_input(4, _wx_e, _wx_s, "skewed size n_%d" % (_k + 1), _v, 0) for _k, _v in enumerate((80, 5, 5))]
assert len(_wx4_sk) == _wx4_K and sum(_wx4_sk) == _wx4_N, "the skewed partition is not K subgroups summing to N"
assert _wx4_sk[1] == _wx4_sk[2], "the text groups the two small subgroups as 2 x n_2 (n_2 - 1)"


def _wx4_pairs(sizes):
    """sum_k n_k (n_k - 1): the ordered pairs i != j inside the subgroups (the fairness
    measure of def:collective-fairness sums over ordered pairs)."""
    return sum(n * (n - 1) for n in sizes)


def _wx4_labels(sizes):
    return np.repeat(np.arange(len(sizes)), sizes)


def _wx4_pairs_enum(sizes):
    """Cross-check: the ordered pairs i != j that carry the same subgroup label, counted on
    the N x N label-equality matrix."""
    lab = _wx4_labels(sizes)
    return int(np.sum(lab[:, None] == lab[None, :])) - lab.size


assert _wx4_N % _wx4_K == 0, "N / K is not an integer, so no partition is exactly balanced"
_wx4_split = [len(_a) for _a in np.array_split(np.arange(_wx4_N), _wx4_K)]
_wx4_bal = [_wx_add(4, _wx_e, _wx_s, "balanced size n_%d" % (_k + 1), _wx4_N // _wx4_K, 0, cross=_wx4_split[_k],
                    rule="N / K; cross-check: numpy.array_split of the N experts into K parts")
            for _k in range(_wx4_K)]
_wx4_lb, _wx4_ls = _wx4_labels(_wx4_split), _wx4_labels(_wx4_sk)
_wx_add(4, _wx_e, _wx_s, "balanced n_k - 1", _wx4_bal[0] - 1, 0, cross=int(np.sum(_wx4_lb == _wx4_lb[0])) - 1,
        rule="the other members of a balanced subgroup; cross-check: the experts sharing the first expert's label, less one")
_wx4_cb = _wx_add(4, _wx_e, _wx_s, "balanced intra-group count", _wx4_K * _wx4_bal[0] * (_wx4_bal[0] - 1), 0,
                  cross=_wx4_pairs_enum(_wx4_split),
                  rule="K n (n - 1) = sum_k n_k (n_k - 1) at the balanced sizes; cross-check: the ordered pairs "
                       "enumerated on the label matrix")
assert _wx4_cb == _wx4_pairs(_wx4_bal)
_wx4_off = ~np.eye(_wx4_N, dtype=bool)
_wx_add(4, _wx_e, _wx_s, "N - 1", _wx4_N - 1, 0, cross=int(np.sum(_wx4_off[0])),
        rule="the other experts of the panel; cross-check: the off-diagonal entries of one row of the N x N pair matrix")
_wx4_cf = _wx_add(4, _wx_e, _wx_s, "flat count", _wx4_N * (_wx4_N - 1), 0, cross=int(np.sum(_wx4_off)),
                  rule="N (N - 1), the ordered pairs of the single-stage program; cross-check: the off-diagonal "
                       "entries of the N x N pair matrix")
_wx_add(4, _wx_e, _wx_s, "balanced / flat", _WxF(_wx4_cb, _wx4_cf), None,
        cross=_wx4_pairs_enum(_wx4_split) / float(np.sum(_wx4_off)),
        rule="exact; the text: a reduction to one third; cross-check: the ratio of the enumerated counts")
_wx_add(4, _wx_e, _wx_s, "skewed n_1 - 1", _wx4_sk[0] - 1, 0, cross=int(np.sum(_wx4_ls == 0)) - 1,
        rule="the other members of the dominant subgroup; cross-check: the experts carrying its label, less one")
_wx4_m2 = _wx_add(4, _wx_e, _wx_s, "skewed subgroups of size n_2", sum(1 for _n in _wx4_sk if _n == _wx4_sk[1]), 0,
                  cross=int(np.bincount(_wx4_sk)[_wx4_sk[1]]),
                  rule="the number of subgroups of the size n_2 = n_3 (the factor 2 of the text); cross-check: "
                       "numpy.bincount of the sizes")
_wx_add(4, _wx_e, _wx_s, "skewed n_2 - 1", _wx4_sk[1] - 1, 0, cross=int(np.sum(_wx4_ls == 1)) - 1,
        rule="the other members of a small subgroup; cross-check: the experts carrying the second label, less one")
_wx4_cs = _wx_add(4, _wx_e, _wx_s, "skewed intra-group count",
                  _wx4_sk[0] * (_wx4_sk[0] - 1) + _wx4_m2 * _wx4_sk[1] * (_wx4_sk[1] - 1), 0,
                  cross=_wx4_pairs_enum(_wx4_sk),
                  rule="n_1 (n_1 - 1) + 2 n_2 (n_2 - 1), as the text groups it; cross-check: the ordered pairs "
                       "enumerated on the label matrix")
assert _wx4_cs == _wx4_pairs(_wx4_sk)
_wx_add(4, _wx_e, _wx_s, "1 - skewed / flat", 1 - _WxF(_wx4_cs, _wx4_cf), None,
        cross=1.0 - _wx4_pairs_enum(_wx4_sk) / float(np.sum(_wx4_off)),
        rule="exact; the text: only a fifth below the flat cost; cross-check: from the enumerated counts")
_wx_add(4, _wx_e, _wx_s, "skewed / balanced", _WxF(_wx4_cs, _wx4_cb), None,
        cross=_wx4_pairs_enum(_wx4_sk) / float(_wx4_pairs_enum(_wx4_split)),
        rule="exact; the text: far above the balanced figure; cross-check: from the enumerated counts")
_wx_add(4, _wx_e, _wx_s, "share of the largest subgroup in the skewed count",
        _WxF(_wx4_sk[0] * (_wx4_sk[0] - 1), _wx4_cs), None,
        cross=float(np.sum((_wx4_ls[:, None] == 0) & (_wx4_ls[None, :] == 0)) - _wx4_sk[0]) / _wx4_pairs_enum(_wx4_sk),
        rule="exact; the text: the count is governed by the largest subgroup; cross-check: its pairs on the label matrix")

# ---- Chapter 4: the fuzzy C-means memberships of sec:scalability-hierarchical
_wx_e = "fuzzy C-means memberships"
_wx4_m = _wx_input(4, _wx_e, _wx_s, "m", _WxF(2), 0)
_wx4_Kf = _wx_input(4, _wx_e, _wx_s, "number of centroids", 2, None, rule="typed, stated in words (one centroid and the other)")
_wx4_d = [_wx_input(4, _wx_e, _wx_s, "distance to the %s centroid" % _w, _WxF(_v), 0)
          for _w, _v in (("nearer", 1), ("farther", 2))]
assert len(_wx4_d) == _wx4_Kf and _wx4_d[0] < _wx4_d[1]


def _wx4_fcm_update(D, m):
    """The membership update of the fuzzy C-means program as displayed,
    u_ik = 1 / sum_k' (d_ik / d_ik')^(2/(m-1)), vectorized by numpy over the experts
    (rows of D) and the subgroups (columns)."""
    D = np.asarray(D, float)
    return 1.0 / np.sum((D[:, :, None] / D[:, None, :]) ** (2.0 / (float(m) - 1.0)), axis=2)


def _wx4_fcm_root(d, m):
    """The membership in the first of two subgroups that minimizes the expert's term of J_m,
    u^m d_1^2 + (1 - u)^m d_2^2 with the centroids fixed: the root in (0, 1) of its
    derivative, by brentq (the update formula is not used)."""
    a, b, m = float(d[0]) ** 2, float(d[1]) ** 2, float(m)
    return _wx_brentq(lambda u: m * u ** (m - 1.0) * a - m * (1.0 - u) ** (m - 1.0) * b, 0.0, 1.0, xtol=1e-15)


_wx4_ustar = _wx4_fcm_root(_wx4_d, _wx4_m)
_wx4_e = _wx_add(4, _wx_e, _wx_s, "exponent 2/(m - 1)", 2 / (_wx4_m - 1), 0,
                 cross=math.log(_wx4_ustar / (1.0 - _wx4_ustar)) / math.log(float(_wx4_d[1] / _wx4_d[0])),
                 rule="2 / (m - 1), exact; cross-check: log(u / (1 - u)) / log(d_2 / d_1) at the u that minimizes "
                      "the expert's term of J_m (brentq on its derivative)", tol=1e-9)
_wx4_D = np.array([[float(_wx4_d[0]), float(_wx4_d[1])], [float(_wx4_d[0]), float(_wx4_d[0])]])
_wx4_U = _wx4_fcm_update(_wx4_D, _wx4_m)
_wx4_r = _wx_add(4, _wx_e, _wx_s, "ratio d_near / d_far", _wx4_d[0] / _wx4_d[1], "fraction",
                 cross=float(np.divide(*_wx4_D[0])),
                 rule="exact, the distance ratio inside the update formula; cross-check: numpy.divide of the two distances")
_wx4_un = _wx_add(4, _wx_e, _wx_s, "membership in the nearer subgroup", 1 / (1 + _wx4_r ** _wx4_e), 1,
                  cross=_wx4_U[0, 0],
                  rule="1 / (1 + (d_1 / d_2)^(2/(m-1))), exact; cross-check: the update formula vectorized by numpy")
_wx4_uf = _wx_add(4, _wx_e, _wx_s, "membership in the farther subgroup", 1 / (1 + (1 / _wx4_r) ** _wx4_e), 1,
                  cross=_wx4_U[0, 1],
                  rule="1 / (1 + (d_2 / d_1)^(2/(m-1))), exact; cross-check: the update formula vectorized by numpy")
_wx_add(4, _wx_e, _wx_s, "membership of an equidistant expert", 1 / (1 + _WxF(1) ** _wx4_e), 1, cross=_wx4_U[1, 0],
        rule="1 / (1 + 1^(2/(m-1))) = 1/K, exact; cross-check: the update formula at two equal distances (numpy)")
assert _wx4_un + _wx4_uf == 1, "the two memberships of the expert do not sum to one"

# ---- Chapter 5: the equicorrelated ceiling, Example ex:neff, Example ex:corr-lineage
_wx_e, _wx_s = "equicorrelated ceiling", "sec:collab-dependence"
_wx_r7 = _wx_input(5, _wx_e, _wx_s, "r", 0.7, 1)
_wx_add(5, _wx_e, _wx_s, "ceiling 1/r", 1.0 / _wx_r7, 1, cross=1e12 / (1.0 + (1e12 - 1.0) * _wx_r7),
        rule="the limit of N / [1 + (N - 1) r] as N grows; cross-check: that ratio at N = 1e12", tol=1e-9)
_wx_e = "ex:neff"
_wx_N10 = _wx_input(5, _wx_e, _wx_s, "N", 10, 0)
_wx_input(5, _wx_e, _wx_s, "w_i", _WxF(1, _wx_N10), "fraction")
_wx_w10 = np.full(_wx_N10, 1.0 / _wx_N10)
_wx_add(5, _wx_e, _wx_s, "N_eff at R = I", _wx_neff(np.eye(_wx_N10), _wx_w10), 0, cross=1.0 / float(np.sum(_wx_w10 ** 2)),
        rule="def:neff; cross-check: the participation ratio 1 / sum w_i^2")
_wx_re = _wx_input(5, _wx_e, _wx_s, "r", 0.7, 1)
_wx_Req = (1.0 - _wx_re) * np.eye(_wx_N10) + _wx_re * np.ones((_wx_N10, _wx_N10))
_wx_add(5, _wx_e, _wx_s, "1 + (N - 1) r", 1.0 + (_wx_N10 - 1) * _wx_re, 1, cross=_wx_N10 * float(_wx_w10 @ _wx_Req @ _wx_w10),
        rule="cross-check: N w' R w at the uniform weight")
_wx_add(5, _wx_e, _wx_s, "N_eff, equicorrelated", _wx_neff(_wx_Req, _wx_w10), 1, cross=_wx_N10 / (1.0 + (_wx_N10 - 1) * _wx_re),
        rule="def:neff with the equicorrelated R; cross-check: N / [1 + (N - 1) r]")
_wx_sizes = [_wx_input(5, _wx_e, _wx_s, "group size %d" % (_k + 1), _v, None, rule="typed, stated in words")
             for _k, _v in enumerate((5, 3, 2))]
_wx_rw = _wx_input(5, _wx_e, _wx_s, "r_w", 0.8, 1)
_wx_rb = _wx_input(5, _wx_e, _wx_s, "r_b", 0.2, 1)
_wx_g = np.repeat(np.arange(len(_wx_sizes)), _wx_sizes)
_wx_Rg = np.where(_wx_g[:, None] == _wx_g[None, :], _wx_rw, _wx_rb)
np.fill_diagonal(_wx_Rg, 1.0)


def _wx_neff_blocks(shares, sizes, rw, rb):
    """N_eff with equal weights inside each group: each group of n members at share W
    contributes W^2 [1/n + (1 - 1/n) r_w], each pair of groups r_b W_k W_l."""
    W, n = np.asarray(shares, float), np.asarray(sizes, float)
    return 1.0 / float(np.sum(W ** 2 * (1.0 / n + (1.0 - 1.0 / n) * rw)) + rb * (1.0 - np.sum(W ** 2)))


_wx_add(5, _wx_e, _wx_s, "N_eff, three lineages, equal weights", _wx_neff(_wx_Rg, _wx_w10), 1,
        cross=_wx_neff_blocks(np.asarray(_wx_sizes) / _wx_N10, _wx_sizes, _wx_rw, _wx_rb),
        rule="def:neff with r_w within and r_b across groups; cross-check: the group-block formula")
for _k in range(3):
    _wx_add(5, _wx_e, _wx_s, "group share %d, equal weights" % (_k + 1), float(_wx_w10[_wx_g == _k].sum()), 2,
            cross=_wx_sizes[_k] / _wx_N10, rule="the group's share of the equal weights, n_k / N")
_wx_Wk = np.linalg.solve(np.block([[2.0 * np.eye(3), np.ones((3, 1))], [np.ones((1, 3)), np.zeros((1, 1))]]),
                         np.array([0.0, 0.0, 0.0, 1.0]))[:3]
for _k in range(3):
    _wx_add(5, _wx_e, _wx_s, "group share %d, penalty minimum" % (_k + 1), _WxF(1, 3), "fraction", cross=_wx_Wk[_k],
            rule="the minimizer of P_grp = sum_k W_k^2 on the simplex, 1/K; cross-check: its KKT system solved by numpy")
_wx_wpen = np.concatenate([np.full(_n, 1.0 / (3.0 * _n)) for _n in _wx_sizes])
_wx_add(5, _wx_e, _wx_s, "N_eff, equal shares, equal weights within", _wx_neff(_wx_Rg, _wx_wpen), 1,
        cross=_wx_neff_blocks(np.full(3, 1.0 / 3.0), _wx_sizes, _wx_rw, _wx_rb),
        rule="def:neff at the equal group shares; cross-check: the group-block formula")
_wx_add(5, _wx_e, _wx_s, "N_eff of the five-member lineage alone", _wx_neff(_wx_Rg[:5, :5], np.full(5, 0.2)), None,
        cross=5.0 / (1.0 + 4.0 * _wx_rw), rule="the text: the five variants amount to nearly one expert")
_wx_e = "ex:corr-lineage"
_wx_nl = _wx_input(5, _wx_e, _wx_s, "n", 5, 0)
_wx_rl = _wx_input(5, _wx_e, _wx_s, "r", 0.85, 2)
_wx_add(5, _wx_e, _wx_s, "n - 1", _wx_nl - 1, 0, cross=float(np.sum(~np.eye(_wx_nl, dtype=bool)[0])),
        rule="the number of other members of the lineage")
_wx_Rl = (1.0 - _wx_rl) * np.eye(_wx_nl) + _wx_rl * np.ones((_wx_nl, _wx_nl))
_wx_add(5, _wx_e, _wx_s, "N_eff of the lineage", _wx_nl / (1.0 + (_wx_nl - 1) * _wx_rl), 1,
        cross=_wx_neff(_wx_Rl, np.full(_wx_nl, 1.0 / _wx_nl)),
        rule="n / [1 + (n - 1) r]; cross-check: def:neff with the equicorrelated R at equal weights")
print("\n-- worked examples of Chapters 2 to 5: %d rows (%s)" % (
    len(_WX_ROWS), ", ".join("Chapter %d %d" % (c, sum(1 for r in _WX_ROWS if r["chapter"] == c)) for c in (2, 3, 4, 5))))
_WX_CFG = ("sheet worked_examples: the arithmetic that the examples of Chapter %d work by hand, computed in "
           "closed form from their stated inputs (block WORKED EXAMPLES OF THE THEORY CHAPTERS), no random "
           "draws; column printed rounds half-up to the decimals the text prints (the exact fraction where it "
           "prints one), column cross_check is an independent route to the same value")


# =====================================================================
# CHAPTER 2 RUNNING EXAMPLE
#   Table 2.1 (tab:elicitation), Figure 2.1 (fig:lp-vs-qa),
#   Figure 2.2 (fig:lp-lop-qa), Example ex:heterogeneous (sheet
#   ex_heterogeneous of the Figure 2.2 workbook), Table 2.5 (tab:biobj-example)
# ---------------------------------------------------------------------
# Stated inputs (the only typed numbers of this block):
#   * the five elicited percentiles of each expert in Table 2.1 (synthetic),
#   * the two-expert panel of Figure 2.1 (means -2.2 and 2.2, sd 0.75, equal
#     weights),
#   * the equal weights of Figure 2.2 and the trade-off alpha = 0.75,
#   * the three experts of Example ex:heterogeneous (log-normal with log x ~
#     N(1.4, 0.45^2), its reflection about 5.25, N(5.25, 0.60^2), equal weights),
#   * the seed 202603 and the numerical settings recorded in the config sheets.
# Everything else is computed here from the definitions of chapter2.tex:
#   Table 2.1     least squares of the elicited x_hat_p on Phi^{-1}(p), the
#                 quantile-matching fit of Section 2.2.2, rounded to one
#                 decimal before use (as the book prints them)
#   Fig 2.1/2.2   quantile average by prop:qa-properties(iii), mu = w.mu and
#                 sigma = w.sigma, cross-checked against alg:qa on a p-grid;
#                 logarithmic pool by def:log-pool, the normalized product,
#                 whose closed form for normal experts is the precision
#                 average, cross-checked against the numerical normalization
#   ex:heterog.   the experts' moments and the quantile-averaged mean and
#                 variance by numerical integration of the averaged quantile
#                 function (scipy quad), cross-checked by closed forms
#   Table 2.5     Algorithm 2.2 (alg:biobjective): A(w) by def:consensus,
#                 V(w) by def:confidence, the four bounds and the scalarized
#                 program solved by multi-start SLSQP under seed 202603, an
#                 interior optimizer polished by Nelder-Mead in (w.mu, w.sigma);
#                 the two edges of the interior-compromise window located by
#                 bisection on the first-order optimality test of the vertex
# =====================================================================
from decimal import Decimal as _Decimal, ROUND_HALF_UP as _ROUND_HALF_UP
from scipy.stats import norm as _norm
from scipy.optimize import minimize as _minimize
from scipy.optimize import minimize_scalar as _minimize_scalar
from scipy.integrate import quad as _quad
from scipy.interpolate import PchipInterpolator as _Pchip
from scipy.special import ndtr as _ndtr
from scipy.special import ndtri as _ndtri

CH2_SEED = NEW_SEED
CH2_P5 = np.array([0.05, 0.25, 0.50, 0.75, 0.95])
CH2_Z5 = _norm.ppf(CH2_P5)
# numerical settings (recorded in every config sheet of the block)
CH2_X_GRID = np.linspace(-2.0, 12.0, 7001)          # trapezoid cross-check of the overlap, dx = 0.002, tails > 5 sd; the log-pool normalization
CH2_P_GRID = np.linspace(1e-6, 1 - 1e-6, 20001)     # alg:qa / eq:confidence-quantile cross-checks
CH2_N_STARTS = 64                                   # Dirichlet(1,...,1) starts, plus vertices and centroid
CH2_FTOL = 1e-10
CH2_MAXITER = 500
CH2_ALPHA_STAR = 0.75
CH2_ALPHA_GRID = np.unique(np.round(np.concatenate([np.arange(0.0, 1.0001, 0.05),
                                                     np.arange(0.70, 0.9001, 0.01)]), 2))   # 37 distinct levels
CH2_N_PROBES = 2000                                 # Dirichlet(1,...,1) probe points of the rule-12 optimality check
CH2_POLISH_XATOL, CH2_POLISH_FATOL = 1e-9, 1e-14    # Nelder-Mead polish of an interior optimizer in (w.mu, w.sigma)
CH2_POLISH_MAXITER = 4000
CH2_EDGE_TOL = 1e-12                                # bisection stops when the alpha bracket is this narrow
CH2_EDGE_N_DIR = 4001                               # directions of the vertex tangent cone scanned before a bounded Brent refinement
CH2_EDGE_BRENT_XATOL = 1e-13                        # xatol of that bounded Brent refinement (in the cone angle)
CH2_EDGE_N_DIR_RATIO = 200001                       # directions of the ratio cross-check of each edge
CH2_EDGE_DELTA = 1e-3                               # the full multi-start solve is repeated at edge -/+ this offset
CH2_EDGE_DELTA_VERTEX = 1e-4                        # ... and on the vertex side of each edge at this closer offset
CH2_QUAD = {"epsabs": 1e-13, "epsrel": 1e-12, "limit": 400}   # scipy quad settings of Example ex:heterogeneous


def _ch2_printed(v, places="0.01"):
    """v rounded half-up to the decimals the book prints (as a string)."""
    return str(_Decimal(repr(round(float(v), 9))).quantize(_Decimal(places), rounding=_ROUND_HALF_UP))


# ---- Table 2.1: the elicited percentiles (stated inputs) and their LS fits
CH2_ELICITED = [
    ("Expert 1", [2.4, 3.3, 4.0, 4.7, 5.6]),
    ("Expert 2", [3.0, 4.2, 5.0, 5.8, 7.0]),
    ("Expert 3", [4.9, 5.8, 6.5, 7.2, 8.1]),
    ("Expert 4", [4.2, 5.0, 5.5, 6.0, 6.8]),
]


def _ch2_ls_fit(xhat):
    """Least squares of x_hat_p = mu + sigma * Phi^{-1}(p) over the five levels."""
    X = np.column_stack([np.ones(5), CH2_Z5])
    coef, res, _, _ = np.linalg.lstsq(X, np.asarray(xhat, float), rcond=None)
    rss = float(np.sum((X @ coef - xhat) ** 2))
    return float(coef[0]), float(coef[1]), rss


_t21 = []
for _name, _xh in CH2_ELICITED:
    _mu_ls, _sd_ls, _rss = _ch2_ls_fit(_xh)
    _t21.append({"expert": _name, "p05": _xh[0], "p25": _xh[1], "p50": _xh[2], "p75": _xh[3], "p95": _xh[4],
                 "fitted_mu": round(_mu_ls, 1), "fitted_sigma": round(_sd_ls, 1),
                 "ls_mu_exact": _mu_ls, "ls_sigma_exact": _sd_ls, "ls_rss": _rss})
CH2_MU = np.array([r["fitted_mu"] for r in _t21])       # the rounded fits are the panel of ex:running
CH2_SD = np.array([r["fitted_sigma"] for r in _t21])
assert np.all(CH2_SD > 0)
write("tab2-1_running_example_percentiles.xlsx", {
    "config": cfg({
        "table": "2.1", "label": "tab:elicitation",
        "inputs": "elicited percentiles p in {0.05,0.25,0.50,0.75,0.95} of four synthetic experts (typed, stated in the book)",
        "fit": "least squares of x_hat_p on Phi^{-1}(p): x_hat_p = mu + sigma * Phi^{-1}(p), numpy.linalg.lstsq",
        "rounding": "fitted mu and sigma rounded to one decimal before use; the rounded values are the panel of ex:running",
        "seed": "none -- deterministic least squares, no random draws",
    }),
    "table_2_1": _t21,
})


# ---- helpers shared by Figures 2.1/2.2 and Table 2.5 ------------------
def _ch2_qa_closed(w, mu, sd):
    """prop:qa-properties(iii): the quantile average of normal experts is N(w.mu, (w.sigma)^2)."""
    w = np.asarray(w, float)
    return float(w @ mu), float(w @ sd)


def _ch2_qa_density_numeric(w, mu, sd, x):
    """alg:qa on CH2_P_GRID: aggregate quantiles, monotone (PCHIP) interpolation,
    inversion, density by the inverse-function rule (cross-check only)."""
    q = sum(wi * (m + s * _norm.ppf(CH2_P_GRID)) for wi, m, s in zip(w, mu, sd))
    Fx = _Pchip(q, CH2_P_GRID, extrapolate=False)(x)
    dQ = _Pchip(CH2_P_GRID, q, extrapolate=False).derivative()
    with np.errstate(divide="ignore", invalid="ignore"):
        f = 1.0 / dQ(Fx)
    return np.where(np.isfinite(f), f, 0.0)


def _ch2_logpool_closed(w, mu, sd):
    """def:log-pool for normal experts: precision = sum w_i/sd_i^2,
    mean = sum(w_i mu_i/sd_i^2)/precision."""
    w = np.asarray(w, float)
    prec = float(np.sum(w / sd ** 2))
    return float(np.sum(w * mu / sd ** 2) / prec), prec ** -0.5


def _ch2_logpool_numeric(w, mu, sd, x):
    """def:log-pool literally: the weighted geometric mean normalized by its integral on CH2_X_GRID."""
    lg = sum(wi * _norm.logpdf(CH2_X_GRID, m, s) for wi, m, s in zip(w, mu, sd))
    g = np.exp(lg)
    Z = _trapz(g, CH2_X_GRID)
    lgx = sum(wi * _norm.logpdf(x, m, s) for wi, m, s in zip(w, mu, sd))
    return np.exp(lgx) / Z, float(_trapz(CH2_X_GRID * g, CH2_X_GRID) / Z)


def _ch2_min_pieces(m1, s1, m2, s2):
    """The two distinct densities N(m1,s1^2) and N(m2,s2^2) cross where log f_1 = log f_2,
    a quadratic in x (one root when s1 = s2).  Returns the intervals between consecutive
    crossing points, each with the index (0 or 1) of the smaller density on it."""
    if abs(s1 - s2) < 1e-14:
        roots = [0.5 * (m1 + m2)]
    else:
        a = 0.5 / s2 ** 2 - 0.5 / s1 ** 2
        b = m1 / s1 ** 2 - m2 / s2 ** 2
        c = 0.5 * m2 ** 2 / s2 ** 2 - 0.5 * m1 ** 2 / s1 ** 2 + math.log(s2 / s1)
        disc = b * b - 4.0 * a * c
        assert disc > 0, "two normal densities with different scales always cross twice"
        r1, r2 = (-b - math.sqrt(disc)) / (2.0 * a), (-b + math.sqrt(disc)) / (2.0 * a)
        roots = sorted([r1, r2])
    edges = [-math.inf] + roots + [math.inf]
    pieces = []
    for lo, hi in zip(edges[:-1], edges[1:]):
        xm = 0.5 * (lo + hi) if math.isfinite(lo) and math.isfinite(hi) else (hi - 1.0 if math.isfinite(hi) else lo + 1.0)
        # the smaller density on this interval (log densities compared at its midpoint)
        if -0.5 * ((xm - m1) / s1) ** 2 - math.log(s1) <= -0.5 * ((xm - m2) / s2) ** 2 - math.log(s2):
            pieces.append((lo, hi, 0))
        else:
            pieces.append((lo, hi, 1))
    return pieces


def _ch2_overlap_normal(m1, s1, m2, s2):
    """int min{N(m1,s1^2), N(m2,s2^2)} dx exactly.  On each interval between consecutive
    crossing points the smaller density is a fixed one of the two, so the integral is a
    sum of normal probabilities (scipy.special.ndtr).  Grid-free, so the optimizer of
    Algorithm 2.2 does not inherit a discretization error."""
    if abs(s1 - s2) < 1e-14 and abs(m1 - m2) < 1e-14:
        return 1.0
    total = 0.0
    for lo, hi, j in _ch2_min_pieces(m1, s1, m2, s2):
        m, s = (m1, s1) if j == 0 else (m2, s2)
        total += _ndtr((hi - m) / s) - _ndtr((lo - m) / s)
    return total


def _ch2_A_ms(m, s):
    """def:consensus in the aggregate coordinates: sum_i int min{f_i, N(m, s^2)} dx, exact."""
    return float(sum(_ch2_overlap_normal(mi, si, m, s) for mi, si in zip(CH2_MU, CH2_SD)))


def _ch2_A(w, mu, sd):
    """def:consensus: A(w) = sum_i int min{f_i, f_QA(.; w)} dx with f_QA the closed-form
    normal aggregate of prop:qa-properties(iii); each overlap evaluated exactly."""
    m, s = _ch2_qa_closed(w, mu, sd)
    return float(sum(_ch2_overlap_normal(mi, si, m, s) for mi, si in zip(mu, sd)))


def _ch2_A_grid(w, mu, sd):
    """def:consensus by the trapezoid rule on CH2_X_GRID with the closed-form aggregate
    (cross-check only; its O(h^2) error at the density crossings is about 1e-7)."""
    m, s = _ch2_qa_closed(w, mu, sd)
    fqa = _norm.pdf(CH2_X_GRID, m, s)
    return float(sum(_trapz(np.minimum(fqa, fi), CH2_X_GRID) for fi in _CH2_FI))


def _ch2_A_numeric(w, mu, sd):
    """def:consensus with f_QA from alg:qa numerically (cross-check only)."""
    fqa = _ch2_qa_density_numeric(w, mu, sd, CH2_X_GRID)
    return float(sum(_trapz(np.minimum(fqa, fi), CH2_X_GRID) for fi in _CH2_FI))


def _ch2_V(w, mu, sd):
    """def:confidence: V(w) = (w.sigma)^2, the equality case of thm:qa-moments
    for the location-scale panel (prop:qa-properties(iii))."""
    return float(np.asarray(w, float) @ sd) ** 2


def _ch2_V_numeric(w, mu, sd):
    """eq:confidence-quantile on CH2_P_GRID by the trapezoid rule (cross-check only)."""
    q = sum(wi * (m + s * _norm.ppf(CH2_P_GRID)) for wi, m, s in zip(w, mu, sd))
    mq = _trapz(q, CH2_P_GRID)
    return float(_trapz((q - mq) ** 2, CH2_P_GRID))


def _ch2_fibre(m, s, mu, sd):
    """The set of simplex weight vectors with (w.mu, w.sigma) = (m, s): the fibre of
    the affine map w -> (w.mu, w.sigma), a segment (or a point) of the simplex.
    Returns its endpoints (each has at least one zero coordinate)."""
    Aeq = np.vstack([np.ones(len(mu)), mu, sd])
    beq = np.array([1.0, m, s])
    ends = []
    for j in range(len(mu)):
        idx = [i for i in range(len(mu)) if i != j]
        try:
            w3 = np.linalg.solve(Aeq[:, idx], beq)
        except np.linalg.LinAlgError:
            continue
        if np.all(w3 >= -1e-9):
            w = np.zeros(len(mu)); w[idx] = np.clip(w3, 0, None); w /= w.sum()
            if not any(np.allclose(w, e, atol=1e-7) for e in ends):
                ends.append(w)
    assert 1 <= len(ends) <= 2, "fibre of (%.4f, %.4f) has %d endpoints" % (m, s, len(ends))
    return ends


def _ch2_tiebreak(w, mu, sd):
    """Selection rule for a non-unique optimizer: the midpoint of the optimal segment
    (the fibre of the best point's (w.mu, w.sigma)); a unique optimizer is returned as is."""
    m, s = _ch2_qa_closed(w, mu, sd)
    ends = _ch2_fibre(m, s, mu, sd)
    mid = np.mean(ends, axis=0)
    return mid, ends


CH2_SHAPE_N = 20001                                   # grid points of the linear-pool shape summary over each figure's domain


def _ch2_lp_shape(w, mu, sd, lo, hi):
    """Shape of the linear pool (def:linear-pool) that the figure script draws: the
    local maxima of the mixture density on CH2_SHAPE_N points of the plotted domain,
    their locations, the peak density, and the lowest density between the outer modes."""
    x = np.linspace(lo, hi, CH2_SHAPE_N)
    f = sum(wi * _norm.pdf(x, m, s) for wi, m, s in zip(w, mu, sd))
    k = np.where((f[1:-1] > f[:-2]) & (f[1:-1] >= f[2:]))[0] + 1
    trough = float(np.min(f[k[0]:k[-1] + 1])) if len(k) > 1 else float("nan")
    return len(k), [float(v) for v in x[k]], float(np.max(f)), trough


# ---- Figure 2.1: two experts, QA row computed ---------------------------
_f21_mu, _f21_sd, _f21_w = np.array([-2.2, 2.2]), np.array([0.75, 0.75]), np.array([0.5, 0.5])
_f21_qa_mu, _f21_qa_sd = _ch2_qa_closed(_f21_w, _f21_mu, _f21_sd)
_f21_x = np.linspace(-5.5, 5.5, 600)                  # the plotting grid of density_figure()
_f21_qa_dev = float(np.max(np.abs(_ch2_qa_density_numeric(_f21_w, _f21_mu, _f21_sd, _f21_x)
                                  - _norm.pdf(_f21_x, _f21_qa_mu, _f21_qa_sd))))
assert _f21_qa_dev < 1e-3, "Figure 2.1: alg:qa disagrees with the closed form (%.2e)" % _f21_qa_dev
# the text describes this linear pool as two separated peaks (bimodal)
_f21_lp_nmodes, _f21_lp_modes, _f21_lp_peak, _f21_lp_trough = _ch2_lp_shape(_f21_w, _f21_mu, _f21_sd, -5.5, 5.5)
assert _f21_lp_nmodes == 2, "Figure 2.1: the linear pool is not bimodal"
write("fig2-1_linear_pool_vs_qa.xlsx", {
    "config": cfg({
        "figure": "2.1", "label": "fig:lp-vs-qa",
        "x_label": "x", "y_label": "density",
        "domain_min": -5.5, "domain_max": 5.5, "ymax": 0.6,
        "draw_linear_pool": 1, "title": "Linear pool versus quantile averaging",
        "inputs": "two normal experts, means -2.2 and 2.2, sd 0.75, equal weights (typed, stated in the text that first cites Figure 2.1; the caption names only two equally weighted normal experts)",
        "linear_pool": "computed by the plotting script from the expert rows (def:linear-pool, weight-normalized mixture)",
        "quantile_average": "computed here by prop:qa-properties(iii): mu = w.mu, sigma = w.sigma",
        "qa_cross_check": "alg:qa on %d probability levels in [1e-6, 1-1e-6], PCHIP interpolation, inverse-function rule" % len(CH2_P_GRID),
        "qa_max_abs_density_deviation_on_plot_grid": _f21_qa_dev,
        "linear_pool_shape": "local maxima of the mixture density on %d points of [-5.5, 5.5] (the shape the text describes)" % CH2_SHAPE_N,
        "linear_pool_modes": _f21_lp_nmodes,
        "linear_pool_mode_locations": "[%s]" % ", ".join("%.4f" % v for v in _f21_lp_modes),
        "linear_pool_peak_density": _f21_lp_peak,
        "linear_pool_trough_density": _f21_lp_trough,
        "seed": "none -- closed forms, no random draws",
    }),
    "gaussians": [
        {"legend": "individual experts", "mu": _f21_mu[0], "sigma": _f21_sd[0], "weight": _f21_w[0], "role": "expert", "linestyle": "dotted"},
        {"legend": "_nolegend_",         "mu": _f21_mu[1], "sigma": _f21_sd[1], "weight": _f21_w[1], "role": "expert", "linestyle": "dotted"},
        {"legend": "quantile average",   "mu": _f21_qa_mu, "sigma": _f21_qa_sd, "weight": 1.0, "role": "aggregate", "linestyle": "dashed"},
    ],
})

# ---- Figure 2.2: the running panel under equal weights ------------------
_CH2_FI = [_norm.pdf(CH2_X_GRID, m, s) for m, s in zip(CH2_MU, CH2_SD)]   # expert densities on the overlap grid
_f22_w = np.full(4, 0.25)
_f22_x = np.linspace(1.0, 10.0, 600)                   # the plotting grid of density_figure()
_f22_qa_mu, _f22_qa_sd = _ch2_qa_closed(_f22_w, CH2_MU, CH2_SD)
_f22_qa_dev = float(np.max(np.abs(_ch2_qa_density_numeric(_f22_w, CH2_MU, CH2_SD, _f22_x)
                                  - _norm.pdf(_f22_x, _f22_qa_mu, _f22_qa_sd))))
_f22_lo_mu, _f22_lo_sd = _ch2_logpool_closed(_f22_w, CH2_MU, CH2_SD)
_f22_lo_num, _f22_lo_num_mu = _ch2_logpool_numeric(_f22_w, CH2_MU, CH2_SD, _f22_x)
_f22_lo_dev = float(np.max(np.abs(_f22_lo_num - _norm.pdf(_f22_x, _f22_lo_mu, _f22_lo_sd))))
assert _f22_qa_dev < 1e-3, "Figure 2.2: alg:qa disagrees with the closed form (%.2e)" % _f22_qa_dev
assert _f22_lo_dev < 1e-3 and abs(_f22_lo_num_mu - _f22_lo_mu) < 1e-3, \
    "Figure 2.2: normalized product disagrees with the precision-average closed form"
_f22_lp_sd = float(np.sqrt(np.sum(_f22_w * CH2_SD ** 2) + np.sum(_f22_w * (CH2_MU - _f22_w @ CH2_MU) ** 2)))
# the text describes this linear pool as wider than every expert because the spread of the
# experts' locations adds to their average variance (the mixture variance of def:linear-pool
# is the average expert variance plus the variance of the locations, and the average
# variance alone is at most the largest expert variance), and as nonetheless a single hump,
# lower and broader than the two normal aggregates; the shape is found on a grid, no
# sufficient condition for unimodality is claimed
_f22_avg_var = float(_f22_w @ CH2_SD ** 2)
_f22_loc_var = float(_f22_w @ (CH2_MU - _f22_w @ CH2_MU) ** 2)
assert abs(_f22_avg_var + _f22_loc_var - _f22_lp_sd ** 2) < 1e-12, "Figure 2.2: mixture variance is not average variance plus location variance"
assert _f22_avg_var <= float(np.max(CH2_SD)) ** 2 < _f22_avg_var + _f22_loc_var, \
    "Figure 2.2: the location term is not what makes the linear pool wider than every expert"
_f22_lp_nmodes, _f22_lp_modes, _f22_lp_peak, _ = _ch2_lp_shape(_f22_w, CH2_MU, CH2_SD, 1.0, 10.0)
assert _f22_lp_nmodes == 1, "Figure 2.2: the linear pool is not unimodal"
assert _f22_lp_sd > max(float(np.max(CH2_SD)), _f22_qa_sd, _f22_lo_sd), "Figure 2.2: the linear pool is not the widest"
assert _f22_lp_peak < min(_norm.pdf(0.0, 0.0, _f22_qa_sd), _norm.pdf(0.0, 0.0, _f22_lo_sd)), "Figure 2.2: the linear pool is not the lowest"

# ---- Example ex:heterogeneous: three experts of different shape, equal weights ----
# Stated inputs (typed, stated in the example): a log-normal expert whose logarithm is
# N(1.4, 0.45^2), its reflection about x = 5.25, and a sharp expert N(5.25, 0.60^2).
# Computed here: each expert's mean and sd, and the quantile-averaged mean and variance,
# all by numerical integration of the (averaged) quantile function over p in (0, 1)
# (scipy quad, settings CH2_QUAD), with the location-scale value (sum_i w_i sigma_i)^2
# against which the variance bound of thm:qa-moments is strict.  Cross-checks: the
# closed forms of the log-normal moments, exp(m + s^2/2) and sqrt(exp(s^2) - 1) times
# that mean, and of V_QA for this mirror-symmetric panel, whose averaged quantile
# function is 5.25 + [2 e^m sinh(s z_p) + sigma_3 z_p]/3, so that
#     V_QA = [2 e^{2m} (e^{2 s^2} - 1) + 4 e^m sigma_3 s e^{s^2/2} + sigma_3^2] / 9.
CH2_HET_LN_M, CH2_HET_LN_S = 1.4, 0.45
CH2_HET_REFLECT = 5.25
CH2_HET_SHARP_MU, CH2_HET_SHARP_SD = 5.25, 0.60
CH2_HET_W = np.full(3, 1.0 / 3.0)
# the quantile functions written in z = Phi^{-1}(p); Q_i(p) = q_i(ndtri(p))
_het_qz = [
    lambda z: math.exp(CH2_HET_LN_M + CH2_HET_LN_S * z),                              # log-normal
    lambda z: 2.0 * CH2_HET_REFLECT - math.exp(CH2_HET_LN_M - CH2_HET_LN_S * z),      # its reflection, Q(p) = 2c - Q_LN(1 - p)
    lambda z: CH2_HET_SHARP_MU + CH2_HET_SHARP_SD * z,                                 # the sharp normal
]
_het_qbarz = lambda z: sum(wi * q(z) for wi, q in zip(CH2_HET_W, _het_qz))
_het_q = [lambda p, q=q: q(_ndtri(p)) for q in _het_qz]
_het_qbar = lambda p: _het_qbarz(_ndtri(p))


def _ch2_quantile_moments(q):
    """mean and variance of the distribution with quantile function q: int_0^1 q dp and
    int_0^1 (q - mean)^2 dp, by scipy quad with the settings CH2_QUAD."""
    m = _quad(q, 0.0, 1.0, **CH2_QUAD)[0]
    return m, _quad(lambda p: (q(p) - m) ** 2, 0.0, 1.0, **CH2_QUAD)[0]


_het_mom = [_ch2_quantile_moments(q) for q in _het_q]
_het_mu = np.array([m for m, _ in _het_mom]); _het_sd = np.sqrt([v for _, v in _het_mom])
_het_mu_qa, _het_V_qa = _ch2_quantile_moments(_het_qbar)
_het_ls = float(CH2_HET_W @ _het_sd) ** 2
# shape of the aggregate (not printed; they back 'symmetric' and 'no longer normal'): the
# third and fourth central moments, integrated in z against the normal density (quad on
# [-CH2_HET_ZMAX, CH2_HET_ZMAX]), and the symmetry Qbar(p) + Qbar(1 - p) = 2c at z and -z
CH2_HET_ZMAX = 12.0                                   # normal mass beyond is below 1e-32
_het_cm = lambda k: _quad(lambda z: (float(_het_qbarz(z)) - _het_mu_qa) ** k * math.exp(-0.5 * z * z) / math.sqrt(2.0 * math.pi),
                          -CH2_HET_ZMAX, CH2_HET_ZMAX, **CH2_QUAD)[0]
_het_skew = _het_cm(3) / _het_V_qa ** 1.5
_het_exkurt = _het_cm(4) / _het_V_qa ** 2 - 3.0
_het_zsym = np.linspace(0.0, 5.0, 2001)
_het_symdev = max(abs(_het_qbarz(z) + _het_qbarz(-z) - 2.0 * CH2_HET_REFLECT) for z in _het_zsym)
# closed forms (cross-check)
_het_ln_mean_cf = math.exp(CH2_HET_LN_M + 0.5 * CH2_HET_LN_S ** 2)
_het_ln_sd_cf = _het_ln_mean_cf * math.sqrt(math.exp(CH2_HET_LN_S ** 2) - 1.0)
_het_V_cf = (2.0 * math.exp(2.0 * CH2_HET_LN_M) * (math.exp(2.0 * CH2_HET_LN_S ** 2) - 1.0)
             + 4.0 * math.exp(CH2_HET_LN_M) * CH2_HET_SHARP_SD * CH2_HET_LN_S * math.exp(0.5 * CH2_HET_LN_S ** 2)
             + CH2_HET_SHARP_SD ** 2) / 9.0
_het_cf = {"mu": [_het_ln_mean_cf, 2.0 * CH2_HET_REFLECT - _het_ln_mean_cf, CH2_HET_SHARP_MU],
           "sd": [_het_ln_sd_cf, _het_ln_sd_cf, CH2_HET_SHARP_SD]}
assert CH2_HET_SHARP_MU == CH2_HET_REFLECT, "the closed form of V_QA assumes the sharp expert sits at the reflection point"
assert np.allclose(_het_mu, _het_cf["mu"], atol=1e-9) and np.allclose(_het_sd, _het_cf["sd"], atol=1e-9), \
    "ex:heterogeneous: quadrature and closed-form expert moments disagree"
assert abs(_het_mu_qa - float(CH2_HET_W @ _het_mu)) < 1e-9, "ex:heterogeneous: the aggregate mean is not the average of the means"
assert abs(_het_V_qa - _het_V_cf) < 1e-9, "ex:heterogeneous: quadrature and closed-form V_QA disagree"
assert _het_V_qa < _het_ls - 1e-3, "ex:heterogeneous: the variance bound of thm:qa-moments is not strict"
assert _het_symdev < 1e-12 and abs(_het_skew) < 1e-8, "ex:heterogeneous: the aggregate is not symmetric about 5.25"
assert abs(_het_exkurt) > 1e-2, "ex:heterogeneous: the aggregate has the kurtosis of a normal"
_het_rows = [
    {"quantity": "log-normal expert, mean of log x", "role": "stated input", "value": CH2_HET_LN_M, "closed_form": CH2_HET_LN_M, "printed": _ch2_printed(CH2_HET_LN_M, "0.1"), "rule": "typed, stated in the example"},
    {"quantity": "log-normal expert, sd of log x", "role": "stated input", "value": CH2_HET_LN_S, "closed_form": CH2_HET_LN_S, "printed": _ch2_printed(CH2_HET_LN_S), "rule": "typed, stated in the example"},
    {"quantity": "reflection point", "role": "stated input", "value": CH2_HET_REFLECT, "closed_form": CH2_HET_REFLECT, "printed": _ch2_printed(CH2_HET_REFLECT), "rule": "typed, stated in the example"},
    {"quantity": "sharp expert, mean", "role": "stated input", "value": CH2_HET_SHARP_MU, "closed_form": CH2_HET_SHARP_MU, "printed": _ch2_printed(CH2_HET_SHARP_MU), "rule": "typed, stated in the example"},
    {"quantity": "sharp expert, sd", "role": "stated input", "value": CH2_HET_SHARP_SD, "closed_form": CH2_HET_SHARP_SD, "printed": _ch2_printed(CH2_HET_SHARP_SD), "rule": "typed, stated in the example"},
    {"quantity": "weight of each expert", "role": "stated input", "value": CH2_HET_W[0], "closed_form": 1.0 / 3.0, "printed": "", "rule": "equal weights"},
    {"quantity": "log-normal expert, mean", "role": "computed", "value": _het_mu[0], "closed_form": _het_cf["mu"][0], "printed": _ch2_printed(_het_mu[0]), "rule": "int_0^1 Q_1(p) dp, Q_1(p) = exp(m + s z_p); closed form exp(m + s^2/2)"},
    {"quantity": "log-normal expert, sd", "role": "computed", "value": _het_sd[0], "closed_form": _het_cf["sd"][0], "printed": _ch2_printed(_het_sd[0]), "rule": "sqrt int_0^1 (Q_1 - mean)^2 dp; closed form mean * sqrt(exp(s^2) - 1)"},
    {"quantity": "reflected expert, mean", "role": "computed", "value": _het_mu[1], "closed_form": _het_cf["mu"][1], "printed": _ch2_printed(_het_mu[1]), "rule": "int_0^1 Q_2(p) dp, Q_2(p) = 2c - Q_1(1 - p); closed form 2c - log-normal mean"},
    {"quantity": "reflected expert, sd", "role": "computed", "value": _het_sd[1], "closed_form": _het_cf["sd"][1], "printed": _ch2_printed(_het_sd[1]), "rule": "sqrt int_0^1 (Q_2 - mean)^2 dp; equals the log-normal sd"},
    {"quantity": "mu_QA", "role": "computed", "value": _het_mu_qa, "closed_form": float(CH2_HET_W @ np.array(_het_cf["mu"])), "printed": _ch2_printed(_het_mu_qa), "rule": "int_0^1 sum_i w_i Q_i(p) dp; closed form sum_i w_i mu_i (thm:qa-moments)"},
    {"quantity": "V_QA", "role": "computed", "value": _het_V_qa, "closed_form": _het_V_cf, "printed": _ch2_printed(_het_V_qa), "rule": "int_0^1 (sum_i w_i Q_i(p) - mu_QA)^2 dp; closed form [2 e^{2m}(e^{2s^2} - 1) + 4 e^m sigma_3 s e^{s^2/2} + sigma_3^2]/9"},
    {"quantity": "(sum_i w_i sigma_i)^2", "role": "computed", "value": _het_ls, "closed_form": float(CH2_HET_W @ np.array(_het_cf["sd"])) ** 2, "printed": _ch2_printed(_het_ls), "rule": "the location-scale equality value of thm:qa-moments"},
    {"quantity": "skewness of the aggregate", "role": "computed", "value": _het_skew, "closed_form": 0.0, "printed": "", "rule": "int (qbar(z) - mu_QA)^3 phi(z) dz / V_QA^1.5, quad on |z| <= 12 (symmetric aggregate)"},
    {"quantity": "excess kurtosis of the aggregate", "role": "computed", "value": _het_exkurt, "closed_form": float("nan"), "printed": "", "rule": "int (qbar(z) - mu_QA)^4 phi(z) dz / V_QA^2 - 3, quad on |z| <= 12 (nonzero, so the aggregate is not normal)"},
    {"quantity": "max |Qbar(p) + Qbar(1 - p) - 2c|", "role": "computed", "value": _het_symdev, "closed_form": 0.0, "printed": "", "rule": "symmetry about c, qbar(z) + qbar(-z) - 2c on 2001 points z in [0, 5]"},
]
print("   ex:heterogeneous: means %.4f / %.4f, sd %.4f, mu_QA %.4f, V_QA %.4f < (sum w sigma)^2 = %.4f" % (
    _het_mu[0], _het_mu[1], _het_sd[0], _het_mu_qa, _het_V_qa, _het_ls))

write("fig2-2_three_pooling_operators.xlsx", {
    "config": cfg({
        "figure": "2.2", "label": "fig:lp-lop-qa",
        "x_label": "x", "y_label": "density",
        "domain_min": 1, "domain_max": 10, "ymax": 0.55,
        "draw_linear_pool": 1, "title": "Linear, logarithmic, and quantile pools",
        "inputs": "the four rounded normal fits of Table 2.1 (tab2-1 workbook) under equal weights 0.25",
        "linear_pool": "computed by the plotting script from the expert rows (def:linear-pool); mixture sd recorded below",
        "linear_pool_sd": _f22_lp_sd,
        "quantile_average": "computed here by prop:qa-properties(iii): mu = w.mu, sigma = w.sigma",
        "qa_cross_check": "alg:qa on %d probability levels, PCHIP interpolation, inverse-function rule" % len(CH2_P_GRID),
        "qa_max_abs_density_deviation_on_plot_grid": _f22_qa_dev,
        "logarithmic_pool": "computed here by def:log-pool; for normal experts the normalized product is normal with precision sum w_i/sigma_i^2 and mean sum(w_i mu_i/sigma_i^2)/precision",
        "logpool_cross_check": "normalized product integrated by the trapezoid rule on x in [-2, 12], step 0.002",
        "logpool_max_abs_density_deviation_on_plot_grid": _f22_lo_dev,
        "linear_pool_shape": "local maxima of the mixture density on %d points of [1, 10] (the shape the text describes)" % CH2_SHAPE_N,
        "linear_pool_modes": _f22_lp_nmodes,
        "linear_pool_mode_locations": "[%s]" % ", ".join("%.4f" % v for v in _f22_lp_modes),
        "linear_pool_peak_density": _f22_lp_peak,
        "linear_pool_average_expert_variance": _f22_avg_var,
        "linear_pool_location_variance": _f22_loc_var,
        "linear_pool_variance_decomposition": "linear_pool_sd^2 = linear_pool_average_expert_variance (sum_i w_i sigma_i^2) + linear_pool_location_variance (sum_i w_i (mu_i - w.mu)^2); the first term alone is at most the largest expert variance, so the location term is what makes the pool wider than every expert (the text's 'because' clause); the single hump is the grid count linear_pool_modes, not a consequence of these terms",
        "ex_heterogeneous": "Example ex:heterogeneous (text only, not plotted): stated inputs and computed moments in sheet ex_heterogeneous",
        "ex_heterogeneous_inputs": "log-normal expert with log x ~ N(1.4, 0.45^2), its reflection about x = 5.25, sharp expert N(5.25, 0.60^2), equal weights 1/3 (typed, stated in the example)",
        "ex_heterogeneous_quadrature": "scipy.integrate.quad of the (averaged) quantile function over p in (0, 1), z_p = scipy.special.ndtri(p), epsabs %.0e, epsrel %.0e, limit %d; the unprinted skewness and kurtosis by the same quad in z = Phi^{-1}(p) on |z| <= %g" % (
            CH2_QUAD["epsabs"], CH2_QUAD["epsrel"], CH2_QUAD["limit"], CH2_HET_ZMAX),
        "ex_heterogeneous_cross_check": "closed forms of the log-normal moments and of V_QA for the mirror-symmetric panel (column closed_form), agreement within 1e-9",
        "ex_heterogeneous_printed": "the column printed rounds half-up to the decimals the example prints",
        "worked_examples": _WX_CFG % 2,
        "seed": "none -- closed forms, no random draws",
    }),
    "gaussians": [
        {"legend": "experts",          "mu": CH2_MU[0], "sigma": CH2_SD[0], "weight": _f22_w[0], "role": "expert", "linestyle": "dotted"},
        {"legend": "_nolegend_",       "mu": CH2_MU[1], "sigma": CH2_SD[1], "weight": _f22_w[1], "role": "expert", "linestyle": "dotted"},
        {"legend": "_nolegend_",       "mu": CH2_MU[2], "sigma": CH2_SD[2], "weight": _f22_w[2], "role": "expert", "linestyle": "dotted"},
        {"legend": "_nolegend_",       "mu": CH2_MU[3], "sigma": CH2_SD[3], "weight": _f22_w[3], "role": "expert", "linestyle": "dotted"},
        {"legend": "quantile average", "mu": _f22_qa_mu, "sigma": _f22_qa_sd, "weight": 1.0, "role": "aggregate", "linestyle": "dashed"},
        {"legend": "logarithmic pool", "mu": _f22_lo_mu, "sigma": _f22_lo_sd, "weight": 1.0, "role": "aggregate", "linestyle": "dashdot"},
    ],
    # not read by the figure script: Example ex:heterogeneous, the non-normal counterpart
    # of this comparison (stated inputs and computed moments, printed roundings)
    "ex_heterogeneous": _het_rows,
    # not read by the figure script: the hand-worked examples of Chapter 2
    "worked_examples": _wx_sheet(2),
})

# ---- Table 2.5: Algorithm 2.2 on the running panel ----------------------
_ch2_rng = np.random.default_rng(CH2_SEED)
# draw order: (1) the CH2_N_STARTS Dirichlet(1,1,1,1) starts, (2) CH2_N_PROBES Dirichlet(1,1,1,1)
# probe points used only by the rule-12 optimality check at the end of the file
CH2_STARTS = np.vstack([_ch2_rng.dirichlet(np.ones(4), size=CH2_N_STARTS), np.eye(4), np.full((1, 4), 0.25)])
CH2_PROBES = _ch2_rng.dirichlet(np.ones(4), size=CH2_N_PROBES)
_ch2_cons = [{"type": "eq", "fun": lambda w: np.sum(w) - 1.0}]
_ch2_bnds = [(0.0, 1.0)] * 4


def _ch2_multistart(fun):
    """Minimize fun over the simplex by SLSQP from every start in CH2_STARTS;
    returns (best value, best w, number of starts within 1e-6 of the best, number of starts used)."""
    sols = []
    for w0 in CH2_STARTS:
        r = _minimize(fun, w0, method="SLSQP", bounds=_ch2_bnds, constraints=_ch2_cons,
                      options={"ftol": CH2_FTOL, "maxiter": CH2_MAXITER})
        w = np.clip(r.x, 0.0, 1.0); w = w / w.sum()
        sols.append((float(fun(w)), w))
    fs = np.array([s[0] for s in sols])
    k = int(np.argmin(fs))
    return sols[k][0], sols[k][1], int(np.sum(fs - fs[k] < 1e-6)), len(sols)


_A = lambda w: _ch2_A(w, CH2_MU, CH2_SD)
_V = lambda w: _ch2_V(w, CH2_MU, CH2_SD)
_Amin, _wAmin, _nAmin, _nS = _ch2_multistart(_A)
_negAmax, _wAmax, _nAmax, _ = _ch2_multistart(lambda w: -_A(w)); _Amax = -_negAmax
_Vmin, _wVmin, _nVmin, _ = _ch2_multistart(_V)
_negVmax, _wVmax, _nVmax, _ = _ch2_multistart(lambda w: -_V(w)); _Vmax = -_negVmax
assert _Amax > _Amin and _Vmax > _Vmin


def _ch2_objective(w, alpha):
    """alg:biobjective: alpha * A~(w) - (1 - alpha) * V~(w) with min-max normalized objectives."""
    At = (_A(w) - _Amin) / (_Amax - _Amin)
    Vt = (_V(w) - _Vmin) / (_Vmax - _Vmin)
    return alpha * At - (1.0 - alpha) * Vt


def _ch2_objective_ms(ms, alpha):
    """The same scalarized objective written in the aggregate coordinates (m, s) = (w.mu, w.sigma),
    through which alone both objectives depend on w on this panel."""
    m, s = float(ms[0]), float(ms[1])
    if s <= 0:
        return -np.inf
    A = _ch2_A_ms(m, s)
    return alpha * (A - _Amin) / (_Amax - _Amin) - (1.0 - alpha) * (s ** 2 - _Vmin) / (_Vmax - _Vmin)


def _ch2_solve_alpha(alpha):
    """One grid point of alg:biobjective: multi-start SLSQP over the simplex, then, for an
    interior optimizer, a Nelder-Mead polish of (m*, s*) in the aggregate plane (exact for
    this panel, since the objective depends on w only through that pair), then the fibre
    of (m*, s*) and its midpoint as the reported weight vector."""
    negf, wb, nagree, _ = _ch2_multistart(lambda w: -_ch2_objective(w, alpha))
    m0, s0 = _ch2_qa_closed(wb, CH2_MU, CH2_SD)
    if np.max(wb) < 0.999:
        r = _minimize(lambda ms: -_ch2_objective_ms(ms, alpha), np.array([m0, s0]), method="Nelder-Mead",
                      options={"xatol": CH2_POLISH_XATOL, "fatol": CH2_POLISH_FATOL, "maxiter": CH2_POLISH_MAXITER})
        m1, s1 = float(r.x[0]), float(r.x[1])
        assert -r.fun >= -negf - 1e-9, "polish worsened the objective at alpha=%.2f" % alpha
        assert abs(m1 - m0) < 5e-3 and abs(s1 - s0) < 5e-3, "polish moved (m, s) by more than 5e-3 at alpha=%.2f" % alpha
        ends = _ch2_fibre(m1, s1, CH2_MU, CH2_SD)
        f = -r.fun
    else:
        ends = _ch2_fibre(m0, s0, CH2_MU, CH2_SD)
        f = -negf
    wmid = np.mean(ends, axis=0)
    return wb, wmid, ends, f, nagree


_sweep = []
_sweep_sol = {}
for _al in CH2_ALPHA_GRID:
    _wb, _wmid, _ends, _f, _nagree = _ch2_solve_alpha(float(_al))
    _m, _s = _ch2_qa_closed(_wmid, CH2_MU, CH2_SD)
    _sweep_sol[float(_al)] = (_wb, _wmid, _ends, _f, _nagree)
    _sweep.append({"alpha": float(_al), "w1": _wmid[0], "w2": _wmid[1], "w3": _wmid[2], "w4": _wmid[3],
                   "A": _A(_wmid), "V": _V(_wmid), "fQA_mu": _m, "fQA_sigma": _s,
                   "objective": _f, "interior": int(np.max(_wmid) < 0.999),
                   "segment_length": float(np.linalg.norm(_ends[-1] - _ends[0])),
                   "starts_within_1e-6_of_best": _nagree, "starts_total": _nS})
_wb75, _w75, _ends75, _f75, _n75 = _sweep_sol[CH2_ALPHA_STAR]
_m75, _s75 = _ch2_qa_closed(_w75, CH2_MU, CH2_SD)
_alphas_interior = [r["alpha"] for r in _sweep if r["interior"] == 1]
assert _alphas_interior, "no interior compromise found on the alpha grid"

# ---- the two edges of the interior-compromise window ----------------------
# On the grid the optimizer is a vertex of the simplex below and above a window of
# interior compromises.  Each edge is located by bisection on the first-order optimality
# test of the adjacent vertex, not on the grid flag 'interior' (max w < 0.999), which would
# bias an edge because close to it the optimizer lies within 0.01 in weight of the vertex.
# Both objectives depend on w only through (m, s) = (w.mu, w.sigma), whose range is the
# convex hull of the experts' points (mu_k, sigma_k).  At the point of expert k the
# one-sided derivative of the scalarized objective along a unit direction d = (cos t, sin t)
# of the hull's tangent cone is  D_d f = alpha a(d) - (1 - alpha) b(d)  with
#     a(d) = [sum_{i != k} grad O_i . d  -  (1/(2 sigma_k)) int |d_m z + d_s (z^2 - 1)| phi(z) dz] / (A_max - A_min),
#     b(d) = 2 sigma_k d_s / (V_max - V_min),
# O_i the overlap of expert i with N(m, s^2); the integral is the one-sided derivative of
# expert k's own overlap, which equals 1 at the vertex and has a kink there.  Both pieces
# are closed forms (normal-density antiderivatives between the crossing points).  The
# vertex is locally optimal iff max_d D_d f <= 0.  D_d f is affine in alpha for every d,
# so that maximum is convex in alpha and changes sign exactly once in each bracket; the
# bracket is the pair of grid levels on either side of the edge of the grid window.
# Cross-check: the edge is also the extreme over d of the level b/(a + b) at which D_d f
# vanishes.  The global statement (the vertex outside the window, a strictly better
# interior optimizer inside) is confirmed by the full multi-start solve at edge -/+ CH2_EDGE_DELTA
# and, on the vertex side, at the closer offset CH2_EDGE_DELTA_VERTEX.  These solves bound the
# global edge, the level at which the vertex stops being the global optimizer: the optimal
# value g(alpha) is a maximum of affine functions of alpha, so g - f_vertex is convex and
# nonnegative, and the levels at which the vertex is optimal form an interval that contains
# alpha = 0 (minimum variance) or alpha = 1 (maximum consensus); past the bisection edge the
# vertex is not even a local maximizer.  Hence the global edge lies between the vertex-side
# solve and the bisection edge, within CH2_EDGE_DELTA_VERTEX of the latter, and the printed
# rounding is asserted for both ends of that interval.
_SQRT2PI = math.sqrt(2.0 * math.pi)


def _ch2_phi_s(z):
    return 0.0 if not math.isfinite(z) else math.exp(-0.5 * z * z) / _SQRT2PI


def _ch2_overlap_grad(mi, si, m, s):
    """gradient in (m, s) of int min{N(mi, si^2), N(m, s^2)} dx, the two densities distinct:
    the integral of dN(m, s^2)/d(m, s) over the intervals on which N(m, s^2) is the smaller,
    which with z = (x - m)/s is [phi(z_lo) - phi(z_hi)]/s for m and [z_lo phi(z_lo) - z_hi phi(z_hi)]/s for s."""
    gm = gs = 0.0
    for lo, hi, j in _ch2_min_pieces(mi, si, m, s):
        if j == 1:
            za, zb = (lo - m) / s, (hi - m) / s
            gm += (_ch2_phi_s(za) - _ch2_phi_s(zb)) / s
            gs += ((za * _ch2_phi_s(za) if math.isfinite(za) else 0.0) - (zb * _ch2_phi_s(zb) if math.isfinite(zb) else 0.0)) / s
    return gm, gs


def _ch2_self_slope(s, dm, ds):
    """one-sided derivative at t = 0+ of int min{N(m, s^2), N(m + t dm, (s + t ds)^2)} dx, i.e.
    -(1/(2s)) int |dm z + ds (z^2 - 1)| phi(z) dz, from the antiderivative -(dm + ds z) phi(z) on
    the intervals cut by the roots of ds z^2 + dm z - ds (vectorized over directions)."""
    dm = np.asarray(dm, float); ds = np.asarray(ds, float)
    r = np.sqrt(dm * dm + 4.0 * ds * ds)
    nz = np.abs(ds) > 1e-15
    den = 2.0 * np.where(nz, ds, 1.0)
    z1 = np.where(nz, (-dm - r) / den, 0.0); z2 = np.where(nz, (-dm + r) / den, 0.0)
    G = lambda z: -(dm + ds * z) * np.exp(-0.5 * z * z) / _SQRT2PI
    g_lo, g_hi = G(np.minimum(z1, z2)), G(np.maximum(z1, z2))
    return -(np.abs(g_lo) + np.abs(g_hi - g_lo) + np.abs(g_hi)) / (2.0 * s)


def _ch2_vertex_cone(k):
    """angular range [t0, t1] of the tangent cone of the (mu, sigma) hull at the point of expert k"""
    D = np.array([[CH2_MU[j] - CH2_MU[k], CH2_SD[j] - CH2_SD[k]] for j in range(len(CH2_MU)) if j != k])
    th = np.arctan2(D[:, 1], D[:, 0])
    base = math.atan2(float(np.sum(np.sin(th))), float(np.sum(np.cos(th))))
    rel = (th - base + np.pi) % (2.0 * np.pi) - np.pi
    assert rel.max() - rel.min() < np.pi, "expert %d is not an extreme point of the (mu, sigma) hull" % (k + 1)
    return base + float(rel.min()), base + float(rel.max())


def _ch2_vertex_slopes(k, t):
    """(a, b) of D_d f = alpha a - (1 - alpha) b at the point of expert k, d = (cos t, sin t)"""
    t = np.asarray(t, float); dm, ds = np.cos(t), np.sin(t)
    gm = gs = 0.0
    for i in range(len(CH2_MU)):
        if i != k:
            g1, g2 = _ch2_overlap_grad(CH2_MU[i], CH2_SD[i], CH2_MU[k], CH2_SD[k])
            gm += g1; gs += g2
    a = (gm * dm + gs * ds + _ch2_self_slope(CH2_SD[k], dm, ds)) / (_Amax - _Amin)
    b = 2.0 * CH2_SD[k] * ds / (_Vmax - _Vmin)
    return a, b


def _ch2_cone_max(k, alpha):
    """max over unit directions of the tangent cone at the point of expert k of D_d f:
    CH2_EDGE_N_DIR directions, the best refined by a bounded Brent search"""
    t0, t1 = _ch2_vertex_cone(k)
    ts = np.linspace(t0, t1, CH2_EDGE_N_DIR)
    a, b = _ch2_vertex_slopes(k, ts)
    v = alpha * a - (1.0 - alpha) * b
    j = int(np.argmax(v))

    def _neg(t):
        aa, bb = _ch2_vertex_slopes(k, np.array([t]))
        return -float(alpha * aa[0] - (1.0 - alpha) * bb[0])
    r = _minimize_scalar(_neg, bounds=(ts[max(j - 1, 0)], ts[min(j + 1, len(ts) - 1)]), method="bounded",
                         options={"xatol": CH2_EDGE_BRENT_XATOL})
    return max(float(v[j]), -float(r.fun))


def _ch2_window_edge(k, lo, hi):
    """bisection on the sign of _ch2_cone_max(k, alpha) over the bracket [lo, hi]"""
    c_lo, c_hi = _ch2_cone_max(k, lo), _ch2_cone_max(k, hi)
    assert (c_lo > 0.0) != (c_hi > 0.0), "no sign change of the vertex test on [%.2f, %.2f]" % (lo, hi)
    s_lo, n = c_lo > 0.0, 0
    while hi - lo > CH2_EDGE_TOL:
        mid = 0.5 * (lo + hi)
        if (_ch2_cone_max(k, mid) > 0.0) == s_lo:
            lo = mid
        else:
            hi = mid
        n += 1
    return 0.5 * (lo + hi), n, c_lo, c_hi


_grid_alphas = sorted(_sweep_sol)
_win_lo, _win_hi = min(_alphas_interior), max(_alphas_interior)
assert all(r["interior"] == 1 for r in _sweep if _win_lo <= r["alpha"] <= _win_hi), "the interior levels of the grid are not contiguous"
_k_minvar, _k_maxcons = int(np.argmin(CH2_SD)), int(np.argmax(_wAmax))
for _rs in _sweep:   # the text: the minimum-variance vertex below the window, the maximum-consensus vertex above it
    _wr = np.array([_rs["w1"], _rs["w2"], _rs["w3"], _rs["w4"]])
    if _rs["alpha"] < _win_lo:
        assert _wr[_k_minvar] > 0.999, "alpha=%.2f: below the window but not the minimum-variance vertex" % _rs["alpha"]
    elif _rs["alpha"] > _win_hi:
        assert _wr[_k_maxcons] > 0.999, "alpha=%.2f: above the window but not the maximum-consensus vertex" % _rs["alpha"]
_edges = []
for _side, _k, _lo, _hi in (("lower", _k_minvar, max(a for a in _grid_alphas if a < _win_lo), _win_lo),
                            ("upper", _k_maxcons, _win_hi, min(a for a in _grid_alphas if a > _win_hi))):
    _e, _nb, _c_lo, _c_hi = _ch2_window_edge(_k, _lo, _hi)
    # cross-check without bisection: the extreme over directions of the level b/(a + b) at
    # which D_d f vanishes (b > 0 on the whole cone at the minimum-variance point, b < 0 at the
    # maximum-variance point, so the edge is the min, resp. the max, of that level)
    _t0, _t1 = _ch2_vertex_cone(_k)
    _aa, _bb = _ch2_vertex_slopes(_k, np.linspace(_t0, _t1, CH2_EDGE_N_DIR_RATIO))
    if _side == "lower":
        assert np.all(_bb > 0), "the minimum-variance point does not raise the variance in every hull direction"
        _ok = _aa + _bb > 0
        _e_ratio = float(np.min(_bb[_ok] / (_aa[_ok] + _bb[_ok])))
    else:
        assert np.all(_bb < 0), "the maximum-consensus point does not lower the variance in every hull direction"
        _ok = _aa + _bb < 0
        _e_ratio = float(np.max(_bb[_ok] / (_aa[_ok] + _bb[_ok])))
    assert abs(_e - _e_ratio) < 1e-8, "%s edge: bisection %.12f and ratio cross-check %.12f disagree" % (_side, _e, _e_ratio)
    assert _lo < _e < _hi
    _edges.append({"edge": _side, "vertex_expert": _k + 1,
                   "vertex": "minimum variance" if _side == "lower" else "maximum consensus",
                   "bracket_lo": _lo, "bracket_hi": _hi,
                   "cone_max_at_bracket_lo": _c_lo, "cone_max_at_bracket_hi": _c_hi,
                   "alpha_edge": _e, "bisection_steps": _nb, "alpha_edge_ratio_cross_check": _e_ratio,
                   "printed": _ch2_printed(_e)})
_alpha_lo_edge, _alpha_hi_edge = _edges[0]["alpha_edge"], _edges[1]["alpha_edge"]
_edge_checks = []
for _ed in _edges:
    _k = _ed["vertex_expert"] - 1
    _sgn_v = -1.0 if _ed["edge"] == "lower" else 1.0      # direction from the edge toward the vertex side
    for _pos, _off, _al in sorted((("edge - delta", CH2_EDGE_DELTA, _ed["alpha_edge"] - CH2_EDGE_DELTA),
                                   ("edge + delta", CH2_EDGE_DELTA, _ed["alpha_edge"] + CH2_EDGE_DELTA),
                                   ("edge %s delta_vertex" % ("-" if _sgn_v < 0 else "+"), CH2_EDGE_DELTA_VERTEX,
                                    _ed["alpha_edge"] + _sgn_v * CH2_EDGE_DELTA_VERTEX)), key=lambda _t: _t[2]):
        _inside = (_al - _ed["alpha_edge"]) * _sgn_v < 0.0
        _wbx, _wmx, _enx, _fx, _nx = _ch2_solve_alpha(_al)
        _fv = _ch2_objective(np.eye(4)[_k], _al)
        _fm = _ch2_objective(_wmx, _al)
        if _inside:
            assert np.max(_wmx) < 0.999 and _fm - _fv > 1e-9, "alpha=%.5f: no interior optimizer strictly better than vertex %d" % (_al, _k + 1)
        else:
            assert _wmx[_k] > 0.999 and abs(_fm - _fv) < 1e-9, "alpha=%.5f: the optimizer is not vertex %d" % (_al, _k + 1)
        if _off == CH2_EDGE_DELTA_VERTEX:
            # the global edge lies between this level and the bisection edge: both must print alike
            assert _ch2_printed(_al) == _ed["printed"], "%s edge: the global edge may print as %s, not %s" % (
                _ed["edge"], _ch2_printed(_al), _ed["printed"])
        _edge_checks.append({"edge": _ed["edge"], "position": _pos, "offset": _off, "alpha": _al,
                             "expected": "interior" if _inside else "vertex %d (%s)" % (_k + 1, _ed["vertex"]),
                             "w1": _wmx[0], "w2": _wmx[1], "w3": _wmx[2], "w4": _wmx[3], "max_weight": float(np.max(_wmx)),
                             "objective": _fm, "vertex_objective": _fv, "gap_to_vertex": _fm - _fv,
                             "starts_within_1e-6_of_best": _nx, "alpha_printed": _ch2_printed(_al)})
_win_printed = "[%s, %s]" % (_ch2_printed(_alpha_lo_edge), _ch2_printed(_alpha_hi_edge))

_ch2_rows = [
    ("equal weights",                    np.full(4, 0.25)),
    ("maximum consensus",                _wAmax),
    ("minimum variance",                 _wVmin),
    ("balance, alpha=%.2f" % CH2_ALPHA_STAR, _w75),
]
_t25 = []
_t25_checks = []
for _emph, _w in _ch2_rows:
    _m, _s = _ch2_qa_closed(_w, CH2_MU, CH2_SD)
    _t25.append({"weight_vector": "(%s)" % ",".join("%.3f" % v for v in _w), "emphasis": _emph,
                 "A": _A(_w), "V": _V(_w), "fQA_mu": _m, "fQA_sigma": _s,
                 "w1": _w[0], "w2": _w[1], "w3": _w[2], "w4": _w[3],
                 "A_tilde": (_A(_w) - _Amin) / (_Amax - _Amin), "V_tilde": (_V(_w) - _Vmin) / (_Vmax - _Vmin),
                 "objective_alpha_star": _ch2_objective(_w, CH2_ALPHA_STAR)})
    _t25_checks.append({"emphasis": _emph, "A_exact": _A(_w),
                        "A_trapezoid_grid": _ch2_A_grid(_w, CH2_MU, CH2_SD),
                        "A_alg_qa_numeric": _ch2_A_numeric(_w, CH2_MU, CH2_SD),
                        "V_closed_form": _V(_w), "V_eq_confidence_quantile_numeric": _ch2_V_numeric(_w, CH2_MU, CH2_SD)})
for _c in _t25_checks:
    assert abs(_c["A_exact"] - _c["A_trapezoid_grid"]) < 1e-6, _c
    assert abs(_c["A_exact"] - _c["A_alg_qa_numeric"]) < 2e-3, _c
    assert abs(_c["V_closed_form"] - _c["V_eq_confidence_quantile_numeric"]) < 2e-3, _c

write("tab2-5_biobjective_example.xlsx", {
    "config": cfg({
        "table": "2.5", "label": "tab:biobj-example",
        "panel": "the four rounded normal fits of Table 2.1 (tab2-1 workbook): mu = %s, sigma = %s" % (
            CH2_MU.tolist(), CH2_SD.tolist()),
        "aggregate": "quantile average, closed form N(w.mu, (w.sigma)^2) by prop:qa-properties(iii)",
        "consensus_A": "def:consensus, A(w) = sum_i int min{f_i, f_QA} dx, each overlap of two normal densities evaluated exactly as the normal probability mass on the intervals cut by their crossing points (roots of log f_i = log f_QA), scipy.special.ndtr",
        "consensus_A_cross_check": "trapezoid rule on x in [-2, 12] with %d points (step %.3f), sheet cross_checks; its error at the density crossings is about 1e-7, enough to move the flat alpha* optimizer in its third decimal, which is why the exact form is the primary rule" % (
            len(CH2_X_GRID), CH2_X_GRID[1] - CH2_X_GRID[0]),
        "confidence_V": "def:confidence, V(w) = (w.sigma)^2 (equality case of thm:qa-moments on the location-scale panel)",
        "cross_checks": "A with f_QA from alg:qa on %d probability levels and V by eq:confidence-quantile on the same levels, sheet cross_checks" % len(CH2_P_GRID),
        "solver": "scipy.optimize.minimize, method SLSQP, bounds [0,1], equality sum w = 1, finite-difference gradients",
        "starts": "%d Dirichlet(1,1,1,1) starts from numpy.random.default_rng(%d) (first draw of the block) plus the 4 vertices and the centroid, %d starts per solve, sheet starts" % (
            CH2_N_STARTS, CH2_SEED, _nS),
        "ftol": "%.0e" % CH2_FTOL, "maxiter": CH2_MAXITER,
        "normalization": "min-max, bounds solved over the simplex by the same multi-start (sheet bounds)",
        "A_min": _Amin, "A_max": _Amax, "V_min": _Vmin, "V_max": _Vmax,
        "alpha_grid": "0.00 to 1.00 step 0.05 and 0.70 to 0.90 step 0.01, %d distinct levels (sheet alpha_sweep)" % len(CH2_ALPHA_GRID),
        "probes": "%d Dirichlet(1,1,1,1) points from the same generator (second draw of the block), used by the rule-12 optimality check of the alpha* row, not stored" % CH2_N_PROBES,
        "alpha_star": CH2_ALPHA_STAR,
        "polish": "for an interior optimizer the pair (m*, s*) = (w.mu, w.sigma) of the best SLSQP point is polished by Nelder-Mead in the aggregate plane, xatol %.0e, fatol %.0e, maxiter %d (exact for this panel: the objective depends on w only through that pair)" % (
            CH2_POLISH_XATOL, CH2_POLISH_FATOL, CH2_POLISH_MAXITER),
        "polish_xatol": "%.0e" % CH2_POLISH_XATOL, "polish_fatol": "%.0e" % CH2_POLISH_FATOL, "polish_maxiter": CH2_POLISH_MAXITER,
        "tie_break": "the optimizer is a segment of the simplex (the fibre of (m*, s*) under w -> (w.mu, w.sigma)); the reported vector is the midpoint of that segment, sheet level_set_alpha_star",
        "alpha_star_slsqp_point": "(%s)" % ",".join("%.6f" % v for v in _wb75),
        "alpha_star_objective": _f75,
        "alpha_star_starts_within_1e-6_of_best": _n75,
        "interior_alpha_window_on_grid": "[%.2f, %.2f]" % (min(_alphas_interior), max(_alphas_interior)),
        "interior_alpha_lower_edge": _alpha_lo_edge,
        "interior_alpha_upper_edge": _alpha_hi_edge,
        "interior_alpha_window_printed": _win_printed,
        "interior_alpha_edge_method": "bisection to an alpha bracket of %.0e on the sign of the largest one-sided directional derivative of the scalarized objective at the adjacent vertex, over the tangent cone of the (w.mu, w.sigma) hull (closed-form derivatives, %d directions refined by a bounded Brent search, xatol %.0e); bracket = the grid levels on either side of the edge of the grid window; cross-check = the extreme over %d directions of the level at which the directional derivative vanishes; sheet window_edges" % (
            CH2_EDGE_TOL, CH2_EDGE_N_DIR, CH2_EDGE_BRENT_XATOL, CH2_EDGE_N_DIR_RATIO),
        "interior_alpha_edge_check": "the full multi-start solve at each edge -/+ %.0e returns the vertex outside the window and an interior optimizer strictly better than the vertex inside it, and on the vertex side of each edge it returns the vertex again at the closer offset %.0e, sheet window_edge_checks. Because g(alpha) - f_vertex(alpha) is convex and nonnegative (g the optimal value, a maximum of functions affine in alpha), the vertex is optimal on an interval of alpha containing 0 (minimum variance) or 1 (maximum consensus), and past the bisection edge the vertex is not even locally optimal, so these solves place each global edge within %.0e of its bisection edge, on the vertex side, where it prints to the same two decimals" % (
            CH2_EDGE_DELTA, CH2_EDGE_DELTA_VERTEX, CH2_EDGE_DELTA_VERTEX),
        "edge_tolerance": "%.0e" % CH2_EDGE_TOL, "edge_directions": CH2_EDGE_N_DIR, "edge_brent_xatol": "%.0e" % CH2_EDGE_BRENT_XATOL,
        "edge_directions_ratio_cross_check": CH2_EDGE_N_DIR_RATIO, "edge_check_delta": "%.0e" % CH2_EDGE_DELTA,
        "edge_check_delta_vertex": "%.0e" % CH2_EDGE_DELTA_VERTEX,
        "seed": CH2_SEED,
        "synthetic": 1,
    }),
    "table_2_5": _t25,
    "bounds": [
        {"bound": "A_min", "value": _Amin, "w1": _wAmin[0], "w2": _wAmin[1], "w3": _wAmin[2], "w4": _wAmin[3], "starts_within_1e-6_of_best": _nAmin, "starts_total": _nS},
        {"bound": "A_max", "value": _Amax, "w1": _wAmax[0], "w2": _wAmax[1], "w3": _wAmax[2], "w4": _wAmax[3], "starts_within_1e-6_of_best": _nAmax, "starts_total": _nS},
        {"bound": "V_min", "value": _Vmin, "w1": _wVmin[0], "w2": _wVmin[1], "w3": _wVmin[2], "w4": _wVmin[3], "starts_within_1e-6_of_best": _nVmin, "starts_total": _nS},
        {"bound": "V_max", "value": _Vmax, "w1": _wVmax[0], "w2": _wVmax[1], "w3": _wVmax[2], "w4": _wVmax[3], "starts_within_1e-6_of_best": _nVmax, "starts_total": _nS},
    ],
    "alpha_sweep": _sweep,
    "window_edges": _edges,
    "window_edge_checks": _edge_checks,
    "level_set_alpha_star": [
        {"point": "endpoint 1", "w1": _ends75[0][0], "w2": _ends75[0][1], "w3": _ends75[0][2], "w4": _ends75[0][3], "objective": _ch2_objective(_ends75[0], CH2_ALPHA_STAR)},
        {"point": "endpoint 2", "w1": _ends75[-1][0], "w2": _ends75[-1][1], "w3": _ends75[-1][2], "w4": _ends75[-1][3], "objective": _ch2_objective(_ends75[-1], CH2_ALPHA_STAR)},
        {"point": "midpoint (reported)", "w1": _w75[0], "w2": _w75[1], "w3": _w75[2], "w4": _w75[3], "objective": _ch2_objective(_w75, CH2_ALPHA_STAR)},
        {"point": "SLSQP best", "w1": _wb75[0], "w2": _wb75[1], "w3": _wb75[2], "w4": _wb75[3], "objective": _f75},
    ],
    "cross_checks": _t25_checks,
    "starts": [{"start": k, "w1": s[0], "w2": s[1], "w3": s[2], "w4": s[3],
                "law": ("Dirichlet(1,1,1,1)" if k < CH2_N_STARTS else ("vertex" if k < CH2_N_STARTS + 4 else "centroid"))}
               for k, s in enumerate(CH2_STARTS)],
})
print("   Table 2.5 solved: A in [%.4f, %.4f], V in [%.4f, %.4f]; alpha=%.2f -> w = (%s), A %.4f, V %.4f, N(%.3f, %.3f^2), objective %.5f (%d/%d starts agree); interior for alpha in [%.2f, %.2f] on the grid, exact window (%.6f, %.6f), printed %s" % (
    _Amin, _Amax, _Vmin, _Vmax, CH2_ALPHA_STAR, ", ".join("%.3f" % v for v in _w75), _A(_w75), _V(_w75), _m75, _s75, _f75, _n75, _nS,
    min(_alphas_interior), max(_alphas_interior), _alpha_lo_edge, _alpha_hi_edge, _win_printed))



# ------------------------------------------------------------------ numerics --
def _erf(x):
    """Abramowitz-Stegun 7.1.26 (vectorized, |err| < 1.5e-7). Deterministic."""
    x = np.asarray(x, float)
    s = np.sign(x)
    a = np.abs(x)
    t = 1.0 / (1.0 + 0.3275911 * a)
    y = 1.0 - (((((1.061405429 * t - 1.453152027) * t) + 1.421413741) * t
                - 0.284496736) * t + 0.254829592) * t * np.exp(-a * a)
    return s * y


def _npdf(x):
    return np.exp(-0.5 * np.asarray(x, float) ** 2) / math.sqrt(2.0 * math.pi)


def _ncdf(x):
    """Approximate normal cdf via _erf (Abramowitz-Stegun 7.1.26).  No block
    uses it any more: the ex:hmpanel and vignette V solvers, their
    cross-checks and _crps_gauss all use the exact scipy.special.ndtr."""
    return 0.5 * (1.0 + _erf(np.asarray(x, float) / math.sqrt(2.0)))


def _norm_ppf(p):
    """Acklam's rational approximation to the standard normal quantile."""
    p = np.asarray(p, float)
    a = [-3.969683028665376e+01, 2.209460984245205e+02, -2.759285104469687e+02,
         1.383577518672690e+02, -3.066479806614716e+01, 2.506628277459239e+00]
    b = [-5.447609879822406e+01, 1.615858368580409e+02, -1.556989798598866e+02,
         6.680131188771972e+01, -1.328068155288572e+01]
    c = [-7.784894002430293e-03, -3.223964580411365e-01, -2.400758277161838e+00,
         -2.549732539343734e+00, 4.374664141464968e+00, 2.938163982698783e+00]
    d = [7.784695709041462e-03, 3.224671290700398e-01, 2.445134137142996e+00,
         3.754408661907416e+00]
    plow, phigh = 0.02425, 1 - 0.02425
    q = np.empty_like(p)
    lo, hi = p < plow, p > phigh
    mid = ~(lo | hi)
    if lo.any():
        u = np.sqrt(-2.0 * np.log(p[lo]))
        q[lo] = (((((c[0] * u + c[1]) * u + c[2]) * u + c[3]) * u + c[4]) * u + c[5]) / \
                ((((d[0] * u + d[1]) * u + d[2]) * u + d[3]) * u + 1.0)
    if hi.any():
        u = np.sqrt(-2.0 * np.log(1.0 - p[hi]))
        q[hi] = -(((((c[0] * u + c[1]) * u + c[2]) * u + c[3]) * u + c[4]) * u + c[5]) / \
                 ((((d[0] * u + d[1]) * u + d[2]) * u + d[3]) * u + 1.0)
    if mid.any():
        u = p[mid] - 0.5
        r = u * u
        q[mid] = (((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * u / \
                 (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1.0)
    return q


def _project_simplex(W):
    """Row-wise Euclidean projection onto the probability simplex."""
    W = np.atleast_2d(np.asarray(W, float))
    n = W.shape[1]
    U = np.sort(W, axis=1)[:, ::-1]
    css = np.cumsum(U, axis=1) - 1.0
    idx = np.arange(1, n + 1)
    rho = (U - css / idx > 0).sum(axis=1)
    theta = css[np.arange(W.shape[0]), rho - 1] / rho
    return np.maximum(W - theta[:, None], 0.0)


def _ascend(obj, starts, iters=220, step0=0.35):
    """Deterministic projected-gradient ascent (batched forward differences)."""
    best_w, best_v = None, -np.inf
    for w0 in starts:
        w = np.asarray(w0, float).copy()
        n = w.size
        eye = np.eye(n)
        for r in range(iters):
            h = 1e-6
            W = np.vstack([w[None, :], w[None, :] + h * eye])
            V = obj(W)
            g = (V[1:] - V[0]) / h
            eta = step0 / (1.0 + r / 40.0)
            w = _project_simplex(w + eta * g)[0]
        v = float(obj(w[None, :])[0])
        if v > best_v:
            best_v, best_w = v, w
    return best_w, best_v


def _gini_vec(u):
    """Gini coefficient of a nonnegative vector (0 for an even profile)."""
    u = np.asarray(u, float)
    n = u.size
    m = u.mean()
    if m <= 0:
        return 0.0
    return float(np.abs(u[:, None] - u[None, :]).sum() / (2.0 * n * n * m))


def _fs_phi(A, rho, theta):
    """Batched Fehr-Schmidt utilities -> power normalization -> Bonferroni
    Phi_{1,1}; A is (B x n) overlaps.  Returns (Phi (B,), Util (B x n))."""
    A = np.atleast_2d(A)
    n = A.shape[1]
    M = A[:, None, :] - A[:, :, None]          # M[b,i,j] = A_j - A_i
    D = np.clip(M, 0.0, None).sum(axis=2) / (n - 1)
    S = np.clip(-M, 0.0, None).sum(axis=2) / (n - 1)
    U = A - rho * D - theta * S
    with np.errstate(divide="ignore", invalid="ignore"):
        ratio = np.where(A > 0, U / np.where(A > 0, A, 1.0), 0.0)
    Ut = np.clip(ratio, 0.0, 1.0) ** 2
    S1 = Ut.sum(axis=1)
    S2 = (Ut ** 2).sum(axis=1)
    Phi = np.sqrt(np.clip((S1 ** 2 - S2) / (n * (n - 1)), 0.0, None))
    return Phi, Ut


# =====================================================================
# CASE STUDY I, FAIRNESS PANEL (Section 6.2) -- SOLVED HERE, NOT TYPED
#   -> Figure 6.3  fig6-3_bim_densities.xlsx  (fig:bim-densities)
#   -> Figure 6.4  fig6-4_bim_tradeoff.xlsx   (fig:bim-tradeoff)
#   -> Table  6.9  tab6-9_bim_results.xlsx    (tab:bim)
# The seven-expert tri-objective program eq:bim-triobjective is solved
# from the stated inputs (eq:bim-panel-mu, eq:bim-panel-sigma, rho, theta)
# with the book's conventions:
#   * F_QA(.;w) = N(w.mu, (w.sigma)^2)                   Prop. qa-properties(iii)
#   * A_i(w) = int_R min{f_i, f_QA} dx, whole real line   def:consensus
#     (closed form through the two crossing points of the densities)
#   * Abar = (1/N) sum_i A_i,  V = (w.sigma)^2            def:confidence
#   * U_i^FS on the overlap profile, rho = 0.4, vartheta = 0.2,
#     divisor N-1                                        def:ifcu
#   * power normalization, kappa = 2                     eq:parabola-norm
#   * Bonferroni mean of order (1,1)                     def:collective-fairness
#   * min-max normalization with the bounds of each objective SOLVED over
#     the simplex by the same multi-start procedure      Algorithm 2.2
#   * max alpha*At - beta*Vt + delta*Pt, alpha = 2(1-delta)/3,
#     beta = (1-delta)/3, delta on the grid {0, 0.1, ..., 0.6}
#   * multi-start SQP = scipy SLSQP from 200 starting points drawn
#     uniformly on the simplex (Dirichlet(1,...,1)) by ONE generator
#     numpy.random.default_rng(202603), consumed in the documented order
#     (bounds first, then the sweep); the first start attaining the
#     best objective is retained.
#   * certificate: a deterministic dense grid of the convex hull of the
#     (mu_i, sigma_i), over which (w.mu, w.sigma) ranges, must not improve
#     on any recorded bound or sweep optimum (sheet hull_scan of fig6-4);
#     it draws nothing, so the solved numbers above are unaffected.
# Everything the book prints for the fairness panel (weights, aggregates,
# overlaps, fairness, the sweep, the start agreement of the bound programs,
# the derived sweep quantities of sheet text_values, the band probabilities
# of the managerial paragraph in sheet band_probabilities of tab6-9, and
# the tail masses outside [0,10] in sheet tail_mass of fig6-3) is read
# back from the workbooks written by this block, and the rule-12 pass at
# the end checks the solved outputs against optimality,
# closed-form invariants and the claims the text makes of them.
# =====================================================================
from scipy.optimize import minimize as _sp_minimize
from scipy.special import ndtr as _ndtr
import scipy as _scipy

_BIM_MU = np.array([7.1, 7.6, 6.8, 7.3, 4.7, 6.6, 7.0])   # eq:bim-panel-mu
_BIM_SD = np.array([1.2, 1.3, 1.1, 1.2, 1.3, 1.1, 1.2])   # eq:bim-panel-sigma
_BIM_N = len(_BIM_MU)
_BIM_RHO, _BIM_THETA, _BIM_KAPPA = 0.4, 0.2, 2          # ch6 sec:bim-data, eq:parabola-norm
_BIM_DELTAS = [0.0, 0.1, 0.2, 0.3, 0.4, 0.5, 0.6]        # the sweep grid
_BIM_N_STARTS = 200
_BIM_FTOL = 1e-12
_BIM_MAXITER = 1000
_BIM_TIE_TOL = 1e-6      # starts whose objective lies within this of the best are counted as ties
_BIM_DISSENTER = 4       # expert 5 (0-based index)


def _bim_overlap(m1, s1, m2, s2):
    """A = int_R min{N(m1,s1^2), N(m2,s2^2)} dx in closed form.
    Equal spreads: 2*Phi(-|m1-m2|/(2 s)).  Otherwise the two densities cross
    at the roots of a quadratic; the narrower density is the larger one
    between the roots, so the minimum is the wider one there and the
    narrower one outside."""
    if abs(s1 - s2) < 1e-12:
        return float(2.0 * _ndtr(-abs(m1 - m2) / (2.0 * s1)))
    if s1 > s2:                      # make (m1, s1) the narrower density
        m1, s1, m2, s2 = m2, s2, m1, s1
    a = 1.0 / s1 ** 2 - 1.0 / s2 ** 2
    b = -2.0 * (m1 / s1 ** 2 - m2 / s2 ** 2)
    c = m1 ** 2 / s1 ** 2 - m2 ** 2 / s2 ** 2 - 2.0 * math.log(s2 / s1)
    disc = max(b * b - 4.0 * a * c, 0.0)
    r1 = (-b - math.sqrt(disc)) / (2.0 * a)
    r2 = (-b + math.sqrt(disc)) / (2.0 * a)
    x1, x2 = min(r1, r2), max(r1, r2)
    F1 = lambda x: _ndtr((x - m1) / s1)
    F2 = lambda x: _ndtr((x - m2) / s2)
    return float((F2(x2) - F2(x1)) + F1(x1) + (1.0 - F1(x2)))


def _bim_overlap_trapz(m1, s1, m2, s2, n=200001):
    """The same overlap by the trapezoid rule on a wide grid (check only)."""
    lo = min(m1 - 8 * s1, m2 - 8 * s2)
    hi = max(m1 + 8 * s1, m2 + 8 * s2)
    x = np.linspace(lo, hi, n)
    f1 = np.exp(-0.5 * ((x - m1) / s1) ** 2) / (s1 * math.sqrt(2 * math.pi))
    f2 = np.exp(-0.5 * ((x - m2) / s2) ** 2) / (s2 * math.sqrt(2 * math.pi))
    return float(_trapz(np.minimum(f1, f2), x))


def _bim_aggregate(w):
    return float(w @ _BIM_MU), float(w @ _BIM_SD)


def _bim_overlaps(w):
    m, s = _bim_aggregate(w)
    return np.array([_bim_overlap(_BIM_MU[i], _BIM_SD[i], m, s) for i in range(_BIM_N)])


def _bim_Abar(w):
    return float(_bim_overlaps(w).mean())


def _bim_V(w):
    return float(w @ _BIM_SD) ** 2


def _bim_Phi(w, return_util=False):
    A = _bim_overlaps(w)
    Phi, Ut = _fs_phi(A[None, :], _BIM_RHO, _BIM_THETA)   # kappa = 2 inside _fs_phi
    return (float(Phi[0]), Ut[0]) if return_util else float(Phi[0])


_BIM_CONS = ({"type": "eq", "fun": lambda w: float(w.sum() - 1.0)},)
_BIM_BNDS = [(0.0, 1.0)] * _BIM_N


def _bim_multistart(fun, rng, sense="max"):
    """Multi-start SLSQP over the simplex.  Draws _BIM_N_STARTS starting
    points from rng (Dirichlet(1,...,1)); returns (w_best, value, n_ties,
    n_converged).  The first start attaining the best value is kept."""
    sign = -1.0 if sense == "max" else 1.0
    starts = rng.dirichlet(np.ones(_BIM_N), size=_BIM_N_STARTS)
    best_w, best_v, vals = None, np.inf, []
    for w0 in starts:
        r = _sp_minimize(lambda w: sign * fun(w), w0, method="SLSQP",
                         bounds=_BIM_BNDS, constraints=_BIM_CONS,
                         options={"ftol": _BIM_FTOL, "maxiter": _BIM_MAXITER})
        w = np.clip(r.x, 0.0, 1.0)
        w = w / w.sum()
        v = sign * fun(w)
        vals.append(v)
        if v < best_v:
            best_v, best_w = v, w
    vals = np.asarray(vals)
    n_ties = int(np.sum(np.abs(vals - best_v) < _BIM_TIE_TOL))
    return best_w, sign * best_v, n_ties


print("\n-- Case I fairness panel: solving eq:bim-triobjective (seed %d, %d SLSQP starts per program) ..."
      % (NEW_SEED, _BIM_N_STARTS))
_bim_rng = np.random.default_rng(NEW_SEED)
_bim_draw_order = []

# ----- Algorithm 2.2: min and max of each objective over the simplex ----------
_bim_bounds = {}
for _name, _fun in (("Abar", _bim_Abar), ("V", _bim_V), ("Phi", _bim_Phi)):
    _wmin, _vmin, _tmin = _bim_multistart(_fun, _bim_rng, sense="min")
    _bim_draw_order.append("%s min" % _name)
    _wmax, _vmax, _tmax = _bim_multistart(_fun, _bim_rng, sense="max")
    _bim_draw_order.append("%s max" % _name)
    _bim_bounds[_name] = dict(min=_vmin, max=_vmax, argmin=_wmin, argmax=_wmax,
                              ties_min=_tmin, ties_max=_tmax)
    print("   %-4s in [%.6f, %.6f]" % (_name, _vmin, _vmax))
_bim_Amin, _bim_Amax = _bim_bounds["Abar"]["min"], _bim_bounds["Abar"]["max"]
_bim_Vmin, _bim_Vmax = _bim_bounds["V"]["min"], _bim_bounds["V"]["max"]
_bim_Pmin, _bim_Pmax = _bim_bounds["Phi"]["min"], _bim_bounds["Phi"]["max"]


def _bim_scalarized(w, delta):
    alpha, beta = 2.0 * (1.0 - delta) / 3.0, (1.0 - delta) / 3.0
    At = (_bim_Abar(w) - _bim_Amin) / (_bim_Amax - _bim_Amin)
    Vt = (_bim_V(w) - _bim_Vmin) / (_bim_Vmax - _bim_Vmin)
    Pt = (_bim_Phi(w) - _bim_Pmin) / (_bim_Pmax - _bim_Pmin)
    return alpha * At - beta * Vt + delta * Pt


# ----- the delta sweep ---------------------------------------------------------
_bim_rows, _bim_wrows, _bim_arows, _bim_sol = [], [], [], {}
for _d in _BIM_DELTAS:
    _w, _val, _ties = _bim_multistart(lambda w, d=_d: _bim_scalarized(w, d), _bim_rng, sense="max")
    _bim_draw_order.append("sweep delta=%.1f" % _d)
    _A = _bim_overlaps(_w)
    _P, _Ut = _bim_Phi(_w, return_util=True)
    _m, _s = _bim_aggregate(_w)
    _bim_sol[_d] = dict(w=_w, Abar=float(_A.mean()), Phi=_P, minA=float(_A.min()),
                        mu=_m, sd=_s, obj=_val, ties=_ties)
    _bim_rows.append({
        "delta": _d, "alpha": 2.0 * (1.0 - _d) / 3.0, "beta": (1.0 - _d) / 3.0,
        "consensus_Abar": float(_A.mean()), "fairness_Phibar": _P,
        "minimum_overlap": float(_A.min()),
        "argmin_overlap_expert": int(np.argmin(_A)) + 1,
        "objective": _val, "starts_at_best": _ties, "n_starts": _BIM_N_STARTS,
        "agg_mu": _m, "agg_sd": _s, "weight_on_dissenter": float(_w[_BIM_DISSENTER]),
    })
    _bim_wrows.append(dict({"delta": _d}, **{"w%d" % (i + 1): float(_w[i]) for i in range(_BIM_N)}))
    _bim_arows.append(dict({"delta": _d},
                           **{"A%d" % (i + 1): float(_A[i]) for i in range(_BIM_N)},
                           **{"Unorm%d" % (i + 1): float(_Ut[i]) for i in range(_BIM_N)}))
    print("   delta=%.1f  obj=%.6f  (%d/%d starts at best)  w=%s  agg=N(%.3f,%.3f^2)  Abar=%.4f Phi=%.4f minA=%.4f"
          % (_d, _val, _ties, _BIM_N_STARTS, np.round(_w, 4), _m, _s, _A.mean(), _P, _A.min()))

# closed-form overlap against the trapezoid rule at the sweep solutions
_bim_overlap_err = 0.0
for _d in _BIM_DELTAS:
    _m, _s = _bim_sol[_d]["mu"], _bim_sol[_d]["sd"]
    for _i in range(_BIM_N):
        _bim_overlap_err = max(_bim_overlap_err,
                               abs(_bim_overlap(_BIM_MU[_i], _BIM_SD[_i], _m, _s)
                                   - _bim_overlap_trapz(_BIM_MU[_i], _BIM_SD[_i], _m, _s)))
print("   closed-form overlap vs trapezoid at the sweep solutions: max |diff| = %.2e" % _bim_overlap_err)


# ----- certificate: a dense scan of the attainable (mean, spread) region ---------
# Every objective depends on w only through the aggregate's (m, s) = (w.mu,
# w.sigma) (Prop. qa-properties(iii)), and (m, s) ranges exactly over the
# convex hull of the seven points (mu_i, sigma_i).  A dense grid of that
# planar region (deterministic, no draws, so the generator above is not
# touched) must not improve on any recorded bound or sweep optimum, and the
# gap it leaves measures its resolution.  This is the evidence the text cites
# for the lower fairness bound, which few of the random starts reach.
_BIM_SCAN_S_LEVELS = 801      # spread levels, sigma_min .. sigma_max (step 0.00025)
_BIM_SCAN_M_POINTS = 1601     # equally spaced means across the hull at each level
_BIM_SCAN_CHUNK = 50000       # rows per batched fairness evaluation (memory only)
_BIM_SCAN_IMPROVE_TOL = 1e-9  # a grid point may not beat a recorded value by more
_BIM_SCAN_GAP_TOL = 1e-4      # the grid must come this close to every recorded value
_BIM_SCAN_LOC_TOL = (0.01, 0.005)   # grid argmax within (|dm|, |ds|) of the recorded sweep optimum


def _bim_overlap_vec(mi, si, m, s):
    """Vectorized closed-form overlap of N(mi, si^2) with N(m, s^2) over the
    whole real line (the same two-crossing rule as _bim_overlap, with the
    numerically stable root formula)."""
    m = np.asarray(m, float)
    s = np.asarray(s, float)
    out = np.empty_like(m)
    eq = np.abs(s - si) < 1e-9
    out[eq] = 2.0 * _ndtr(-np.abs(m[eq] - mi) / (2.0 * si))
    ne = ~eq
    mm, ss = m[ne], s[ne]
    nar = ss < si
    m1 = np.where(nar, mm, mi)
    s1 = np.where(nar, ss, si)
    m2 = np.where(nar, mi, mm)
    s2 = np.where(nar, si, ss)
    a = 1.0 / s1 ** 2 - 1.0 / s2 ** 2
    b = -2.0 * (m1 / s1 ** 2 - m2 / s2 ** 2)
    c = m1 ** 2 / s1 ** 2 - m2 ** 2 / s2 ** 2 - 2.0 * np.log(s2 / s1)
    q = -0.5 * (b + np.copysign(np.sqrt(np.maximum(b * b - 4.0 * a * c, 0.0)), b))
    r1, r2 = q / a, c / q
    x1, x2 = np.minimum(r1, r2), np.maximum(r1, r2)
    out[ne] = ((_ndtr((x2 - m2) / s2) - _ndtr((x1 - m2) / s2))
               + _ndtr((x1 - m1) / s1) + (1.0 - _ndtr((x2 - m1) / s1)))
    return out


def _bim_hull_grid():
    """Grid of the convex hull of the (mu_i, sigma_i): at each spread level the
    extreme means lie on hull edges, i.e. on segments between two experts."""
    s_lv = np.linspace(_BIM_SD.min(), _BIM_SD.max(), _BIM_SCAN_S_LEVELS)
    lo = np.full(s_lv.size, np.inf)
    hi = np.full(s_lv.size, -np.inf)
    for i in range(_BIM_N):
        for j in range(_BIM_N):
            si, sj, mi, mj = _BIM_SD[i], _BIM_SD[j], _BIM_MU[i], _BIM_MU[j]
            if abs(sj - si) < 1e-12:
                cand = np.where(np.abs(s_lv - si) < 1e-9, mi, np.nan)
            elif si < sj:
                on = (s_lv >= si - 1e-12) & (s_lv <= sj + 1e-12)
                cand = np.where(on, mi + (mj - mi) * (s_lv - si) / (sj - si), np.nan)
            else:
                continue
            ok = ~np.isnan(cand)
            lo[ok] = np.minimum(lo[ok], cand[ok])
            hi[ok] = np.maximum(hi[ok], cand[ok])
    t = np.linspace(0.0, 1.0, _BIM_SCAN_M_POINTS)
    M = lo[:, None] + (hi - lo)[:, None] * t[None, :]
    S = np.repeat(s_lv[:, None], _BIM_SCAN_M_POINTS, axis=1)
    return M.ravel(), S.ravel()


def _bim_hull_scan(bounds, sweep):
    """bounds: {name: (min, max)} for Abar, V, Phi; sweep: {delta: (objective,
    agg_mu, agg_sd)}.  Returns one row per program with the recorded value, the
    best grid value, the grid's improvement on the recorded value (must be
    <= _BIM_SCAN_IMPROVE_TOL) and its gap (resolution)."""
    m, s = _bim_hull_grid()
    A = np.column_stack([_bim_overlap_vec(_BIM_MU[i], _BIM_SD[i], m, s) for i in range(_BIM_N)])
    vals = {"Abar": A.mean(axis=1), "V": s ** 2,
            "Phi": np.concatenate([_fs_phi(A[k:k + _BIM_SCAN_CHUNK], _BIM_RHO, _BIM_THETA)[0]
                                   for k in range(0, A.shape[0], _BIM_SCAN_CHUNK)])}
    rows = []
    for name in ("Abar", "V", "Phi"):
        lo, hi = bounds[name]
        for sense, rec in (("min", lo), ("max", hi)):
            k = int(np.argmin(vals[name]) if sense == "min" else np.argmax(vals[name]))
            g = float(vals[name][k])
            imp = (rec - g) if sense == "min" else (g - rec)
            rows.append({"program": "%s %s" % (name, sense), "recorded": float(rec), "grid_best": g,
                         "grid_improvement": imp, "grid_gap": -imp,
                         "grid_m": float(m[k]), "grid_s": float(s[k]), "recorded_m": np.nan, "recorded_s": np.nan})
    tl = {n: (vals[n] - bounds[n][0]) / (bounds[n][1] - bounds[n][0]) for n in ("Abar", "V", "Phi")}
    for d, (obj, mu_d, sd_d) in sweep.items():
        S = 2.0 * (1.0 - d) / 3.0 * tl["Abar"] - (1.0 - d) / 3.0 * tl["V"] + d * tl["Phi"]
        k = int(np.argmax(S))
        rows.append({"program": "sweep delta=%.1f" % d, "recorded": float(obj), "grid_best": float(S[k]),
                     "grid_improvement": float(S[k]) - float(obj), "grid_gap": float(obj) - float(S[k]),
                     "grid_m": float(m[k]), "grid_s": float(s[k]), "recorded_m": float(mu_d), "recorded_s": float(sd_d)})
    return rows, int(m.size)


_bim_scan_rows, _bim_scan_n = _bim_hull_scan(
    {"Abar": (_bim_Amin, _bim_Amax), "V": (_bim_Vmin, _bim_Vmax), "Phi": (_bim_Pmin, _bim_Pmax)},
    {_d: (_bim_sol[_d]["obj"], _bim_sol[_d]["mu"], _bim_sol[_d]["sd"]) for _d in _BIM_DELTAS})
_bim_scan_imp = max(r["grid_improvement"] for r in _bim_scan_rows)
_bim_scan_gap = max(r["grid_gap"] for r in _bim_scan_rows)
print("   hull-scan certificate on %d grid points: largest improvement on a recorded value %.2e,"
      " largest gap %.2e" % (_bim_scan_n, _bim_scan_imp, _bim_scan_gap))

_bim_solver_cfg = {
    "seed": NEW_SEED,
    "rng": "numpy.random.default_rng(202603), one generator for the whole block",
    "draw_order": " -> ".join(_bim_draw_order) + " (%d Dirichlet(1,...,1) starts each)" % _BIM_N_STARTS,
    "solver": "scipy.optimize.minimize(method='SLSQP') with bounds [0,1]^7 and the equality sum(w)=1",
    "scipy_version": _scipy.__version__, "numpy_version": np.__version__,
    "n_starts": _BIM_N_STARTS,
    "start_law": "uniform on the simplex, Dirichlet(1,...,1)",
    "ftol": _BIM_FTOL, "maxiter": _BIM_MAXITER,
    "tie_rule": "the first start attaining the best objective (highest in a maximization, lowest in a minimization) is retained; starts_at_best, starts_at_min and starts_at_max count starts within %g of it" % _BIM_TIE_TOL,
    "post_processing": "clip to [0,1] and renormalize to sum 1, then re-evaluate",
    "overlap_rule": "A_i = int_R min{f_i, f_QA} dx over the whole real line, closed form via the crossing points (def:consensus); checked against the trapezoid rule on [min mu-8sd, max mu+8sd] with 200001 points",
    "overlap_check_max_abs_diff": _bim_overlap_err,
    "consensus": "Abar = mean_i A_i", "confidence": "V = (w.sigma)^2 (def:confidence, Prop. qa-properties(iii))",
    "fairness": "Fehr-Schmidt on the overlap profile, divisor N-1 (def:ifcu), power normalization (eq:parabola-norm), Bonferroni (1,1) (def:collective-fairness)",
    "rho": _BIM_RHO, "vartheta": _BIM_THETA, "kappa": _BIM_KAPPA,
    "scalarization": "alpha*At - beta*Vt + delta*Pt, alpha = 2(1-delta)/3, beta = (1-delta)/3, min-max normalized objectives",
    "delta_grid": ", ".join("%.1f" % d for d in _BIM_DELTAS),
    "bounds_method": "min and max of each objective solved over the simplex by the same multi-start SLSQP (Algorithm 2.2)",
    "Abar_min": _bim_Amin, "Abar_max": _bim_Amax,
    "V_min": _bim_Vmin, "V_max": _bim_Vmax,
    "Phi_min": _bim_Pmin, "Phi_max": _bim_Pmax,
    "panel_mu": ", ".join("%.1f" % v for v in _BIM_MU),
    "panel_sigma": ", ".join("%.1f" % v for v in _BIM_SD),
    "synthetic": 1,
}
_bim_bounds_sheet = [
    {"objective": k, "min": v["min"], "max": v["max"],
     "argmin_w": ", ".join("%.4f" % x for x in v["argmin"]),
     "argmax_w": ", ".join("%.4f" % x for x in v["argmax"]),
     "starts_at_min": v["ties_min"], "starts_at_max": v["ties_max"]}
    for k, v in _bim_bounds.items()
]

# ---- derived quantities the text reads off the solved sweep (bookkeeping) ----
# Differences, subset extremes and identities of the solved rows that the
# results and robustness paragraphs print; every entry is recomputed from the
# series/weights/overlaps sheets by the rule-12 pass at the end.
_bim_Aof = {_d: _bim_overlaps(_bim_sol[_d]["w"]) for _d in _BIM_DELTAS}
_bim_vtx = int(np.argmax(_bim_sol[0.0]["w"]))                      # the delta = 0 vertex (0-based)
_bim_major = [i for i in range(_BIM_N) if i not in (_bim_vtx, _BIM_DISSENTER)]
_bim_lo_i = min(_bim_major, key=lambda i: _bim_Aof[0.0][i])
_bim_hi_i = max(_bim_major, key=lambda i: _bim_Aof[0.0][i])
_bim_w04 = _bim_sol[0.4]["w"].copy()
_bim_off_i = max((i for i in range(_BIM_N) if i != _bim_vtx), key=lambda i: _bim_w04[i])
_bim_two05 = list(np.argsort(_bim_Aof[0.5])[:2])
_bim_min06 = int(np.argmin(_bim_Aof[0.6]))


def _bim_tv(q, v, expert=None, nd=3):
    return {"quantity": q, "value": float(v), "expert": ("" if expert is None else int(expert) + 1),
            "printed": round(float(v), nd)}


_bim_text_values = [
    _bim_tv("delta=0: overlap of the vertex expert with the aggregate", _bim_Aof[0.0][_bim_vtx], _bim_vtx),
    _bim_tv("delta=0: lowest overlap among the other majority experts", _bim_Aof[0.0][_bim_lo_i], _bim_lo_i, 2),
    _bim_tv("delta=0: highest overlap among the other majority experts", _bim_Aof[0.0][_bim_hi_i], _bim_hi_i, 2),
    _bim_tv("delta=0.4: weight moved off the vertex expert, onto this expert", _bim_w04[_bim_off_i], _bim_off_i),
    _bim_tv("delta=0.4: weight on the dissenter", _bim_w04[_BIM_DISSENTER], _BIM_DISSENTER),
    _bim_tv("delta=0.4 minus delta=0: mean overlap (consensus)", _bim_sol[0.4]["Abar"] - _bim_sol[0.0]["Abar"], None, 4),
    _bim_tv("delta=0.4 minus delta=0: collective fairness", _bim_sol[0.4]["Phi"] - _bim_sol[0.0]["Phi"], None, 4),
    _bim_tv("delta=0.4 minus delta=0: aggregate mean", _bim_sol[0.4]["mu"] - _bim_sol[0.0]["mu"], None, 4),
    _bim_tv("delta=0.4 minus delta=0: minimum overlap", _bim_sol[0.4]["minA"] - _bim_sol[0.0]["minA"], None, 4),
    _bim_tv("delta=0.4: dissenter's overlap", _bim_Aof[0.4][_BIM_DISSENTER], _BIM_DISSENTER),
    _bim_tv("delta=0.5: overlap of the least-represented expert", _bim_Aof[0.5][_bim_two05[0]], _bim_two05[0]),
    _bim_tv("delta=0.5: overlap of the second least-represented expert", _bim_Aof[0.5][_bim_two05[1]], _bim_two05[1]),
    _bim_tv("delta=0.6: dissenter's overlap", _bim_Aof[0.6][_BIM_DISSENTER], _BIM_DISSENTER),
    _bim_tv("delta=0.6: minimum overlap, attained by this expert", _bim_Aof[0.6][_bim_min06], _bim_min06),
    _bim_tv("delta=0.6 minus delta=0.5: minimum overlap", _bim_sol[0.6]["minA"] - _bim_sol[0.5]["minA"], None, 4),
    _bim_tv("delta=0.6 minus delta=0.5: mean overlap (consensus)", _bim_sol[0.6]["Abar"] - _bim_sol[0.5]["Abar"], None, 4),
    _bim_tv("delta=0.6 minus delta=0.5: collective fairness", _bim_sol[0.6]["Phi"] - _bim_sol[0.5]["Phi"], None, 4),
]

# ---- Figure 6.4: the sweep ---------------------------------------------------
# The three plotted curves are the mean overlap (consensus), the collective
# fairness and the minimum overlap, each an unnormalized diagnostic value, not
# the scalarized objective, hence the y-axis label.
write("fig6-4_bim_tradeoff.xlsx", {
    "config": cfg(dict({
        "figure": "6.4", "label": "fig:bim-tradeoff",
        "x_label": "fairness weight delta", "y_label": "overlap or fairness value",
        "ymin": 0.3, "ymax": 1.0, "title": "BIM fairness-consensus trade-off",
        "generator": "build_workbooks.py, Case I fairness-panel block (solved)",
    }, **_bim_solver_cfg, **{
        "certificate": "dense grid of the convex hull of the seven (mu_i, sigma_i), over which (w.mu, w.sigma) ranges and through which alone Abar, V and Phi depend on w; no grid point may improve on a recorded bound or sweep optimum by more than hull_scan_improve_tol (sheet hull_scan); deterministic, no draws",
        "hull_scan_grid": "%d spread levels from sigma_min to sigma_max, %d equally spaced means across the hull at each level" % (_BIM_SCAN_S_LEVELS, _BIM_SCAN_M_POINTS),
        "hull_scan_points": _bim_scan_n,
        "hull_scan_improve_tol": _BIM_SCAN_IMPROVE_TOL,
        "hull_scan_gap_tol": _BIM_SCAN_GAP_TOL,
        "hull_scan_max_improvement": _bim_scan_imp,
        "hull_scan_max_gap": _bim_scan_gap,
        "bound_start_agreement": "starts_at_min / starts_at_max of the bounds sheet count the starts within %g of each bound" % _BIM_TIE_TOL,
    })),
    "series": pd.DataFrame(_bim_rows),
    "weights": pd.DataFrame(_bim_wrows),
    "overlaps": pd.DataFrame(_bim_arows),
    "bounds": pd.DataFrame(_bim_bounds_sheet),
    "hull_scan": pd.DataFrame(_bim_scan_rows),
    "text_values": pd.DataFrame(_bim_text_values),
})

# ---- Figure 6.3: the seven experts and the two solved aggregates -------------
_bim_expert_rows = []
for _i in range(_BIM_N):
    _bim_expert_rows.append({
        "legend": "dissenter (expert 5)" if _i == _BIM_DISSENTER else "_nolegend_",
        "mu": float(_BIM_MU[_i]), "sigma": float(_BIM_SD[_i]), "weight": 1,
        "role": "expert_highlight" if _i == _BIM_DISSENTER else "expert",
        "linestyle": "solid"})
# The seven judgments are untruncated normals on the latent maturity score
# (sec:bim-elicit), so part of each lies outside the [0,10] rubric and is left
# in the tails.  The panel paragraph of sec:bim-data prints the largest share
# (expert 2, above 10) and names the expert with the next largest share
# (expert 4) with that share rounded, not as a bound; all are read from sheet
# tail_mass (exact normal cdf, no draws).
_BIM_SCALE_LO, _BIM_SCALE_HI = 0.0, 10.0      # the anchoring rubric [0,10] of tab:bim-levels (stated)
_bim_tail_rows = []
for _i in range(_BIM_N):
    _plo = float(_ndtr((_BIM_SCALE_LO - _BIM_MU[_i]) / _BIM_SD[_i]))
    _phi = float(_ndtr((_BIM_MU[_i] - _BIM_SCALE_HI) / _BIM_SD[_i]))
    _bim_tail_rows.append({"expert": _i + 1, "mu": float(_BIM_MU[_i]), "sigma": float(_BIM_SD[_i]),
                           "P_below_0": _plo, "P_above_10": _phi, "P_outside_0_10": _plo + _phi,
                           "P_outside_printed": round(_plo + _phi, 3)})
_bim_tail_out = np.array([r["P_outside_0_10"] for r in _bim_tail_rows])
_bim_tail_top = int(np.argmax(_bim_tail_out))                             # 0-based
_bim_tail_second = int(np.argsort(_bim_tail_out)[-2])                    # 0-based, the next largest share
_bim_tail_rest = float(_bim_tail_out[_bim_tail_second])

# The delta = 0 solution is a vertex, so its aggregate IS that expert's
# distribution and the thick aggregate curve hides the expert's dotted one;
# the legend names the expert (computed from the solved weights).  The
# legends write delta as the symbol (matplotlib mathtext), as the caption,
# the text and the axis label of Figure 6.4 do.
_bim_d0_is_vertex = bool(_bim_sol[0.0]["w"].max() > 1.0 - 1e-9)
_bim_agg_rows = [
    {"legend": (r"aggregate, $\delta=0$ (= expert %d)" % (_bim_vtx + 1)) if _bim_d0_is_vertex else r"aggregate, $\delta=0$",
     "mu": _bim_sol[0.0]["mu"], "sigma": _bim_sol[0.0]["sd"],
     "weight": 1, "role": "aggregate", "linestyle": "solid"},
    {"legend": r"aggregate, $\delta=0.5$", "mu": _bim_sol[0.5]["mu"], "sigma": _bim_sol[0.5]["sd"],
     "weight": 1, "role": "aggregate", "linestyle": "dashed"},
]
write("fig6-3_bim_densities.xlsx", {
    "config": cfg({
        "figure": "6.3", "label": "fig:bim-densities",
        "x_label": "maturity score", "y_label": "density",
        "domain_min": 0, "domain_max": 10, "ymax": 0.42,
        "draw_linear_pool": 0, "title": "BIM maturity: experts and aggregates",
        "generator": "build_workbooks.py, Case I fairness-panel block (solved)",
        "seed": NEW_SEED,
        "experts": "stated inputs eq:bim-panel-mu / eq:bim-panel-sigma (no random draws)",
        "aggregates": "N(w.mu, (w.sigma)^2) at the solved w(delta=0) and w(delta=0.5) of fig6-4_bim_tradeoff.xlsx",
        "delta0_vertex_expert": (_bim_vtx + 1) if _bim_d0_is_vertex else "none",
        "delta0_legend_note": "the delta = 0 solution puts all weight on one expert, so its aggregate coincides with, and is drawn over, that expert's dotted curve; the legend entry names the expert",
        "tail_rule": "sheet tail_mass: P(x < %g) and P(x > %g) of each untruncated expert judgment N(mu_i, sigma_i^2), exact normal cdf, the total printed to 3 decimals (panel paragraph of sec:bim-data)" % (_BIM_SCALE_LO, _BIM_SCALE_HI),
        "tail_largest_expert": _bim_tail_top + 1,
        "tail_largest_outside": float(_bim_tail_out[_bim_tail_top]),
        "tail_largest_outside_printed": round(float(_bim_tail_out[_bim_tail_top]), 3),
        "tail_second_expert": _bim_tail_second + 1,
        "tail_second_rule": "the expert with the next largest mass outside [0,10] is named in the text with that mass rounded to 3 decimals (a value, not a bound: the unrounded share may exceed the printed one)",
        "tail_largest_among_others": _bim_tail_rest,
        "tail_largest_among_others_printed": round(_bim_tail_rest, 3),
        "solver_settings": "see the config sheet of fig6-4_bim_tradeoff.xlsx (same block, same seed)",
        "synthetic": 1,
    }),
    "gaussians": _bim_expert_rows + _bim_agg_rows,
    "aggregates": [
        dict({"delta": d, "mu": _bim_sol[d]["mu"], "sigma": _bim_sol[d]["sd"]},
             **{"w%d" % (i + 1): float(_bim_sol[d]["w"][i]) for i in range(_BIM_N)})
        for d in (0.0, 0.5)],
    "annotations": [{"text": "dissenter", "x": 4.7, "y": 0.31, "ha": "center"}],
    "tail_mass": _bim_tail_rows,
})

# ---- Table 6.9: the two solutions compared -------------------------------------
_bim_tab_rows, _bim_tab_printed = [], []
for _q, _key, _nd in (("aggregate distribution", None, 2),
                      ("weight on dissenter (exp 5)", "wdis", 3),
                      ("minimum overlap", "minA", 3),
                      ("mean overlap (consensus)", "Abar", 3),
                      ("collective fairness", "Phi", 3)):
    _vals, _prt = {}, {}
    for _col, _d in (("consensus_confidence_d0", 0.0), ("fairness_weighted_d0.5", 0.5)):
        _S = _bim_sol[_d]
        if _key is None:
            _vals[_col] = "N(%.6f, %.6f^2)" % (_S["mu"], _S["sd"])
            _prt[_col] = "N(%.2f, %.2f^2)" % (round(_S["mu"], 2), round(_S["sd"], 2))
        else:
            _v = float(_S["w"][_BIM_DISSENTER]) if _key == "wdis" else float(_S[_key])
            _vals[_col] = _v
            _prt[_col] = round(_v, _nd)
    _bim_tab_rows.append(dict({"quantity": _q}, **_vals))
    _bim_tab_printed.append(dict({"quantity": _q}, **_prt))

# the managerial reading of the two solutions: where each aggregate sits in the
# standard level L7 = [6,7) of tab:bim-levels, and the probability it places
# below the standard level (x < 6) and below the defined level L6 = [5,6) (x < 5)
_BIM_L6_LO, _BIM_L7_LO, _BIM_L7_HI = 5.0, 6.0, 7.0      # band edges of tab:bim-levels (stated)
_bim_band_rows = []
for _col, _d in (("consensus_confidence_d0", 0.0), ("fairness_weighted_d0.5", 0.5)):
    _S = _bim_sol[_d]
    _pb7 = float(_ndtr((_BIM_L7_LO - _S["mu"]) / _S["sd"]))
    _pb6 = float(_ndtr((_BIM_L6_LO - _S["mu"]) / _S["sd"]))
    _bim_band_rows.append({
        "solution": _col, "delta": _d, "agg_mu": _S["mu"], "agg_sd": _S["sd"],
        "mean_minus_standard_floor": _S["mu"] - _BIM_L7_LO,
        "P_below_standard_x_lt_6": _pb7, "P_below_defined_x_lt_5": _pb6,
        "P_below_standard_printed": round(_pb7, 2), "P_below_defined_printed": round(_pb6, 2)})
write("tab6-9_bim_results.xlsx", {
    "config": cfg(dict({
        "table": "6.9", "label": "tab:bim",
        "generator": "build_workbooks.py, Case I fairness-panel block (solved)",
        "source": "the delta = 0 and delta = 0.5 rows of the solved sweep in fig6-4_bim_tradeoff.xlsx",
        "rounding": "table_6_9 carries full precision; table_6_9_printed rounds as the book prints (aggregate to 2 decimals, the rest to 3)",
        "band_edges": "defined level L6 = [%g,%g), standard level L7 = [%g,%g) (tab:bim-levels)" % (_BIM_L6_LO, _BIM_L7_LO, _BIM_L7_LO, _BIM_L7_HI),
        "band_probabilities": "sheet band_probabilities: P(x < 6) and P(x < 5) under N(agg_mu, agg_sd^2) of each solution, exact normal cdf, printed to 2 decimals (managerial paragraph of sec:bim-results)",
    }, **_bim_solver_cfg)),
    "table_6_9": _bim_tab_rows,
    "table_6_9_printed": _bim_tab_printed,
    "band_probabilities": _bim_band_rows,
})


# =====================================================================
# FIGURE 3.1  Lorenz curve with Gini area and Hoover index (fig:lorenz)
#   SOLVED, not drawn from a formula: the profile is the normalized fairness
#   utilities (power normalization eq:parabola-norm, kappa = 2) of the seven
#   experts of the Case I fairness panel at its consensus-confidence solution
#   delta = 0, recomputed here by _bim_Phi(w, return_util=True) at the solved
#   weights _bim_sol[0.0]["w"] of the Case I block above.  From that profile
#   the block computes
#   * the Lorenz ordinates L_k (cumulative utility share of the k least
#     satisfied experts, k = 0..N) at the abscissae k/N;
#   * the Gini coefficient twice over, by the pairwise formula of Table
#     tab:swf (_gini_vec) and as twice the area between the piecewise-linear
#     Lorenz curve and the diagonal, which must agree;
#   * the Hoover index twice over, as half the relative mean deviation and
#     as the largest vertical gap k/N - L_k, which must agree.
#   No random draws; the profile is a deterministic function of the Case I
#   solution (whose own seed is recorded in fig6-4_bim_tradeoff.xlsx).
# =====================================================================
_lz_w = np.asarray(_bim_sol[0.0]["w"], float)
_lz_Phi, _lz_U = _bim_Phi(_lz_w, return_util=True)
_lz_U = np.asarray(_lz_U, float)
_lz_A = _bim_overlaps(_lz_w)
_lz_order = np.argsort(_lz_U, kind="stable")
_lz_x = np.arange(_BIM_N + 1) / float(_BIM_N)
_lz_L = np.concatenate([[0.0], np.cumsum(_lz_U[_lz_order]) / _lz_U.sum()])
_lz_gini_pair = _gini_vec(_lz_U)
_lz_gini_area = 1.0 - float(np.sum(np.diff(_lz_x) * (_lz_L[1:] + _lz_L[:-1])))
_lz_hoover_rmd = float(np.abs(_lz_U - _lz_U.mean()).sum() / (2.0 * _BIM_N * _lz_U.mean()))
_lz_gapv = _lz_x - _lz_L
_lz_khoover = int(np.argmax(_lz_gapv))
_lz_hoover_gap = float(_lz_gapv[_lz_khoover])
assert abs(_lz_gini_pair - _lz_gini_area) < 1e-12, "Gini: pairwise formula != twice the Lorenz area"
assert abs(_lz_hoover_rmd - _lz_hoover_gap) < 1e-12, "Hoover: half the relative mean deviation != largest gap"
assert abs(_lz_Phi - _bim_sol[0.0]["Phi"]) < 1e-12
print("\n-- Figure 3.1: Lorenz curve of the Case I delta=0 utilities: Gini %.4f, Hoover %.4f at k/N = %d/%d"
      % (_lz_gini_pair, _lz_hoover_gap, _lz_khoover, _BIM_N))
write("fig3-1_lorenz_gini_hoover.xlsx", {
    "config": cfg({
        "figure": "3.1", "label": "fig:lorenz",
        "generator": "build_workbooks.py, FIGURE 3.1 block (after the Case I fairness-panel block)",
        "seed": "none -- deterministic function of the Case I delta = 0 solution, no random draws of its own "
                "(the Case I solve itself uses seed %d, see fig6-4_bim_tradeoff.xlsx)" % NEW_SEED,
        "synthetic": 1,
        "profile": ("normalized fairness utilities (power normalization eq:parabola-norm, kappa = %d) of the "
                    "seven experts of the Case I fairness panel at the solved delta = 0 weights" % _BIM_KAPPA),
        "rho": _BIM_RHO, "vartheta": _BIM_THETA, "kappa": _BIM_KAPPA,
        "N": _BIM_N,
        "gini_rule": "G = sum_ij |U_i - U_j| / (2 N^2 mean U) (tab:swf, _gini_vec) = 1 - sum_k (x_k - x_{k-1})(L_k + L_{k-1})",
        "hoover_rule": "H = sum_i |U_i - mean U| / (2 N mean U) (tab:swf) = max_k (k/N - L_k)",
        "x_label": "cumulative share of experts",
        "y_label": "cumulative share of fairness utility",
        "title": "Gini coefficient and Hoover index on the Lorenz curve",
        "worked_examples": _WX_CFG % 3,
    }),
    "profile": pd.DataFrame({
        "expert": np.arange(1, _BIM_N + 1),
        "weight_delta0": _lz_w,
        "overlap_A": _lz_A,
        "utility_U": _lz_U,
        "rank_ascending": np.argsort(_lz_order) + 1,
    }),
    "lorenz": pd.DataFrame({
        "k": np.arange(_BIM_N + 1),
        "expert_added": [0] + [int(i) + 1 for i in _lz_order],
        "x_share_of_experts": _lz_x,
        "L_share_of_utility": _lz_L,
        "gap_to_equality": _lz_gapv,
    }),
    "summary": cfg({
        "mean_utility": float(_lz_U.mean()),
        "min_utility": float(_lz_U.min()),
        "min_utility_expert": int(np.argmin(_lz_U)) + 1,
        "max_utility": float(_lz_U.max()),
        "second_lowest_utility": float(np.sort(_lz_U)[1]),
        "gini_pairwise": _lz_gini_pair,
        "gini_twice_area": _lz_gini_area,
        "hoover_half_rmd": _lz_hoover_rmd,
        "hoover_max_gap": _lz_hoover_gap,
        "hoover_at_k": _lz_khoover,
        "hoover_at_x": float(_lz_x[_lz_khoover]),
        "gini_printed": round(_lz_gini_pair, 3),
        "hoover_printed": round(_lz_hoover_gap, 3),
        "collective_fairness_Phi": _lz_Phi,
    }),
    # not read by the figure script: the hand-worked examples of Chapter 3
    "worked_examples": _wx_sheet(3),
})


# =====================================================================
# FIGURE 4.4  Consensus-fairness Pareto front of the Case I fairness panel
#             (fig:pareto) -- SOLVED here, not drawn from a formula
#   The two objectives are the mean overlap g1 = Abar(w) and the collective
#   fairness g2 = Phi(w) of the seven-expert panel of Section 6.2, evaluated
#   by the functions _bim_Abar and _bim_Phi of the Case I block above (no
#   re-implementation), both maximized over the simplex.  The confidence
#   term of eq:bim-triobjective is set aside, so the front is the
#   bi-objective consensus-fairness trade-off of that panel.
#   * epsilon-constraint: the program eq:eps-constraint with the roles of the
#     objectives exchanged, g2 = Phi retained and g1 = Abar bounded, i.e.
#     maximize Phi(w) subject to Abar(w) >= b.  Grid: _PAR_N_FLOORS equally
#     spaced floors b from Abar at the fairness maximizer to the maximum of
#     Abar (both maximizers from the Case I bounds solve).  Each floor is
#     solved by SLSQP (the Case I ftol and maxiter) from the previous floor's
#     solution and _PAR_K_RANDOM Dirichlet(1,...,1) starts, floors
#     ascending; the two single-objective maximizers enter as candidates in
#     their own right whenever they meet the floor (the consensus maximizer
#     is the only feasible point at the top floor), and the best feasible
#     point (Abar >= b - _PAR_FEAS_TOL) is kept; the spread of the SLSQP runs
#     is recorded per floor.
#   * weighted sum: maximize lam*At + (1-lam)*Pt, At and Pt min-max
#     normalized with the Case I bounds, for lam on _PAR_N_LAMBDA equally
#     spaced values in [0, 1], from the previous weighting's solution and
#     _PAR_K_RANDOM random starts, the maximizers and (backward) the next
#     weighting's solution entering as candidates.
#   * concave majorant: the upper hull of ALL computed efficient points, the
#     epsilon-constraint points and the weighted-sum solutions together.  A
#     weighted-sum optimum is a supported efficient point, and the front
#     has kinks and bends between two grid floors, so a hull of the floor
#     samples alone would pass below supported points (several weighted sums
#     return the same kink).  A hull edge that skips
#     epsilon points more than _PAR_GAP_TOL below it bridges an unsupported
#     stretch; the bridge with the deepest point is the one the figure and
#     the text show, and its ends are the hull vertices on either side.  The
#     text also mentions the other bridges, printing only their count (the
#     key other_dents), as further dents far shallower than the stretch that
#     adjoin its ends and are invisible at the figure's scale; self-check
#     (9b) asserts that they share a hull vertex with the stretch, one at
#     each end, and that each is less than _PAR_DENT_RATIO as deep.
#   * refinement: the deepest point is then resolved on a finer grid.  Each
#     of the two grid intervals next to the deepest grid floor is split into
#     _PAR_REFINE_SPLIT equal parts and the epsilon program is solved at the
#     new floors (SLSQP from the previous refinement floor's solution, the
#     first from the grid solution below, and _PAR_K_RANDOM random starts;
#     the grid solutions of the two floors bounding each refinement floor
#     from above and the two maximizers as candidates).  A backward pass over
#     all floors, grid and refinement merged in ascending order, offers each
#     floor the next floor's solution.  The majorant, the gaps and the
#     stretch are then recomputed on the merged set.
#   Random starts: ONE generator numpy.random.default_rng(202603), drawn in
#   the order (1) grid epsilon sweep, _PAR_K_RANDOM per floor, b ascending,
#   (2) weighted-sum sweep, _PAR_K_RANDOM per weighting, lambda ascending,
#   (3) refinement floors, _PAR_K_RANDOM per floor, b ascending.
# =====================================================================
_PAR_N_FLOORS = 161
_PAR_K_RANDOM = 2
_PAR_N_LAMBDA = 51
_PAR_REFINE_SPLIT = 21
_PAR_FEAS_TOL = 1e-9
_PAR_GAP_TOL = 1e-8      # well above the solver noise (spread across starts and weighted-sum gaps < 1e-9)
_PAR_DENT_RATIO = 0.01   # a dent "far shallower" than the stretch, invisible at the inset's scale

print("\n-- Figure 4.4: consensus-fairness Pareto front of the Case I panel (%d floors, %d weightings) ..."
      % (_PAR_N_FLOORS, _PAR_N_LAMBDA))
_par_rng = np.random.default_rng(NEW_SEED)
_par_wA = np.asarray(_bim_bounds["Abar"]["argmax"], float)    # consensus maximizer
_par_wP = np.asarray(_bim_bounds["Phi"]["argmax"], float)     # fairness maximizer
_par_grid = np.linspace(_bim_Abar(_par_wP), _bim_Amax, _PAR_N_FLOORS)
_par_step = float(_par_grid[1] - _par_grid[0])
_par_nruns = [0]


def _par_polish(w):
    w = np.clip(np.asarray(w, float), 0.0, 1.0)
    return w / w.sum()


def _par_eps_run(b, w0):
    _par_nruns[0] += 1
    r = _sp_minimize(lambda w: -_bim_Phi(w), w0, method="SLSQP", bounds=_BIM_BNDS,
                     constraints=_BIM_CONS + ({"type": "ineq", "fun": lambda w: _bim_Abar(w) - b},),
                     options={"ftol": _BIM_FTOL, "maxiter": _BIM_MAXITER})
    return _par_polish(r.x)


def _par_ws_value(w, lam):
    At = (_bim_Abar(w) - _bim_Amin) / (_bim_Amax - _bim_Amin)
    Pt = (_bim_Phi(w) - _bim_Pmin) / (_bim_Pmax - _bim_Pmin)
    return lam * At + (1.0 - lam) * Pt


def _par_ws_run(lam, w0):
    _par_nruns[0] += 1
    r = _sp_minimize(lambda w: -_par_ws_value(w, lam), w0, method="SLSQP", bounds=_BIM_BNDS,
                     constraints=_BIM_CONS, options={"ftol": _BIM_FTOL, "maxiter": _BIM_MAXITER})
    return _par_polish(r.x)


# one record per epsilon floor: floor b, kind, best (w, Abar, Phi, source), Phi of the feasible SLSQP runs
def _par_new_floor(b, kind):
    return {"b": float(b), "kind": kind, "sol": None, "runs": []}


def _par_offer(e, w, label, run=False):
    a, p = _bim_Abar(w), _bim_Phi(w)
    if a < e["b"] - _PAR_FEAS_TOL:
        return False
    if run:
        e["runs"].append(p)
    if e["sol"] is None or p > e["sol"][2]:
        e["sol"] = (w, a, p, label)
        return True
    return False


def _par_eps_sweep(floors, w_first, candidates):
    """Solve the floors (ascending) in turn; candidates(e) lists extra (label, w) to offer."""
    prev = w_first
    for e in floors:
        starts = [("previous floor", prev)]
        starts += [("random", s) for s in _par_rng.dirichlet(np.ones(_BIM_N), size=_PAR_K_RANDOM)]
        for lab, w0 in starts:
            _par_offer(e, _par_eps_run(e["b"], w0), "SLSQP from " + lab, run=True)
        for lab, w in candidates(e):
            _par_offer(e, w, lab)
        assert e["sol"] is not None, "no feasible point at floor %.6f" % e["b"]
        prev = e["sol"][0]


# ----- (1) grid epsilon-constraint sweep ------------------------------------------
_par_floors_grid = [_par_new_floor(b, "grid") for b in _par_grid]
_par_eps_sweep(_par_floors_grid, _par_wP,
               lambda e: [("consensus maximizer", _par_wA), ("fairness maximizer", _par_wP)])

# ----- (2) weighted-sum sweep --------------------------------------------------------
_par_lams = np.linspace(0.0, 1.0, _PAR_N_LAMBDA)
_par_ws = [None] * _PAR_N_LAMBDA


def _par_ws_offer(j, w, label):
    v = _par_ws_value(w, _par_lams[j])
    if _par_ws[j] is None or v > _par_ws[j][1]:
        _par_ws[j] = (w, v, label)
        return True
    return False


_par_prev = _par_wP
for _j, _lam in enumerate(_par_lams):
    _starts = [("previous weighting", _par_prev)]
    _starts += [("random", _s) for _s in _par_rng.dirichlet(np.ones(_BIM_N), size=_PAR_K_RANDOM)]
    for _lab, _w0 in _starts:
        _par_ws_offer(_j, _par_ws_run(_lam, _w0), "SLSQP from " + _lab)
    _par_ws_offer(_j, _par_wA, "consensus maximizer")
    _par_ws_offer(_j, _par_wP, "fairness maximizer")
    _par_prev = _par_ws[_j][0]
for _j in range(_PAR_N_LAMBDA - 2, -1, -1):
    _par_ws_offer(_j, _par_ws[_j + 1][0], "next weighting's solution")
_par_wsW = np.vstack([s[0] for s in _par_ws])
_par_wsA = np.array([_bim_Abar(w) for w in _par_wsW])
_par_wsP = np.array([_bim_Phi(w) for w in _par_wsW])


# ----- concave majorant of all efficient points --------------------------------------
def _par_upper_hull(x, y):
    """Indices of the vertices of the least concave majorant of (x, y), x ascending."""
    h = []
    for i in range(len(x)):
        while len(h) >= 2:
            i1, i2 = h[-2], h[-1]
            if (x[i2] - x[i1]) * (y[i] - y[i1]) - (y[i2] - y[i1]) * (x[i] - x[i1]) >= 0:
                h.pop()
            else:
                break
        h.append(i)
    return h


def _par_majorant(A_eps, P_eps, A_ws, P_ws):
    """Upper hull of the union of the epsilon points and the weighted-sum solutions.
    Returns the hull vertices as (source, row, Abar, Phi), source 'epsilon' or
    'weighted sum' and row its index in that list, x strictly ascending.  The
    points are taken in the order Abar ascending, Phi descending, epsilon
    before weighted sum, so of two identical points the later one is the vertex
    and of two points with equal Abar the lower one is popped."""
    A = np.r_[A_eps, A_ws]
    P = np.r_[P_eps, P_ws]
    src = ["epsilon"] * len(A_eps) + ["weighted sum"] * len(A_ws)
    row = list(range(len(A_eps))) + list(range(len(A_ws)))
    o = np.lexsort((np.arange(len(A)), -P, A))
    h = [o[i] for i in _par_upper_hull(A[o], P[o])]
    hv = [(src[i], row[i], float(A[i]), float(P[i])) for i in h]
    assert np.all(np.diff([v[2] for v in hv]) > 0), "majorant abscissae not strictly ascending"
    return hv


def _par_analyse(A_eps, P_eps):
    """Majorant over (epsilon points, weighted sums), the gaps, the bridges and the deepest one."""
    hv = _par_majorant(A_eps, P_eps, _par_wsA, _par_wsP)
    hx = np.array([v[2] for v in hv])
    hy = np.array([v[3] for v in hv])
    gap = np.interp(A_eps, hx, hy) - P_eps
    bridges = []
    for i in range(len(hv) - 1):
        inside = np.where((A_eps > hx[i]) & (A_eps < hx[i + 1]))[0]
        if inside.size and gap[inside].max() > _PAR_GAP_TOL:
            d = int(inside[np.argmax(gap[inside])])
            bridges.append((i, float(gap[d]), d, int(inside.size)))
    return hv, hx, hy, gap, bridges


_par_A0 = np.array([e["sol"][1] for e in _par_floors_grid])
_par_P0 = np.array([e["sol"][2] for e in _par_floors_grid])
_, _, _, _par_gap0, _par_bridges0 = _par_analyse(_par_A0, _par_P0)
assert len(_par_bridges0) >= 1, "the computed front is concave: nothing is unsupported"
_par_k0 = int(np.argmax(_par_gap0))                       # deepest grid floor
assert 0 < _par_k0 < _PAR_N_FLOORS - 1 and _par_gap0[_par_k0] > 100 * _PAR_GAP_TOL

# ----- (3) refinement around the deepest grid floor -----------------------------------
_par_refine = []
for _lo in (_par_k0 - 1, _par_k0):
    for _t in range(1, _PAR_REFINE_SPLIT):
        _par_refine.append(_par_grid[_lo] + (_par_grid[_lo + 1] - _par_grid[_lo]) * _t / _PAR_REFINE_SPLIT)
_par_floors_ref = [_par_new_floor(b, "refinement") for b in _par_refine]


def _par_ref_candidates(e):
    # the grid solutions of the next two grid floors above b are feasible for b
    up = [k for k in (_par_k0, _par_k0 + 1) if _par_grid[k] > e["b"]]
    return ([("grid floor solution %d" % k, _par_floors_grid[k]["sol"][0]) for k in up]
            + [("consensus maximizer", _par_wA), ("fairness maximizer", _par_wP)])


_par_eps_sweep(_par_floors_ref, _par_floors_grid[_par_k0 - 1]["sol"][0], _par_ref_candidates)

# ----- merged front, backward pass, final analysis --------------------------------------
_par_floors = sorted(_par_floors_grid + _par_floors_ref, key=lambda e: e["b"])
_PAR_N_FRONT = len(_par_floors)
_par_back = 0
for _k in range(_PAR_N_FRONT - 2, -1, -1):      # backward pass: the next floor's solution
    _par_back += int(_par_offer(_par_floors[_k], _par_floors[_k + 1]["sol"][0], "next floor's solution"))
_par_B = np.array([e["b"] for e in _par_floors])
_par_A = np.array([e["sol"][1] for e in _par_floors])
_par_P = np.array([e["sol"][2] for e in _par_floors])
_par_W = np.vstack([e["sol"][0] for e in _par_floors])
_par_isgrid = np.array([e["kind"] == "grid" for e in _par_floors])
_par_spread = np.array([(max(e["runs"]) - min(e["runs"])) if len(e["runs"]) > 1 else 0.0 for e in _par_floors])
assert np.all(np.diff(_par_B) > 0) and np.all(np.diff(_par_A) >= -1e-12), "front abscissae not ascending"
assert np.all(np.diff(_par_P) <= 0.0), "front fairness rises with the consensus floor"

_par_hv, _par_hx, _par_hy, _par_gap, _par_bridges = _par_analyse(_par_A, _par_P)
_par_major = np.interp(_par_A, _par_hx, _par_hy)
_par_ws_major = np.interp(_par_wsA, _par_hx, _par_hy)
_par_ws_gap = _par_ws_major - _par_wsP
_par_unsup = _par_gap > _PAR_GAP_TOL
assert len(_par_bridges) >= 1
_par_bi, _par_bgap, _par_kdeep, _par_bn = max(_par_bridges, key=lambda t: t[1])
assert _par_kdeep == int(np.argmax(_par_gap))
_par_lo, _par_hi = _par_hv[_par_bi], _par_hv[_par_bi + 1]     # the stretch ends (hull vertices)
_par_main = (_par_A > _par_lo[2]) & (_par_A < _par_hi[2])
assert int(_par_main.sum()) == _par_bn and _par_unsup[_par_main].all(), "a point inside the stretch touches the majorant"
assert _par_grid[_par_k0 - 1] < _par_A[_par_kdeep] < _par_grid[_par_k0 + 1], "the refined deepest point left the refined interval"
_par_other_gap = max([t[1] for t in _par_bridges if t[0] != _par_bi] or [0.0])
# the weighting at which the two ends of the bridging chord tie (normalized slope)
_par_sn = ((_par_hi[3] - _par_lo[3]) / (_bim_Pmax - _bim_Pmin)) / ((_par_hi[2] - _par_lo[2]) / (_bim_Amax - _bim_Amin))
_par_lamstar = -_par_sn / (1.0 - _par_sn)

# weighted sums against the stretch: none inside, and the ones on either side of lambda* are its ends
_par_in_stretch = (_par_wsA > _par_lo[2] + 1e-7) & (_par_wsA < _par_hi[2] - 1e-7)
_par_ws_isleft = _par_wsA <= _par_lo[2] + 1e-7
_par_ws_isright = _par_wsA >= _par_hi[2] - 1e-7
assert np.all(_par_ws_isleft | _par_ws_isright), "a weighted sum lands in the stretch"
_par_jl = int(np.max(np.where(_par_ws_isleft)[0]))       # last weighting left of the stretch
_par_jr = int(np.min(np.where(_par_ws_isright)[0]))      # first weighting right of it
assert _par_jr == _par_jl + 1, "the weighted-sum solutions cross the stretch more than once"
assert _par_lams[_par_jl] < _par_lamstar < _par_lams[_par_jr]
_par_end_dist = max(abs(_par_wsA[_par_jl] - _par_lo[2]), abs(_par_wsP[_par_jl] - _par_lo[3]),
                    abs(_par_wsA[_par_jr] - _par_hi[2]), abs(_par_wsP[_par_jr] - _par_hi[3]))
assert _par_end_dist < 1e-7, "the stretch ends are not the weighted sums on either side of lambda*"
assert _par_ws_gap.max() < 1e-9, "a weighted-sum solution lies below the majorant of the efficient points"
# a weighted-sum optimum is Pareto optimal, so no front point may beat it on its own weighting,
# and it may not beat the epsilon-front at its own consensus
_par_ws_excess = max(float(_par_wsP[_j] - _par_P[np.searchsorted(_par_B, _par_wsA[_j] + _PAR_FEAS_TOL) - 1])
                     for _j in range(_PAR_N_LAMBDA))
_par_front_beats_ws = max(float(max(_par_ws_value(_par_W[_k], _par_lams[_j]) for _k in range(_PAR_N_FRONT))
                                - _par_ws[_j][1]) for _j in range(_PAR_N_LAMBDA))
_par_ws_at_pmax = int(np.sum(np.abs(_par_wsA - _par_A[0]) < 1e-6))
_par_ws_at_amax = int(np.sum(np.abs(_par_wsA - _par_A[-1]) < 1e-6))

print("   front: Abar in [%.4f, %.4f], Phi in [%.4f, %.4f]; unsupported stretch Abar in (%.6f, %.6f),"
      " deepest %.6f below the majorant at Abar = %.6f (grid alone: %.6f at %.6f; other dips <= %.1e);"
      " lambda* = %.4f; %d/%d weighted-sum solutions inside the stretch; max spread across starts %.1e;"
      " %d backward improvements; %d SLSQP runs"
      % (_par_A.min(), _par_A.max(), _par_P.min(), _par_P.max(), _par_lo[2], _par_hi[2],
         _par_bgap, _par_A[_par_kdeep], _par_gap0[_par_k0], _par_A0[_par_k0], _par_other_gap, _par_lamstar,
         int(_par_in_stretch.sum()), _PAR_N_LAMBDA, _par_spread.max(), _par_back, _par_nruns[0]))

_par_front_rows = []
for _k, _e in enumerate(_par_floors):
    _par_front_rows.append(dict({
        "floor_b": float(_par_B[_k]), "floor_kind": _e["kind"],
        "Abar": float(_par_A[_k]), "Phi": float(_par_P[_k]),
        "majorant": float(_par_major[_k]), "gap_below_majorant": float(_par_gap[_k]),
        "on_hull": int(any(v[0] == "epsilon" and v[1] == _k for v in _par_hv)),
        "unsupported": int(_par_unsup[_k]), "in_main_stretch": int(_par_main[_k]),
        "constraint_active": int(_par_A[_k] - _par_B[_k] < 1e-6),
        "slsqp_feasible_runs": len(_e["runs"]), "spread_across_starts": float(_par_spread[_k]),
        "source": _e["sol"][3],
    }, **{"w%d" % (i + 1): float(_par_W[_k, i]) for i in range(_BIM_N)}))
_par_ws_rows = []
for _j in range(_PAR_N_LAMBDA):
    _par_ws_rows.append(dict({
        "lambda_consensus": float(_par_lams[_j]), "Abar": float(_par_wsA[_j]), "Phi": float(_par_wsP[_j]),
        "objective": float(_par_ws[_j][1]),
        "majorant": float(_par_ws_major[_j]), "gap_below_majorant": float(_par_ws_gap[_j]),
        "on_hull": int(any(v[0] == "weighted sum" and v[1] == _j for v in _par_hv)),
        "inside_unsupported_stretch": int(_par_in_stretch[_j]),
        "source": _par_ws[_j][2],
    }, **{"w%d" % (i + 1): float(_par_wsW[_j, i]) for i in range(_BIM_N)}))
_par_summary = {
    "front_Abar_min": float(_par_A.min()), "front_Abar_max": float(_par_A.max()),
    "front_Phi_min": float(_par_P.min()), "front_Phi_max": float(_par_P.max()),
    "fairness_range": float(_par_P.max() - _par_P.min()),
    "unsupported_Abar_lo": _par_lo[2], "unsupported_Abar_hi": _par_hi[2],
    "unsupported_Phi_at_lo": _par_lo[3], "unsupported_Phi_at_hi": _par_hi[3],
    "unsupported_lo_source": _par_lo[0], "unsupported_lo_row": int(_par_lo[1]),
    "unsupported_hi_source": _par_hi[0], "unsupported_hi_row": int(_par_hi[1]),
    "unsupported_lo_hull_vertex": int(_par_bi), "unsupported_hi_hull_vertex": int(_par_bi + 1),
    "main_stretch_points": int(_par_main.sum()),
    "front_points": _PAR_N_FRONT, "grid_floors": _PAR_N_FLOORS, "refinement_floors": len(_par_refine),
    "max_gap": float(_par_bgap), "max_gap_index": int(_par_kdeep), "max_gap_floor_kind": _par_floors[_par_kdeep]["kind"],
    "max_gap_Abar": float(_par_A[_par_kdeep]), "max_gap_Phi": float(_par_P[_par_kdeep]),
    "max_gap_share_of_fairness_range": float(_par_bgap / (_par_P.max() - _par_P.min())),
    "grid_deepest_index_before_refinement": _par_k0,
    "grid_max_gap_before_refinement": float(_par_gap0[_par_k0]),
    "grid_max_gap_Abar_before_refinement": float(_par_A0[_par_k0]),
    "other_dips_max_gap": float(_par_other_gap), "n_bridging_edges": len(_par_bridges),
    "chord_slope_normalized": float(_par_sn), "lambda_star": float(_par_lamstar),
    "ws_solutions": _PAR_N_LAMBDA, "ws_inside_stretch": int(_par_in_stretch.sum()),
    "ws_last_left_lambda": float(_par_lams[_par_jl]), "ws_last_left_Abar": float(_par_wsA[_par_jl]),
    "ws_first_right_lambda": float(_par_lams[_par_jr]), "ws_first_right_Abar": float(_par_wsA[_par_jr]),
    "stretch_ends_vs_ws_max_distance": float(_par_end_dist),
    "ws_max_gap_below_majorant": float(_par_ws_gap.max()),
    "ws_at_fairness_maximizer": _par_ws_at_pmax, "ws_at_consensus_maximizer": _par_ws_at_amax,
    "ws_max_excess_over_front": _par_ws_excess, "front_max_excess_over_ws_objective": _par_front_beats_ws,
    "max_spread_across_starts": float(_par_spread.max()), "backward_improvements": _par_back,
    "floor_step": _par_step, "refinement_step": _par_step / _PAR_REFINE_SPLIT, "slsqp_runs": _par_nruns[0],
}
# the numbers chapter4.tex prints (and Figure 4.4 draws), rounded as printed
_par_printed = {
    "grid_floors": _PAR_N_FLOORS, "refinement_floors": len(_par_refine),
    "Abar_min_3dp": round(_par_A.min(), 3), "Abar_max_3dp": round(_par_A.max(), 3),
    "Phi_max_3dp": round(_par_P.max(), 3), "Phi_min_3dp": round(_par_P.min(), 3),
    "fairness_range_3dp": round(_par_P.max() - _par_P.min(), 3),
    "stretch_lo_3dp": round(_par_lo[2], 3), "stretch_hi_3dp": round(_par_hi[2], 3),
    "max_gap_4dp": round(_par_bgap, 4), "max_gap_Abar_3dp": round(_par_A[_par_kdeep], 3),
    "other_dents": len(_par_bridges) - 1,
    "lambda_star_3dp": round(_par_lamstar, 3),
    "ws_solutions": _PAR_N_LAMBDA,
    "ws_last_left_lambda_2dp": round(float(_par_lams[_par_jl]), 2),
    "ws_first_right_lambda_2dp": round(float(_par_lams[_par_jr]), 2),
    "ws_at_fairness_maximizer": _par_ws_at_pmax,
}
write("fig4-4_pareto_front.xlsx", {
    "config": cfg({
        "figure": "4.4", "label": "fig:pareto",
        "generator": "build_workbooks.py, FIGURE 4.4 block (after the Case I fairness-panel block)",
        "seed": NEW_SEED,
        "rng": "numpy.random.default_rng(202603), one generator for this block",
        "draw_order": ("(1) grid epsilon sweep: %d Dirichlet(1,...,1) starts per floor, floors ascending; "
                       "(2) weighted-sum sweep: %d per weighting, lambda ascending; "
                       "(3) refinement floors: %d per floor, floors ascending"
                       % (_PAR_K_RANDOM, _PAR_K_RANDOM, _PAR_K_RANDOM)),
        "synthetic": 1,
        "panel": ("the seven-expert Case I fairness panel (eq:bim-panel-mu, eq:bim-panel-sigma, "
                  "rho %.1f, vartheta %.1f, kappa %d)" % (_BIM_RHO, _BIM_THETA, _BIM_KAPPA)),
        "objective_g1": "mean overlap Abar(w) = (1/N) sum_i A_i(w) (_bim_Abar of the Case I block), maximized",
        "objective_g2": ("collective fairness Phi(w), Bonferroni (1,1) of the power-normalized Fehr-Schmidt "
                         "utilities (_bim_Phi of the Case I block), maximized"),
        "confidence_term": "set aside: the front is the consensus-fairness trade-off of the panel",
        "method": ("epsilon-constraint (eq:eps-constraint with g_2 = Phi retained and g_1 = Abar bounded): "
                   "max Phi(w) s.t. Abar(w) >= b, w on the simplex"),
        "floors": ("%d equally spaced b (the grid) from Abar at the fairness maximizer to the maximum of Abar "
                   "(both maximizers from the Case I bounds solve), then %d refinement floors splitting each of "
                   "the two grid intervals next to the deepest grid point into %d equal parts"
                   % (_PAR_N_FLOORS, len(_par_refine), _PAR_REFINE_SPLIT)),
        "starts_per_floor": ("SLSQP from the previous floor's solution and %d random starts; the consensus and "
                             "fairness maximizers (and, at a refinement floor, the grid solutions of the grid "
                             "floors above it within the refined interval) enter as candidates; a backward pass "
                             "over the merged floors offers each floor the next floor's solution; the best "
                             "feasible point is kept" % _PAR_K_RANDOM),
        "solver": ("scipy.optimize.minimize(method='SLSQP'), bounds [0,1]^7, sum(w) = 1, ftol %g, maxiter %d; "
                   "clip to [0,1] and renormalize, then re-evaluate" % (_BIM_FTOL, _BIM_MAXITER)),
        "feasibility_tolerance": _PAR_FEAS_TOL,
        "refine_split": _PAR_REFINE_SPLIT,
        "majorant": ("least concave majorant = upper hull of all computed efficient points, the epsilon-constraint "
                     "points and the weighted-sum solutions together (sorted by Abar ascending, Phi descending, "
                     "epsilon before weighted sum; of two identical points the later is the vertex)"),
        "unsupported_rule": ("an epsilon point is unsupported when it lies more than %g below the majorant; the "
                             "unsupported stretch is the hull edge bridging the deepest such point, its ends the two "
                             "hull vertices of that edge" % _PAR_GAP_TOL),
        "other_dents": ("the remaining bridging edges (sheet 'bridges', main = 0), printed only as a count; each "
                        "shares a hull vertex with the stretch, one at each end, and is less than shallow_dent_ratio "
                        "times the stretch's depth (self-check 9b)"),
        "shallow_dent_ratio": _PAR_DENT_RATIO,
        "weighted_sum": ("max lambda*At + (1-lambda)*Pt, At and Pt min-max normalized by the Case I bounds, the "
                         "extremes of each objective over the whole simplex (Abar_min..Phi_max below), not the nadir; "
                         "%d equally spaced lambda in [0,1]; SLSQP from the previous weighting's solution and %d "
                         "random starts, the maximizers and (backward) the next weighting's solution as candidates"
                         % (_PAR_N_LAMBDA, _PAR_K_RANDOM)),
        "Abar_min": _bim_Amin, "Abar_max": _bim_Amax, "Phi_min": _bim_Pmin, "Phi_max": _bim_Pmax,
        "scipy_version": _scipy.__version__, "numpy_version": np.__version__,
        "x_label": "consensus, $g_1 = \\,\\bar{A}(\\mathbf{w})$",
        "y_label": "collective fairness, $g_2 = \\,\\bar{\\Phi}(\\mathbf{w})$",
        "title": "Consensus-fairness Pareto front of the Case I panel",
        "front_legend": "Pareto front ($\\varepsilon$-constraint)",
        "arc_legend": "unsupported stretch",
        "hull_legend": "concave majorant",
        "ws_legend": "weighted-sum solutions",
        "marked_label": "deepest unsupported point",
        "inset_title": "gap below the majorant",
    }),
    "front": pd.DataFrame(_par_front_rows),
    "hull": pd.DataFrame([{"source": v[0], "row": int(v[1]), "Abar": v[2], "Phi": v[3]} for v in _par_hv]),
    "bridges": pd.DataFrame([{"from_hull_vertex": i, "to_hull_vertex": i + 1,
                              "Abar_from": float(_par_hx[i]), "Abar_to": float(_par_hx[i + 1]),
                              "epsilon_points_inside": n, "max_gap": g, "deepest_front_row": d,
                              "main": int(i == _par_bi)} for i, g, d, n in _par_bridges]),
    "weighted_sum": pd.DataFrame(_par_ws_rows),
    "summary": cfg(_par_summary),
    "printed": cfg(_par_printed),
})


# =====================================================================
# FIGURE 4.3  Scalability of the hierarchical paradigm (fig:scalability)
#   Pairwise-interaction counts from the stated formulas -- no timings.
#   Flat program: N(N-1) ordered pairs.  Two-stage program with balanced
#   subgroups at K* = round((N^2/2)^(1/3)) ~ N^(2/3):
#   sum_k n_k(n_k-1) intra-group pairs + K(K-1) inter-group pairs.
#   The workbook also carries the epigraph-size count of sec:scalability-convex
#   (sheet lp_reformulation_size) and the worked examples of Chapter 4 (sheet
#   worked_examples, from the block WORKED EXAMPLES OF THE THEORY CHAPTERS).
# =====================================================================
def _hier_count(N):
    K = int(round((N * N / 2.0) ** (1.0 / 3.0)))
    K = max(1, min(K, N))
    base, rem = divmod(N, K)
    sizes = [base + 1] * rem + [base] * (K - rem)
    return K, sum(n * (n - 1) for n in sizes) + K * (K - 1)


_Ns = np.unique(np.round(np.logspace(math.log10(10), math.log10(10000), 28)).astype(int))
_scal_rows = []
for _N in _Ns:
    _K, _hc = _hier_count(int(_N))
    _scal_rows.append({"N": int(_N), "flat_pairwise": int(_N) * (int(_N) - 1),
                       "K_star": _K, "hier_pairwise": _hc})
# the saving the text prints for the largest panel (flat count over the
# two-stage count at N = 10,000, rounded to an integer as printed)
_scal_last = _scal_rows[-1]
_scal_ratio = _scal_last["flat_pairwise"] / float(_scal_last["hier_pairwise"])
# the illustrative size of the epigraph reformulation (Remark rem:lp-qp and the
# paragraph after it): stated N, T, L; the counts follow from the stated formulas
# (N weights, TL pinball variables, N(N-1)/2 pair variables; 2TL pinball rows with
# about N + 1 nonzeros each, N(N-1) pair rows), and the pair block at a large panel
_LPQ = dict(N=30, T=40, L=19, N_large=1000)
_lpq_TL = _LPQ["T"] * _LPQ["L"]
_lpq_pairs = _LPQ["N"] * (_LPQ["N"] - 1) // 2
_lpq_rows = {
    "N": _LPQ["N"], "T": _LPQ["T"], "L": _LPQ["L"],
    "pinball_variables_TL": _lpq_TL,
    "pair_variables": _lpq_pairs,
    "variables": _LPQ["N"] + _lpq_TL + _lpq_pairs,
    "pinball_rows": 2 * _lpq_TL,
    "pair_rows": _LPQ["N"] * (_LPQ["N"] - 1),
    "nonzeros_per_pinball_row": _LPQ["N"] + 1,
    "pinball_nonzeros": 2 * _lpq_TL * (_LPQ["N"] + 1),
    "N_large": _LPQ["N_large"],
    "pair_variables_at_N_large": _LPQ["N_large"] * (_LPQ["N_large"] - 1) // 2,
}
write("fig4-3_scalability.xlsx", {
    "config": cfg({
        "figure": "4.3", "label": "fig:scalability",
        "seed": "none -- closed-form interaction counts from the stated formulas, no random draws",
        "synthetic": 1,
        "x_label": "panel size $N$",
        "y_label": "pairwise interaction count",
        "flat_legend": "flat single-stage program, $\\Theta(N^2)$",
        "hier_legend": "representative-level hierarchy, $K^{*}\\propto N^{2/3}$, $\\Theta(N^{4/3})$",
        "note": ("interaction counts computed from the stated formulas -- "
                 "flat N(N-1) ordered pairs vs sum_k n_k(n_k-1) + K(K-1) "
                 "(inter-group stage posed over the K representatives) "
                 "with balanced subgroups at K* = round((N^2/2)^(1/3)); "
                 "normalized cost units (counts, not timings)"),
        "worked_examples": _WX_CFG % 4,
    }),
    "series": pd.DataFrame(_scal_rows),
    "summary": cfg({
        "N_max": _scal_last["N"],
        "flat_pairwise_at_N_max": _scal_last["flat_pairwise"],
        "hier_pairwise_at_N_max": _scal_last["hier_pairwise"],
        "K_star_at_N_max": _scal_last["K_star"],
        "ratio_flat_over_hier_at_N_max": _scal_ratio,
        "ratio_printed": int(round(_scal_ratio)),
    }),
    # the illustrative size count of the epigraph reformulation printed in Section 4
    # (stated N, T, L; every count from the formula in its 'rule' column)
    "lp_reformulation_size": pd.DataFrame([
        {"quantity": k, "value": int(v), "printed": "{:,}".format(int(v)),
         "rule": {"N": "stated panel size", "T": "stated number of outcomes", "L": "stated grid levels",
                  "pinball_variables_TL": "T L", "pair_variables": "N (N - 1) / 2",
                  "variables": "N + T L + N (N - 1) / 2", "pinball_rows": "2 T L",
                  "pair_rows": "N (N - 1)", "nonzeros_per_pinball_row": "N + 1 (N weights, one epigraph variable)",
                  "pinball_nonzeros": "2 T L (N + 1)", "N_large": "stated large panel",
                  "pair_variables_at_N_large": "N_large (N_large - 1) / 2"}[k]}
        for k, v in _lpq_rows.items()]),
    # the arithmetic of Example ex:unbalanced-clustering and of the fuzzy C-means
    # memberships (block WORKED EXAMPLES OF THE THEORY CHAPTERS)
    "worked_examples": _wx_sheet(4),
})


# =====================================================================
# FIGURE 5.3  Leave-one-out contributions of the synthetic human-machine
#             panel of Example ex:hmpanel  (fig5-3_loo)
#             FIGURE 5.2 and TABLE 5.3 are derived from the same block.
#
# The panel of tab:hmpanel is GENERATED and SOLVED here, from the stated
# inputs (seed 202603, the outcome model and its constants below, the
# penalties lambda_G and lambda_C) by the formulas the book states:
#   * objective  eq:convex-reformulation, mean CRPS + lambda_G G + lambda_C C,
#     with G the Gini coefficient of the weight vector and C the
#     representativeness penalty (squared quantile deviation from the
#     equal-weight aggregate, continuous limit, see _hm_program);
#   * solver     Algorithm 4.2 (alg:subgradient) exactly as printed --
#     equal-weight start, analytic subgradient with the sign(0) = 0 tie
#     rule for the Gini kinks, diminishing steps eta_r = eta0/(1 + r/tau),
#     best iterate returned, stop when |w(r) - w(r-1)|_inf < tol and
#     r >= r_min; every solve is then VERIFIED against a multistart SLSQP
#     solve of the same objective (scipy) and the build aborts if the two
#     disagree (_HM_VERIFY tolerances);
#   * leave-one-out re-solves (def:loo) with the same solver and settings;
#   * the calibration/resolution split (prop:loo-decomposition) by the
#     Hersbach (2000) decomposition on the L = 19 quantile representation
#     of every aggregate, held identical across the full and the
#     leave-one-out aggregates (_hersbach_X, convention documented there).
# No printed number is typed anywhere in this block: Table 5.3 is written
# from the rounded solved values, and the rule-12 self-checks at the end
# of this file test the written workbooks against invariants (weights on
# the simplex, objective recomputed from the stored data, optimality
# certificate, exact channel identity), never against typed constants.
#
# Outcome model:  y_t = h_t + sum_i u_{i,t} + m_t + sum_k v_{k,t} + g_t + eps_t
# where h is seen by the five humans (each with its own unique signal
# u_i), m by all machines (each calibrated machine with its own unique
# signal v_k), and g only by the two overconfident machines M4/M5, whose
# reported predictive scale is shrunk far below their honest one.  Every
# forecast is Gaussian, so the quantile-averaged aggregate is Gaussian
# with mean sum(w mu) and sd sum(w sigma), and its CRPS has the exact
# closed form of _crps_gauss (no grid, no truncation).
# =====================================================================
from scipy.special import ndtr as _ndtr          # exact standard normal cdf
from scipy.optimize import minimize as _sp_minimize


def _crps_gauss(mu, sig, y):
    """Exact CRPS of N(mu, sig^2) against y (closed form, exact normal cdf)."""
    z = (y - mu) / sig
    return sig * (z * (2.0 * _ndtr(z) - 1.0) + 2.0 * _npdf(z) - 1.0 / math.sqrt(math.pi))


def _hersbach_X(X, y):
    """Hersbach (2000) decomposition of a forecast sequence represented by
    its T x L matrix X of sorted quantiles at the levels k/(L+1), k = 1..L.
    CONVENTION (Hersbach's original ensemble treatment, recorded in the
    workbook config and stated in the book): the L quantiles are read as
    an L-member ensemble; interior bin i (between members i and i+1) has
    nominal level i/L; the two outer regions, below member 1 and above
    member L, are scored as outlier bins at nominal levels 0 and 1.  The
    same routine is applied to every sequence (full aggregate,
    leave-one-out aggregates, individual experts, climatology), which is
    the identical-convention requirement of prop:loo-decomposition.
    Returns (crps_h, REL, POT) with crps_h = REL + POT exactly."""
    X = np.asarray(X, float)
    T, L = X.shape
    rel = 0.0
    pot = 0.0
    for i in range(1, L):
        x1, x2 = X[:, i - 1], X[:, i]
        alpha = np.where(y > x2, x2 - x1, np.where(y > x1, y - x1, 0.0))
        beta = np.where(y < x1, x2 - x1, np.where(y < x2, x2 - y, 0.0))
        ab, bb = alpha.mean(), beta.mean()
        g = ab + bb
        p = i / float(L)
        if g > 0:
            o = bb / g
            rel += g * (o - p) ** 2
            pot += g * o * (1.0 - o)
    below = y < X[:, 0]
    O0 = below.mean()
    if O0 > 0:
        g0 = np.where(below, X[:, 0] - y, 0.0).mean() / O0
        rel += g0 * O0 ** 2
        pot += g0 * O0 * (1.0 - O0)
    above = y > X[:, -1]
    OL = above.mean()
    if OL > 0:
        gL = np.where(above, y - X[:, -1], 0.0).mean() / OL
        rel += gL * OL ** 2
        pot += gL * OL * (1.0 - OL)
    return rel + pot, rel, pot


def _hersbach(mu, sig, y, zlev):
    """_hersbach_X for a Gaussian sequence N(mu_t, sig_t^2) read through its
    L quantiles mu_t + sig_t z_k, z_k the standard normal quantiles of the
    levels k/(L+1)."""
    X = mu[:, None] + sig[:, None] * zlev[None, :]
    return _hersbach_X(X, y)


def _hm_outer_terms(X, y):
    """The two outer-bin terms of _hersbach_X, written through the outlier
    frequencies O0 (outcome below member 1) and OL (outcome above member L)
    and the mean outlier distances d0 = mean (x_1 - y)^+, dL = mean (y - x_L)^+.
    Hersbach's outlier rule sets g0 = d0/O0 and gL = dL/OL, so its outer
    reliability is g0 O0^2 + gL OL^2 = d0 O0 + dL OL and its outer potential
    score d0 (1 - O0) + dL (1 - OL).  Returns (outer REL, outer POT, d0 + dL);
    d0 + dL is what the outer bins contribute to the score under ANY split,
    and is the whole outer reliability if the outer regions are read with the
    interior width-weighted definitions (o-bar = 1 below, 0 above, so
    REL = h-bar and POT = 0), the sensitivity variant of rem:binning-sensitivity."""
    below = y < X[:, 0]
    above = y > X[:, -1]
    O0, OL = below.mean(), above.mean()
    d0 = np.where(below, X[:, 0] - y, 0.0).mean()
    dL = np.where(above, y - X[:, -1], 0.0).mean()
    return (float(d0 * O0 + dL * OL), float(d0 * (1.0 - O0) + dL * (1.0 - OL)),
            float(d0 + dL))


def _hm_rel_floor(L):
    """Population reliability term of a PERFECTLY CALIBRATED N(0, 1) forecast
    under the convention of _hersbach_X (members at the standard normal
    quantiles of k/(L+1), interior bin i at nominal level i/L, the outer
    regions as outlier bins at 0 and 1), in closed form.  With the outcome
    y ~ N(0, 1) and constant members x_1 < ... < x_L, interior bin i has
    width g_i = x_{i+1} - x_i and width-weighted frequency
    o_i = [I(x_{i+1}) - I(x_i)] / g_i, where I(s) = s Phi(s) + phi(s) is the
    integral of Phi up to s; the lower outlier bin has O0 = Phi(x_1) and
    mean outlier distance g0 = x_1 + phi(x_1)/Phi(x_1), the upper one
    OL = 1 - Phi(x_L) and gL = phi(x_L)/OL - x_L.  The term is not zero
    because the step-function representation of a continuous forecast is
    not itself reliable at finite L.  It scales with the forecast's standard
    deviation, so sigma x REL is the floor of a calibrated N(mu_t, sigma^2)
    forecast.  Returns (REL per unit sd, share of it from the two outer bins)."""
    x = _norm_ppf(np.arange(1, L + 1) / (L + 1.0))
    I = x * _ndtr(x) + _npdf(x)
    g = np.diff(x)
    o = np.diff(I) / g
    p = np.arange(1, L) / float(L)
    rel_in = float((g * (o - p) ** 2).sum())
    O0 = float(_ndtr(x[0]))
    g0 = float(x[0] + _npdf(x[0]) / O0)
    OL = float(1.0 - _ndtr(x[-1]))
    gL = float(_npdf(x[-1]) / OL - x[-1])
    rel_out = g0 * O0 ** 2 + gL * OL ** 2
    return rel_in + rel_out, rel_out / (rel_in + rel_out)


def _hm_sig2(v):
    """v at two significant figures, as the Section 5.5 prose prints it."""
    d = max(0, 1 - int(math.floor(math.log10(abs(v)))))
    return "%.*f" % (d, v)


# generator constants (stated inputs; the seed drives all draws).  The
# outcome decomposes into a human-shared signal h, one unique signal u_i
# per human, a machine-shared signal m, one unique signal v_k per
# calibrated machine, a signal g seen only by the two overconfident
# machines, and an irreducible eps.  _HM_SCALE multiplies every scale
# below; since the CRPS is 1-homogeneous it sets the level of every score
# and was chosen so that the human-only aggregate's mean CRPS prints as
# 0.45 (the relative gain and the weight shares do not depend on it).
_HM_SCALE = 0.5567
_HM_SIG = {"h": 1.00, "u": 0.30, "m": 0.85, "v": 0.35, "g": 0.65,
           "eps": 0.60, "tau_H": 0.30, "tau_M": 0.22}
_HM_MULT = np.array([1.00, 1.03, 0.98, 1.05, 1.01, 0.99, 1.03, 1.00, 0.36, 0.32])
_HM_T = 250
# Algorithm 4.2 settings (recorded in the workbook config)
_HM_ALG = {"eta0": 0.5, "tau": 100.0, "tol": 1e-6, "r_min": 1000, "r_max": 200000}
# SLSQP verification of every Algorithm 4.2 solve: equal-weight start plus
# n_random Dirichlet(1) starts drawn from default_rng(seed); the build
# aborts if the best SLSQP value beats the returned iterate by more than
# obj_gap or the two weight vectors differ by more than w_gap (inf norm)
_HM_VERIFY = {"n_random": 30, "seed": NEW_SEED, "ftol": 1e-14, "maxiter": 2000,
              "obj_gap": 1e-8, "w_gap": 1e-3}


def _hm_generate():
    """Draw order under default_rng(202603): h (T), u (T x 5), m (T),
    v (T x 3), g (T), eps (T), then the forecast noises tau_H (T) for
    H1..H5, tau_M (T) for M1..M3, tau_M (T) for M4, M5, in that order."""
    rng = np.random.default_rng(NEW_SEED)
    T = _HM_T
    c = _HM_SCALE
    s = {k: v * c for k, v in _HM_SIG.items()}
    h = rng.normal(0.0, s["h"], T)
    u = rng.normal(0.0, s["u"], (T, 5))
    m = rng.normal(0.0, s["m"], T)
    v = rng.normal(0.0, s["v"], (T, 3))
    g = rng.normal(0.0, s["g"], T)
    eps = rng.normal(0.0, s["eps"], T)
    y = h + u.sum(axis=1) + m + v.sum(axis=1) + g + eps
    ids = ["H1", "H2", "H3", "H4", "H5", "M1", "M2", "M3", "M4", "M5"]
    groups = ["human"] * 5 + ["machine calibrated"] * 3 + ["machine overconfident"] * 2
    vary = (s["h"] ** 2 + 5 * s["u"] ** 2 + s["m"] ** 2 + 3 * s["v"] ** 2
            + s["g"] ** 2 + s["eps"] ** 2)
    # honest predictive scales for each information set
    sd_H = math.sqrt(vary - s["h"] ** 2 - s["u"] ** 2 + s["tau_H"] ** 2)
    sd_M = math.sqrt(vary - s["m"] ** 2 - s["v"] ** 2 + s["tau_M"] ** 2)
    sd_O = math.sqrt(vary - s["m"] ** 2 - s["g"] ** 2 + s["tau_M"] ** 2)
    mult = _HM_MULT
    MU = np.empty((T, 10))
    SD = np.empty(10)
    for i in range(5):
        MU[:, i] = h + u[:, i] + rng.normal(0.0, s["tau_H"], T)
        SD[i] = mult[i] * sd_H
    for i in range(5, 8):
        MU[:, i] = m + v[:, i - 5] + rng.normal(0.0, s["tau_M"], T)
        SD[i] = mult[i] * sd_M
    for i in range(8, 10):
        MU[:, i] = m + g + rng.normal(0.0, s["tau_M"], T)
        SD[i] = mult[i] * sd_O
    params = {"T": T, "scale": c,
              "outcome_model": "y_t = h_t + sum_i u_it + m_t + sum_k v_kt + g_t + eps_t",
              "sig_h": s["h"], "sig_u": s["u"],
              "sig_m": s["m"], "sig_v": s["v"], "sig_g": s["g"],
              "sig_eps": s["eps"], "tau_H": s["tau_H"], "tau_M": s["tau_M"],
              "sd_honest_human": sd_H, "sd_honest_machine": sd_M,
              "sd_honest_overconf": sd_O,
              "sd_multipliers": ",".join("%.2f" % v for v in mult),
              "forecast_family": "Gaussian N(mu_it, sd_i^2), sd_i = multiplier_i x honest sd",
              "draw_order": ("default_rng(202603): h, u, m, v, g, eps, "
                             "then noise tau_H for H1..H5, tau_M for M1..M3, tau_M for M4..M5")}
    return y, MU, SD, ids, groups, mult, params


def _hm_program(y, MU, SD, idx, lam_G, lam_C):
    """The convex program eq:convex-reformulation restricted to the experts
    in idx, as three closures: terms(w) -> (objective, mean CRPS, G, C),
    obj(w), and subgrad(w) (analytic, sign(0) = 0 at Gini ties).
      mean CRPS : exact Gaussian closed form of the QA aggregate
                  N(sum w mu, (sum w sd)^2), averaged over t
      G         : Gini of w, (1/(2 n^2 wbar)) sum_ij |w_i - w_j|, wbar = 1/n
      C         : (1/T) sum_t E_p [q_t,p(w) - q_t,p(w0)]^2 with w0 = 1/n
                  the equal-weight vector; for Gaussians q_t,p = mu_t + sd z_p,
                  so C = mean_t (dmu_t)^2 + dsd^2 E z^2 with E z^2 = 1 in
                  the continuous (untruncated) limit used here
    Subgradient: d CRPS_t/d mu_a = 1 - 2 Phi(z_t), d CRPS_t/d sd_a =
    2 phi(z_t) - 1/sqrt(pi), z_t = (y_t - mu_a)/sd_a, chained through
    mu_a = mu w, sd_a = sd . w; d G/d w_i = (1/(n^2 wbar)) sum_j sign(w_i - w_j);
    d C/d w_i = 2 mean_t dmu_t mu_it + 2 dsd sd_i."""
    mu = MU[:, idx]
    sd = SD[idx]
    n = len(idx)
    T = len(y)
    w0 = np.full(n, 1.0 / n)
    mu0, sd0 = mu @ w0, float(sd @ w0)
    sqrt_pi = math.sqrt(math.pi)

    def terms(w):
        Ma = mu @ w
        Sa = float(sd @ w)
        crps = float(_crps_gauss(Ma, Sa, y).mean())
        gini = _gini_vec(w)
        Cw = float(((Ma - mu0) ** 2).mean() + (Sa - sd0) ** 2)
        return crps + lam_G * gini + lam_C * Cw, crps, gini, Cw

    def obj(w):
        return terms(w)[0]

    def subgrad(w):
        Ma = mu @ w
        Sa = float(sd @ w)
        z = (y - Ma) / Sa
        d_mu = 1.0 - 2.0 * _ndtr(z)
        d_sd = 2.0 * _npdf(z) - 1.0 / sqrt_pi
        g_fid = (d_mu @ mu) / T + d_sd.mean() * sd
        sgn = np.sign(w[:, None] - w[None, :])          # sign(0) = 0: the tie rule
        g_G = sgn.sum(axis=1) / (n * n * (1.0 / n))
        g_C = 2.0 * ((Ma - mu0) @ mu) / T + 2.0 * (Sa - sd0) * sd
        return g_fid + lam_G * g_G + lam_C * g_C

    return {"terms": terms, "obj": obj, "subgrad": subgrad, "n": n}


def _hm_alg42(prog, eta0, tau, tol, r_min, r_max):
    """Algorithm 4.2 (alg:subgradient) as printed: projected subgradient
    descent from the equal-weight vector with eta_r = eta0/(1 + r/tau),
    incumbent-best bookkeeping, stop when |w(r) - w(r-1)|_inf < tol and
    r >= r_min.  r_max is a safety cap only; hitting it aborts the build."""
    n = prog["n"]
    w = np.full(n, 1.0 / n)
    L_min, w_best = prog["obj"](w), w.copy()
    r = 0
    while True:
        g = prog["subgrad"](w)
        eta = eta0 / (1.0 + r / tau)
        w_new = _project_simplex(w - eta * g)[0]
        L_new = prog["obj"](w_new)
        if L_new < L_min:
            L_min, w_best = L_new, w_new.copy()
        dw = float(np.abs(w_new - w).max())
        w = w_new
        r += 1
        if dw < tol and r >= r_min:
            break
        if r >= r_max:
            raise RuntimeError("Algorithm 4.2 hit the iteration cap %d (last dw %.2e)"
                               % (r_max, dw))
    return w_best, L_min, r, dw


def _hm_slsqp(prog, n_random, seed, ftol, maxiter):
    """Multistart SLSQP on the same objective (verification only)."""
    n = prog["n"]
    rng = np.random.default_rng(seed)
    starts = [np.full(n, 1.0 / n)] + [rng.dirichlet(np.ones(n)) for _ in range(n_random)]
    best = None
    for s in starts:
        res = _sp_minimize(prog["obj"], s, method="SLSQP", bounds=[(0.0, 1.0)] * n,
                           constraints=[{"type": "eq", "fun": lambda w: w.sum() - 1.0}],
                           options={"ftol": ftol, "maxiter": maxiter})
        if best is None or res.fun < best.fun:
            best = res
    w = np.clip(best.x, 0.0, None)
    return w / w.sum(), float(best.fun)


_hm_solver_log = []


def _hm_solve(y, MU, SD, idx, lam_G, lam_C, tag, log=None):
    """Minimize mean CRPS + lam_G G(w) + lam_C C(w) over the simplex of the
    experts in idx by Algorithm 4.2, verify against multistart SLSQP, log
    both (in _hm_solver_log, or in the list log if given), and return
    (w, mean CRPS, aggregate mu_t, aggregate sd)."""
    prog = _hm_program(y, MU, SD, idx, lam_G, lam_C)
    w, L_alg, r, dw = _hm_alg42(prog, _HM_ALG["eta0"], _HM_ALG["tau"], _HM_ALG["tol"],
                                _HM_ALG["r_min"], _HM_ALG["r_max"])
    w_chk, L_chk = _hm_slsqp(prog, _HM_VERIFY["n_random"], _HM_VERIFY["seed"],
                             _HM_VERIFY["ftol"], _HM_VERIFY["maxiter"])
    gap = L_alg - L_chk
    w_gap = float(np.abs(w - w_chk).max())
    assert gap <= _HM_VERIFY["obj_gap"], \
        "%s: Algorithm 4.2 value %.10f exceeds the SLSQP value %.10f by %.2e" % (tag, L_alg, L_chk, gap)
    assert w_gap <= _HM_VERIFY["w_gap"], \
        "%s: Algorithm 4.2 weights differ from SLSQP by %.2e (inf norm)" % (tag, w_gap)
    L_w, crps, gini, Cw = prog["terms"](w)
    (_hm_solver_log if log is None else log).append({
        "solve": tag, "n_experts": len(idx), "experts": ",".join(str(i) for i in idx),
        "alg42_iterations": r, "alg42_last_dw_inf": dw,
        "objective_alg42": L_alg, "objective_slsqp": L_chk, "objective_gap": gap,
        "weights_inf_gap_vs_slsqp": w_gap,
        "mean_crps": crps, "gini_G": gini, "representativeness_C": Cw,
        # full precision (repr round trip), so the self-checks can rebuild every
        # leave-one-out aggregate exactly from this sheet
        "weights": ",".join("%.17g" % v for v in w)})
    mu_a, sd_a = MU[:, idx] @ w, float(SD[idx] @ w)
    return w, crps, mu_a, sd_a


print("\n-- generating and solving the ex:hmpanel study "
      "(fig5-3_loo.xlsx, fig5-2_calibration_sharpness.xlsx, tab5-3_hm_panel.xlsx) ...")
_y, _MU, _SD, _ids, _hm_groups, _mult, _hm_params = _hm_generate()
_LAM_G, _LAM_C = 0.010 * _HM_SCALE, 0.080 / _HM_SCALE
_all = list(range(10))
_hum = list(range(5))
_w_full, _crps_full, _mua_full, _sda_full = _hm_solve(_y, _MU, _SD, _all, _LAM_G, _LAM_C, "full panel")
_w_hum, _crps_hum, _mua_hum, _sda_hum = _hm_solve(_y, _MU, _SD, _hum, _LAM_G, _LAM_C, "human-only")
_WH = float(_w_full[:5].sum())
_WM = float(_w_full[5:].sum())
_gain_pct = 100.0 * (_crps_full / _crps_hum - 1.0)

# decomposition grid: L = 19 levels k/20, the same for every sequence
_L_LEV = 19
_hm_plev = np.arange(1, _L_LEV + 1) / (_L_LEV + 1.0)
_zlev = _norm_ppf(_hm_plev)
_ch_full, _rel_full, _pot_full = _hersbach(_mua_full, np.full(len(_y), _sda_full), _y, _zlev)
_ch_hum, _rel_hum, _pot_hum = _hersbach(_mua_hum, np.full(len(_y), _sda_hum), _y, _zlev)
# uncertainty on the SAME grid and convention: the sample climatology
# represented by the empirical outcome quantiles at the 19 levels (numpy
# linear interpolation), scored by the same routine; UNC is its potential
# CRPS, so the climatology has resolution exactly zero
_X_clim = np.tile(np.quantile(_y, _hm_plev), (len(_y), 1))
_ch_clim, _rel_clim, _pot_clim = _hersbach_X(_X_clim, _y)
_UNC_grid = float(_pot_clim)
_U_cont = float(0.5 * np.abs(_y[:, None] - _y[None, :]).mean())   # continuous reference only

# leave-one-out re-solves (def:loo, re-optimized construction), split by
# prop:loo-decomposition on the L = 19 grid; the binning-sensitivity grid
# of rem:binning-sensitivity re-evaluates the SAME solved aggregates at
# other L (the solve does not depend on L)
_L_SENS = [9, 19, 39, 99]
_zsens = {L: _norm_ppf(np.arange(1, L + 1) / (L + 1.0)) for L in _L_SENS}
_full_sens = {L: _hersbach(_mua_full, np.full(len(_y), _sda_full), _y, _zsens[L]) for L in _L_SENS}
# the outer-bin terms of the same aggregates, for the outlier-rule
# sensitivity of rem:binning-sensitivity (columns *_ww_outer below: the
# outer regions read with the interior width-weighted definitions, which
# moves the whole outlier distance into REL and leaves every total unchanged)
_full_outer = {L: _hm_outer_terms(_mua_full[:, None] + _sda_full * _zsens[L][None, :], _y)
               for L in _L_SENS}
_loo_rows = []
_sens_rows = []
for _i in range(10):
    _sub = [j for j in _all if j != _i]
    _w_i, _crps_i, _mua_i, _sda_i = _hm_solve(_y, _MU, _SD, _sub, _LAM_G, _LAM_C, "leave-out " + _ids[_i])
    _ch_i, _rel_i, _pot_i = _hersbach(_mua_i, np.full(len(_y), _sda_i), _y, _zlev)
    _d_cal = _rel_i - _rel_full                 # eq:loo-split-cal
    _d_res = _pot_i - _pot_full                 # eq:loo-split-res (UNC cancels)
    _loo_rows.append({
        "id": _ids[_i], "group": _hm_groups[_i],
        "weight_full": float(_w_full[_i]),
        "delta_cal": float(_d_cal), "delta_res": float(_d_res),
        "delta_total": float(_ch_i - _ch_full),       # step-score total = cal + res exactly
        "delta_cont_defloo": float(_crps_i - _crps_full),   # def:loo continuous mean-CRPS difference
        # convention-free limit of the resolution channel: both aggregates are
        # Gaussian of constant spread, whose potential score tends (L -> inf) to
        # half the mean absolute difference of the errors, so the limit depends
        # on the aggregate means alone and not on any reported spread
        "delta_res_limit": float(
            0.5 * np.abs((_y - _mua_i)[:, None] - (_y - _mua_i)[None, :]).mean()
            - 0.5 * np.abs((_y - _mua_full)[:, None] - (_y - _mua_full)[None, :]).mean()),
        "crps_loo_hersbach": float(_ch_i),
        "crps_loo_continuous": float(_crps_i),
    })
    for _L in _L_SENS:
        _cL, _rL, _pL = _hersbach(_mua_i, np.full(len(_y), _sda_i), _y, _zsens[_L])
        _cF, _rF, _pF = _full_sens[_L]
        _oI = _hm_outer_terms(_mua_i[:, None] + _sda_i * _zsens[_L][None, :], _y)
        _oF = _full_outer[_L]
        _sens_rows.append({"id": _ids[_i], "L": _L,
                           "delta_cal": float(_rL - _rF), "delta_res": float(_pL - _pF),
                           "delta_total": float(_cL - _cF),
                           "delta_cal_ww_outer": float((_rL - _oI[0] + _oI[2]) - (_rF - _oF[0] + _oF[2])),
                           "delta_res_ww_outer": float((_pL - _oI[1]) - (_pF - _oF[1]))})

# group leave-outs (sheet group_loo): the same program re-solved, with the
# same solver, settings and SLSQP verification, on the panel without a GROUP
# of machines, and on the humans joined by one calibrated machine alone.
# Section 5.5 reads from them where the machines' credit goes: removing one
# calibrated machine leaves the other machines, which all carry the signal
# m, to supply what it offered, and the two groups of machines substitute
# for one another, so the value the machines add appears in full only at
# the level of the subpanel.  delta_vs_full is the def:loo orientation
# (score of the reduced panel minus that of the full panel, positive = the
# group helps); gain_vs_human_only is the human-only score minus the score
# of the reduced panel.  The all-machines row is the human-only solve.
# Printed values: two significant figures (_hm_sig2, defined above).
_HM_GROUP_SETS = [("calibrated machines M1-M3", [5, 6, 7]),
                  ("overconfident machines M4, M5", [8, 9]),
                  ("all machines M1-M5", [5, 6, 7, 8, 9]),
                  ("all machines but M1 (humans + M1)", [6, 7, 8, 9]),
                  ("all machines but M2 (humans + M2)", [5, 7, 8, 9]),
                  ("all machines but M3 (humans + M3)", [5, 6, 8, 9])]
_hm_group_log = []
_group_rows = []
for _gname, _grem in _HM_GROUP_SETS:
    _gsub = [j for j in _all if j not in _grem]
    if _gsub == _hum:
        # the human-only solve above is this re-solve; reuse it and its log entry
        _gc, _gmu, _gsd = _crps_hum, _mua_hum, _sda_hum
        _glog = dict(_hm_solver_log[1])
    else:
        _gw, _gc, _gmu, _gsd = _hm_solve(_y, _MU, _SD, _gsub, _LAM_G, _LAM_C,
                                         "group leave-out " + _gname, log=_hm_group_log)
        _glog = dict(_hm_group_log[-1])
    _gch, _grel, _gpot = _hersbach(_gmu, np.full(len(_y), _gsd), _y, _zlev)
    _grow = {"removed_set": _gname,
             "removed": ",".join(_ids[j] for j in _grem),
             "mean_crps": float(_gc),
             "delta_vs_full": float(_gc - _crps_full),
             "gain_vs_human_only": float(_crps_hum - _gc),
             "delta_step_L%d" % _L_LEV: float(_gch - _ch_full),
             "delta_cal_L%d" % _L_LEV: float(_grel - _rel_full),
             "delta_res_L%d" % _L_LEV: float(_gpot - _pot_full),
             "printed_delta_vs_full": _hm_sig2(_gc - _crps_full),
             "printed_gain_vs_human_only": (_hm_sig2(_crps_hum - _gc) if _gsub != _hum else "")}
    for _k in ("solve", "experts", "alg42_iterations", "alg42_last_dw_inf", "objective_alg42",
               "objective_slsqp", "objective_gap", "weights_inf_gap_vs_slsqp", "weights"):
        _grow[_k] = _glog[_k]
    _group_rows.append(_grow)

# per-expert individual reliability / potential CRPS (Figure 5.2 reads these)
_ind_rows = []
for _i in range(10):
    _c_i, _r_i, _p_i = _hersbach(_MU[:, _i], np.full(len(_y), _SD[_i]), _y, _zlev)
    _ind_rows.append({"id": _ids[_i], "group": _hm_groups[_i],
                      "REL_i": float(_r_i), "POT_i": float(_p_i),
                      "RES_i": float(_UNC_grid - _p_i),
                      "crps_i_hersbach": float(_c_i),
                      "crps_i_continuous": float(_crps_gauss(_MU[:, _i], _SD[_i], _y).mean()),
                      "reported_sd": float(_SD[_i]),
                      "scale_multiplier": float(_mult[_i])})

# reliability floor of a perfectly calibrated normal forecast on the grids of
# rem:binning-sensitivity (the qualification of REL in Section 5.5 and the
# reference level of Figure 5.2): exact population values per unit sd
_rel_floor = {L: _hm_rel_floor(L) for L in _L_SENS}
_rel_floor_rows = [{"L": L, "REL_floor_per_unit_sd": _rel_floor[L][0],
                    "share_from_outer_bins": _rel_floor[L][1],
                    "printed": "%.2g" % _rel_floor[L][0]} for L in _L_SENS]

# individual resolution across the binning grid, at an honest spread, and in
# the convention-free limit.  For a Gaussian forecast of constant spread the
# potential score tends, as L grows, to half the mean absolute difference of
# its errors e_t = y_t - mu_t, which does not depend on the reported spread,
# and the uncertainty term tends to 0.5 mean|y_t - y_s|, so
#     RES_i^lim = 0.5 mean|y_t - y_s| - 0.5 mean|e_it - e_is|.
# The honest-spread column re-scores the same means with the spread the
# expert's information set warrants (reported sd / multiplier) on the L = 19
# grid.  These columns carry what Section 5.5 reports about the L = 19
# resolution of the overconfident machines (see the Figure 5.2 block).
_UNC_sens = {}
for _L in _L_SENS:
    _pL_lev = np.arange(1, _L + 1) / (_L + 1.0)
    _UNC_sens[_L] = float(_hersbach_X(np.tile(np.quantile(_y, _pL_lev), (len(_y), 1)), _y)[2])
assert abs(_UNC_sens[_L_LEV] - _UNC_grid) < 1e-15
for _i, _r in enumerate(_ind_rows):
    for _L in _L_SENS:
        _r["RES_i_L%d" % _L] = float(
            _UNC_sens[_L] - _hersbach(_MU[:, _i], np.full(len(_y), _SD[_i]), _y, _zsens[_L])[2])
    _r["honest_sd"] = float(_SD[_i] / _mult[_i])
    _r["RES_i_honest_spread"] = float(
        _UNC_grid - _hersbach(_MU[:, _i], np.full(len(_y), _r["honest_sd"]), _y, _zlev)[2])
    _e_i = _y - _MU[:, _i]
    _r["RES_i_limit"] = float(_U_cont - 0.5 * np.abs(_e_i[:, None] - _e_i[None, :]).mean())
    _r["REL_floor_if_calibrated"] = float(_SD[_i] * _rel_floor[_L_LEV][0])

# error dependence of the panel under the instrument of Section 5.6 (the
# standardized point-forecast errors, median minus outcome, the median of a
# Gaussian forecast being its mean) and the forecast correlations that
# locate the panel's diversity.  Section 5.5 and Example ex:corr-lineage
# report these: each subpanel is internally redundant by this instrument,
# while the two subpanels read different components of the outcome.
_hmE = _MU - _y[:, None]
_hmRerr = np.corrcoef(_hmE, rowvar=False)
_hmRfc = np.corrcoef(_MU, rowvar=False)
_hmI = {"human": [0, 1, 2, 3, 4], "calibrated": [5, 6, 7], "overconfident": [8, 9],
        "machine": [5, 6, 7, 8, 9], "panel": list(range(10))}


def _hm_block(Rm, a, b):
    """Off-diagonal entries of Rm between index sets a and b (each pair once)."""
    if a == b:
        return np.array([Rm[i, j] for i in a for j in a if i < j])
    return np.array([Rm[i, j] for i in a for j in b])


def _hm_neff(Rm, idx, w=None):
    """def:neff at the weight vector w (equal weights if None)."""
    w = np.full(len(idx), 1.0 / len(idx)) if w is None else np.asarray(w, float)
    return float(1.0 / (w @ Rm[np.ix_(idx, idx)] @ w))


_hm_r2 = np.array([np.corrcoef(_MU[:, i], _y)[0, 1] ** 2 for i in range(10)])
_hm_dep = []


def _hm_dep_add(quantity, value, dp=2):
    _hm_dep.append({"quantity": quantity, "value": float(value), "printed": round(float(value), dp)})


for _nm, _a, _b in (("human-human", "human", "human"), ("machine-machine", "machine", "machine"),
                    ("calibrated-calibrated", "calibrated", "calibrated"),
                    ("overconfident-overconfident", "overconfident", "overconfident"),
                    ("calibrated-overconfident", "calibrated", "overconfident"),
                    ("human-machine", "human", "machine")):
    _v = _hm_block(_hmRerr, _hmI[_a], _hmI[_b])
    _hm_dep_add("error r %s min" % _nm, _v.min())
    _hm_dep_add("error r %s max" % _nm, _v.max())
    _v = _hm_block(_hmRfc, _hmI[_a], _hmI[_b])
    _hm_dep_add("forecast r %s min" % _nm, _v.min())
    _hm_dep_add("forecast r %s max" % _nm, _v.max())
for _nm in ("human", "machine", "calibrated", "panel"):
    _hm_dep_add("N_eff equal weights %s" % _nm, _hm_neff(_hmRerr, _hmI[_nm]))
_hm_dep_add("N_eff full-panel optimized weights", _hm_neff(_hmRerr, _hmI["panel"], _w_full))
_hm_dep_add("squared corr(median, outcome) min", _hm_r2.min())
_hm_dep_add("squared corr(median, outcome) max", _hm_r2.max())
_hm_depd = {d["quantity"]: d for d in _hm_dep}

# the further numbers the Section 5.5 prose prints about Figure 5.2, with the
# rounding it prints them at (sheet text_numbers; the individual sheet holds
# the per-expert values they summarize)
_hm_txt = []


def _hm_txt_add(quantity, value, dp):
    _hm_txt.append({"quantity": quantity, "value": float(value), "printed": round(float(value), dp)})


_hm_calib = [i for i in range(10) if _hm_groups[i] != "machine overconfident"]
_hm_txt_add("REL_i min, humans and calibrated machines (L = %d)" % _L_LEV,
            min(_ind_rows[i]["REL_i"] for i in _hm_calib), 4)
_hm_txt_add("REL_i max, humans and calibrated machines (L = %d)" % _L_LEV,
            max(_ind_rows[i]["REL_i"] for i in _hm_calib), 4)
_hm_txt_add("REL floor min, humans and calibrated machines (L = %d)" % _L_LEV,
            min(_ind_rows[i]["REL_floor_if_calibrated"] for i in _hm_calib), 4)
_hm_txt_add("REL floor max, humans and calibrated machines (L = %d)" % _L_LEV,
            max(_ind_rows[i]["REL_floor_if_calibrated"] for i in _hm_calib), 4)
for _i in range(10):
    if _hm_groups[_i] == "machine overconfident":
        _hm_txt_add("RES_i limit %s" % _ids[_i], _ind_rows[_i]["RES_i_limit"], 3)
        _hm_txt_add("RES_i honest spread %s (L = %d)" % (_ids[_i], _L_LEV), _ind_rows[_i]["RES_i_honest_spread"], 3)

# sign stability across the binning grid (rem:binning-sensitivity)
_sens_df = pd.DataFrame(_sens_rows)


def _sign_stable(col):
    out = {}
    for _id in _ids:
        v = _sens_df.loc[_sens_df["id"] == _id, col].to_numpy(float)
        out[_id] = bool(len(set(np.sign(v))) == 1)
    return out


_stab_tot = _sign_stable("delta_total")
_stab_cal = _sign_stable("delta_cal")
_stab_res = _sign_stable("delta_res")
_unstable_total = [k for k in _ids if not _stab_tot[k]]


def _sv(eid, L, col):
    m = (_sens_df["id"] == eid) & (_sens_df["L"] == L)
    return float(_sens_df.loc[m, col].iloc[0])


# which L keep "H3-H5 chiefly resolution" (res > cal), "M1-M3 chiefly
# calibration" (cal > res), and "M4, M5 in the (-, +) cell"
_L_h3h5_res = [L for L in _L_SENS if all(
    _sv(k, L, "delta_res") > _sv(k, L, "delta_cal") for k in ("H3", "H4", "H5"))]
_L_m1m3_cal = [L for L in _L_SENS if all(
    _sv(k, L, "delta_cal") > _sv(k, L, "delta_res") for k in ("M1", "M2", "M3"))]
_L_m4m5_cell = [L for L in _L_SENS if all(
    _sv(k, L, "delta_cal") < 0 and _sv(k, L, "delta_res") > 0 for k in ("M4", "M5"))]
# the same two readings under the width-weighted outer-bin variant
_L_h3h5_res_ww = [L for L in _L_SENS if all(
    _sv(k, L, "delta_res_ww_outer") > _sv(k, L, "delta_cal_ww_outer") for k in ("H3", "H4", "H5"))]
_L_m4m5_cell_ww = [L for L in _L_SENS if all(
    _sv(k, L, "delta_cal_ww_outer") < 0 and _sv(k, L, "delta_res_ww_outer") > 0 for k in ("M4", "M5"))]


def _Ljoin(Ls):
    return ",".join(str(L) for L in Ls) if Ls else "none"

print("   hm panel: human-only %.6f  collaborative %.6f  gain %.2f%%  W_H %.4f  W_M %.4f"
      % (_crps_hum, _crps_full, _gain_pct, _WH, _WM))
for _r in _loo_rows:
    print("   loo %-3s cal %+.6f res %+.6f tot %+.6f defloo %+.6f  w %.4f"
          % (_r["id"], _r["delta_cal"], _r["delta_res"], _r["delta_total"],
             _r["delta_cont_defloo"], _r["weight_full"]))
for _r in _group_rows:
    print("   group leave-out %-34s delta %+.6f  gain over humans %+.6f"
          % (_r["removed_set"], _r["delta_vs_full"], _r["gain_vs_human_only"]))
print("   sign-unstable totals across L=%s: %s" % (_L_SENS, _unstable_total))
print("   UNC on grid %.6f (climatology REL %.2e), continuous U %.6f" % (_UNC_grid, _rel_clim, _U_cont))

# the design invariant of the example (what Figure 5.3 is built to show):
# the two overconfident machines sit in the (-, +) cell of tab:loo-decomposition
for _r in _loo_rows[8:]:
    assert _r["delta_cal"] < 0 and _r["delta_res"] > 0, \
        "overconfident machine %s left the (-,+) cell: %r" % (_r["id"], _r)

write("fig5-3_loo.xlsx", {
    "config": cfg({
        "figure": "5.3", "label": "fig:loo-chart",
        "example": "ex:hmpanel (Section 5.5)",
        "synthetic": 1, "seed": NEW_SEED,
        "T": _hm_params["T"], "n_human": 5, "n_machine": 5,
        "lambda_G": _LAM_G, "lambda_C": _LAM_C,
        "lambda_G_rule": "0.010 x scale", "lambda_C_rule": "0.080 / scale",
        "objective": ("eq:convex-reformulation: mean exact Gaussian CRPS of the QA aggregate "
                      "+ lambda_G Gini(w) + lambda_C C(w); C in the continuous limit (E z^2 = 1), "
                      "no grid and no truncation enter the objective"),
        "solver": "Algorithm 4.2 (alg:subgradient), projected subgradient descent",
        "solver_start": "equal weights 1/N (one start; the program is convex)",
        "solver_step_schedule": "eta_r = %.3g / (1 + r / %.3g)" % (_HM_ALG["eta0"], _HM_ALG["tau"]),
        "solver_subgradient": ("analytic (closed-form CRPS, Gini, C derivatives); "
                               "tie rule sign(0) = 0 at Gini kinks"),
        "solver_returned_iterate": "best iterate visited (incumbent), as printed",
        "solver_tolerance": _HM_ALG["tol"],
        "solver_r_min": _HM_ALG["r_min"],
        "solver_r_max_cap": _HM_ALG["r_max"],
        "solver_projection": "Euclidean projection onto the simplex (sort-based)",
        "normal_cdf": "scipy.special.ndtr (exact)",
        "verification": ("every solve re-run by SLSQP (scipy) from the equal-weight start plus %d "
                         "Dirichlet(1) starts under default_rng(%d), ftol %.0e; build aborts if the "
                         "Algorithm 4.2 value exceeds the SLSQP value by more than %.0e or the weights "
                         "differ by more than %.0e (inf norm); see sheet solver_log"
                         % (_HM_VERIFY["n_random"], _HM_VERIFY["seed"], _HM_VERIFY["ftol"],
                            _HM_VERIFY["obj_gap"], _HM_VERIFY["w_gap"])),
        "loo_construction": "re-optimized (def:loo): the program is re-solved on each reduced panel with the same solver and settings",
        "human_only_row": "the same program solved on the human subpanel alone (Gini and C on the reduced panel)",
        "hersbach_levels_L": _L_LEV,
        "hersbach_levels": "k/(L+1), k = 1..L, standard normal quantiles of the Gaussian aggregate",
        "hersbach_convention": ("Hersbach (2000) L-member ensemble: interior bin i at nominal level i/L; "
                                "the two outer regions scored as outlier bins at nominal levels 0 and 1 by "
                                "Hersbach's outlier rule (hersbach_outlier_rule below; an empty outer bin "
                                "contributes nothing); identical routine for every sequence"),
        "hersbach_outlier_rule": ("lower bin: o_0 = (1/T) #{t: y_t < x_1t}, g_0 = sum_t (x_1t - y_t)^+ / (T o_0), "
                                  "REL_0 = g_0 o_0^2, POT_0 = g_0 o_0 (1 - o_0); upper bin: "
                                  "o_L = (1/T) #{t: y_t <= x_Lt} (outcome below the bin, as for interior bins), "
                                  "g_L = sum_t (y_t - x_Lt)^+ / (T (1 - o_L)), REL_L = g_L (1 - o_L)^2, "
                                  "POT_L = g_L o_L (1 - o_L); g_0 o_0 and g_L (1 - o_L) are the mean outlier "
                                  "distances, so the score identity CRPS = REL + POT is unchanged and an "
                                  "outlying outcome is charged partly to REL and partly to POT"),
        "outer_bin_sensitivity": ("columns *_ww_outer of sheet binning_sensitivity: the outer regions read with "
                                  "the interior width-weighted definitions (h-bar = mean outlier distance, "
                                  "o-bar = 1 below and 0 above), which charges the whole outlier distance to REL; "
                                  "totals unchanged"),
        "L_with_H3H5_chiefly_resolution_ww_outer": _Ljoin(_L_h3h5_res_ww),
        "L_with_M4M5_in_minus_plus_cell_ww_outer": _Ljoin(_L_m4m5_cell_ww),
        "rel_floor_note": ("sheet reliability_floor: exact population REL of a perfectly calibrated N(0,1) "
                           "forecast under this convention, its L members at the standard normal quantiles "
                           "of the levels k/(L+1), k = 1..L (closed form, _hm_rel_floor; the floor depends on "
                           "this level placement); sd x value is the "
                           "floor of a calibrated normal forecast of that sd; individual column "
                           "REL_floor_if_calibrated = reported_sd x floor at L = %d" % _L_LEV),
        "group_loo_note": ("sheet group_loo: the program re-solved (same solver, settings and SLSQP "
                           "verification) without a group of machines, and on the humans joined by one "
                           "calibrated machine alone; delta_vs_full = mean exact CRPS of the reduced panel "
                           "minus that of the full panel (def:loo orientation); gain_vs_human_only = "
                           "human-only mean CRPS minus that of the reduced panel; delta_*_L%d = the step "
                           "score and its channels on the decomposition grid; printed = two significant "
                           "figures; the all-machines row is the human-only solve of sheet solver_log"
                           % _L_LEV),
        "individual_resolution_note": ("individual columns RES_i_L<L> re-score each expert on the binning grid; "
                                       "RES_i_honest_spread re-scores the same means at sd = reported_sd / "
                                       "multiplier on the L = %d grid; RES_i_limit = 0.5 mean|y - y'| - "
                                       "0.5 mean|e - e'|, e = y - mu_i, the L -> inf limit for a Gaussian "
                                       "forecast of constant spread, which depends on the means alone" % _L_LEV),
        "error_dependence_note": ("sheets error_dependence, error_correlation, forecast_correlation: the "
                                  "instrument of Section 5.6 (sample correlation of the point-forecast errors "
                                  "median - outcome, the median of a Gaussian forecast being its mean; the "
                                  "correlation standardizes each series) and the correlations of the forecast "
                                  "medians; N_eff = 1 / (w' R w) of def:neff"),
        "uncertainty_UNC_grid": _UNC_grid,
        "uncertainty_note": ("potential CRPS of the sample climatology represented by the empirical "
                             "outcome quantiles at the same L levels (numpy linear interpolation), "
                             "same routine and convention; climatology REL %.3e, total %.6f; "
                             "continuous 0.5 E|y - y'| = %.6f for reference" % (_rel_clim, _ch_clim, _U_cont)),
        "binning_sensitivity_L": ",".join(str(L) for L in _L_SENS),
        "total_sign_unstable_across_L": ",".join(_unstable_total) if _unstable_total else "none",
        "L_with_H3H5_chiefly_resolution": ",".join(str(L) for L in _L_h3h5_res),
        "L_with_M1M3_chiefly_calibration": ",".join(str(L) for L in _L_m1m3_cal),
        "L_with_M4M5_in_minus_plus_cell": ",".join(str(L) for L in _L_m4m5_cell),
        "x_label": "leave-one-out contribution $\\Delta_i$ to the mean CRPS",
        "cal_legend": "calibration channel $\\Delta_i^{\\mathrm{cal}}$",
        "res_legend": "resolution channel $\\Delta_i^{\\mathrm{res}}$",
        "tot_legend": "total $\\Delta_i$ (step score)",
        "note": ("panel generated and convex program solved here by Algorithm 4.2; "
                 "decomposition = Hersbach on the L-quantile representation, identical "
                 "convention for full and leave-one-out aggregates; the plotted total is "
                 "the step-function score difference, which equals Delta_cal + Delta_res "
                 "exactly; the def:loo continuous difference is the column delta_cont_defloo"),
    }),
    "generator_params": cfg({k: v for k, v in _hm_params.items()}),
    "experts": pd.DataFrame({
        "id": _ids, "group": _hm_groups,
        "signal": (["h+u_i"] * 5 + ["m+v_k"] * 3 + ["m+g"] * 2),
        "scale_multiplier": _mult,
        "reported_sd": _SD,
    }),
    "outcomes": pd.DataFrame({"t": np.arange(1, len(_y) + 1), "y": _y}),
    "forecast_mu": pd.DataFrame(_MU, columns=_ids),
    "loo": pd.DataFrame(_loo_rows),
    "group_loo": pd.DataFrame(_group_rows),
    "individual": pd.DataFrame(_ind_rows),
    "binning_sensitivity": _sens_df,
    "reliability_floor": pd.DataFrame(_rel_floor_rows),
    "error_dependence": pd.DataFrame(_hm_dep),
    "text_numbers": pd.DataFrame(_hm_txt),
    "error_correlation": pd.DataFrame(_hmRerr, columns=_ids).assign(id=_ids)[["id"] + _ids],
    "forecast_correlation": pd.DataFrame(_hmRfc, columns=_ids).assign(id=_ids)[["id"] + _ids],
    "solver_log": pd.DataFrame(_hm_solver_log),
    "results": pd.DataFrame([
        {"quantity": "human-only mean CRPS", "value": _crps_hum,
         "printed": round(_crps_hum, 2)},
        {"quantity": "collaborative mean CRPS", "value": _crps_full,
         "printed": round(_crps_full, 2)},
        {"quantity": "human weight share W_H", "value": _WH, "printed": round(_WH, 2)},
        {"quantity": "machine weight share W_M", "value": _WM, "printed": round(_WM, 2)},
        {"quantity": "CRPS gain (collab vs human-only, pct)",
         "value": _gain_pct, "printed": round(_gain_pct)},
        {"quantity": "full-panel Hersbach REL", "value": _rel_full, "printed": round(_rel_full, 3)},
        {"quantity": "full-panel Hersbach POT", "value": _pot_full, "printed": round(_pot_full, 3)},
        {"quantity": "full-panel Hersbach CRPS", "value": _ch_full, "printed": round(_ch_full, 3)},
        {"quantity": "human-only Hersbach CRPS", "value": _ch_hum, "printed": round(_ch_hum, 3)},
        {"quantity": "uncertainty UNC on grid", "value": _UNC_grid, "printed": round(_UNC_grid, 3)},
        {"quantity": "full-panel objective", "value": _hm_solver_log[0]["objective_alg42"],
         "printed": round(_hm_solver_log[0]["objective_alg42"], 4)},
        {"quantity": "human-only objective", "value": _hm_solver_log[1]["objective_alg42"],
         "printed": round(_hm_solver_log[1]["objective_alg42"], 4)},
    ]),
    "weights": pd.DataFrame({"id": _ids, "weight_full_panel": _w_full,
                             "weight_human_only": list(_w_hum) + [0.0] * 5}),
})

# ---- TABLE 5.3 (tab:hmpanel): the rounded outputs of the solve above -------
write("tab5-3_hm_panel.xlsx", {
    "table_5_3": [
        {"row": "Human-only aggregate", "optimized_weight_share": None,
         "mean_CRPS": round(_crps_hum, 2)},
        {"row": "Collaborative aggregate", "optimized_weight_share": None,
         "mean_CRPS": round(_crps_full, 2)},
        {"row": "  human subpanel (5 experts)", "optimized_weight_share": round(_WH, 2),
         "mean_CRPS": None},
        {"row": "  machine subpanel (5 experts)", "optimized_weight_share": round(_WM, 2),
         "mean_CRPS": None},
    ],
    "provenance": cfg({
        "generator": "build_workbooks.py (ex:hmpanel block; values are round(x, 2) of fig5-3_loo.xlsx sheet results)",
        "seed": NEW_SEED,
        "synthetic": 1,
        "note": ("rounded outputs of the seeded ex:hmpanel solve, asserted in build_workbooks.py "
                 "(rule-12 self-check: this table must equal the rounded results sheet of fig5-3_loo.xlsx)"),
        "gain_pct_printed": round(_gain_pct),
    }),
})

# =====================================================================
# FIGURE 5.2  Calibration-sharpness profile of the mixed panel
#             (fig:calsharp, Example ex:hmpanel)
#
# DERIVED from the same generated panel as fig5-3_loo.xlsx above, not
# hand-set: each expert's reliability error REL_i and potential CRPS
# POT_i are the per-expert Hersbach values of the 'individual' sheet,
# and the resolution follows Hersbach's own definition
#     RES_i = UNC - POT_i,
# where UNC is the uncertainty term formed on the SAME L = 19 grid and
# convention as REL_i and POT_i (the potential CRPS of the sample
# climatology represented by its empirical quantiles at the 19 levels),
# as prop:loo-decomposition requires.  Figures 5.2 and 5.3 are therefore
# two views of ONE dataset.  The workbook remains the figure's own data
# source (one workbook per figure), and this block records the
# cross-reference.
# NOTE on the horizontal position of the two overconfident machines: on
# this L = 19 grid they show the highest RES_i of the panel as well as by
# far the worst reliability error, but that resolution lead is mainly a
# response of the finite grid to their shrunk spread, not superior
# informativeness.  Re-scored with the same means at an honest spread, or
# in the convention-free limit (which depends on the means alone), their
# RES_i falls below that of every human expert; in this sample their means
# also correlate with y less than every human's, although by design their
# population correlation is higher.  The fig5-3_loo.xlsx 'individual'
# columns RES_i_L<L>, RES_i_honest_spread and RES_i_limit carry this, and
# the x-axis label names the grid.  Their positive resolution CHANNEL in
# the pool (delta_res, and delta_res_limit of the 'loo' sheet) is a
# separate property of the aggregate: it survives the limit and reflects
# the signal g that no other expert sees.
# The calibrated experts' REL_i scatter about the reliability floor that a
# perfectly calibrated normal forecast of their spread carries on this grid
# (REL_floor column, from _hm_rel_floor).
# =====================================================================
_GRP51 = {"human": "human experts",
          "machine calibrated": "calibrated machines",
          "machine overconfident": "overconfident machines"}
# the marker convention of Figure 6.12 (the vignette block's _VG_COLOR and
# _VG_MARKER, checked at the tail): humans blue circles, calibrated machines
# teal squares, overconfident machines red triangles
_STYLE51 = {"human": ("blue", "o"),
            "machine calibrated": ("teal", "s"),
            "machine overconfident": ("red", "^")}
_pts51 = [{"id": r["id"], "group": _GRP51[r["group"]],
           "color": _STYLE51[r["group"]][0], "marker": _STYLE51[r["group"]][1],
           "RES": r["RES_i"], "REL": r["REL_i"],
           "REL_floor": r["REL_floor_if_calibrated"], "RES_limit": r["RES_i_limit"]}
          for r in _ind_rows]
_res51 = [p["RES"] for p in _pts51]
_rel51 = [p["REL"] for p in _pts51]
_xspan51 = max(_res51) - min(_res51)
_x0_51 = round(min(_res51) - 0.10 * _xspan51, 3)
_x1_51 = round(max(_res51) + 0.10 * _xspan51, 3)
_y1_51 = round(1.42 * max(_rel51), 3)      # headroom keeps the legend clear
_rel_over51 = [p["REL"] for p in _pts51 if p["group"] == "overconfident machines"]
assert min(_rel_over51) > 3.0 * max(
    p["REL"] for p in _pts51 if p["group"] != "overconfident machines"), \
    "the overconfident machines no longer stand apart in REL_i"
# the callout of the nearly coincident calibrated machines: the group it names, its
# text (from the ids of that group) and its offsets from the group's mean point are
# config entries the drawing script reads; the coincidence it asserts is checked at
# the tail against callout_coincident_max_axis_fraction of each axis range
_call51_group = _GRP51["machine calibrated"]
_call51_ids = [p["id"] for p in _pts51 if p["group"] == _call51_group]
_call51_text = "%s\u2013%s\n(nearly coincident)" % (_call51_ids[0], _call51_ids[-1])
write("fig5-2_calibration_sharpness.xlsx", {
    "config": cfg({
        "figure": "5.2", "label": "fig:calsharp",
        "example": "ex:hmpanel (Section 5.5)",
        "synthetic": 1, "seed": NEW_SEED,
        "derived_from": ("fig5-3_loo.xlsx sheet 'individual' -- same generated "
                         "panel and Hersbach split; RES_i = UNC_grid - POT_i"),
        "uncertainty_UNC_grid": _UNC_grid,
        "uncertainty_note": ("potential CRPS of the sample climatology on the same L = 19 grid "
                             "and Hersbach convention as REL_i and POT_i (see fig5-3_loo.xlsx config); "
                             "continuous 0.5 E|y - y'| = %.6f for reference" % _U_cont),
        "hersbach_levels_L": _L_LEV,
        "hersbach_convention": ("Hersbach (2000) L-member ensemble: interior bin i at nominal level i/L; "
                                "outer regions scored as outlier bins at nominal levels 0 and 1; "
                                "outer bins by Hersbach's outlier rule: O_0 = freq(y < x_1), "
                                "g_0 = mean (x_1 - y)^+ / O_0, REL_0 = g_0 O_0^2, POT_0 = g_0 O_0 (1 - O_0); "
                                "the upper bin mirrored with O_L = freq(y > x_L), g_L = mean (y - x_L)^+ / O_L; "
                                "an empty outer bin contributes nothing (as fig5-3_loo.xlsx config "
                                "hersbach_outlier_rule)"),
        "rel_floor_per_unit_sd": _rel_floor[_L_LEV][0],
        "rel_floor_note": ("points column REL_floor = reported sd x the reliability floor of a perfectly "
                           "calibrated normal forecast on the L = %d grid (fig5-3_loo.xlsx sheet "
                           "reliability_floor); column RES_limit = the convention-free resolution "
                           "(fig5-3_loo.xlsx individual RES_i_limit); neither is plotted" % _L_LEV),
        # mathtext, so the subscripts typeset as they do in the caption and
        # prose; the label names the grid, on which the x positions depend
        "x_label": "resolution $\\mathrm{RES}_i$ on the $L = %d$ quantile grid" % _L_LEV,
        "y_label": "reliability error $\\mathrm{REL}_i$ (lower is better)",
        # a small negative floor keeps the low-REL cluster off the axis line
        "xmin": _x0_51, "xmax": _x1_51,
        "ymin": round(-0.04 * _y1_51, 3), "ymax": _y1_51,
        "title": "Calibration profile of a synthetic human-machine panel",
        "marker_rule": ("points columns color and marker: the convention of Figure 6.12 "
                        "(fig6-12_hm_calsharp.xlsx forecasters), humans blue circles, calibrated "
                        "machines teal squares, overconfident machines red triangles (checked at the tail)"),
        "callout_group": _call51_group,
        "callout_text": _call51_text,
        "callout_anchor": ("mean RES and mean REL of the callout_group points; the arrow points at "
                           "anchor + (callout_point_dx, callout_point_dy), the text sits at anchor + "
                           "(callout_text_dx, callout_text_dy); the group's points are drawn as open "
                           "markers of graded size at their exact positions"),
        "callout_point_dx": 0.0, "callout_point_dy": 0.003,
        "callout_text_dx": -0.004, "callout_text_dy": 0.040,
        "callout_coincident_max_axis_fraction": 0.02,
        "callout_coincident_rule": ("the callout's 'nearly coincident' holds when the RES range and the REL "
                                    "range of the callout_group points are each below "
                                    "callout_coincident_max_axis_fraction of the axis range (checked at the tail)"),
    }),
    "points": pd.DataFrame(_pts51),
    "annotations": [
        {"text": "poorly\ncalibrated",
         "x": round(_x0_51 + 0.03 * (_x1_51 - _x0_51), 4),
         "y": round(sum(_rel_over51) / len(_rel_over51), 4), "ha": "left"},
        {"text": "calibrated\nand informative",
         "x": round(_x1_51 - 0.03 * (_x1_51 - _x0_51), 4),
         "y": round(0.11 * _y1_51, 4), "ha": "right"},
    ],
})


# =====================================================================
# VIGNETTE V (Section 6.6, human-machine forecast aggregation)
#   -> Figure fig6-12_hm_calsharp (fig:hm-calsharp), Table tab:hm-vignette,
#      the vignette row of tab:robustness, and every number the vignette
#      prose reports.
#
# GENERATED AND SOLVED here.  Redesigned in R24 so that the conclusions the
# text draws are properties of the DESIGN and not of one lucky draw: the
# design-level (population) solution is computed in closed form, the record
# is regenerated under _VG_RESEED["count"] further seeds, and the seed-202603
# draw is placed in that across-seed distribution (percentiles written to
# the workbook).  Under seed 202603:
#   * target: a stationary AR(1) over T = _VG_T quarters, whose innovation is
#     split into a leading signal x_t, already observable when the forecasts
#     for quarter t are issued, and a residual u_t that no forecaster sees,
#         y_0 ~ N(c, s^2/(1-phi^2)) (pre-sample value),
#         y_t = c + phi (y_{t-1} - c) + x_t + u_t,
#         x_t ~ N(0, s_x^2), u_t ~ N(0, s_u^2), s_x^2 + s_u^2 = s^2,
#     c = 2.0, phi = 0.6, s = 1.2, s_u = 0.35, s_x = sqrt(s^2 - s_u^2) (a
#     quarterly growth-rate series in percent; the AR(1) law of y is
#     unchanged by the split);
#   * panel: N_H = 4 human forecasters, N_M = 4 machine experts (two
#     calibrated, two overconfident).  Forecaster i issues N(mu_it, spread_i^2)
#     from the information available at issue time, the previous value and
#     its own reading of the leading signal,
#         mu_it = c + phi (y_{t-1} - c) + x_t + bias_i + tau_i * xi_it,
#     xi_it iid N(0,1); the realized y_t never enters a forecast.  The humans
#     read the signal with modest mixed-sign biases and large reading errors,
#     the calibrated machines M1, M2 with small biases and smaller reading
#     errors; the overconfident machines M3 and M4 share the upward bias 0.60
#     and report narrow spreads, M3 reading the signal poorly and M4 reading
#     it well.  Humans and calibrated machines report the spread commensurate
#     with their error, spread_i = sqrt(s_u^2 + bias_i^2 + tau_i^2); M3, M4
#     report a spread far below it.  The eight (bias, tau, spread) triples
#     are the constants _VG_BIAS / _VG_TAU / _VG_SPREAD_OVER below and are
#     written to the 'forecasters' sheet;
#   * draw order: (1) one standard normal for the pre-sample value y_0, (2) a
#     T x 2 matrix of standard normals for (x_t, u_t), quarter-major, (3) a
#     T x 8 matrix of forecast noise, quarter-major / forecaster-minor.
#     Nothing else is drawn from this rng;
#   * pool: quantile averaging (eq:hm-qa); for normals the aggregate is
#     N(w.mu_t, (w.spread)^2), and the CRPS is the closed form of def:crps
#     (_crps_gauss, with the exact normal cdf scipy.special.ndtr);
#   * program: eq:hm-crps, mean CRPS + lambda_G Gini(w) + eps/2 |w|^2 over the
#     simplex, lambda_G = 0.02, eps_ridge = 1e-3, solved by Algorithm 4.2
#     (projected subgradient descent from the equal-weight vector, analytic
#     subgradient with the same exact cdf, diminishing steps eta_r =
#     eta0/(1 + r/r0), the tie rule sign(0) = 0 for the Gini term, best
#     iterate returned, stop when the sup-norm change of the iterate is below
#     1e-6 after at least r_min iterations, with an iteration guard).  Every
#     solve is cross-checked by scipy SLSQP applied to the equivalent smooth
#     epigraph form of the same program, G(w) = (1/N) sum_{i<j} t_ij with
#     t_ij >= |w_i - w_j| (convex, so a KKT point of it is the minimizer),
#     from the equal-weight vector plus K = 20 Dirichlet(1) starts (a fresh
#     default_rng(202603) per solve draws the starts), each end point scored
#     by the exact objective; the objective gap and the weight difference are
#     written to the 'solves' sheet;
#   * comparators: the same program on the humans alone, the equal-weight
#     pool, and the eight leave-one-out re-solves (Def. def:loo: the reduced
#     program re-forms its Gini term over the N - 1 survivors, so it is not
#     the full program restricted to the face w_i = 0);
#   * in-sample and on new quarters: every score of a fitted pool is reported
#     twice, as its mean CRPS on the T quarters the weights were fitted to
#     (in-sample) and as its expected CRPS on a new quarter of the same
#     design, which is closed form for a normal pool because its error
#     y - w.mu = u - w.b - sum_i w_i tau_i xi_i does not depend on the AR(1)
#     state:
#         E CRPS(w) = R sqrt(2/pi) exp(-m^2/(2R^2)) + m (2 Phi(m/R) - 1) - S/sqrt(pi),
#         m = w.b, S = w.spread, R^2 = S^2 + s_u^2 + sum_i w_i^2 tau_i^2;
#   * design level (no draws): the program with E CRPS in place of the
#     sample mean is the design-level program, solved for the full panel,
#     the humans alone, the machines alone (the comparator Section 6.6.4
#     prints), each leave-one-out subpanel, and the M4 counterfactual, in the
#     same epigraph form from the same 1 + K starts (sheet 'design_solves',
#     which records the spread of the objective over the starts).  The Hersbach split
#     of a normal forecast with constant spread and iid normal error is also
#     closed form (its bin averages are normal partial moments,
#     _vg_pop_hersbach), which gives the design-level reliability,
#     potential score, leave-one-out channels and resolution; the
#     design-level UNC is the continuous climatological score of the
#     stationary law, sigma_y / sqrt(pi), sigma_y = s / sqrt(1 - phi^2);
#   * across seeds: the record is regenerated under the seeds
#     202604 .. 202603 + K with the same draw order, and each regenerated
#     record is solved (full, human-only, eight leave-one-out) in the smooth
#     epigraph form of the same convex program, G(w) = (1/N) sum_{i<j} t_ij
#     with t_ij >= |w_i - w_j| on the simplex, by SLSQP from the equal-weight
#     vector (the program is convex, so this is its exact solution; on the
#     seed-202603 record the epigraph solve reproduces Algorithm 4.2).  The
#     frequency of each claim the text makes is written to 'reseed_summary',
#     and the percentile of every printed or plotted seed-202603 statistic
#     within the across-seed distribution to 'typicality' (each printed one
#     must lie in the central 80 percent, _VG_TYPICAL_RANGE, and each plotted
#     one too unless it is one of the atypical values the text discloses,
#     _VG_TYPICAL_DISCLOSED, checked at the tail);
#   * channels of M4: its leave-one-out calibration and resolution channels
#     are recomputed at the same fitted weights (full panel and without M4)
#     with M4's bias removed and with M4's spread set to its design error, in
#     the draw and at the design level (sheet 'channel_counterfactuals'): the
#     resolution channel comes from its reading of the signal (nearly unchanged
#     by either change), the calibration channel nets the gain from narrowing
#     the over-dispersed pool against the cost of the bias;
#   * weight of M4 at its error: the full-panel program is re-solved with M4's
#     spread set to its design error, on the seed-202603 record (Algorithm 4.2,
#     certified like every other solve), at the design level, and on every
#     regenerated record (epigraph solve); the ratio of M4's weight to its
#     weight at its error is about two at the design level and in the median
#     record (the text's "about twice", scoped to the design), its value in
#     the draw lying inside the central 80 percent of the regenerated records;
#   * diagnostics: the fraction of outcomes below each forecaster's median,
#     with its two-sided binomial p-value against one-half, and the Hersbach
#     reliability / potential split on the L = 19 ensemble-quantile
#     representation of each forecast sequence (outer bins scored), the
#     binning of the ex:hmpanel block above.  The leave-one-out contribution
#     the text reports is the difference of the nineteen-member step-function
#     scores (column delta_loo_hersbach), so that it is exactly the sum of
#     its two channels (eq:loo-split); the exact-CRPS difference of Def.
#     def:loo is kept beside it (delta_loo).  The resolution is RES_i = UNC -
#     POT_i with UNC = 0.5 mean|y - y'| the continuous sample climatological
#     CRPS of the T outcomes.  Section 4 and ex:hmpanel form UNC on the same
#     L-member grid instead (only the binning is shared); the grid value is
#     recorded beside it (uncertainty_UNC_grid), and the departure shifts
#     every member's RES_i by the same amount, which the tail checks, (6c').  Because
#     the outer bins charge a narrow biased forecast's one-sided misses to
#     reliability, the plotted resolution of M3, M4 overstates what they
#     carry; the resolution of each member after recalibration (mean error
#     removed, spread set to the error sd) is written beside it, with the two
#     single corrections (mean error removed with the spread kept, spread set
#     to the error sd with the mean kept) that separate the part due to the
#     bias from the part due to the narrow spread, and the plotted and
#     recalibrated values are recomputed on finer representations (L in
#     _VG_L_GRID, sheet 'resolution_by_L'), where the excess shrinks but
#     persists and the order holds.
# All settings are written to the workbook's config sheet.
# =====================================================================
from scipy.optimize import minimize as _sp_minimize
from scipy.stats import binom as _vg_binom

_VG_T = 72
_VG_AR = {"c": 2.0, "phi": 0.6, "sigma": 1.2}
_VG_SIGMA_U = 0.35
_VG_SIGNAL = {"sigma_x": math.sqrt(_VG_AR["sigma"] ** 2 - _VG_SIGMA_U ** 2),   # innovation split
              "sigma_u": _VG_SIGMA_U}
_VG_IDS = ["H1", "H2", "H3", "H4", "M1", "M2", "M3", "M4"]
_VG_GROUPS = (["human forecasters"] * 4 + ["calibrated machines"] * 2
              + ["overconfident machines"] * 2)
_VG_COLOR = ["blue"] * 4 + ["teal"] * 2 + ["red"] * 2
_VG_MARKER = ["o"] * 4 + ["s"] * 2 + ["^"] * 2
_VG_BIAS = np.array([0.20, -0.15, 0.25, -0.10, 0.05, -0.05, 0.60, 0.60])
_VG_TAU = np.array([0.85, 0.90, 0.80, 0.95, 0.50, 0.55, 1.10, 0.10])
_VG_SPREAD_OVER = np.array([0.15, 0.12])           # reported by M3, M4
_VG_RMSE = np.sqrt(_VG_SIGNAL["sigma_u"] ** 2 + _VG_BIAS ** 2 + _VG_TAU ** 2)   # design error of each mean
_VG_SPREAD = np.concatenate([_VG_RMSE[:6], _VG_SPREAD_OVER])
_VG_LAM_G, _VG_EPS = 0.02, 1e-3
_VG_SOLVER = {"eta0": 0.5, "r0": 50.0, "tol": 1e-6, "r_min": 500, "r_max": 500000}
_VG_CERT = {"K": 20}      # cross-check starts: the equal-weight vector plus K Dirichlet(1) draws
_VG_CERT_GAP = 1e-7      # every solve must agree with its SLSQP cross-check to this bound (the book prints it)
_VG_L = 19
_VG_RESEED = {"first": NEW_SEED + 1, "count": 500}
_VG_EPI = {"ftol": 1e-13, "maxiter": 1000}         # SLSQP on the smooth epigraph form (cross-check, design level, re-seeds)
_VG_EPI_AGREE = {"weights": 1e-3, "score": 1e-6}   # epigraph vs Algorithm 4.2 on the seed-202603 record
_VG_EPI_FAILURES = []                               # SLSQP runs that end without success (counted per record)
_VG_BAND = 0.2           # well-calibrated band: REL_i below a fifth of UNC (the figure's shaded band)
_VG_BIAS_ALPHA = 0.01    # the below-median count exposes a bias when its two-sided binomial p < alpha
# every seed-202603 statistic the text prints, and the across-seed percentile range
# (the central 80 percent) inside which each must lie for the draw to count as
# typical (checked at the tail)
_VG_TYPICAL_PRINTED = ["ins_hm", "ins_h", "ins_eq", "oos_hm", "oos_h", "max_loo_ins", "W_M",
                       "w_M1", "w_M2", "w_M3", "w_M4", "w_H3", "crps_M4",
                       "delta_M1", "delta_M2", "delta_M4", "delta_H2", "delta_M3", "delta_H3",
                       "dcal_M4", "dres_M4", "below_M4", "below_M3",
                       "UNC", "RES_M4", "RES_M3", "RES_M1", "RESrec_M4", "RESrec_M3", "RESrec_M1", "rmse_cal6",
                       "RES_finest_M4"]
_VG_TYPICAL_RANGE = (10.0, 90.0)
# the plotted coordinates of Figure 6.12 that fall outside that range, each disclosed
# in the text with its direction (below the 10th / above the 90th percentile); every
# other plotted coordinate must lie inside it (checked at the tail)
_VG_TYPICAL_DISCLOSED = {"max_REL_calibrated_over_UNC": "below", "REL_H3": "below", "REL_H4": "below",
                         "REL_M3": "above"}
_VG_TIE_TOL = 1e-4        # two weights count as level when they agree to this bound (frequency recorded, not a claim)
_VG_CMP_EPS = 1e-12       # margin of the strict comparisons of the recorded-only claims, so that a claim evaluates
                          # alike on the in-memory values and on their workbook round-trip (last-digit rounding)
_VG_CF_RES_TOL = 0.1      # M4's resolution channel moves by less than this share of itself when its bias or its narrowness is removed
_VG_CF_RATIO_BAND = (1.8, 2.2)   # "about twice": the ratio of M4's weight to its weight at its error, at the design level and in the median re-seeded record
_VG_FOCUS = 7                                      # M4 (sharp, biased, reads the signal well)
_VG_L_GRID = [19, 39, 99, 499, 1999]               # finer ensemble representations for the resolution check
# the draw conforms to its design: each member's realized error within these factors of
# its design error, and no member's reading noise correlated with the unobserved
# residual beyond CORR_Z / sqrt(T) (checked at the tail)
_VG_CONFORM = {"rmse_ratio": (0.75, 1.25), "corr_z": 3.0}
_SQRT_PI_INV = 1.0 / math.sqrt(math.pi)


def _vg_generate(seed=NEW_SEED):
    rng = np.random.default_rng(seed)
    c, phi, s = _VG_AR["c"], _VG_AR["phi"], _VG_AR["sigma"]
    sx, su = _VG_SIGNAL["sigma_x"], _VG_SIGNAL["sigma_u"]
    assert abs(sx ** 2 + su ** 2 - s ** 2) < 1e-12, "the innovation split must preserve the AR(1) law"
    y0 = c + s / math.sqrt(1.0 - phi ** 2) * rng.normal(0.0, 1.0)   # draw (1)
    xu = rng.normal(0.0, 1.0, (_VG_T, 2))                  # draw (2)
    x, u = sx * xu[:, 0], su * xu[:, 1]
    y, m_pub = np.empty(_VG_T), np.empty(_VG_T)
    prev = y0
    for t in range(_VG_T):
        m_pub[t] = c + phi * (prev - c)                    # conditional mean given y_{t-1}
        y[t] = m_pub[t] + x[t] + u[t]
        prev = y[t]
    xi = rng.normal(0.0, 1.0, (_VG_T, len(_VG_IDS)))       # draw (3)
    # issue-time information only: the previous value and a reading of the signal
    MU = (m_pub + x)[:, None] + _VG_BIAS[None, :] + _VG_TAU[None, :] * xi
    return y, MU, _VG_SPREAD.copy(), y0, x, u, m_pub


def _vg_gini_term(w):
    """Gini of a simplex vector and its subgradient (tie rule sign(0) = 0)."""
    n = w.size
    D = w[:, None] - w[None, :]
    mbar = w.mean()
    G = float(np.abs(D).sum() / (2.0 * n * n * mbar))      # == _gini_vec(w)
    return G, np.sign(D).sum(axis=1) / (n * n * mbar) - G / (n * mbar)


def _vg_val_grad(w, mu, sd, y, lam_G, eps):
    """Objective of eq:hm-crps and one of its subgradients, with the exact
    normal cdf (the objective the SLSQP cross-check evaluates)."""
    m = mu @ w
    s = float(sd @ w)
    z = (y - m) / s
    Phi = _ndtr(z)
    phi = _npdf(z)
    crps = float((s * (z * (2.0 * Phi - 1.0) + 2.0 * phi - _SQRT_PI_INV)).mean())
    G, gG = _vg_gini_term(w)
    v = crps + lam_G * G + 0.5 * eps * float(w @ w)
    g = ((1.0 - 2.0 * Phi)[:, None] * mu).mean(axis=0) + (2.0 * phi - _SQRT_PI_INV).mean() * sd
    return v, g + lam_G * gG + eps * w


def _vg_objective(w, mu, sd, y, lam_G, eps):
    w = np.asarray(w, float)
    return float(_crps_gauss(mu @ w, float(sd @ w), y).mean()
                 + lam_G * _gini_vec(w) + 0.5 * eps * float(w @ w))


def _vg_psd_core(val_grad, n, eta0, r0, tol, r_min, r_max):
    """Algorithm 4.2: projected subgradient descent from the equal-weight
    vector, best iterate returned."""
    w = np.full(n, 1.0 / n)
    best_v, g = val_grad(w)
    best_w = w.copy()
    r = 0
    while True:
        eta = eta0 / (1.0 + r / r0)
        w_new = _project_simplex(w - eta * g)[0]
        v, g = val_grad(w_new)
        if v < best_v:
            best_v, best_w = v, w_new.copy()
        r += 1
        step = float(np.abs(w_new - w).max())
        w = w_new
        if (step < tol and r >= r_min) or r >= r_max:
            break
    return best_w, best_v, r, step, bool(step < tol)


def _vg_epi_solve(fun, grad, n, w0=None):
    """min fun(w) + lambda_G G(w) + eps/2 |w|^2 on the simplex in its smooth
    epigraph form, G(w) = (1/n) sum_{i<j} t_ij with t_ij >= |w_i - w_j|
    (convex, so any KKT point is the minimizer), by SLSQP from w0 (default
    the equal-weight vector) with the analytic gradient."""
    I, J = np.triu_indices(n, 1)
    P = I.size
    w0 = np.full(n, 1.0 / n) if w0 is None else np.asarray(w0, float)
    z0 = np.concatenate([w0, np.abs(w0[I] - w0[J])])
    A = np.zeros((2 * P, n + P))
    r = np.arange(P)
    A[r, n + r] = 1.0
    A[r, I] -= 1.0
    A[r, J] += 1.0
    A[P + r, n + r] = 1.0
    A[P + r, I] += 1.0
    A[P + r, J] -= 1.0
    Aeq = np.zeros((1, n + P))
    Aeq[0, :n] = 1.0

    def F(z):
        w = z[:n]
        return fun(w) + _VG_LAM_G / n * float(z[n:].sum()) + 0.5 * _VG_EPS * float(w @ w)

    def dF(z):
        w = z[:n]
        return np.concatenate([grad(w) + _VG_EPS * w, np.full(P, _VG_LAM_G / n)])

    res = _sp_minimize(F, z0, jac=dF, method="SLSQP", bounds=[(0.0, 1.0)] * (n + P),
                       constraints=({"type": "ineq", "fun": lambda z: A @ z, "jac": lambda z: A},
                                    {"type": "eq", "fun": lambda z: Aeq @ z - 1.0, "jac": lambda z: Aeq}),
                       options={"ftol": _VG_EPI["ftol"], "maxiter": _VG_EPI["maxiter"]})
    if not res.success:
        _VG_EPI_FAILURES.append((n, int(res.status), str(res.message)))
    w = np.clip(res.x[:n], 0.0, None)
    return w / w.sum()


def _vg_certify_epi(fun, grad, obj, n, K):
    """The epigraph SLSQP solve from the equal-weight vector and K Dirichlet(1)
    starts drawn from a fresh default_rng(NEW_SEED); each end point is scored
    by the exact (nonsmooth) objective obj.  Returns the best point, its
    value, and the spread of the values over the starts."""
    rng = np.random.default_rng(NEW_SEED)
    starts = [np.full(n, 1.0 / n)] + [rng.dirichlet(np.ones(n)) for _ in range(K)]
    pts = [_vg_epi_solve(fun, grad, n, w0) for w0 in starts]
    vals = np.array([obj(w) for w in pts])
    k = int(np.argmin(vals))
    return pts[k], float(vals[k]), float(vals.max() - vals.min())


def _vg_solve(y, MU, SD, idx, label):
    """Algorithm 4.2 on the record, cross-checked by the epigraph SLSQP."""
    mu, sd = MU[:, idx], SD[idx]
    n = len(idx)
    w, v, iters, step, met = _vg_psd_core(lambda q: _vg_val_grad(q, mu, sd, y, _VG_LAM_G, _VG_EPS),
                                          n, **_VG_SOLVER)
    f, g = _vg_sample_fg(mu, sd, y)
    wc, vc, vspread = _vg_certify_epi(f, g, lambda q: _vg_objective(q, mu, sd, y, _VG_LAM_G, _VG_EPS),
                                      n, _VG_CERT["K"])
    return w, {"solve": label, "n_experts": n, "iterations": iters,
               "final_step_sup_norm": step, "tolerance_met": int(met),
               "objective_best_iterate": v, "objective_certificate_slsqp": vc,
               "objective_gap": v - vc, "max_abs_weight_diff_vs_certificate": float(np.abs(w - wc).max()),
               "certificate_value_spread_over_starts": vspread,
               "gini_w": _gini_vec(w), "mean_CRPS": float(_crps_gauss(mu @ w, float(sd @ w), y).mean())}


# ---- design level: expected CRPS on a new quarter, closed form ------------
def _vg_ecrps(w, idx, spread=_VG_SPREAD):
    """Expected CRPS on a new quarter of the pool of members idx with weights w."""
    w = np.asarray(w, float)
    b, tau, s = _VG_BIAS[idx], _VG_TAU[idx], spread[idx]
    S = float(w @ s)
    m = float(w @ b)
    R = math.sqrt(S * S + _VG_SIGMA_U ** 2 + float((w * w) @ (tau * tau)))
    return (R * math.sqrt(2.0 / math.pi) * math.exp(-m * m / (2.0 * R * R))
            + m * (2.0 * float(_ndtr(m / R)) - 1.0) - S * _SQRT_PI_INV)


def _vg_ecrps_grad(w, idx, spread=_VG_SPREAD):
    b, tau, s = _VG_BIAS[idx], _VG_TAU[idx], spread[idx]
    S = float(w @ s)
    m = float(w @ b)
    R = math.sqrt(S * S + _VG_SIGMA_U ** 2 + float((w * w) @ (tau * tau)))
    return ((2.0 * float(_ndtr(m / R)) - 1.0) * b
            + 2.0 * float(_npdf(m / R)) * (S * s + w * tau * tau) / R - s * _SQRT_PI_INV)


def _vg_pop_objective(w, idx, spread=_VG_SPREAD):
    w = np.asarray(w, float)
    return _vg_ecrps(w, idx, spread) + _VG_LAM_G * _gini_vec(w) + 0.5 * _VG_EPS * float(w @ w)


def _vg_pop_solve(idx, label, spread=_VG_SPREAD):
    """The design-level program, solved in its epigraph form from the
    equal-weight vector and K Dirichlet(1) starts (best of the starts)."""
    w, v, vspread = _vg_certify_epi(lambda q: _vg_ecrps(q, idx, spread),
                                    lambda q: _vg_ecrps_grad(q, idx, spread),
                                    lambda q: _vg_pop_objective(q, idx, spread), len(idx), _VG_CERT["K"])
    return w, {"solve": label, "n_experts": len(idx), "objective": v,
               "objective_spread_over_starts": vspread, "gini_w": _gini_vec(w),
               "expected_CRPS": _vg_ecrps(w, idx, spread)}


def _vg_pool_error(w, idx, bias=_VG_BIAS, spread=_VG_SPREAD):
    """(mean, sd) of the error y - w.mu of the pool, and its spread w.s."""
    w = np.asarray(w, float)
    return (-float(w @ bias[idx]),
            math.sqrt(_VG_SIGMA_U ** 2 + float((w * w) @ (_VG_TAU[idx] ** 2))),
            float(w @ spread[idx]))


def _vg_pop_hersbach(mean_err, sd_err, S, zlev):
    """Design-level (T -> infinity) Hersbach split, same convention as
    _hersbach_X, of a normal forecast N(mu_t, S^2) whose error y_t - mu_t is
    iid N(mean_err, sd_err^2): the ensemble is mu_t + S z_l, so every bin
    average is S times a partial moment of v = (y - mu)/S ~ N(a, k^2).
    Returns (crps_h, REL, POT)."""
    a, k = mean_err / S, sd_err / S
    z = np.asarray(zlev, float)
    L = z.size
    rel = pot = 0.0
    for i in range(1, L):
        z1, z2 = z[i - 1], z[i]
        d = z2 - z1
        l1, l2 = (z1 - a) / k, (z2 - a) / k
        P_in = float(_ndtr(l2) - _ndtr(l1))
        E_in = (a - z1) * P_in + k * float(_npdf(l1) - _npdf(l2))     # E[(v - z1) 1{z1 < v < z2}]
        alpha = d * float(1.0 - _ndtr(l2)) + E_in
        beta = d * float(_ndtr(l1)) + (d * P_in - E_in)
        g = alpha + beta
        o = beta / g
        rel += g * (o - i / float(L)) ** 2
        pot += g * o * (1.0 - o)
    l1 = (z[0] - a) / k
    O0 = float(_ndtr(l1))
    g0 = (k * float(_npdf(l1)) + (z[0] - a) * O0) / O0                  # E[z1 - v | v < z1]
    rel += g0 * O0 ** 2
    pot += g0 * O0 * (1.0 - O0)
    lL = (z[-1] - a) / k
    OL = float(1.0 - _ndtr(lL))
    gL = (k * float(_npdf(lL)) + (a - z[-1]) * OL) / OL                 # E[v - zL | v > zL]
    rel += gL * OL ** 2
    pot += gL * OL * (1.0 - OL)
    return S * (rel + pot), S * rel, S * pot


def _vg_sample_fg(mu, sd, y):
    def f(w):
        return float(_crps_gauss(mu @ w, float(sd @ w), y).mean())

    def g(w):
        z = (y - mu @ w) / float(sd @ w)
        return (((1.0 - 2.0 * _ndtr(z))[:, None] * mu).mean(axis=0)
                + float((2.0 * _npdf(z) - _SQRT_PI_INV).mean()) * sd)
    return f, g


def _vg_binom_p(k, n):
    """Two-sided binomial p-value of k successes in n trials against 1/2."""
    k = int(k)
    return float(min(1.0, 2.0 * min(_vg_binom.cdf(k, n, 0.5), _vg_binom.sf(k - 1, n, 0.5))))


def _vg_record_stats(y, MU, u):
    """Every statistic the text reports for one record, the fits made by the
    epigraph solve (used identically for seed 202603 and each re-seed)."""
    N = len(_VG_IDS)
    allx = list(range(N))
    hum = [i for i in allx if _VG_IDS[i].startswith("H")]
    nfail0 = len(_VG_EPI_FAILURES)
    f, g = _vg_sample_fg(MU, _VG_SPREAD, y)
    w = _vg_epi_solve(f, g, N)
    fh, gh = _vg_sample_fg(MU[:, hum], _VG_SPREAD[hum], y)
    wh = _vg_epi_solve(fh, gh, len(hum))
    weq = np.full(N, 1.0 / N)
    zl = _norm_ppf(np.arange(1, _VG_L + 1) / (_VG_L + 1.0))
    unc = float(0.5 * np.abs(y[:, None] - y[None, :]).mean())
    chf, relf, potf = _hersbach(MU @ w, np.full(y.size, float(_VG_SPREAD @ w)), y, zl)
    st = {"sd_u": float(np.sqrt(np.mean(u ** 2))), "UNC": unc,
          "ins_hm": f(w), "ins_h": fh(wh), "ins_eq": f(weq),
          "oos_hm": _vg_ecrps(w, allx), "oos_h": _vg_ecrps(wh, hum), "oos_eq": _vg_ecrps(weq, allx),
          "W_M": float(w[[i for i in allx if i not in hum]].sum())}
    # the full panel re-solved with the focus machine's spread set to its design error
    scf = _VG_SPREAD.copy()
    scf[_VG_FOCUS] = _VG_RMSE[_VG_FOCUS]
    fc, gc = _vg_sample_fg(MU, scf, y)
    st["w_" + _VG_IDS[_VG_FOCUS] + "_spread_at_error"] = float(_vg_epi_solve(fc, gc, N)[_VG_FOCUS])
    loo_ins, loo_oos = [], []
    for i in allx:
        sub = [j for j in allx if j != i]
        fi, gi = _vg_sample_fg(MU[:, sub], _VG_SPREAD[sub], y)
        wi = _vg_epi_solve(fi, gi, N - 1)
        loo_ins.append(fi(wi))
        loo_oos.append(_vg_ecrps(wi, sub))
        chi, reli, poti = _hersbach(MU[:, sub] @ wi, np.full(y.size, float(_VG_SPREAD[sub] @ wi)), y, zl)
        _, rel_ind, pot_ind = _hersbach(MU[:, i], np.full(y.size, _VG_SPREAD[i]), y, zl)
        err = y - MU[:, i]
        _, _, pot_rec = _hersbach(MU[:, i] + err.mean(), np.full(y.size, float(err.std())), y, zl)
        k = int((y < MU[:, i]).sum())
        st.update({"w_" + _VG_IDS[i]: float(w[i]),
                   "delta_" + _VG_IDS[i]: float(chi - chf),
                   "dcal_" + _VG_IDS[i]: float(reli - relf),
                   "dres_" + _VG_IDS[i]: float(poti - potf),
                   "below_" + _VG_IDS[i]: k,
                   "REL_" + _VG_IDS[i]: float(rel_ind),
                   "RES_" + _VG_IDS[i]: float(unc - pot_ind),
                   "RESrec_" + _VG_IDS[i]: float(unc - pot_rec),
                   "crps_" + _VG_IDS[i]: float(_crps_gauss(MU[:, i], _VG_SPREAD[i], y).mean())})
    cal = [i for i in allx if _VG_GROUPS[i] != "overconfident machines"]
    st["rmse_cal6"] = float(np.sqrt(np.mean((y - MU[:, cal].mean(axis=1)) ** 2)))
    Lf = _VG_L_GRID[-1]
    st["RES_finest_" + _VG_IDS[_VG_FOCUS]] = float(unc - _hersbach(
        MU[:, _VG_FOCUS], np.full(y.size, _VG_SPREAD[_VG_FOCUS]), y,
        _norm_ppf(np.arange(1, Lf + 1) / (Lf + 1.0)))[2])
    st["max_loo_ins"] = float(max(loo_ins))
    st["max_loo_oos"] = float(max(loo_oos))
    st["max_REL_calibrated_over_UNC"] = float(max(st["REL_" + _VG_IDS[i]] for i in allx
                                                  if _VG_GROUPS[i] != "overconfident machines") / unc)
    fid = _VG_IDS[_VG_FOCUS]
    st["below_p_" + fid] = _vg_binom_p(st["below_" + fid], y.size)
    st["epigraph_failures"] = len(_VG_EPI_FAILURES) - nfail0
    return st


# the claims of Section 6.6, each a predicate on one record's statistics
_VG_CLAIMS = [
    ("C1_in_sample_hm_beats_human_only", "in-sample mean CRPS: optimized human-machine < optimized human-only",
     lambda s: s["ins_hm"] < s["ins_h"]),
    ("C1_in_sample_hm_beats_equal", "in-sample mean CRPS: optimized human-machine < equal-weight pool",
     lambda s: s["ins_hm"] < s["ins_eq"]),
    ("C1_new_quarters_hm_beats_human_only", "expected CRPS on new quarters of the fitted weights: human-machine < human-only",
     lambda s: s["oos_hm"] < s["oos_h"]),
    ("C1_new_quarters_hm_beats_equal", "expected CRPS on new quarters of the fitted weights: human-machine < equal-weight",
     lambda s: s["oos_hm"] < s["oos_eq"]),
    ("C2_machine_share_above_half", "machine weight share W_M > 1/2",
     lambda s: s["W_M"] > 0.5),
    ("C3_M4_weight_above_equal_share", "w_M4 > 1/N (the optimizer raises the sharp, biased machine above its equal share)",
     lambda s: s["w_M4"] > 1.0 / len(_VG_IDS)),
    ("C3_M4_positive_contribution", "leave-one-out contribution of M4 (step score) > 0",
     lambda s: s["delta_M4"] > 0),
    ("C3_M4_resolution_plus_calibration_minus", "M4: resolution channel > 0 and calibration channel < 0",
     lambda s: s["dres_M4"] > 0 and s["dcal_M4"] < 0),
    ("C3_M4_bias_exposed", "M4's below-median count differs from one-half, two-sided binomial p < alpha",
     lambda s: s["below_p_M4"] < _VG_BIAS_ALPHA),
    ("C4_calibrated_in_band", "the six calibrated members (humans, M1, M2) have REL_i < band * UNC",
     lambda s: s["max_REL_calibrated_over_UNC"] < _VG_BAND),
    ("robust_every_loo_beats_human_only_in_sample", "every leave-one-out re-solve beats the human-only pool in-sample",
     lambda s: s["max_loo_ins"] < s["ins_h"]),
    ("robust_every_loo_beats_human_only_new_quarters", "every leave-one-out pool beats the human-only pool on new quarters",
     lambda s: s["max_loo_oos"] < s["oos_h"]),
    # further statements of the text about the draw, recorded for their frequency (not core claims)
    ("S_M3_below_half_equal_share", "w_M3 < 1/(2N) (the equally narrow, equally biased M3 gets almost no weight)",
     lambda s: s["w_M3"] < 0.5 / len(_VG_IDS)),
    ("S_calibrated_machines_both_channels", "M1 and M2 each contribute through both channels (calibration and resolution channels > 0)",
     lambda s: min(s["dcal_M1"], s["dres_M1"], s["dcal_M2"], s["dres_M2"]) > 0),
    ("S_M4_highest_resolution", "M4 has the highest resolution in the panel, as plotted and after recalibration",
     lambda s: all(s["RES_M4"] >= s["RES_" + k] and s["RESrec_M4"] >= s["RESrec_" + k] for k in _VG_IDS)),
    ("S_M3_recalibrated_below_humans", "after recalibration M3 resolves less than every human forecaster",
     lambda s: all(s["RESrec_M3"] < s["RESrec_" + k] for k in _VG_IDS if k.startswith("H"))),
    ("S_M1_M2_M4_three_largest_weights", "M1, M2 and M4 hold the three largest weights (each above every other weight by more than the tie tolerance)",
     lambda s: min(s["w_M1"], s["w_M2"], s["w_M4"]) > max(s["w_" + k] for k in ("H1", "H2", "H3", "H4", "M3")) + _VG_TIE_TOL),
    ("S_M1_M2_M4_weights_level", "the weights of M1, M2 and M4 agree to within the tie tolerance (a feature of the solution, not a claim)",
     lambda s: max(abs(s["w_M1"] - s["w_M4"]), abs(s["w_M2"] - s["w_M4"])) < _VG_TIE_TOL),
    ("S_three_largest_weights_contribute_most", "M1, M2 and M4 have the three largest leave-one-out contributions (a feature of the draw)",
     lambda s: min(s["delta_M1"], s["delta_M2"], s["delta_M4"]) > max(s["delta_" + k] for k in ("H1", "H2", "H3", "H4", "M3")) + _VG_CMP_EPS),
    ("S_contribution_order_M2_M1_M4_H2", "contributions ordered M2 > M1 > M4 > H2 > the rest (a feature of the draw)",
     lambda s: (s["delta_M2"] - _VG_CMP_EPS > s["delta_M1"] > s["delta_M4"] + _VG_CMP_EPS
                and s["delta_M4"] > s["delta_H2"] + _VG_CMP_EPS > max(s["delta_" + k] for k in ("H1", "H3", "H4", "M3")) + 2 * _VG_CMP_EPS)),
    ("S_M4_weight_falls_with_spread_at_error", "M4's weight falls when the full panel is re-solved with its spread set to its design error",
     lambda s: s["w_M4_spread_at_error"] < s["w_M4"] - _VG_CMP_EPS),
    ("S_M4_weight_more_than_halves_at_error", "M4's weight more than halves at its error, w_M4 > 2 w_M4_spread_at_error (recorded, not a claim; the text says about twice, scoped to the design)",
     lambda s: s["w_M4"] > 2.0 * s["w_M4_spread_at_error"] + _VG_CMP_EPS),
]


def _vg_cf_ratio(w_narrow, w_at_error):
    """Ratio of M4's weight to its weight at its error (inf when the latter vanishes)."""
    w_narrow, w_at_error = np.asarray(w_narrow, float), np.asarray(w_at_error, float)
    return np.where(w_at_error > 0, w_narrow / np.where(w_at_error > 0, w_at_error, 1.0), np.inf)
_VG_CORE = [k for k, _, _ in _VG_CLAIMS if k.startswith("C")]


print("\n-- generating and solving the Section 6.6 vignette "
      "(fig6-12_hm_calsharp.xlsx) ...")
_vy, _vMU, _vSD, _vpre0, _vx, _vu, _vmpub = _vg_generate()
_vN = len(_VG_IDS)
_vall = list(range(_vN))
_vhum = [i for i in _vall if _VG_IDS[i].startswith("H")]
_vmac = [i for i in _vall if _VG_IDS[i].startswith("M")]
_vcal = [i for i in _vall if _VG_GROUPS[i] != "overconfident machines"]
_vg_solves = []
_vw_full, _r = _vg_solve(_vy, _vMU, _vSD, _vall, "full panel")
_vg_solves.append(_r)
_vcrps_full = _r["mean_CRPS"]
_vw_hum_sub, _r = _vg_solve(_vy, _vMU, _vSD, _vhum, "human-only")
_vg_solves.append(_r)
_vcrps_hum = _r["mean_CRPS"]
_vw_hum = np.zeros(_vN)
_vw_hum[_vhum] = _vw_hum_sub
_vw_eq = np.full(_vN, 1.0 / _vN)
_vcrps_eq = float(_crps_gauss(_vMU @ _vw_eq, float(_vSD @ _vw_eq), _vy).mean())
_vWH = float(_vw_full[_vhum].sum())
_vWM = float(_vw_full[_vmac].sum())
# the same fitted weights scored on new quarters of the design (closed form)
_vnew_full = _vg_ecrps(_vw_full, _vall)
_vnew_hum = _vg_ecrps(_vw_hum_sub, _vhum)
_vnew_eq = _vg_ecrps(_vw_eq, _vall)

# Hersbach split (L-member ensemble quantiles, outer bins scored) of the
# full aggregate, the leave-one-out aggregates, and each individual forecaster
_vzlev = _norm_ppf(np.arange(1, _VG_L + 1) / (_VG_L + 1.0))
_vUNC = float(0.5 * np.abs(_vy[:, None] - _vy[None, :]).mean())
# the uncertainty term under the convention of Section 4 and ex:hmpanel, the
# climatology represented on the same L-member grid (the empirical outcome
# quantiles at the levels l/(L+1), numpy linear interpolation) and scored by the
# same routine; recorded for comparison, the vignette's RES_i use _vUNC
_vUNC_grid = float(_hersbach_X(np.tile(np.quantile(_vy, np.arange(1, _VG_L + 1) / (_VG_L + 1.0)),
                                       (_VG_T, 1)), _vy)[2])
_vch_full, _vrel_full, _vpot_full = _hersbach(
    _vMU @ _vw_full, np.full(_VG_T, float(_vSD @ _vw_full)), _vy, _vzlev)

# ---- design level (no draws) ---------------------------------------------
_vsig_y = _VG_AR["sigma"] / math.sqrt(1.0 - _VG_AR["phi"] ** 2)
_vUNC_design = _vsig_y * _SQRT_PI_INV
_vd_solves = []
_vdw_full, _r = _vg_pop_solve(_vall, "design full panel")
_vd_solves.append(_r)
_vdw_hum, _r = _vg_pop_solve(_vhum, "design human-only")
_vd_solves.append(_r)
# the same design-level program over the four machine experts alone: the comparator that
# shows what the human forecasters add (Section 6.6.4 prints its expected score)
_vdw_mac, _r = _vg_pop_solve(_vmac, "design machine-only")
_vd_solves.append(_r)
_vdE_full = _vg_ecrps(_vdw_full, _vall)
_vdE_hum = _vg_ecrps(_vdw_hum, _vhum)
_vdE_mac = _vg_ecrps(_vdw_mac, _vmac)
_vdE_eq = _vg_ecrps(_vw_eq, _vall)
_vd_ch_full, _vd_rel_full, _vd_pot_full = _vg_pop_hersbach(*_vg_pool_error(_vdw_full, _vall), _vzlev)
_vd_loo_w = np.full((_vN, _vN), np.nan)
_vd_rows = {}
for _i in _vall:
    _sub = [j for j in _vall if j != _i]
    _w_i, _r = _vg_pop_solve(_sub, "design leave-one-out %s" % _VG_IDS[_i])
    _vd_solves.append(_r)
    _vd_loo_w[_i, _sub] = _w_i
    _ch_i, _rel_i, _pot_i = _vg_pop_hersbach(*_vg_pool_error(_w_i, _sub), _vzlev)
    _sd_i = math.sqrt(_VG_SIGMA_U ** 2 + _VG_TAU[_i] ** 2)
    _, _r_ind, _p_ind = _vg_pop_hersbach(-_VG_BIAS[_i], _sd_i, _VG_SPREAD[_i], _vzlev)
    _, _, _p_rec = _vg_pop_hersbach(0.0, _sd_i, _sd_i, _vzlev)
    _, _, _p_mrem = _vg_pop_hersbach(0.0, _sd_i, _VG_SPREAD[_i], _vzlev)       # mean error removed, spread kept
    _, _, _p_scor = _vg_pop_hersbach(-_VG_BIAS[_i], _sd_i, _sd_i, _vzlev)      # spread set to the error sd, bias kept
    _vd_rows[_i] = {
        "weight_design": float(_vdw_full[_i]),
        "expected_CRPS_loo_design": _vg_ecrps(_w_i, _sub),
        "delta_design_exact": _vg_ecrps(_w_i, _sub) - _vdE_full,
        "delta_design": float(_ch_i - _vd_ch_full),
        "delta_cal_design": float(_rel_i - _vd_rel_full),
        "delta_res_design": float(_pot_i - _vd_pot_full),
        "p_below_median_design": float(_ndtr(_VG_BIAS[_i] / _sd_i)),
        "REL_design": float(_r_ind),
        "resolution_design": float(_vUNC_design - _p_ind),
        "resolution_recalibrated_design": float(_vUNC_design - _p_rec),
        "resolution_mean_error_removed_design": float(_vUNC_design - _p_mrem),
        "resolution_spread_corrected_design": float(_vUNC_design - _p_scor),
    }
# why a narrow member earns weight: the equal-weight pool of the six calibrated
# members reports the average of their spreads, which overstates the error of the
# averaged mean (quantile averaging does not shrink the spread when the means'
# idiosyncratic errors cancel); design values and the realized ones
_vw6 = np.full(len(_vcal), 1.0 / len(_vcal))
_vcal6_spread = float(_vSD[_vcal] @ _vw6)
_vcal6_rmse_design = math.sqrt(_VG_SIGMA_U ** 2 + float((_vw6 ** 2) @ (_VG_TAU[_vcal] ** 2))
                               + float(_vw6 @ _VG_BIAS[_vcal]) ** 2)
_vcal6_rmse = float(np.sqrt(np.mean((_vy - _vMU[:, _vcal] @ _vw6) ** 2)))
# counterfactual: M4 reporting the spread commensurate with its error (design level)
_vcf_spread = _VG_SPREAD.copy()
_vcf_spread[_VG_FOCUS] = _VG_RMSE[_VG_FOCUS]
_vcf_w, _r = _vg_pop_solve(_vall, "design full panel, M4 spread set to its error", _vcf_spread)
_vd_solves.append(_r)

# ---- the seed-202603 record: leave-one-out re-solves and diagnostics ------
_vloo_w = np.full((_vN, _vN), np.nan)
_vrows = []
for _i in _vall:
    _sub = [j for j in _vall if j != _i]
    _w_i, _r = _vg_solve(_vy, _vMU, _vSD, _sub, "leave-one-out %s" % _VG_IDS[_i])
    _vg_solves.append(_r)
    _vloo_w[_i, _sub] = _w_i
    _ch_i, _rel_i, _pot_i = _hersbach(
        _vMU[:, _sub] @ _w_i, np.full(_VG_T, float(_vSD[_sub] @ _w_i)), _vy, _vzlev)
    _c_ind, _r_ind, _p_ind = _hersbach(_vMU[:, _i], np.full(_VG_T, _vSD[_i]), _vy, _vzlev)
    _err = _vy - _vMU[:, _i]
    _mrec = _vMU[:, _i] + _err.mean()
    _srec = float(np.sqrt(np.mean((_err - _err.mean()) ** 2)))
    _, _, _p_rec = _hersbach(_mrec, np.full(_VG_T, _srec), _vy, _vzlev)
    _, _, _p_mrem = _hersbach(_mrec, np.full(_VG_T, _vSD[_i]), _vy, _vzlev)          # mean error removed only
    _, _, _p_scor = _hersbach(_vMU[:, _i], np.full(_VG_T, _srec), _vy, _vzlev)       # spread corrected only
    _kb = int((_vy < _vMU[:, _i]).sum())
    _vrows.append(dict({
        "id": _VG_IDS[_i], "group": _VG_GROUPS[_i],
        "color": _VG_COLOR[_i], "marker": _VG_MARKER[_i],
        "bias": float(_VG_BIAS[_i]), "error_scale_tau": float(_VG_TAU[_i]),
        "spread": float(_vSD[_i]),
        "design_rmse_of_mean": float(_VG_RMSE[_i]),
        "realized_rmse_of_mean": float(np.sqrt(np.mean(_err ** 2))),
        "realized_over_design_rmse": float(np.sqrt(np.mean(_err ** 2)) / _VG_RMSE[_i]),
        "realized_corr_residual_noise": float(np.corrcoef(
            _vu, (_vMU[:, _i] - _vmpub - _vx - _VG_BIAS[_i]) / _VG_TAU[_i])[0, 1]),
        "spread_over_rmse": float(_vSD[_i] / _VG_RMSE[_i]),
        "mean_CRPS_individual": float(_crps_gauss(_vMU[:, _i], _vSD[_i], _vy).mean()),
        "weight_full": float(_vw_full[_i]),
        "weight_human_only": float(_vw_hum[_i]),
        "weight_equal": float(_vw_eq[_i]),
        "crps_loo_resolve": _r["mean_CRPS"],
        "expected_CRPS_loo_new_quarters": _vg_ecrps(_w_i, _sub),
        "delta_loo": _r["mean_CRPS"] - _vcrps_full,
        "delta_cal": float(_rel_i - _vrel_full),
        "delta_res": float(_pot_i - _vpot_full),
        "delta_loo_hersbach": float(_ch_i - _vch_full),
        "count_below_median": _kb,
        "frac_below_median": _kb / float(_VG_T),
        "below_median_binomial_p": _vg_binom_p(_kb, _VG_T),
        "REL_i": float(_r_ind), "POT_i": float(_p_ind),
        "crps_i_hersbach": float(_c_ind),
        # the two columns the figure script plots
        "resolution": float(_vUNC - _p_ind), "reliability_error": float(_r_ind),
        "resolution_recalibrated": float(_vUNC - _p_rec),
        "resolution_mean_error_removed": float(_vUNC - _p_mrem),
        "resolution_spread_corrected": float(_vUNC - _p_scor),
    }, **_vd_rows[_i]))
_vF = pd.DataFrame(_vrows)
# counterfactual in the draw: the full panel re-solved on the seed-202603 record with
# M4's spread set to its design error (Algorithm 4.2, certified like every other solve)
_vcfr_sd = _vSD.copy()
_vcfr_sd[_VG_FOCUS] = _VG_RMSE[_VG_FOCUS]
_vcfr_w, _r = _vg_solve(_vy, _vMU, _vcfr_sd, _vall, "full panel, M4 spread set to its error")
_vg_solves.append(_r)
_vcf_ratio_draw = float(_vg_cf_ratio(_vw_full[_VG_FOCUS], _vcfr_w[_VG_FOCUS]))
_vcf_ratio_design = float(_vg_cf_ratio(_vdw_full[_VG_FOCUS], _vcf_w[_VG_FOCUS]))
# the resolution of each member on finer ensemble representations (L members,
# same convention), as plotted and after recalibration, in the draw and at the
# design level: the outer-bin excess of a narrow biased forecast is not an
# artifact of the coarse L = 19 grid, and the order the text states holds on every grid
_vresL = []
for _L in _VG_L_GRID:
    _zL = _norm_ppf(np.arange(1, _L + 1) / (_L + 1.0))
    for _i in _vall:
        _err = _vy - _vMU[:, _i]
        _sdd = math.sqrt(_VG_SIGMA_U ** 2 + _VG_TAU[_i] ** 2)
        _vresL.append({
            "L": _L, "id": _VG_IDS[_i],
            "resolution": _vUNC - _hersbach(_vMU[:, _i], np.full(_VG_T, _vSD[_i]), _vy, _zL)[2],
            "resolution_recalibrated": _vUNC - _hersbach(_vMU[:, _i] + _err.mean(),
                                                         np.full(_VG_T, float(_err.std())), _vy, _zL)[2],
            "resolution_design": _vUNC_design - _vg_pop_hersbach(-_VG_BIAS[_i], _sdd, _VG_SPREAD[_i], _zL)[2],
            "resolution_recalibrated_design": _vUNC_design - _vg_pop_hersbach(0.0, _sdd, _sdd, _zL)[2]})
_vResL = pd.DataFrame(_vresL)
# why M4 contributes: its leave-one-out channels at the same fitted weights (the
# full-panel weights and the weights without M4, both held fixed) with M4's bias
# removed or with its spread set to its design error, in the draw and at the design
# level.  The resolution channel comes from M4's reading of the signal and barely
# moves; the calibration channel nets the gain from narrowing the over-dispersed
# pool (positive once the bias is removed) against the cost of the bias
_vcf_rows = []
_vwl_f = np.nan_to_num(_vloo_w[_VG_FOCUS])
_vdwl_f = np.nan_to_num(_vd_loo_w[_VG_FOCUS])
for _set, _db, _sp in (("as generated", 0.0, float(_VG_SPREAD[_VG_FOCUS])),
                       ("M4 bias removed", -float(_VG_BIAS[_VG_FOCUS]), float(_VG_SPREAD[_VG_FOCUS])),
                       ("M4 spread set to its design error", 0.0, float(_VG_RMSE[_VG_FOCUS]))):
    _MUc = _vMU.copy()
    _MUc[:, _VG_FOCUS] += _db
    _SDc = _vSD.copy()
    _SDc[_VG_FOCUS] = _sp
    _bc = _VG_BIAS.copy()
    _bc[_VG_FOCUS] += _db
    _chf, _rf, _pf = _hersbach(_MUc @ _vw_full, np.full(_VG_T, float(_SDc @ _vw_full)), _vy, _vzlev)
    _chl, _rl, _pl = _hersbach(_MUc @ _vwl_f, np.full(_VG_T, float(_SDc @ _vwl_f)), _vy, _vzlev)
    _dchf, _drf, _dpf = _vg_pop_hersbach(*_vg_pool_error(_vdw_full, _vall, _bc, _SDc), _vzlev)
    _dchl, _drl, _dpl = _vg_pop_hersbach(*_vg_pool_error(_vdwl_f, _vall, _bc, _SDc), _vzlev)
    for _lev, _d, _dc, _dr in (("draw", _chl - _chf, _rl - _rf, _pl - _pf),
                               ("design", _dchl - _dchf, _drl - _drf, _dpl - _dpf)):
        _vcf_rows.append({"level": _lev, "setting": _set, "M4_bias": float(_bc[_VG_FOCUS]), "M4_spread": _sp,
                          "weights": "fitted full-panel and without-M4 weights of the %s, held fixed" % _lev,
                          "delta": float(_d), "delta_cal": float(_dc), "delta_res": float(_dr)})
_vCF = pd.DataFrame(_vcf_rows)
# robustness of the comparison to withholding any one member (Table tab:robustness)
_vloo_crps = _vF["crps_loo_resolve"].to_numpy(float)
_vloo_max = float(_vloo_crps.max())
_vloo_below_hum = int(np.sum(_vloo_crps < _vcrps_hum))
_vloo_new = _vF["expected_CRPS_loo_new_quarters"].to_numpy(float)
_vloo_new_max = float(_vloo_new.max())
_vloo_new_below_hum = int(np.sum(_vloo_new < _vnew_hum))

# ---- across seeds ----------------------------------------------------------
_vS0 = _vg_record_stats(_vy, _vMU, _vu)
# the epigraph solve reproduces Algorithm 4.2 on the seed-202603 record
assert abs(_vS0["ins_hm"] - _vcrps_full) < _VG_EPI_AGREE["score"], (_vS0["ins_hm"], _vcrps_full)
assert abs(_vS0["ins_h"] - _vcrps_hum) < _VG_EPI_AGREE["score"], (_vS0["ins_h"], _vcrps_hum)
assert max(abs(_vS0["w_" + k] - float(_vw_full[j])) for j, k in enumerate(_VG_IDS)) < _VG_EPI_AGREE["weights"]
assert np.allclose([_vS0["delta_" + k] for k in _VG_IDS], _vF["delta_loo_hersbach"], atol=10 * _VG_EPI_AGREE["score"])
assert abs(_vS0["w_M4_spread_at_error"] - float(_vcfr_w[_VG_FOCUS])) < _VG_EPI_AGREE["weights"]
_vseeds = list(range(_VG_RESEED["first"], _VG_RESEED["first"] + _VG_RESEED["count"]))
_vreseed = []
for _vseed in _vseeds:
    _yk, _MUk, _, _, _, _uk, _ = _vg_generate(_vseed)
    _vreseed.append(dict({"seed": _vseed}, **_vg_record_stats(_yk, _MUk, _uk)))
_vR = pd.DataFrame(_vreseed)
_vR0 = pd.DataFrame([dict({"seed": NEW_SEED}, **_vS0)])
for _k, _, _pred in _VG_CLAIMS:
    _vR[_k] = [int(bool(_pred(r))) for r in _vreseed]
    _vR0[_k] = int(bool(_pred(_vS0)))
_vR["all_core_claims"] = _vR[_VG_CORE].min(axis=1)
_vR0["all_core_claims"] = int(_vR0[_VG_CORE].min(axis=1).iloc[0])
_vRsum = pd.DataFrame([{"claim": _k, "definition": _dsc, "frequency": float(_vR[_k].mean()),
                        "count": int(_vR[_k].sum()), "seeds": len(_vR), "holds_in_seed_202603": int(_vR0[_k].iloc[0])}
                       for _k, _dsc, _ in _VG_CLAIMS]
                      + [{"claim": "all_core_claims", "definition": "every C1-C4 claim at once",
                          "frequency": float(_vR["all_core_claims"].mean()), "count": int(_vR["all_core_claims"].sum()),
                          "seeds": len(_vR), "holds_in_seed_202603": int(_vR0["all_core_claims"].iloc[0])}])
# the ratio of M4's weight to its weight at its error across the re-seeded records
# (median, and the percentile of the draw, both from the stored weight columns)
_vcf_ratios = _vg_cf_ratio(_vR["w_M4"], _vR["w_M4_spread_at_error"])
_vcf_ratio_median = float(np.median(_vcf_ratios))
_vcf_ratio_pct = float(100.0 * np.mean(_vcf_ratios <= float(_vg_cf_ratio(_vS0["w_M4"], _vS0["w_M4_spread_at_error"]))))
# where the seed-202603 draw lies in the across-seed distribution
_VG_TYPICAL_PLOTTED = ["REL_" + k for k in _VG_IDS] + ["RES_" + k for k in _VG_IDS]   # Figure 6.12
_VG_TYPICAL = (["sd_u", "UNC", "ins_hm", "ins_h", "ins_eq", "oos_hm", "oos_h", "max_loo_ins", "max_loo_oos", "W_M"]
               + ["w_" + k for k in _VG_IDS] + ["delta_" + k for k in _VG_IDS]
               + ["dcal_M4", "dres_M4", "below_M4", "below_M3", "RESrec_M4", "RESrec_M3", "RESrec_M1",
                  "rmse_cal6", "RES_finest_M4", "w_M4_spread_at_error"] + ["crps_" + k for k in _VG_IDS]
               + _VG_TYPICAL_PLOTTED + ["max_REL_calibrated_over_UNC"])
_vTyp = pd.DataFrame([{
    "statistic": _k, "seed_202603": float(_vS0[_k]),
    "across_seed_median": float(_vR[_k].median()),
    "across_seed_p10": float(_vR[_k].quantile(0.10)), "across_seed_p90": float(_vR[_k].quantile(0.90)),
    "percentile_of_seed_202603": float(100.0 * (_vR[_k] <= _vS0[_k]).mean()),
} for _k in _VG_TYPICAL])

# ---- console summary ------------------------------------------------------
print("   vignette (in-sample): human-only %.4f  equal-weight %.4f  optimized %.4f  W_H %.4f  W_M %.4f"
      % (_vcrps_hum, _vcrps_eq, _vcrps_full, _vWH, _vWM))
print("   vignette (new quarters): human-only %.4f  equal-weight %.4f  optimized %.4f"
      % (_vnew_hum, _vnew_eq, _vnew_full))
print("   design level: human-only %.4f  equal-weight %.4f  machine-only %.4f  optimized %.4f  w = %s"
      % (_vdE_hum, _vdE_eq, _vdE_mac, _vdE_full, np.round(_vdw_full, 3)))
for _r in _vrows:
    print("   %-3s w %.3f (design %.3f)  Delta %+.4f (cal %+.4f res %+.4f; design %+.4f = %+.4f %+.4f)  below %2d/%d  REL %.4f RES %.4f (recal %.4f)"
          % (_r["id"], _r["weight_full"], _r["weight_design"], _r["delta_loo_hersbach"], _r["delta_cal"], _r["delta_res"],
             _r["delta_design"], _r["delta_cal_design"], _r["delta_res_design"],
             _r["count_below_median"], _VG_T, _r["REL_i"], _r["resolution"], _r["resolution_recalibrated"]))
for _r in _vg_solves:
    print("   solve %-46s iters %6d  tol met %d  gap %+.1e  |dw| %.1e"
          % (_r["solve"], _r["iterations"], _r["tolerance_met"], _r["objective_gap"],
             _r["max_abs_weight_diff_vs_certificate"]))
for _r in _vd_solves:
    print("   solve %-46s objective %.10f  spread over starts %.1e"
          % (_r["solve"], _r["objective"], _r["objective_spread_over_starts"]))
for _r in _vg_solves:
    assert abs(_r["objective_gap"]) < _VG_CERT_GAP, "vignette solve %r not certified: gap %.2e" % (
        _r["solve"], _r["objective_gap"])
for _r in _vd_solves:
    assert _r["objective_spread_over_starts"] < _VG_CERT_GAP, "design solve %r: starts disagree" % _r["solve"]
print("   across %d re-seeded records:" % len(_vR))
for _, _r in _vRsum.iterrows():
    print("     %-50s %.3f  (seed 202603: %d)" % (_r["claim"], _r["frequency"], _r["holds_in_seed_202603"]))
for _L in _VG_L_GRID:
    _q = _vResL[_vResL["L"] == _L].set_index("id")
    print("   L = %4d  RES %s  recalibrated %s" % (_L, np.round(_q["resolution"].to_numpy(float), 3),
                                                   np.round(_q["resolution_recalibrated"].to_numpy(float), 3)))
for _, _r in _vCF.iterrows():
    print("   M4 channels at fixed weights, %-6s %-34s Delta %+.4f  cal %+.4f  res %+.4f"
          % (_r["level"], _r["setting"], _r["delta"], _r["delta_cal"], _r["delta_res"]))
print("   M3 resolution: plotted %.3f, mean error removed %.3f, spread corrected %.3f, recalibrated %.3f"
      % tuple(float(_vF.loc[_VG_IDS.index("M3"), _c]) for _c in ("resolution", "resolution_mean_error_removed",
                                               "resolution_spread_corrected", "resolution_recalibrated")))
print("   M4 weight with its spread at its error: draw %.4f -> %.4f (ratio %.3f, percentile %.1f), design %.4f -> %.4f (ratio %.3f), across-seed median ratio %.3f"
      % (_vw_full[_VG_FOCUS], _vcfr_w[_VG_FOCUS], _vcf_ratio_draw, _vcf_ratio_pct,
         _vdw_full[_VG_FOCUS], _vcf_w[_VG_FOCUS], _vcf_ratio_design, _vcf_ratio_median))
print("   epigraph SLSQP runs without success: %d" % len(_VG_EPI_FAILURES))
print("   seed-202603 percentiles: " + ", ".join("%s %.0f" % (r["statistic"], r["percentile_of_seed_202603"])
                                             for _, r in _vTyp.iterrows()))

# ---- figure geometry, derived from the solved points ----------------------
_vres = _vF["resolution"].to_numpy(float)
_vrel = _vF["reliability_error"].to_numpy(float)
_vxspan = float(_vres.max() - _vres.min())
_vx0 = round(float(_vres.min() - 0.12 * _vxspan), 3)
_vx1 = round(float(_vres.max() + 0.12 * _vxspan), 3)
_vy1 = round(1.42 * float(_vrel.max()), 3)
_vy0 = round(-0.04 * _vy1, 3)
_vcut = _VG_BAND * _vUNC               # visual guide: REL below a fifth of UNC
_vover = _vF[_vF["group"] == "overconfident machines"]
_vcall_x = float(_vover["resolution"].mean())
_vcall_y = float(_vover["reliability_error"].mean())


def _vpr(v, d):
    return round(float(v), d)


write("fig6-12_hm_calsharp.xlsx", {
    "config": cfg({
        "case": "vignette V -- human-machine forecast aggregation (Section 6.6)",
        "figures": "fig6-12_hm_calsharp",
        "labels": "fig:hm-calsharp; tab:hm-vignette; tab:robustness (vignette row)",
        "synthetic": 1,
        "seed": NEW_SEED,
        "generator": "build_workbooks.py, vignette block (_vg_generate / _vg_solve / _vg_pop_solve / _vg_record_stats / _hersbach)",
        "design_note": "R24 redesign: the design parameters were chosen so that the claims of Section 6.6 hold at the design level (closed-form expected CRPS and Hersbach split) and in most re-seeded records (sheet reseed_summary); the seed was not tuned, and the percentile of every printed seed-202603 statistic is in sheet typicality",
        "n_human": len(_vhum), "n_machine": len(_vmac), "horizon_T": _VG_T,
        "target_law": "stationary AR(1): y_0 ~ N(c, sigma^2/(1-phi^2)), y_t = c + phi (y_{t-1} - c) + x_t + u_t, x_t ~ N(0, sigma_x^2) leading signal observable at issue time, u_t ~ N(0, sigma_u^2) unobserved residual, sigma_x^2 + sigma_u^2 = sigma^2",
        "target_c": _VG_AR["c"], "target_phi": _VG_AR["phi"], "target_sigma": _VG_AR["sigma"],
        "target_sigma_x": _VG_SIGNAL["sigma_x"], "target_sigma_u": _VG_SIGNAL["sigma_u"],
        "target_sigma_x_rule": "sigma_x = sqrt(sigma^2 - sigma_u^2)",
        "presample_y0": float(_vpre0),
        "forecast_model": "mu_it = c + phi (y_{t-1} - c) + x_t + bias_i + tau_i xi_it, xi iid N(0,1), issue-time information only (y_t never enters a forecast); predictive N(mu_it, spread_i^2)",
        "spread_rule_humans_and_calibrated": "spread_i = sqrt(sigma_u^2 + bias_i^2 + tau_i^2), the design root-mean-square error of the mean (commensurate with the error)",
        "spread_rule_overconfident": "spread_i = %s for M3, M4 (far below the error)" % list(map(float, _VG_SPREAD_OVER)),
        "draw_order": "(1) one standard normal for the pre-sample value y_0, scaled to the stationary law; (2) T x 2 standard normals for (x_t, u_t), quarter-major; (3) T x 8 forecast-noise matrix, quarter-major",
        "pool": "quantile averaging eq:hm-qa; normal aggregate N(w.mu_t, (w.spread)^2)",
        "crps": "closed form for normals (def:crps), _crps_gauss with the exact normal cdf scipy.special.ndtr; the Algorithm 4.2 objective and subgradient (_vg_val_grad) use the same exact cdf, so solver and cross-check evaluate one objective",
        "objective": "mean CRPS + lambda_G Gini(w) + eps_ridge/2 ||w||^2 over the simplex (eq:hm-crps)",
        "lambda_G": _VG_LAM_G, "eps_ridge": _VG_EPS,
        "solver": "Algorithm 4.2 projected subgradient descent, analytic subgradient, equal-weight start, best iterate returned",
        "step_size_rule": "eta_r = eta0 / (1 + r / r0)",
        "eta0": _VG_SOLVER["eta0"], "r0": _VG_SOLVER["r0"],
        "tie_breaking_rule": "sign(0) = 0 for every vanishing weight difference in the Gini subgradient",
        "stopping_rule": "sup-norm change of the iterate < tol after at least r_min iterations; r_max is a guard",
        "tol": _VG_SOLVER["tol"], "r_min": _VG_SOLVER["r_min"], "r_max": _VG_SOLVER["r_max"],
        "certificate": "cross-check by scipy SLSQP on the equivalent smooth epigraph form of the same program (G(w) = (1/N) sum_{i<j} t_ij, t_ij >= |w_i - w_j|; convex, so a KKT point is the minimizer), from the equal-weight vector plus K Dirichlet(1) starts drawn from a fresh default_rng(seed) per solve, each end point scored by the exact objective; best value over starts; gap and weight difference in sheet 'solves'",
        "certificate_K": _VG_CERT["K"], "epigraph_ftol": _VG_EPI["ftol"], "epigraph_maxiter": _VG_EPI["maxiter"],
        "certificate_gap_bound": _VG_CERT_GAP,
        "comparators": "human-only re-solve (same program), equal-weight pool, eight leave-one-out re-solves; at the design level also the machine-only re-solve (same program over the four machine experts)",
        "loo_definition": "delta_loo_hersbach = nineteen-member step-function (Hersbach) score of the re-solved pool without i minus that of the full optimized pool, = delta_cal + delta_res exactly (the value the text reports); delta_loo = the exact-CRPS difference of def:loo (reference); delta_cal / delta_res are the Hersbach channels REL_-i - REL and POT_-i - POT",
        "loo_reduced_program": "the reduced program re-forms the Gini term over the N - 1 survivors (w-bar = 1/(N-1)), so it is not the full program restricted to the face w_i = 0 and a member with weight 0 can have a nonzero contribution",
        "scores_in_sample": "mean CRPS on the T quarters the weights were fitted to (Table tab:hm-vignette, first score column)",
        "scores_new_quarters": "expected CRPS on a new quarter of the same design of the fitted weights, closed form E CRPS = R sqrt(2/pi) exp(-m^2/(2R^2)) + m (2 Phi(m/R) - 1) - S/sqrt(pi), m = w.b, S = w.spread, R^2 = S^2 + sigma_u^2 + sum w_i^2 tau_i^2 (second score column)",
        "design_level": "the regularized program eq:hm-crps (expected CRPS + lambda_G Gini(w) + eps_ridge/2 ||w||^2, the same lambda_G and eps_ridge) with the expected CRPS on a new quarter in place of the in-sample mean, solved in the epigraph form by SLSQP from the equal-weight vector plus K Dirichlet(1) starts (sheet design_solves records the spread of the objective over the starts); design-level Hersbach split in closed form (_vg_pop_hersbach, normal partial moments, the same bin convention); design UNC = sigma_y / sqrt(pi), sigma_y = sigma / sqrt(1 - phi^2)",
        "design_counterfactual": "design full panel re-solved with M4's spread set to its design error (weight in aggregates)",
        "draw_counterfactual": "full panel re-solved on the seed-202603 record with M4's spread set to its design error, by Algorithm 4.2 with the same settings and certificate (sheet solves; weights column weight_full_M4_spread_at_error); the same re-solve on every re-seeded record by the epigraph solve (reseed column w_M4_spread_at_error)",
        "counterfactual_ratio_band": "%g,%g" % _VG_CF_RATIO_BAND,
        "counterfactual_ratio_rule": "ratio = w_M4 / w_M4 at its error; the text says M4 held, as its design dictates, about twice the weight it would have held at its error: the design-level ratio and the across-seed median ratio lie in counterfactual_ratio_band, the draw ratio inside typical_range_percent of the re-seeded records, and the weight falls in the draw, at the design level and in at least nine re-seeded records in ten (checked at the tail)",
        "reseed_first": _VG_RESEED["first"], "reseed_count": _VG_RESEED["count"],
        "reseed_rule": "seeds reseed_first .. reseed_first + reseed_count - 1, each regenerating the whole record by _vg_generate with the same draw order",
        "reseed_solver": "the same convex program in smooth epigraph form, G(w) = (1/N) sum_{i<j} t_ij with t_ij >= |w_i - w_j|, by SLSQP from the equal-weight vector (exact for a convex program); ftol %g, maxiter %d; on the seed-202603 record it reproduces Algorithm 4.2 to %g in the weights and %g in the scores" % (_VG_EPI["ftol"], _VG_EPI["maxiter"], _VG_EPI_AGREE["weights"], _VG_EPI_AGREE["score"]),
        "typical_statistics_printed": ",".join(_VG_TYPICAL_PRINTED),
        "typical_statistics_plotted": ",".join(_VG_TYPICAL_PLOTTED),
        "typical_statistics_disclosed": ",".join("%s:%s" % kv for kv in _VG_TYPICAL_DISCLOSED.items()),
        "typical_range_percent": "%g,%g" % _VG_TYPICAL_RANGE,
        "typicality_rule": "percentile of the seed-202603 value = 100 x share of the re-seeded records with a value <= it; every statistic in typical_statistics_printed must lie in typical_range_percent, every plotted coordinate in typical_statistics_plotted too unless it is listed in typical_statistics_disclosed (an atypical value the text discloses, with its side of the range), and every listed one must lie outside the range on its stated side (checked at the tail)",
        "tie_tolerance": _VG_TIE_TOL,
        "tie_rule": "weights count as level when they agree to within tie_tolerance (claim S_M1_M2_M4_weights_level; its frequency is recorded, it is not a claim of the text); S_M1_M2_M4_three_largest_weights requires each of the three to exceed every other weight by more than tie_tolerance, so that a tie with a fourth member does not count",
        "comparison_margin": _VG_CMP_EPS,
        "comparison_margin_rule": "the strict comparisons of the recorded contribution-order claims (S_three_largest_weights_contribute_most, S_contribution_order_M2_M1_M4_H2) carry this margin, so that each claim column evaluates alike on the in-memory values and on the workbook round-trip",
        "channel_counterfactual_rule": "sheet channel_counterfactuals: the leave-one-out channels of M4 (nineteen-member Hersbach split) recomputed at the fitted full-panel and without-M4 weights, held fixed, with M4's bias removed (its mean lowered by its design bias) and with M4's spread set to its design error sqrt(sigma_u^2 + b^2 + tau^2), in the draw and at the design level (closed form); the tail requires the resolution channel to move by less than %g of itself under either change, the calibration channel to be positive with the bias removed, and lower with the spread at its error than as generated" % _VG_CF_RES_TOL,
        "channel_counterfactual_resolution_tolerance": _VG_CF_RES_TOL,
        "resolution_single_corrections_definition": "resolution_mean_error_removed: UNC - POT with the member's mean error over the T quarters added to its mean and its spread kept; resolution_spread_corrected: UNC - POT with its mean kept and its spread set to the sd of its error; design level: the same with the design bias and the error sd sqrt(sigma_u^2 + tau_i^2)",
        "bias_test": "two-sided binomial test of the below-median count against one-half; the count exposes a bias when p < alpha",
        "bias_test_alpha": _VG_BIAS_ALPHA,
        "calibration_diagnostic": "frac_below_median = fraction of the T outcomes below the forecaster's median (its mean)",
        "hersbach_levels_L": _VG_L,
        "hersbach_convention": ("L ensemble members at the normal quantiles l/(L+1); interior bin i between "
                                "members i and i+1 at nominal level i/L; "
                                "outer bins by Hersbach's outlier rule: O_0 = freq(y < x_1), "
                                "g_0 = mean (x_1 - y)^+ / O_0, REL_0 = g_0 O_0^2, POT_0 = g_0 O_0 (1 - O_0); "
                                "the upper bin mirrored with O_L = freq(y > x_L), g_L = mean (y - x_L)^+ / O_L; "
                                "an empty outer bin contributes nothing (the binning "
                                "and outlier rule of the ex:hmpanel block, fig5-3_loo.xlsx config "
                                "hersbach_outlier_rule, at the design level with probabilities in place of "
                                "frequencies; the UNC definitions differ, see resolution_definition)"),
        "uncertainty_UNC": _vUNC,
        "uncertainty_UNC_design": _vUNC_design,
        "uncertainty_UNC_grid": _vUNC_grid,
        "uncertainty_UNC_grid_rule": ("potential score (_hersbach_X) of the empirical climatology represented on the "
                                      "same L-member grid, the outcome quantiles at the levels l/(L+1) by numpy linear "
                                      "interpolation, the convention of Section 4 and ex:hmpanel; recorded for "
                                      "comparison, not used by RES_i"),
        "resolution_definition": ("RES_i = UNC - POT_i, UNC = 0.5 mean |y - y'| over all T^2 ordered pairs of "
                                  "the T outcomes, self-pairs included (the continuous CRPS of the empirical "
                                  "climatology; Section 4 and ex:hmpanel form UNC on the L-member grid instead, "
                                  "uncertainty_UNC_grid, a shift UNC - uncertainty_UNC_grid common to every RES_i)"),
        "resolution_recalibrated_definition": "UNC - POT of the member's forecast after recalibration (its mean error over the T quarters added to its mean, its spread set to the sd of its error); design level: UNC_design - POT of N(0, e^2) error with spread e, e^2 = sigma_u^2 + tau_i^2",
        "resolution_by_L_levels": ",".join(str(v) for v in _VG_L_GRID),
        "resolution_by_L_rule": "sheet resolution_by_L: RES_i and the recalibrated RES_i of every member on the L-member representation (same convention) for each listed L, in the draw and at the design level; the L = 19 rows are the plotted values; the excess of the narrow biased M3, M4 over their recalibrated values shrinks as L grows (in part the same narrow-spread grid dependence as in Chapter 5) but persists, and M4 keeps the highest resolution on every grid (checked at the tail); the value of M4 at the finest L is printed (aggregates)",
        "design_conformity_rule": "the draw conforms to its design: realized/design root-mean-square error of every member's mean in [%g, %g], |corr(u, xi_i)| < %g / sqrt(T) for every member, and the realized rms of u inside typical_range_percent of the re-seeded records (checked at the tail)" % (_VG_CONFORM["rmse_ratio"] + (_VG_CONFORM["corr_z"],)),
        "design_conformity_rmse_ratio": "%g,%g" % _VG_CONFORM["rmse_ratio"],
        "design_conformity_corr_z": _VG_CONFORM["corr_z"],
        "epigraph_success_rule": "every SLSQP run of the re-seeding study must end with success (column epigraph_failures of sheet reseed, checked at the tail)",
        "x_label": "resolution $\\mathrm{RES}_i$ $\\rightarrow$",
        "y_label": "reliability error $\\mathrm{REL}_i$ (miscalibration) $\\rightarrow$",
        "x_min": _vx0, "x_max": _vx1, "y_min": _vy0, "y_max": _vy1,
        "well_calibrated_cut": _vcut,
        "well_calibrated_rule": "band top = %g * UNC (a visual guide, and the C4 criterion of the re-seeding study)" % _VG_BAND,
        "band_label": "well-calibrated region",
        "band_label_x": round(_vx0 + 0.58 * (_vx1 - _vx0), 4),
        "band_label_y": round(0.55 * _vcut, 4),
        "callout_text": "narrow and biased,\nmiscalibrated",
        "callout_xy_x": round(_vcall_x, 4), "callout_xy_y": round(_vcall_y, 4),
        "callout_text_x": round(_vcall_x - 0.30 * (_vx1 - _vx0), 4),
        "callout_text_y": round(_vcall_y - 0.10 * _vy1, 4),
    }),
    "forecasters": _vF,
    "outcomes": pd.DataFrame({"t": np.arange(1, _VG_T + 1), "y": _vy,
                              "conditional_mean_given_previous": _vmpub,
                              "leading_signal_x": _vx, "unobserved_residual_u": _vu}),
    "forecast_mu": pd.DataFrame(_vMU, columns=_VG_IDS),
    "weights": pd.DataFrame({"id": _VG_IDS, "weight_full": _vw_full,
                             "weight_human_only": _vw_hum, "weight_equal": _vw_eq,
                             "weight_design": _vdw_full,
                             "weight_design_human_only": np.concatenate([_vdw_hum, np.zeros(len(_vmac))]),
                             "weight_design_machine_only": np.concatenate([np.zeros(len(_vhum)), _vdw_mac]),
                             "weight_design_M4_spread_at_error": _vcf_w,
                             "weight_full_M4_spread_at_error": _vcfr_w}),
    "loo_weights": pd.DataFrame(_vloo_w, columns=_VG_IDS).assign(withheld=_VG_IDS)[["withheld"] + _VG_IDS],
    "design_loo_weights": pd.DataFrame(_vd_loo_w, columns=_VG_IDS).assign(withheld=_VG_IDS)[["withheld"] + _VG_IDS],
    "solves": pd.DataFrame(_vg_solves),
    "design_solves": pd.DataFrame(_vd_solves),
    "reseed": pd.concat([_vR0, _vR], ignore_index=True),
    "reseed_summary": _vRsum,
    "typicality": _vTyp,
    "resolution_by_L": _vResL,
    "channel_counterfactuals": _vCF,
    "aggregates": pd.DataFrame([
        {"quantity": "optimized human-only mean CRPS", "value": _vcrps_hum, "printed": _vpr(_vcrps_hum, 2)},
        {"quantity": "equal-weight human-machine mean CRPS", "value": _vcrps_eq, "printed": _vpr(_vcrps_eq, 2)},
        {"quantity": "optimized human-machine mean CRPS", "value": _vcrps_full, "printed": _vpr(_vcrps_full, 2)},
        {"quantity": "optimized human-only expected CRPS on new quarters", "value": _vnew_hum, "printed": _vpr(_vnew_hum, 2)},
        {"quantity": "equal-weight human-machine expected CRPS on new quarters", "value": _vnew_eq, "printed": _vpr(_vnew_eq, 2)},
        {"quantity": "optimized human-machine expected CRPS on new quarters", "value": _vnew_full, "printed": _vpr(_vnew_full, 2)},
        {"quantity": "design-level optimized human-only expected CRPS", "value": _vdE_hum, "printed": _vpr(_vdE_hum, 3)},
        {"quantity": "design-level optimized machine-only expected CRPS", "value": _vdE_mac, "printed": _vpr(_vdE_mac, 3)},
        {"quantity": "design-level equal-weight expected CRPS", "value": _vdE_eq, "printed": _vpr(_vdE_eq, 3)},
        {"quantity": "design-level optimized human-machine expected CRPS", "value": _vdE_full, "printed": _vpr(_vdE_full, 3)},
        {"quantity": "human weight share W_H", "value": _vWH, "printed": _vpr(_vWH, 2)},
        {"quantity": "machine weight share W_M", "value": _vWM, "printed": _vpr(_vWM, 2)},
        {"quantity": "design-level machine weight share W_M", "value": float(_vdw_full[_vmac].sum()),
         "printed": _vpr(_vdw_full[_vmac].sum(), 2)},
        {"quantity": "design-level weight of M4 with its spread set to its error", "value": float(_vcf_w[_VG_FOCUS]),
         "printed": _vpr(_vcf_w[_VG_FOCUS], 2)},
        {"quantity": "weight of M4 with its spread set to its error (draw)", "value": float(_vcfr_w[_VG_FOCUS]),
         "printed": _vpr(_vcfr_w[_VG_FOCUS], 2)},
        {"quantity": "ratio of M4's weight to its weight at its error (design level)", "value": _vcf_ratio_design,
         "printed": _vpr(_vcf_ratio_design, 2)},
        {"quantity": "ratio of M4's weight to its weight at its error (draw)", "value": _vcf_ratio_draw,
         "printed": _vpr(_vcf_ratio_draw, 2)},
        {"quantity": "across-seed median ratio of M4's weight to its weight at its error", "value": _vcf_ratio_median,
         "printed": _vpr(_vcf_ratio_median, 2)},
        {"quantity": "percentile of the draw's ratio of M4's weight to its weight at its error", "value": _vcf_ratio_pct,
         "printed": _vpr(_vcf_ratio_pct, 1)},
        {"quantity": "full-pool Hersbach REL", "value": _vrel_full, "printed": _vpr(_vrel_full, 3)},
        {"quantity": "full-pool Hersbach POT", "value": _vpot_full, "printed": _vpr(_vpot_full, 3)},
        {"quantity": "full-pool Hersbach CRPS", "value": _vch_full, "printed": _vpr(_vch_full, 3)},
        {"quantity": "uncertainty UNC (sample climatology)", "value": _vUNC, "printed": _vpr(_vUNC, 2)},
        {"quantity": "uncertainty UNC (design)", "value": _vUNC_design, "printed": _vpr(_vUNC_design, 2)},
        {"quantity": "largest leave-one-out mean CRPS", "value": _vloo_max, "printed": _vpr(_vloo_max, 2)},
        {"quantity": "leave-one-out re-solves scoring below the human-only aggregate", "value": _vloo_below_hum,
         "printed": _vloo_below_hum},
        {"quantity": "largest leave-one-out expected CRPS on new quarters", "value": _vloo_new_max,
         "printed": _vpr(_vloo_new_max, 2)},
        {"quantity": "leave-one-out pools scoring below the human-only pool on new quarters",
         "value": _vloo_new_below_hum, "printed": _vloo_new_below_hum},
        {"quantity": "equal-weight pool of the six calibrated members, spread", "value": _vcal6_spread,
         "printed": _vpr(_vcal6_spread, 2)},
        {"quantity": "equal-weight pool of the six calibrated members, design root-mean-square error of the mean",
         "value": _vcal6_rmse_design, "printed": _vpr(_vcal6_rmse_design, 2)},
        {"quantity": "equal-weight pool of the six calibrated members, realized root-mean-square error of the mean",
         "value": _vcal6_rmse, "printed": _vpr(_vcal6_rmse, 2)},
        {"quantity": "realized rms of the unobserved residual u", "value": _vS0["sd_u"], "printed": _vpr(_vS0["sd_u"], 2)},
        {"quantity": "M4 plotted resolution on the finest representation (L = %d)" % _VG_L_GRID[-1],
         "value": _vS0["RES_finest_M4"], "printed": _vpr(_vS0["RES_finest_M4"], 2)},
        {"quantity": "largest reliability error of the six calibrated members over UNC",
         "value": _vS0["max_REL_calibrated_over_UNC"], "printed": _vpr(_vS0["max_REL_calibrated_over_UNC"], 3)},
        {"quantity": "across-seed median of the largest reliability error of the six calibrated members over UNC",
         "value": float(_vR["max_REL_calibrated_over_UNC"].median()),
         "printed": _vpr(_vR["max_REL_calibrated_over_UNC"].median(), 3)},
        {"quantity": "re-seeded records", "value": len(_vR), "printed": len(_vR)},
    ] + [{"quantity": "re-seeding count: " + _k, "value": int(_vR[_k].sum()), "printed": int(_vR[_k].sum())}
         for _k, _, _ in _VG_CLAIMS]
      + [{"quantity": "re-seeding count: all_core_claims", "value": int(_vR["all_core_claims"].sum()),
          "printed": int(_vR["all_core_claims"].sum())}]),
})


# =====================================================================
# FIGURE 5.4  Expert-dependence network of Example ex:corr-lineage
#             (fig5-4_depnet)
# A STIPULATED error-correlation matrix, not estimated from any data: a
# block of five weakly correlated human experts (pairwise r drawn
# U(0.05, 0.20)), a lineage of five fine-tuned variants of one base model
# (0.85 + U(-0.03, 0.03), then shifted so the block mean is 0.85), and
# weak human-lineage correlations U(0.02, 0.10), each rounded to 3 dp.  The
# human block is NOT the five humans of ex:hmpanel, whose errors correlate
# at about 0.92 under the Section 5.6 instrument (fig5-3_loo.xlsx sheet
# error_dependence); the book presents it as a separately stipulated block.
# Its equal-weight effective number alone (neff_human_equalweight) and with
# the lineage (neff_panel_equalweight) are the counts Section 5.6 compares
# with the ex:hmpanel humans-alone and full-panel counts.
# Node layout is deterministic (coordinates stored).  Every pair is drawn
# as an edge whose width (base + slope x r, in pt) and opacity (r clipped
# to [floor, cap]) increase with r; the encoding constants are config keys
# read by the drawing script, so the weak edges stay faintly visible and
# the lineage block reads as a dense clique.
# =====================================================================
_rngd = np.random.default_rng(NEW_SEED)
# node IDs A1-A5 (stipulated humans) and V1-V5 (lineage variants) are chosen
# so that they do not collide with the H1-H5 / M1-M5 experts of ex:hmpanel
# (Figures 5.2 and 5.3), which are a different panel
_dep_ids = ["A1", "A2", "A3", "A4", "A5", "V1", "V2", "V3", "V4", "V5"]
_dep_grp = ["human"] * 5 + ["lineage"] * 5
_R = np.eye(10)
for _i in range(10):
    for _j in range(_i + 1, 10):
        if _dep_grp[_i] == "lineage" and _dep_grp[_j] == "lineage":
            r = 0.85 + _rngd.uniform(-0.03, 0.03)
        elif _dep_grp[_i] == "human" and _dep_grp[_j] == "human":
            r = _rngd.uniform(0.05, 0.20)
        else:
            r = _rngd.uniform(0.02, 0.10)
        _R[_i, _j] = _R[_j, _i] = round(r, 3)
# exact block mean 0.85 for the lineage clique
_lin = [i for i, gse in enumerate(_dep_grp) if gse == "lineage"]
_pairs = [(i, j) for i in _lin for j in _lin if i < j]
_off = 0.85 - np.mean([_R[i, j] for i, j in _pairs])
for _i, _j in _pairs:
    _R[_i, _j] = _R[_j, _i] = round(_R[_i, _j] + _off, 3)

# deterministic layout: humans on a wide left pentagon, lineage on a
# tight right pentagon (its density is the visual point)
_dep_nodes = []
for _k in range(5):
    _ang = math.pi / 2 + 2 * math.pi * _k / 5
    _dep_nodes.append({"id": _dep_ids[_k], "group": "human",
                       "x": -1.75 + 1.05 * math.cos(_ang),
                       "y": 1.05 * math.sin(_ang)})
for _k in range(5):
    _ang = math.pi / 2 + 2 * math.pi * _k / 5
    _dep_nodes.append({"id": _dep_ids[5 + _k], "group": "lineage",
                       "x": 1.75 + 0.52 * math.cos(_ang),
                       "y": 0.52 * math.sin(_ang)})

_w_eq = np.full(10, 0.1)
_neff_panel = float(1.0 / (_w_eq @ _R @ _w_eq))
_wl = np.full(5, 0.2)
_Rl = _R[np.ix_(_lin, _lin)]
_neff_lineage = float(1.0 / (_wl @ _Rl @ _wl))
_hum_d = [i for i, gse in enumerate(_dep_grp) if gse == "human"]
# the stipulated humans alone at equal weights; Section 5.6 compares the
# count with and without the lineage (adding it at equal weights lowers it)
_neff_human_d = float(1.0 / (_wl @ _R[np.ix_(_hum_d, _hum_d)] @ _wl))
_dep_hh = np.array([_R[i, j] for i in _hum_d for j in _hum_d if i < j])
_dep_hl = np.array([_R[i, j] for i in _hum_d for j in _lin])
_dep_ll = np.array([_R[i, j] for i, j in _pairs])
write("fig5-4_depnet.xlsx", {
    "config": cfg({
        "figure": "5.4", "label": "fig:depnet",
        "example": "ex:corr-lineage (Section 5.6)",
        "synthetic": 1, "seed": NEW_SEED,
        "matrix_status": ("stipulated error-correlation matrix (not estimated from data); the human block "
                          "is a separately stipulated weakly correlated block, not the humans of ex:hmpanel"),
        "human_block_reading": ("Section 5.6 reads the weak human block as experts who each err mainly "
                                "through components of their own (the stipulation, not an estimate)"),
        "node_labels": ("A1-A5 the stipulated human experts, V1-V5 the fine-tuned variants of the lineage; "
                        "chosen so as not to collide with the H1-H5 and M1-M5 experts of ex:hmpanel "
                        "(Figures 5.2 and 5.3), which form a different panel"),
        "draw_order": ("default_rng(202603), upper triangle row by row: lineage pairs 0.85 + U(-0.03, 0.03), "
                       "human pairs U(0.05, 0.20), human-lineage pairs U(0.02, 0.10), each rounded to 3 dp; "
                       "lineage block then shifted to mean 0.85 and re-rounded"),
        "lineage_mean_corr": 0.85,
        "lineage_block_min": float(_dep_ll.min()), "lineage_block_max": float(_dep_ll.max()),
        "human_block_min": float(_dep_hh.min()), "human_block_max": float(_dep_hh.max()),
        "human_block_printed": "%.2f to %.2f" % (_dep_hh.min(), _dep_hh.max()),
        "cross_block_min": float(_dep_hl.min()), "cross_block_max": float(_dep_hl.max()),
        "cross_block_printed": "%.2f to %.2f" % (_dep_hl.min(), _dep_hl.max()),
        "min_eigenvalue": float(np.linalg.eigvalsh(_R).min()),
        "neff_lineage_equalweight": _neff_lineage,
        "neff_panel_equalweight": _neff_panel,
        "neff_human_equalweight": _neff_human_d,
        "neff_human_printed": "%.2f" % _neff_human_d,
        "neff_panel_printed": "%.2f" % _neff_panel,
        # edge encoding, read by the drawing script: width = base + slope x r
        # (pt), opacity = r clipped to [floor, cap]; both increase with r
        "edge_width_base_pt": 0.5, "edge_width_slope_pt": 3.0,
        "edge_alpha_floor": 0.06, "edge_alpha_cap": 0.92,
        "edge_note": "edge width and opacity increase with the pairwise error correlation",
        "human_note": ("weakly correlated human experts\n(stipulated block, pairwise $r$ from %.2f to %.2f)"
                       % (_dep_hh.min(), _dep_hh.max())),
        "human_legend": "human experts (weakly correlated)",
        "lineage_legend": "fine-tuned variants of one base model",
        "worked_examples": _WX_CFG % 5,
    }),
    "nodes": pd.DataFrame(_dep_nodes),
    "correlation": pd.DataFrame(_R, columns=_dep_ids).assign(id=_dep_ids)[
        ["id"] + _dep_ids],
    # not read by the figure script: the hand-worked examples of Chapter 5
    "worked_examples": _wx_sheet(5),
})


# =====================================================================
# CASE STUDY IV (Section 6.5, rail-transit station accessibility)
#   Urban rail-transit station accessibility with stakeholder equity
#   (mirrors chen2024railtransit at the level of MECHANISM only: equity
#   weighting reorders priorities).  One workbook backs the whole case.
#   -> Figure fig6-11_rt_equity (fig:rt-equity), Table tab:rt-results
#
#   Everything the book prints for this case is SOLVED here from the
#   stated inputs (tab:rt-stations, rho, vartheta, sigma, seed 202603)
#   by the stated formulas:
#     overlap      def:consensus, closed form for equal-variance normals
#                  over the real line, 2 Phi(-|mu_i - m| / (2 sigma))
#     consensus    the mean overlap (a positive rescaling of the
#                  cumulative overlap; single objective, same argmax)
#     equity       Sen's abbreviated welfare form Ubar (1 - G(U)) of the
#                  power-normalized Fehr-Schmidt utilities (kappa = 2)
#     confidence   constant over the simplex (common sigma) -> dropped
#     solver       multi-start SQP (scipy SLSQP) over the simplex from the
#                  uniform vector and K = 40 Dirichlet(1) starts, best
#                  objective retained; certified against a dense 1-D scan
#                  of the aggregate mean (every objective depends on w only
#                  through m = w.mu)
#     reporting    the representative of the optimal level set nearest to
#                  the uniform vector (convex QP, KKT cross-check)
#     level sets   what the optimal level set itself fixes, as opposed to
#                  the tie-break: the LP range (HiGHS, cross-checked by
#                  vertex enumeration) of the advocates' weight share over
#                  {w in simplex : w.mu = m*}, per station and regime; the
#                  overlaps and utilities depend on w only through m* and
#                  are therefore properties of the level set
#     consensus    with a common sigma each overlap is convex in m on either
#                  side of mu_i, so the consensus optimum is one member's
#                  own mean (recorded and checked per station)
#     near-ties    the scan also records every other local maximum of
#                  every solve; a local maximum whose RELATIVE gap to the
#                  optimum, (v* - v) / v*, is below _RT_NEAR_TIE_TOL is a
#                  near-tie (scale-free, so the count does not depend on
#                  whether consensus is the mean or the cumulative overlap);
#                  every one is reported, and every combination of reported
#                  and near-tied optima is re-evaluated for the statements
#                  the text makes
#     overlaps     the consensus optimum maximizes the mean overlap, so the
#                  equity optimum lowers it wherever the two differ; the
#                  Gini of the overlap profile is recorded beside the level
#                  so that a leveling-down move in the sense of
#                  prop:consensus-fairness(iii) (Gini index) is visible, on
#                  the reported optima and on every combination of
#                  reported and near-tied optima
#     priority     ascending collective mean
#     robustness   40 re-seeded panels and a 22-point Fehr-Schmidt grid,
#                  solved by the certified 1-D scan
# =====================================================================
print("\n-- generating and solving Case IV (fig6-11_rt_case.xlsx) ...")
import itertools as _rt_itertools
from scipy.optimize import minimize, minimize_scalar, least_squares, linprog
from scipy.special import ndtr

_RT_STATIONS = [
    {"id": "S1",  "name": "Central Exchange",      "ops": 8.0, "pas": 7.8, "adv": 7.6},
    {"id": "S2",  "name": "Civic Plaza",           "ops": 7.4, "pas": 7.2, "adv": 7.0},
    {"id": "S3",  "name": "Riverside",             "ops": 6.7, "pas": 6.6, "adv": 6.4},
    {"id": "S4",  "name": "Northgate Interchange", "ops": 6.5, "pas": 6.3, "adv": 4.5},
    {"id": "S5",  "name": "University Park",       "ops": 6.0, "pas": 5.9, "adv": 5.8},
    {"id": "S6",  "name": "Harbor View",           "ops": 5.4, "pas": 5.3, "adv": 5.2},
    {"id": "S7",  "name": "Meadow East",           "ops": 4.9, "pas": 4.8, "adv": 4.7},
    {"id": "S8",  "name": "Old Town",              "ops": 4.4, "pas": 4.3, "adv": 4.2},
    {"id": "S9",  "name": "Westfield Depot",       "ops": 3.9, "pas": 3.8, "adv": 3.6},
    {"id": "S10", "name": "Airport South",         "ops": 3.3, "pas": 3.2, "adv": 3.1},
]
_RT_CONST = [("operators", "ops", 8), ("passenger representatives", "pas", 10),
             ("accessibility advocates", "adv", 6)]
_RT_RHO, _RT_THETA = 0.4, 0.2            # Fehr-Schmidt rho_i, vartheta_i (chapter 6)
_RT_LENIENCY_SD, _RT_NOISE_SD = 0.15, 0.25
_RT_SIGMA = 1.0                          # the common reported dispersion
# solver settings (all recorded in the config sheet)
_RT_K_STARTS = 40                        # Dirichlet(1) starts in addition to the uniform vector
_RT_SQP_FTOL, _RT_SQP_MAXITER = 1e-12, 500
_RT_SCAN_POINTS, _RT_SCAN_XATOL = 8001, 1e-12
_RT_CERT_OBJ_TOL, _RT_CERT_M_TOL = 1e-8, 1e-4   # observed SQP-vs-scan gaps are ~1e-6
_RT_PRINT_DP = 2                          # the book prints means to 2 dp; SQP and scan must agree there
_RT_GAP_THRESHOLD = 0.2                   # the prose bounds every uncontested gap but the largest below this
_RT_NEAR_TIE_TOL = 1e-4                   # a local maximum within this RELATIVE gap of the optimum is a near-tie
_RT_CLOSEST_TIE_BOUND = 2e-5              # the text bounds the relative gaps of the closest solve's near-ties by this
_RT_WARN_THRESHOLD = 0.5                  # early-warning rule: investigate stations whose regime gap exceeds this
_RT_RESEED_N = 40                        # re-seeded panels: seeds 202604 .. 202643
_RT_FS_RHO = (0.2, 0.3, 0.4, 0.5, 0.6)
_RT_FS_THETA = (0.0, 0.1, 0.2, 0.3, 0.4)  # combinations with vartheta <= rho

_rt_N = sum(c[2] for c in _RT_CONST)
_rt_cid = sum([[c[0]] * c[2] for c in _RT_CONST], [])
_rt_ckey = sum([[c[1]] * c[2] for c in _RT_CONST], [])
_rt_adv = np.array([k == "adv" for k in _rt_ckey])


def _rt_generate_panel(rng):
    """Draw order (recorded in the config sheet): first the 24 leniency
    draws N(0, 0.15^2), one per stakeholder; then, station by station
    (S1 .. S10) and within a station stakeholder by stakeholder (1 .. 24),
    one scalar noise draw N(0, 0.25^2).  Every mean is clipped to [0, 10]
    (the clip does not bind under seed 202603; the config records this)."""
    len_ = rng.normal(0.0, _RT_LENIENCY_SD, _rt_N)
    MU = np.empty((_rt_N, 10))
    for j, st in enumerate(_RT_STATIONS):
        for i in range(_rt_N):
            MU[i, j] = np.clip(st[_rt_ckey[i]] + len_[i]
                               + rng.normal(0.0, _RT_NOISE_SD), 0.0, 10.0)
    return MU, len_


_rngr = np.random.default_rng(NEW_SEED)
_rt_MU, _rt_len = _rt_generate_panel(_rngr)
_rt_clip_binding = bool((_rt_MU <= 0.0).any() or (_rt_MU >= 10.0).any())
# the SQP starts are the next draws of the SAME stream: the uniform vector
# plus K Dirichlet(1, ..., 1) vectors, drawn once and shared by every
# station and both regimes
_rt_starts = [np.full(_rt_N, 1.0 / _rt_N)] + \
    [_rngr.dirichlet(np.ones(_rt_N)) for _ in range(_RT_K_STARTS)]
# a common reported dispersion: the aggregate sd under QA, sum w sigma, is
# the same for every weight vector, so the confidence objective is constant
# over the simplex and is DROPPED from both programs (chapter 6, Section
# 6.5); this also frees the level sets of the aggregate mean, along which
# both regimes' reported representatives are selected (_rt_project_uniform)
_rt_SD = np.full(_rt_N, _RT_SIGMA)


def _rt_overlap(mus, m):
    """The overlap of def:consensus for equal-variance normals, in closed
    form over the real line: int min{f_i, f_QA} dx = 2 Phi(-|mu_i - m| / (2 sigma)).
    m: scalar or (M,) -> (M, n)."""
    m = np.atleast_1d(np.asarray(m, float))
    return 2.0 * ndtr(-np.abs(mus[None, :] - m[:, None]) / (2.0 * _RT_SIGMA))


def _rt_objective(m, mus, regime, rho=_RT_RHO, theta=_RT_THETA):
    """Regime objective as a function of the aggregate mean m (vectorized).
    'eff': the consensus objective, the mean overlap.  'eq4': Sen's
    abbreviated welfare form Ubar (1 - G(U)) of the power-normalized
    Fehr-Schmidt utilities of the overlaps (kappa = 2, via _fs_phi) with
    G the I_4 Gini of the utility profile.  No confidence term."""
    A = _rt_overlap(mus, m)
    if regime == "eff":
        return A.mean(axis=1)
    _, Ut = _fs_phi(A, rho, theta)
    gini = np.array([_gini_vec(u) for u in Ut])
    return Ut.mean(axis=1) * (1.0 - gini)


def _rt_scan(mus, regime, rho=_RT_RHO, theta=_RT_THETA, npts=_RT_SCAN_POINTS,
             runner_up=False):
    """Certificate of global optimality.  Because sigma is common, every
    objective depends on w only through m = w.mu, and m ranges exactly over
    [min mu, max mu] on the simplex, so a dense scan of m (npts points) with
    a bounded refinement around the best grid point is a global solver of
    the one-dimensional problem.  Returns (m*, value, number of local maxima),
    and with runner_up=True also the list of the OTHER interior local maxima
    of the grid, each refined in the same way, distinct in m, as (m, value)
    pairs sorted by value descending (empty when the optimum is the only
    local maximum); its first entry is the runner-up."""
    lo, hi = float(mus.min()), float(mus.max())
    grid = np.linspace(lo, hi, npts)
    vals = np.concatenate([_rt_objective(grid[k:k + 1024], mus, regime, rho, theta)
                           for k in range(0, npts, 1024)])

    def refine(k):
        a, b = grid[max(k - 2, 0)], grid[min(k + 2, npts - 1)]
        res = minimize_scalar(lambda m: -float(_rt_objective([m], mus, regime, rho, theta)[0]),
                              bounds=(a, b), method="bounded", options={"xatol": _RT_SCAN_XATOL})
        if -res.fun >= vals[k]:
            return float(res.x), float(-res.fun)
        return float(grid[k]), float(vals[k])
    k = int(np.argmax(vals))
    m_star, v_star = refine(k)
    locs = [i for i in range(1, npts - 1) if vals[i] > vals[i - 1] and vals[i] >= vals[i + 1]]
    n_loc = len(locs)
    if not runner_up:
        return m_star, v_star, n_loc
    others = []
    for i in locs:
        if abs(i - k) <= 2:          # the optimum's own grid peak
            continue
        m2, v2 = refine(i)
        if abs(m2 - m_star) > 1e-6 and all(abs(m2 - o[0]) > 1e-6 for o in others):
            others.append((m2, v2))
    others.sort(key=lambda o: -o[1])
    return m_star, v_star, n_loc, others


def _rt_rel_gap(v_opt, v):
    """Relative gap of a local maximum to the optimum, (v* - v) / v*; the
    near-tie rule is scale-free (v* > 0 for both regimes)."""
    return (v_opt - v) / v_opt


def _rt_share_range(mus, m_star, mask):
    """What the optimal level set fixes about a group's weight share: the
    min and max of sum_{i in mask} w_i over {w in simplex : w.mu = m*}, two
    LPs solved by HiGHS and cross-checked by enumerating the vertices of the
    set (single members with mu_i = m*, and pairs of members straddling m*)."""
    n = mus.size
    c = mask.astype(float)
    A_eq, b_eq = np.vstack([np.ones(n), mus]), np.array([1.0, m_star])
    lo = linprog(c, A_eq=A_eq, b_eq=b_eq, bounds=[(0.0, None)] * n, method="highs")
    hi = linprog(-c, A_eq=A_eq, b_eq=b_eq, bounds=[(0.0, None)] * n, method="highs")
    assert lo.status == 0 and hi.status == 0, "share-range LP failed at m* %.6f" % m_star
    vx = [float(c[i]) for i in range(n) if abs(mus[i] - m_star) <= 1e-9]
    for i in range(n):
        for j in range(n):
            if mus[i] < m_star < mus[j]:
                t = (mus[j] - m_star) / (mus[j] - mus[i])
                vx.append(float(t * c[i] + (1.0 - t) * c[j]))
    assert abs(min(vx) - lo.fun) < 1e-7 and abs(max(vx) + hi.fun) < 1e-7, \
        "share-range LP disagrees with the vertex enumeration at m* %.6f" % m_star
    return float(lo.fun), float(-hi.fun)


def _rt_sqp(mus, regime):
    """The solver the book states: multi-start sequential quadratic
    programming (scipy SLSQP, finite-difference gradients) over the simplex
    {w >= 0, sum w = 1} from the uniform vector and the K Dirichlet(1)
    starts; every start's terminal point is accepted whatever its exit
    status (the objective is nonsmooth at the Gini kinks) and the best
    objective value over the starts is retained."""
    n = mus.size
    f = lambda w: -float(_rt_objective([float(w @ mus)], mus, regime)[0])
    cons = ({"type": "eq", "fun": lambda w: w.sum() - 1.0,
             "jac": lambda w: np.ones(n)},)
    best = None
    finals = []
    for w0 in _rt_starts:
        r = minimize(f, w0, method="SLSQP", bounds=[(0.0, 1.0)] * n, constraints=cons,
                     options={"ftol": _RT_SQP_FTOL, "maxiter": _RT_SQP_MAXITER})
        w = np.clip(r.x, 0.0, None)
        w = w / w.sum()
        v = -f(w)
        finals.append((v, float(w @ mus)))
        if best is None or v > best[1]:
            best = (w, v)
    return best[0], best[1], finals


def _rt_project_uniform(mus, m_star):
    """The reporting rule of chapter 6: the representative of the optimal
    level set {w in simplex : w.mu = m*} closest to the uniform vector
    (min sum w_i^2), a convex QP solved by SLSQP and cross-checked against
    its KKT form w_i = max(0, 1/n + a + b mu_i)."""
    n = mus.size
    u = np.full(n, 1.0 / n)
    cons = ({"type": "eq", "fun": lambda w: w.sum() - 1.0, "jac": lambda w: np.ones(n)},
            {"type": "eq", "fun": lambda w: w @ mus - m_star, "jac": lambda w: mus})
    r = minimize(lambda w: float(((w - u) ** 2).sum()), u, jac=lambda w: 2.0 * (w - u),
                 method="SLSQP", bounds=[(0.0, 1.0)] * n, constraints=cons,
                 options={"ftol": 1e-16, "maxiter": 1000})
    w = np.clip(r.x, 0.0, None)

    def kkt(p):
        ww = np.clip(1.0 / n + p[0] + p[1] * mus, 0.0, None)
        return [ww.sum() - 1.0, ww @ mus - m_star]
    ls = least_squares(kkt, [0.0, 0.0], xtol=1e-15, ftol=1e-15, gtol=1e-15)
    wk = np.clip(1.0 / n + ls.x[0] + ls.x[1] * mus, 0.0, None)
    kkt_err = float(np.abs(w - wk).max())
    assert abs(w.sum() - 1.0) < 1e-9 and abs(float(w @ mus) - m_star) < 1e-9, \
        "projection left the level set (m* %.6f)" % m_star
    assert kkt_err < 1e-6, "projection disagrees with its KKT form by %.2e" % kkt_err
    return w, kkt_err


def _rt_solve(j, regime):
    """Per-station solve of one regime: multi-start SQP, certification by
    the 1-D scan (the build aborts if the SQP best is not the global
    optimum), then the nearest-to-uniform representative and its
    diagnostics.  regime 'eff' = consensus only; 'eq4' = Sen welfare."""
    mus = _rt_MU[:, j]
    w_sqp, v_sqp, finals = _rt_sqp(mus, regime)
    m_sqp = float(w_sqp @ mus)
    m_scan, v_scan, n_loc, others = _rt_scan(mus, regime, runner_up=True)
    second = others[0] if others else None
    assert abs(v_sqp - v_scan) <= _RT_CERT_OBJ_TOL and abs(m_sqp - m_scan) <= _RT_CERT_M_TOL, \
        "%s %s: multi-start SQP best %.8f at m %.5f is not the 1-D global optimum %.8f at m %.5f" % (
            _RT_STATIONS[j]["id"], regime, v_sqp, m_sqp, v_scan, m_scan)
    assert round(m_sqp, _RT_PRINT_DP) == round(m_scan, _RT_PRINT_DP), \
        "%s %s: SQP mean %.6f and scan mean %.6f print differently at %d dp" % (
            _RT_STATIONS[j]["id"], regime, m_sqp, m_scan, _RT_PRINT_DP)
    w, kkt_err = _rt_project_uniform(mus, m_sqp)
    A = _rt_overlap(mus, m_sqp)[0]
    _, Ut = _fs_phi(A[None, :], _RT_RHO, _RT_THETA)
    n_at_best = int(sum(1 for v, _ in finals if v >= v_sqp - 1e-7))
    # what the level set itself fixes about the advocates' share (LP range)
    share_lo, share_hi = _rt_share_range(mus, m_sqp, _rt_adv)
    assert share_lo - 1e-9 <= float(w[_rt_adv].sum()) <= share_hi + 1e-9
    # the runner-up local maximum and the tie-break representative there
    if second is not None:
        w2, _ = _rt_project_uniform(mus, second[0])
        ru = {"ru_m": second[0], "ru_obj": second[1], "ru_gap": v_scan - second[1],
              "ru_rel_gap": _rt_rel_gap(v_scan, second[1]),
              "ru_adv_share": float(w2[_rt_adv].sum()), "ru_gini_w": _gini_vec(w2)}
    else:
        ru = {"ru_m": np.nan, "ru_obj": np.nan, "ru_gap": np.nan, "ru_rel_gap": np.nan,
              "ru_adv_share": np.nan, "ru_gini_w": np.nan}
    # every near-tied local maximum (relative gap below the tolerance), each
    # with the tie-break representative of its own level set
    ties = []
    for m2, v2 in others:
        if _rt_rel_gap(v_scan, v2) < _RT_NEAR_TIE_TOL:
            w2, _ = _rt_project_uniform(mus, m2)
            ties.append({"m": m2, "obj": v2, "gap": v_scan - v2, "rel_gap": _rt_rel_gap(v_scan, v2),
                         "adv_share": float(w2[_rt_adv].sum()), "gini_w": _gini_vec(w2)})
    ru["ties"] = ties
    # the closest local maximum that is not a near-tie (the margin of the rule)
    ru["first_non_tie_rel_gap"] = (_rt_rel_gap(v_scan, others[len(ties)][1])
                                   if len(others) > len(ties) else np.nan)
    # the consensus optimum is one member's own mean (piecewise convexity)
    k_mem = int(np.argmin(np.abs(mus - m_sqp)))
    return dict({"w": w, "mu": m_sqp, "obj": v_sqp, "sd": float(w @ _rt_SD),
                 "Abar": float(A.mean()), "Amin": float(A.min()),
                 "Abar_adv": float(A[_rt_adv].mean()), "gini_A": _gini_vec(A),
                 "Ubar": float(Ut[0].mean()), "gini_U": _gini_vec(Ut[0]),
                 "Umin": float(Ut[0].min()),
                 "gini_w": _gini_vec(w), "adv_share": float(w[_rt_adv].sum()),
                 "adv_share_min": share_lo, "adv_share_max": share_hi,
                 "m_scan": m_scan, "obj_scan": v_scan, "n_local_maxima": n_loc,
                 "n_starts_at_best": n_at_best,
                 "start_m_min": min(m for _, m in finals), "start_m_max": max(m for _, m in finals),
                 "kkt_err": kkt_err,
                 "member": k_mem + 1, "member_dist": float(abs(mus[k_mem] - m_sqp))}, **ru)


_rt_rows, _rt_cert = [], []
_rt_tie_lists = {}                                   # (station, regime) -> near-tied local maxima
_rt_W = {"eff": [], "eq": []}
for _j, _st in enumerate(_RT_STATIONS):
    _re = _rt_solve(_j, "eff")
    _rq = _rt_solve(_j, "eq4")
    _rt_W["eff"].append(_re["w"])
    _rt_W["eq"].append(_rq["w"])
    _rt_rows.append({"station": _st["id"], "name": _st["name"],
                     "mean_eff": _re["mu"], "mean_eq": _rq["mu"],
                     "sd_eff": _re["sd"], "sd_eq": _rq["sd"],
                     "weight_gini_eff": _re["gini_w"], "weight_gini_eq": _rq["gini_w"],
                     "adv_share_eff": _re["adv_share"], "adv_share_eq": _rq["adv_share"],
                     "Ubar_eff": _re["Ubar"], "Ubar_eq": _rq["Ubar"],
                     "gini_U_eff": _re["gini_U"], "gini_U_eq": _rq["gini_U"],
                     "Abar_eff": _re["Abar"], "Abar_eq": _rq["Abar"],
                     "Amin_eff": _re["Amin"], "Amin_eq": _rq["Amin"],
                     "objective_eff": _re["obj"], "objective_eq": _rq["obj"],
                     # level-set range of the advocates' share (the reported
                     # adv_share_* is the tie-break representative inside it)
                     "adv_share_min_eff": _re["adv_share_min"], "adv_share_max_eff": _re["adv_share_max"],
                     "adv_share_min_eq": _rq["adv_share_min"], "adv_share_max_eq": _rq["adv_share_max"],
                     # functions of the collective mean alone
                     "Abar_adv_eff": _re["Abar_adv"], "Abar_adv_eq": _rq["Abar_adv"],
                     "Umin_eff": _re["Umin"], "Umin_eq": _rq["Umin"],
                     # the Gini of the overlap profile (a function of the mean alone)
                     "gini_A_eff": _re["gini_A"], "gini_A_eq": _rq["gini_A"],
                     # the consensus optimum is this member's own mean
                     "consensus_member": _re["member"],
                     "consensus_member_constituency": _rt_cid[_re["member"] - 1],
                     "consensus_member_dist": _re["member_dist"]})
    for _reg, _r in (("eff", _re), ("eq4", _rq)):
        _rt_cert.append({"station": _st["id"], "regime": _reg,
                         "m_sqp": _r["mu"], "objective_sqp": _r["obj"],
                         "m_scan": _r["m_scan"], "objective_scan": _r["obj_scan"],
                         "n_local_maxima_scan": _r["n_local_maxima"],
                         "n_starts_at_best": _r["n_starts_at_best"],
                         "start_terminal_m_min": _r["start_m_min"],
                         "start_terminal_m_max": _r["start_m_max"],
                         "projection_kkt_maxdiff": _r["kkt_err"],
                         "runner_up_m": _r["ru_m"], "runner_up_objective": _r["ru_obj"],
                         "runner_up_gap": _r["ru_gap"],
                         "runner_up_rel_gap": _r["ru_rel_gap"],
                         "near_tie": int(bool(_r["ru_rel_gap"] < _RT_NEAR_TIE_TOL)),
                         "n_near_tie_maxima": len(_r["ties"]),
                         "first_non_tie_rel_gap": _r["first_non_tie_rel_gap"],
                         "runner_up_adv_share": _r["ru_adv_share"],
                         "runner_up_weight_gini": _r["ru_gini_w"]})
        _rt_tie_lists[(_st["id"], _reg)] = _r["ties"]
        assert _rt_cert[-1]["near_tie"] == int(len(_r["ties"]) > 0)
        for _t in _r["ties"]:
            print("   near-tie: %s %s optimum %.8f at m %.4f, local maximum %.8f at m %.4f (gap %.1e, relative %.1e)"
                  % (_st["id"], _reg, _r["obj_scan"], _r["mu"], _t["obj"], _t["m"], _t["gap"], _t["rel_gap"]))
    print("   %-4s eff %.4f  eq %.4f  (w-gini %.2f -> %.2f, adv share %.2f -> %.2f, "
          "Ubar %.3f -> %.3f)  local maxima %d/%d, starts at best %d/%d"
          % (_st["id"], _re["mu"], _rq["mu"], _re["gini_w"], _rq["gini_w"],
             _re["adv_share"], _rq["adv_share"], _re["Ubar"], _rq["Ubar"],
             _re["n_local_maxima"], _rq["n_local_maxima"],
             _re["n_starts_at_best"], _rq["n_starts_at_best"]))

# priority order for upgrading = ascending collective accessibility
_ord_eff = sorted(range(10), key=lambda j: _rt_rows[j]["mean_eff"])
_ord_eq = sorted(range(10), key=lambda j: _rt_rows[j]["mean_eq"])
for _p, _j in enumerate(_ord_eff, 1):
    _rt_rows[_j]["priority_eff"] = _p
for _p, _j in enumerate(_ord_eq, 1):
    _rt_rows[_j]["priority_eq"] = _p
_swapped = [r["station"] for r in _rt_rows if r["priority_eff"] != r["priority_eq"]]
# the designed contested station is the one whose constituency means are
# furthest apart (tab:rt-stations); the design target of the case is that
# it, and only it, exchanges places with its neighbour in the queue
_rt_contested = max(_RT_STATIONS, key=lambda s: max(s["ops"], s["pas"], s["adv"])
                    - min(s["ops"], s["pas"], s["adv"]))["id"]
assert len(_swapped) == 2 and _rt_contested in _swapped, \
    "the designed priority swap at the contested station is lost: %r" % _swapped
_rt_wg_eff = float(np.mean([r["weight_gini_eff"] for r in _rt_rows]))
_rt_wg_eq = float(np.mean([r["weight_gini_eq"] for r in _rt_rows]))
# "uncontested" as the book uses it: the nine stations other than the
# contested one (the swap partner S5 is uncontested; its means agree)
_rt_unc_gaps = {r["station"]: abs(r["mean_eff"] - r["mean_eq"]) for r in _rt_rows
                if r["station"] != _rt_contested}
_rt_unc_gap = max(_rt_unc_gaps.values())
_rt_unc_gap_station = max(_rt_unc_gaps, key=_rt_unc_gaps.get)
_rt_unc_below = sum(1 for g in _rt_unc_gaps.values() if g < _RT_GAP_THRESHOLD)
# the prose: largest uncontested gap at one station, every other one below the threshold
assert _rt_unc_below == len(_rt_unc_gaps) - 1 and _rt_unc_gaps[_rt_unc_gap_station] >= _RT_GAP_THRESHOLD, \
    "uncontested gaps: %r" % _rt_unc_gaps
print("   swapped stations:", _swapped, " weight-gini eff %.3f vs eq %.3f;"
      " largest uncontested |cons - equity| %.3f at %s, %d/%d below %g"
      % (_rt_wg_eff, _rt_wg_eq, _rt_unc_gap, _rt_unc_gap_station, _rt_unc_below,
         len(_rt_unc_gaps), _RT_GAP_THRESHOLD))
_rt_ids0 = [s["id"] for s in _RT_STATIONS]
_rt_c = _rt_rows[_rt_ids0.index(_rt_contested)]
_rt_contested_gap = abs(_rt_c["mean_eff"] - _rt_c["mean_eq"])

# ---- what the optimal level sets fix about the advocates' share ------------
# the reported shares are tie-break representatives; the LP ranges are what
# every optimum shares.  Headcount share of the advocates = 6/24.
_rt_headcount = float(_rt_adv.mean())
_rt_cap_stations = [r["station"] for r in _rt_rows if r["adv_share_max_eff"] < _rt_headcount]
_rt_floor_stations = [r["station"] for r in _rt_rows if r["adv_share_min_eq"] > 1e-9]
assert _rt_cap_stations == [_rt_contested], \
    "the consensus level set should cap the advocates below headcount at the contested station only: %r" % _rt_cap_stations
assert _rt_c["adv_share_min_eq"] > 0.0
_rt_share_overlap = (max(_rt_c["adv_share_min_eff"], _rt_c["adv_share_min_eq"]),
                     min(_rt_c["adv_share_max_eff"], _rt_c["adv_share_max_eq"]))
assert _rt_share_overlap[0] < _rt_share_overlap[1], "the two optimal sets share no advocate value"
# the level-set properties the text states (functions of m* alone)
assert _rt_c["Abar_adv_eq"] > _rt_c["Abar_adv_eff"] and _rt_c["Amin_eq"] > _rt_c["Amin_eff"]
assert _rt_c["gini_U_eq"] < _rt_c["gini_U_eff"] and _rt_c["Abar_eq"] < _rt_c["Abar_eff"]
# in the overlaps: the consensus optimum maximizes the mean overlap, so the
# equity optimum lowers it at every station (the means differ at all ten);
# where it also lowers the Gini of the overlaps the move is the leveling-down
# case of prop:consensus-fairness(iii), with the Gini index as the measure.
# The text states that on the reported optima this holds at every station
# but one (named), bounds the mean overlap given up at the uncontested
# stations, and gives the contested station's figures, where the smallest
# overlap nevertheless rises; the same classification over the combinations
# of nearly tied optima is recorded after the near-tie block below
assert all(abs(r["mean_eff"] - r["mean_eq"]) > 1e-6 and r["Abar_eq"] < r["Abar_eff"] for r in _rt_rows)
_rt_ovl_down = [r["station"] for r in _rt_rows if r["gini_A_eq"] < r["gini_A_eff"] and r["Abar_eq"] < r["Abar_eff"]]
_rt_ovl_except = [r["station"] for r in _rt_rows if r["station"] not in _rt_ovl_down]
assert _rt_contested in _rt_ovl_down and _rt_c["gini_A_eq"] < _rt_c["gini_A_eff"]
assert len(_rt_ovl_except) == 1, "the text names exactly one station without overlap leveling-down: %r" % _rt_ovl_except
_rt_dA = {r["station"]: r["Abar_eff"] - r["Abar_eq"] for r in _rt_rows if r["station"] != _rt_contested}
_rt_dA_st = max(_rt_dA, key=_rt_dA.get)
assert _rt_dA[_rt_dA_st] < _rt_c["Abar_eff"] - _rt_c["Abar_eq"]
print("   overlaps: mean overlap lower under equity at every station; overlap-level leveling-down"
      " (overlap Gini and mean both fall) at %s (not at %s); largest uncontested fall of the mean overlap"
      " %.4f at %s; contested overlap Gini %.3f -> %.3f"
      % (",".join(_rt_ovl_down), ",".join(_rt_ovl_except), _rt_dA[_rt_dA_st], _rt_dA_st,
         _rt_c["gini_A_eff"], _rt_c["gini_A_eq"]))
print("   level sets at %s: advocate share consensus [%.3f, %.3f], equity [%.3f, %.3f] (headcount %.2f);"
      " consensus caps the advocates below headcount at %s; equity floors them above 0 at %s"
      % (_rt_contested, _rt_c["adv_share_min_eff"], _rt_c["adv_share_max_eff"],
         _rt_c["adv_share_min_eq"], _rt_c["adv_share_max_eq"], _rt_headcount,
         ",".join(_rt_cap_stations), ",".join(_rt_floor_stations)))
# the consensus optimum is one member's own mean at every station
assert max(r["consensus_member_dist"] for r in _rt_rows) < 1e-6
print("   consensus optimum = a member's own mean at every station (max distance %.1e)"
      % max(r["consensus_member_dist"] for r in _rt_rows))

# ---- near-ties: every combination of reported and near-tied optima ---------
# a tied solve with k near-tied local maxima offers 1 + k choices (0 = the
# reported optimum, t = its t-th near-tied local maximum by value)
_rt_ties = [c for c in _rt_cert if c["near_tie"]]
_rt_tie_rows, _rt_combo = [], []
for _c in _rt_ties:
    _col = "eff" if _c["regime"] == "eff" else "eq"
    _r = _rt_rows[_rt_ids0.index(_c["station"])]
    _tl = _rt_tie_lists[(_c["station"], _c["regime"])]
    assert abs(_tl[0]["m"] - _c["runner_up_m"]) < 1e-15      # rank 1 is the runner-up
    for _rank, _t in enumerate(_tl, 1):
        _rt_tie_rows.append({"station": _c["station"], "regime": _c["regime"], "rank": _rank,
                             "m_reported": _c["m_sqp"], "m_local_max": _t["m"],
                             "objective_gap": _t["gap"], "relative_gap": _t["rel_gap"],
                             "adv_share_reported": _r["adv_share_" + _col],
                             "adv_share_local_max": _t["adv_share"],
                             "weight_gini_reported": _r["weight_gini_" + _col],
                             "weight_gini_local_max": _t["gini_w"]})
_rt_n_tie_max = len(_rt_tie_rows)
for _pick in _rt_itertools.product(*[range(1 + c["n_near_tie_maxima"]) for c in _rt_ties]):
    _mm = {k: [r["mean_" + k] for r in _rt_rows] for k in ("eff", "eq")}
    _gg = {k: [r["weight_gini_" + k] for r in _rt_rows] for k in ("eff", "eq")}
    for _c, _p in zip(_rt_ties, _pick):
        if _p:
            _col = "eff" if _c["regime"] == "eff" else "eq"
            _j = _rt_ids0.index(_c["station"])
            _t = _rt_tie_lists[(_c["station"], _c["regime"])][_p - 1]
            _mm[_col][_j] = _t["m"]
            _gg[_col][_j] = _t["gini_w"]
    _same = all(list(np.argsort(np.argsort(_mm[k])) + 1) == [r["priority_" + k] for r in _rt_rows]
                for k in ("eff", "eq"))
    _ug = {s: abs(_mm["eff"][j] - _mm["eq"][j]) for j, s in enumerate(_rt_ids0) if s != _rt_contested}
    _top = max(_ug, key=_ug.get)
    # the overlap classification of prop:consensus-fairness(iii) on this
    # combination: the mean overlap lower under equity, and both the mean and
    # the Gini of the overlaps lower (the leveling-down case, Gini index)
    _rt_cb_low, _rt_cb_lvl = [], []
    for _j, _s in enumerate(_rt_ids0):
        _Ae = _rt_overlap(_rt_MU[:, _j], _mm["eff"][_j])[0]
        _Aq = _rt_overlap(_rt_MU[:, _j], _mm["eq"][_j])[0]
        if _Aq.mean() < _Ae.mean():
            _rt_cb_low.append(_s)
            if _gini_vec(_Aq) < _gini_vec(_Ae):
                _rt_cb_lvl.append(_s)
    _rt_combo.append({"pick": "".join(str(p) for p in _pick), "priorities_unchanged": int(_same),
                      "max_uncontested_gap_station": _top, "max_uncontested_gap": _ug[_top],
                      "others_below_threshold": int(all(g < _RT_GAP_THRESHOLD for s, g in _ug.items() if s != _top)),
                      "n_uncontested_le_0.1": sum(1 for g in _ug.values() if g <= 0.1),
                      "weight_gini_eff_avg": float(np.mean(_gg["eff"])),
                      "weight_gini_eq_avg": float(np.mean(_gg["eq"])),
                      "n_mean_overlap_lower": len(_rt_cb_low),
                      "n_overlap_leveling_down": len(_rt_cb_lvl),
                      "overlap_leveling_down_at": ",".join(_rt_cb_lvl)})
# the equity near-tie whose local maximum lies furthest from the reported optimum
_rt_tie_eq = max((t for t in _rt_tie_rows if t["regime"] == "eq4"),
                 key=lambda t: abs(t["m_local_max"] - t["m_reported"]))
# the closest near-tie of all, by relative gap; the text prints the optimum
# and EVERY near-tied local maximum of that solve, bounds all their relative
# gaps by _RT_CLOSEST_TIE_BOUND, and prints the share and weight Gini of the
# runner-up's representative; on this panel it is the same solve as the
# equity near-tie above
_rt_tie_min = min(_rt_tie_rows, key=lambda t: t["relative_gap"])
assert _rt_tie_min["rank"] == 1
assert (_rt_tie_min["station"], _rt_tie_min["regime"]) == (_rt_tie_eq["station"], _rt_tie_eq["regime"])
_rt_tie_cl = [t for t in _rt_tie_rows
              if (t["station"], t["regime"]) == (_rt_tie_min["station"], _rt_tie_min["regime"])]
assert all(t["relative_gap"] < _RT_CLOSEST_TIE_BOUND for t in _rt_tie_cl)
# the closest solve is THE one solve with more than one near-tied local
# maximum (the text: "the one solve with two of them")
_rt_tie_multi = ["%s %s" % (c["station"], c["regime"]) for c in _rt_ties if c["n_near_tie_maxima"] > 1]
assert _rt_tie_multi == ["%s %s" % (_rt_tie_min["station"], _rt_tie_min["regime"])] and len(_rt_tie_cl) == 2, \
    "the closest near-tie is not at the one solve with two near-tied local maxima: %r" % _rt_tie_multi
_rt_combo_n = len(_rt_combo)
# the overlap leveling-down classification over the combinations: the pick
# of the reported optima reproduces the results sheet; the text states the
# range of the count, that the contested station is always included and the
# reported exception never, and names the stations whose classification
# turns on the choice among nearly tied optima
assert _rt_combo[0]["pick"] == "0" * len(_rt_ties)
assert _rt_combo[0]["overlap_leveling_down_at"] == ",".join(_rt_ovl_down)
assert _rt_combo[0]["n_mean_overlap_lower"] == len(_rt_rows)
_rt_lvl_sets = [set(filter(None, r["overlap_leveling_down_at"].split(","))) for r in _rt_combo]
_rt_lvl_n = (min(len(s) for s in _rt_lvl_sets), max(len(s) for s in _rt_lvl_sets))
_rt_lvl_always = [s for s in _rt_ids0 if all(s in d for d in _rt_lvl_sets)]
_rt_lvl_never = [s for s in _rt_ids0 if all(s not in d for d in _rt_lvl_sets)]
_rt_lvl_varies = [s for s in _rt_ids0 if s not in _rt_lvl_always and s not in _rt_lvl_never]
_rt_lowA_n = (min(r["n_mean_overlap_lower"] for r in _rt_combo), max(r["n_mean_overlap_lower"] for r in _rt_combo))
assert _rt_contested in _rt_lvl_always and _rt_lvl_never == _rt_ovl_except
assert _rt_lvl_n[1] == len(_rt_ovl_down) and _rt_lvl_n[0] < _rt_lvl_n[1] and _rt_lvl_varies
# every station whose classification varies has a near-tied solve
assert all(any(c["station"] == s for c in _rt_ties) for s in _rt_lvl_varies)
print("   overlap leveling-down over the %d combinations: %d..%d stations (always %s, never %s, varies %s);"
      " mean overlap lower under equity at %d..%d"
      % (len(_rt_combo), _rt_lvl_n[0], _rt_lvl_n[1], ",".join(_rt_lvl_always), ",".join(_rt_lvl_never),
         ",".join(_rt_lvl_varies), _rt_lowA_n[0], _rt_lowA_n[1]))
_rt_combo_same = sum(r["priorities_unchanged"] for r in _rt_combo)
_rt_combo_gapok = sum(1 for r in _rt_combo if r["others_below_threshold"]
                      and r["max_uncontested_gap_station"] == _rt_unc_gap_station)
_rt_combo_wg = {k: (min(r["weight_gini_%s_avg" % k] for r in _rt_combo),
                    max(r["weight_gini_%s_avg" % k] for r in _rt_combo)) for k in ("eff", "eq")}
assert _rt_combo_same == _rt_combo_n and _rt_combo_gapok == _rt_combo_n, \
    "a near-tie changes a priority or the uncontested-gap statement"
assert _rt_combo_n == int(np.prod([1 + c["n_near_tie_maxima"] for c in _rt_ties]))
print("   near-ties (relative gap < %g): %s, %d local maxima; over the %d combinations priorities unchanged %d,"
      " gap statement holds %d, average weight Gini eff %.3f..%.3f, eq %.3f..%.3f"
      % (_RT_NEAR_TIE_TOL, ", ".join("%s %s x%d" % (c["station"], c["regime"], c["n_near_tie_maxima"]) for c in _rt_ties),
         _rt_n_tie_max, _rt_combo_n, _rt_combo_same, _rt_combo_gapok,
         _rt_combo_wg["eff"][0], _rt_combo_wg["eff"][1], _rt_combo_wg["eq"][0], _rt_combo_wg["eq"][1]))

# ---- leveling-down record on the reported solutions ------------------------
_rt_dU = {r["station"]: r["Ubar_eq"] - r["Ubar_eff"] for r in _rt_rows if r["station"] != _rt_contested}
_rt_dG = {r["station"]: r["gini_U_eq"] - r["gini_U_eff"] for r in _rt_rows if r["station"] != _rt_contested}
_rt_dU_st = max(_rt_dU, key=lambda s: abs(_rt_dU[s]))
_rt_dG_st = max(_rt_dG, key=lambda s: abs(_rt_dG[s]))
_rt_Ufall = [r["station"] for r in _rt_rows if r["Ubar_eq"] < r["Ubar_eff"]]
assert abs(_rt_dU[_rt_dU_st]) < 0.01, "an uncontested mean utility moves by 0.01 or more"
# wherever the mean utility falls, the smallest utility on the panel rises
assert all(_rt_rows[_rt_ids0.index(s)]["Umin_eq"] > _rt_rows[_rt_ids0.index(s)]["Umin_eff"] for s in _rt_Ufall)
assert _rt_c["Ubar_eq"] > _rt_c["Ubar_eff"]
print("   leveling down: uncontested max |dUbar| %.5f at %s, max |dGini_U| %.5f at %s; Ubar falls at %s"
      % (abs(_rt_dU[_rt_dU_st]), _rt_dU_st, abs(_rt_dG[_rt_dG_st]), _rt_dG_st, ",".join(_rt_Ufall) or "none"))

# ---- the realized panel behind each station's regime gap -------------------
_rt_real = []
for _j, _st in enumerate(_RT_STATIONS):
    _x = _rt_MU[:, _j]
    _base = np.array([_st[k] for k in _rt_ckey])
    _noise = _x - _base - _rt_len                       # the clip does not bind (asserted below)
    _des = _st["adv"] - np.mean([_st[k] for k in _rt_ckey if k != "adv"])
    _o = np.argsort(_x)
    _row = {"station": _st["id"], "panel_mean": float(_x.mean()),
            "adv_mean": float(_x[_rt_adv].mean()), "rest_mean": float(_x[~_rt_adv].mean()),
            "designed_adv_gap": float(_des),
            "realized_adv_gap": float(_x[_rt_adv].mean() - _x[~_rt_adv].mean()),
            "regime_gap": abs(_rt_rows[_j]["mean_eff"] - _rt_rows[_j]["mean_eq"])}
    for _q in (0, 1):
        _i = int(_o[_q])
        _row.update({"low%d_expert" % (_q + 1): _i + 1,
                     "low%d_constituency" % (_q + 1): _rt_cid[_i],
                     "low%d_mu" % (_q + 1): float(_x[_i]),
                     "low%d_below_panel_mean" % (_q + 1): float(_x.mean() - _x[_i]),
                     "low%d_leniency" % (_q + 1): float(_rt_len[_i]),
                     "low%d_noise" % (_q + 1): float(_noise[_i])})
    _rt_real.append(_row)
assert not _rt_clip_binding

# ---- robustness (a): re-seeded panels, same generator and draw order -----
# solved by the certified 1-D scan (the global solver of this
# one-dimensional problem), which is why 40 panels x 20 solves are cheap
_rt_j4 = [s["id"] for s in _RT_STATIONS].index(_rt_contested)
_rt_j5 = _ord_eff[_ord_eff.index(_rt_j4) - 1]      # the station just ahead of it in the consensus queue
_rt_j6 = _ord_eff[_ord_eff.index(_rt_j4) - 2]      # and the next one ahead
assert sorted(_swapped) == sorted([_RT_STATIONS[_rt_j4]["id"], _RT_STATIONS[_rt_j5]["id"]]), \
    "the swap %r is not the contested station and its consensus neighbour" % _swapped
_rt_seed_rows = []
for _s in range(1, _RT_RESEED_N + 1):
    _MUs, _ = _rt_generate_panel(np.random.default_rng(NEW_SEED + _s))
    _mc = [_rt_scan(_MUs[:, j], "eff")[0] for j in range(10)]
    _me = [_rt_scan(_MUs[:, j], "eq4")[0] for j in range(10)]
    _pc = np.argsort(np.argsort(_mc)) + 1
    _pe = np.argsort(np.argsort(_me)) + 1
    _dir = bool(_mc[_rt_j4] > _mc[_rt_j5] and _me[_rt_j4] < _me[_rt_j5])
    _chg = [_RT_STATIONS[k]["id"] for k in range(10) if _pc[k] != _pe[k]]
    _ug = {k: abs(_mc[k] - _me[k]) for k in range(10) if k != _rt_j4}
    _ku = max(_ug, key=_ug.get)
    _rt_seed_rows.append({"seed": NEW_SEED + _s,
                          "S4_mean_eff": _mc[_rt_j4], "S4_mean_eq": _me[_rt_j4],
                          "S5_mean_eff": _mc[_rt_j5], "S5_mean_eq": _me[_rt_j5],
                          "S6_mean_eff": _mc[_rt_j6], "S6_mean_eq": _me[_rt_j6],
                          "swap_direction_preserved": int(_dir),
                          "other_orders_unchanged": int(all(k in _swapped for k in _chg)),
                          "contested_places_moved": int(_pc[_rt_j4] - _pe[_rt_j4]),
                          "S5_S6_order_unchanged": int((_pc[_rt_j5] > _pc[_rt_j6]) == (_pe[_rt_j5] > _pe[_rt_j6])),
                          "stations_reordered": ",".join(_chg),
                          # the equity margin behind the swap, and the regime gaps
                          # behind the early-warning rule (uncontested = all but S4)
                          "eq_swap_margin": _me[_rt_j5] - _me[_rt_j4],
                          "contested_gap": abs(_mc[_rt_j4] - _me[_rt_j4]),
                          "max_uncontested_gap": _ug[_ku],
                          "max_uncontested_gap_station": _RT_STATIONS[_ku]["id"]})
_rt_seed_dir = sum(r["swap_direction_preserved"] for r in _rt_seed_rows)
_rt_seed_oth = sum(r["other_orders_unchanged"] for r in _rt_seed_rows)
_rt_seed_two = sum(1 for r in _rt_seed_rows
                   if r["contested_places_moved"] == 2 and r["S5_S6_order_unchanged"])
print("   re-seeded panels: swap direction preserved %d/%d, every other order unchanged %d/%d"
      % (_rt_seed_dir, _RT_RESEED_N, _rt_seed_oth, _RT_RESEED_N))
# the swap margin against its seed-to-seed spread (sample sd, ddof 1)
_rt_margin = _rt_rows[_rt_j5]["mean_eq"] - _rt_rows[_rt_j4]["mean_eq"]
_rt_seed_margins = np.array([r["eq_swap_margin"] for r in _rt_seed_rows])
_rt_margin_sd = float(np.std(_rt_seed_margins, ddof=1))
_rt_margin_min = float(_rt_seed_margins.min())
assert _rt_margin > 3.0 * _rt_margin_sd and _rt_margin_min > 0.0
# the early-warning rule over the main panel and the re-seeded ones
_rt_warn_cont = [_rt_contested_gap] + [r["contested_gap"] for r in _rt_seed_rows]
_rt_warn_unc = [_rt_unc_gap] + [r["max_uncontested_gap"] for r in _rt_seed_rows]
_rt_warn_sep = sum(1 for a, b in zip(_rt_warn_cont, _rt_warn_unc)
                   if a > _RT_WARN_THRESHOLD > b)
assert _rt_warn_sep == 1 + _RT_RESEED_N, "the early-warning threshold does not separate every panel"
print("   swap margin %.4f vs re-seed sd %.4f (min %.4f); warning threshold %.2f separates %d/%d panels"
      " (contested gap >= %.3f, uncontested <= %.3f; re-seed median of the largest uncontested gap %.3f)"
      % (_rt_margin, _rt_margin_sd, _rt_margin_min, _RT_WARN_THRESHOLD, _rt_warn_sep, 1 + _RT_RESEED_N,
         min(_rt_warn_cont), max(_rt_warn_unc),
         float(np.median([r["max_uncontested_gap"] for r in _rt_seed_rows]))))

# ---- robustness (b): the Fehr-Schmidt grid on the seed-202603 panel -----
_rt_fs_rows = []
_mc0 = [r["mean_eff"] for r in _rt_rows]
for _rho in _RT_FS_RHO:
    for _th in _RT_FS_THETA:
        if _th > _rho:
            continue
        _me = [_rt_scan(_rt_MU[:, j], "eq4", _rho, _th)[0] for j in range(10)]
        _pc = np.argsort(np.argsort(_mc0)) + 1
        _pe = np.argsort(np.argsort(_me)) + 1
        _w4, _ = _rt_project_uniform(_rt_MU[:, _rt_j4], _me[_rt_j4])
        _chg = [_RT_STATIONS[k]["id"] for k in range(10) if _pc[k] != _pe[k]]
        _rt_fs_rows.append({"rho": _rho, "theta": _th,
                            "S4_mean_eq": _me[_rt_j4], "S5_mean_eq": _me[_rt_j5],
                            "swap_direction_preserved": int(_me[_rt_j4] < _me[_rt_j5]),
                            "S4_adv_share_eq": float(_w4[_rt_adv].sum()),
                            "S4_weight_gini_eq": _gini_vec(_w4),
                            "other_orders_unchanged": int(all(k in _swapped for k in _chg)),
                            "stations_reordered": ",".join(_chg)})
_rt_fs_dir = sum(r["swap_direction_preserved"] for r in _rt_fs_rows)
_rt_fs_adv = (min(r["S4_adv_share_eq"] for r in _rt_fs_rows),
              max(r["S4_adv_share_eq"] for r in _rt_fs_rows))
print("   Fehr-Schmidt grid: swap direction preserved %d/%d, S4 advocate share in [%.3f, %.3f]"
      % (_rt_fs_dir, len(_rt_fs_rows), _rt_fs_adv[0], _rt_fs_adv[1]))

_rt_ids = [s["id"] for s in _RT_STATIONS]
write("fig6-11_rt_case.xlsx", {
    "config": cfg({
        "case": "IV -- rail-transit station accessibility with stakeholder equity (Section 6.5)",
        "figures": "fig6-11_rt_equity",
        "labels": "fig:rt-equity, tab:rt-results",
        "source_paper": "chen2024railtransit",
        "mechanism_only": ("synthetic instrument and panel; mirrors the source "
                           "study at the level of MECHANISM only -- equity "
                           "weighting reorders priorities -- and none of its "
                           "findings"),
        "synthetic": 1, "seed": NEW_SEED,
        "n_stakeholders": _rt_N, "n_stations": 10,
        "scale_min": 0.0, "scale_max": 10.0,
        "rho": _RT_RHO, "theta": _RT_THETA, "kappa": 2,
        "leniency_sd": _RT_LENIENCY_SD, "noise_sd": _RT_NOISE_SD,
        "common_sigma": _RT_SIGMA,
        "generator_draw_order": ("numpy default_rng(202603): 24 leniency draws N(0, 0.15^2), "
                                 "one per stakeholder; then 240 scalar noise draws N(0, 0.25^2) "
                                 "station-major (S1 stakeholders 1..24, then S2, ...); then the "
                                 "40 Dirichlet(1,...,1) SQP start vectors"),
        "generator_clip": "each mean clipped to [0, 10]",
        "generator_clip_binding": int(_rt_clip_binding),
        "overlap": ("def:consensus for equal-variance normals, closed form over the real "
                    "line: 2 Phi(-|mu_i - m| / (2 sigma)); no truncation, no quadrature"),
        "regime_eff": ("consensus only -- the mean overlap, i.e. the cumulative overlap of def:consensus "
                       "divided by N as in Case I (same argmax); the confidence objective is dropped"),
        "regime_eq": ("Sen's abbreviated welfare form, mean power-normalized "
                      "Fehr-Schmidt utility times one minus the Gini of the utilities, "
                      "Ubar (1 - G(U)); the confidence objective is dropped (regime id 'eq4' in the "
                      "sheets; not the bare equity objective EQ_4 = -I_4 of Table tab:swf)"),
        "confidence_objective": ("dropped from both programs: with a common sigma the "
                                 "aggregate sd is constant over the simplex"),
        "solver": ("multi-start sequential quadratic programming: scipy.optimize.minimize "
                   "method SLSQP, finite-difference gradients, bounds [0,1], equality "
                   "constraint sum w = 1; both regimes; best objective over the starts retained"),
        "solver_starts": "uniform vector plus %d Dirichlet(1,...,1) draws (see generator_draw_order), shared by every station and regime" % _RT_K_STARTS,
        "solver_ftol": _RT_SQP_FTOL, "solver_maxiter": _RT_SQP_MAXITER,
        "certification": ("every objective depends on w only through m = w.mu (common sigma); "
                          "a dense scan of m over [min mu, max mu] with %d points and a bounded "
                          "refinement (xatol %g) is the global solver of the 1-D problem; the build "
                          "aborts unless the SQP best is within %g (objective) and %g (m) of it"
                          % (_RT_SCAN_POINTS, _RT_SCAN_XATOL, _RT_CERT_OBJ_TOL, _RT_CERT_M_TOL)),
        "reporting_rule": ("each regime reports the point of its optimal level set "
                           "{w in simplex : w.mu = m*} nearest to the uniform vector "
                           "(min sum w_i^2), a convex QP solved by SLSQP (ftol 1e-16) and "
                           "cross-checked against its KKT form w_i = max(0, 1/n + a + b mu_i)"),
        "priority_rule": "ascending collective mean (worst station first)",
        "robustness_reseed": ("%d further panels under seeds %d..%d, same generator and draw "
                              "order, both regimes solved by the certified 1-D scan"
                              % (_RT_RESEED_N, NEW_SEED + 1, NEW_SEED + _RT_RESEED_N)),
        "robustness_reseed_swap_direction": "%d/%d" % (_rt_seed_dir, _RT_RESEED_N),
        "robustness_reseed_other_orders_unchanged": "%d/%d" % (_rt_seed_oth, _RT_RESEED_N),
        "robustness_reseed_contested_moves_two_places": "%d/%d (S5 and S6 keep their order)" % (_rt_seed_two, _RT_RESEED_N),
        "robustness_fs_grid": ("rho in {%s}, vartheta in {%s}, vartheta <= rho, %d combinations, "
                               "seed-202603 panel, equity regime solved by the certified 1-D scan"
                               % (", ".join(str(x) for x in _RT_FS_RHO),
                                  ", ".join(str(x) for x in _RT_FS_THETA), len(_rt_fs_rows))),
        "robustness_fs_swap_direction": "%d/%d" % (_rt_fs_dir, len(_rt_fs_rows)),
        "robustness_fs_S4_adv_share_range": "%.4f .. %.4f" % _rt_fs_adv,
        "x_label": "aggregation regime",
        "y_label": "collective accessibility index",
        "eff_tick": "consensus only",
        # Section 6.5 does not introduce EQ4, and Table tab:swf's EQ_4 = -I_4 is the bare
        # equity objective the case does not use, so the tick names the Sen form
        "eq_tick": "equity aware (Sen form)",
        "contested_station": _rt_contested,
        "swapped": ",".join(_swapped),
        "weight_gini_eff_avg": _rt_wg_eff,
        "weight_gini_eq_avg": _rt_wg_eq,
        "max_uncontested_mean_gap": _rt_unc_gap,
        "max_uncontested_mean_gap_station": _rt_unc_gap_station,
        "uncontested_gaps_below_threshold": "%d/%d below %g (uncontested = every station but the contested one)"
                                            % (_rt_unc_below, len(_rt_unc_gaps), _RT_GAP_THRESHOLD),
        "gap_threshold": _RT_GAP_THRESHOLD,
        "print_dp_check": "SQP and scan means agree when rounded to %d dp at every solve" % _RT_PRINT_DP,
        "contested_mean_gap": _rt_contested_gap,
        # what the optimal level sets fix (the reported shares are the tie-break representatives)
        "level_set_share": ("LP range (scipy linprog, HiGHS; cross-checked by vertex enumeration) of the "
                            "advocates' weight share over each optimal level set {w in simplex : w.mu = m*}; "
                            "results columns adv_share_min_* / adv_share_max_*"),
        "adv_headcount_share": _rt_headcount,
        "contested_adv_share_range_eff": "%.4f .. %.4f" % (_rt_c["adv_share_min_eff"], _rt_c["adv_share_max_eff"]),
        "contested_adv_share_range_eq": "%.4f .. %.4f" % (_rt_c["adv_share_min_eq"], _rt_c["adv_share_max_eq"]),
        "contested_adv_share_common_range": "%.4f .. %.4f" % _rt_share_overlap,
        "contested_adv_share_ratio": _rt_c["adv_share_eq"] / _rt_c["adv_share_eff"],
        "consensus_caps_advocates_below_headcount_at": ",".join(_rt_cap_stations),
        "equity_floors_advocates_above_zero_at": ",".join(_rt_floor_stations),
        "consensus_optimum": ("with a common sigma each overlap is convex in m on either side of mu_i, so the "
                              "consensus optimum is one member's own mean; results column consensus_member "
                              "(max distance %.1e)" % max(r["consensus_member_dist"] for r in _rt_rows)),
        # near-ties of the one-dimensional objectives
        "near_tie_tolerance": _RT_NEAR_TIE_TOL,
        "near_tie_rule": ("every other interior local maximum of the certification scan, refined like the "
                          "optimum, is compared with the optimum by its RELATIVE gap (v* - v) / v*, which does "
                          "not depend on the scale of the objective (consensus as the mean or the cumulative "
                          "overlap); it is a near-tie when that gap is below near_tie_tolerance; sheet "
                          "near_ties has one row per near-tied local maximum (rank 1 = the runner-up of the "
                          "certification sheet), and sheet near_tie_combinations re-evaluates every "
                          "combination of reported and near-tied optima (1 + k choices at a solve with k)"),
        "near_ties": "%d/%d solves, %d local maxima: %s" % (
            len(_rt_ties), len(_rt_cert), _rt_n_tie_max,
            ", ".join("%s %s%s" % (c["station"], c["regime"],
                                   " (%d)" % c["n_near_tie_maxima"] if c["n_near_tie_maxima"] > 1 else "")
                      for c in _rt_ties)),
        "near_tie_combinations_n": _rt_combo_n,
        # the margin of the rule: the closest local maximum that is NOT a near-tie, over all solves
        "near_tie_smallest_non_tie_rel_gap": float(np.nanmin([c["first_non_tie_rel_gap"] for c in _rt_cert])),
        "near_tie_equity_widest_station": _rt_tie_eq["station"],
        "near_tie_equity_widest_runner_up_m": _rt_tie_eq["m_local_max"],
        "near_tie_closest": "%s %s" % (_rt_tie_min["station"], _rt_tie_min["regime"]),
        "near_tie_closest_n_maxima": len(_rt_tie_cl),
        "near_tie_closest_gap": _rt_tie_min["objective_gap"],
        "near_tie_closest_rel_gap": _rt_tie_min["relative_gap"],
        "near_tie_closest_max_rel_gap": max(t["relative_gap"] for t in _rt_tie_cl),
        "near_tie_closest_gap_bound": _RT_CLOSEST_TIE_BOUND,
        "near_tie_closest_gap_bound_rule": ("every near-tied local maximum of the closest solve has a relative "
                                            "gap below near_tie_closest_gap_bound"),
        "near_tie_solves_with_several_maxima": ",".join(_rt_tie_multi),
        "near_tie_closest_m_reported": _rt_tie_min["m_reported"],
        "near_tie_closest_m_runner_up": _rt_tie_min["m_local_max"],
        "near_tie_closest_m_local_maxima": ", ".join("%.6f" % t["m_local_max"] for t in _rt_tie_cl),
        "near_tie_closest_runner_up_adv_share": _rt_tie_min["adv_share_local_max"],
        "near_tie_closest_runner_up_weight_gini": _rt_tie_min["weight_gini_local_max"],
        "near_tie_priorities_unchanged": "%d/%d" % (_rt_combo_same, _rt_combo_n),
        "near_tie_gap_statement_holds": "%d/%d" % (_rt_combo_gapok, _rt_combo_n),
        "near_tie_weight_gini_eff_avg_range": "%.4f .. %.4f" % _rt_combo_wg["eff"],
        "near_tie_weight_gini_eq_avg_range": "%.4f .. %.4f" % _rt_combo_wg["eq"],
        # leveling-down record (uncontested = every station but the contested one)
        "uncontested_max_abs_mean_utility_change": abs(_rt_dU[_rt_dU_st]),
        "uncontested_max_abs_mean_utility_change_station": _rt_dU_st,
        "uncontested_max_abs_utility_gini_change": abs(_rt_dG[_rt_dG_st]),
        "uncontested_max_abs_utility_gini_change_station": _rt_dG_st,
        "mean_utility_falls_at": ",".join(_rt_Ufall),
        # the same comparison in the overlaps (results columns Abar_*, Amin_*, gini_A_*)
        "mean_overlap_lower_under_equity_at": ("every station (the consensus optimum maximizes the mean "
                                               "overlap and the two means differ at all %d)" % len(_rt_rows)),
        "overlap_leveling_down_at": ",".join(_rt_ovl_down),
        "overlap_leveling_down_except": ",".join(_rt_ovl_except),
        "uncontested_max_mean_overlap_fall": _rt_dA[_rt_dA_st],
        "uncontested_max_mean_overlap_fall_station": _rt_dA_st,
        "overlap_leveling_down_rule": ("prop:consensus-fairness(iii) with the Gini index I_4 as the inequality "
                                       "measure (the index of the case's welfare criterion; the classification "
                                       "is index-specific): the Gini of the overlap profile falls and the mean "
                                       "(total) overlap falls from the consensus to the equity optimum; "
                                       "overlap_leveling_down_at/_except classify the reported optima"),
        # the same classification over every combination of reported and
        # near-tied optima (sheet near_tie_combinations, columns
        # n_overlap_leveling_down, overlap_leveling_down_at, n_mean_overlap_lower)
        "overlap_leveling_down_combo_count_range": "%d .. %d" % _rt_lvl_n,
        "overlap_leveling_down_combo_always": ",".join(_rt_lvl_always),
        "overlap_leveling_down_combo_never": ",".join(_rt_lvl_never),
        "overlap_leveling_down_combo_varies": ",".join(_rt_lvl_varies),
        "mean_overlap_lower_combo_count_range": "%d .. %d" % _rt_lowA_n,
        # swap margin and early-warning rule
        "eq_swap_margin": _rt_margin,
        "reseed_eq_swap_margin_sd": _rt_margin_sd,
        "reseed_eq_swap_margin_min": _rt_margin_min,
        "warn_threshold": _RT_WARN_THRESHOLD,
        "warn_contested_gap_min": min(_rt_warn_cont),
        "warn_uncontested_gap_max": max(_rt_warn_unc),
        "warn_panels_separated": "%d/%d (the main panel and the re-seeded ones)" % (_rt_warn_sep, 1 + _RT_RESEED_N),
        "reseed_max_uncontested_gap_median": float(np.median([r["max_uncontested_gap"] for r in _rt_seed_rows])),
    }),
    "constituencies": pd.DataFrame(
        [{"constituency": c[0], "key": c[1], "size": c[2]} for c in _RT_CONST]),
    "stations": pd.DataFrame(_RT_STATIONS),
    "panel": pd.DataFrame({"expert": np.arange(1, _rt_N + 1),
                           "constituency": _rt_cid, "leniency": _rt_len,
                           "sigma": _rt_SD}),
    "panel_mu": pd.DataFrame(_rt_MU, columns=_rt_ids),
    "results": pd.DataFrame(_rt_rows),
    "weights_eff": pd.DataFrame(np.column_stack(_rt_W["eff"]), columns=_rt_ids),
    "weights_eq": pd.DataFrame(np.column_stack(_rt_W["eq"]), columns=_rt_ids),
    "sqp_starts": pd.DataFrame(np.vstack(_rt_starts),
                               columns=["w%d" % i for i in range(1, _rt_N + 1)]),
    "certification": pd.DataFrame(_rt_cert),
    "near_ties": pd.DataFrame(_rt_tie_rows, columns=["station", "regime", "rank", "m_reported", "m_local_max",
                                                     "objective_gap", "relative_gap",
                                                     "adv_share_reported", "adv_share_local_max",
                                                     "weight_gini_reported", "weight_gini_local_max"]),
    "near_tie_combinations": pd.DataFrame(_rt_combo),
    "realized_panel": pd.DataFrame(_rt_real),
    "robustness_seeds": pd.DataFrame(_rt_seed_rows),
    "robustness_fs": pd.DataFrame(_rt_fs_rows),
})

# =====================================================================
# CASE STUDY III (Section 6.4, building digital twin maturity, FA-LSCOG)
#   -> Figure fig6-8_bdt_q1agg        (fig:bdt-q1agg)
#   -> Figure fig6-9_bdt_profile      (fig:bdt-profile)
#   -> Figure fig6-10_bdt_sensitivity (fig:bdt-sensitivity)
#   -> Tables tab:bdt-groups, tab:bdt-results
#
# GENERATED AND SOLVED (since 2026-09-14; before that every number of the
# case was a typed constant).  Every printed number of the case is computed
# here from the stated inputs under seed 202603:
#   1. a 15-expert x 22-indicator panel of truncated Weibull densities
#      (eq:bdt-weibull-trunc) generated from stated dimension design
#      levels, stratum offsets, a per-expert leniency and a per-cell noise,
#      with the shape drawn uniformly on a stated range and the scale
#      solved so that the TRUNCATED mean equals the target;
#   2. a stated 15 x 3 professional-attribute matrix clustered by k-means
#      (Lloyd's algorithm, standardized columns, seeded restarts) into the
#      three strata, and TOPSIS on the raw cluster centroids (vector
#      normalization, equal criterion weights, benefit criteria) giving
#      the inter-group weights Omega = closeness / sum of closeness;
#   3. the intra-group program of eq:bdt-intra as the book prints it,
#         max_w  delta*Phi~(w) - (1-delta)*V~(w) - lambda_G*G(w)
#      (Bonferroni(1,1) collective fairness of the power-normalized
#      Fehr-Schmidt utilities of the overlaps, the quantile-average
#      variance, both min-max normalized over a sampled simplex, and the
#      weight-Gini influence-dispersion regularizer of Chapter 4), solved
#      for each (subgroup, indicator) and each delta in {0, ..., 0.5} by
#      multi-start SLSQP;
#   4. quantile averaging under Omega (inter-group), equal indicator
#      weights inside a dimension, and best-worst dimension weights from
#      stated comparison vectors (linear model of Rezaei 2016; consistency
#      reported by the input-based ratio CR^I of Liang, Brunelli and Rezaei
#      2020), up to the global maturity distribution F_G;
#   5. the delta sweep of Phi, V, Gini(w) and the global maturity, with the
#      reported variances and member standard deviations evaluated exactly
#      (closed form, or a graded midpoint rule accurate to 1e-12) while the
#      solver keeps its 257-level quantile grid;
#   6. the multi-start evidence over DISTINCT starts and the retained
#      solution: a follow-up of further random starts for every program
#      whose runner-up original start falls short of its best original
#      start by more than a stated gap; a second, independent set of random
#      starts for EVERY program, which measures how far further starts can
#      raise the best value of the original starts (the accuracy of a
#      multi-start solve of this size); the retained solution, which is the
#      best further run wherever it ends above the best original start by
#      more than tie_tol and the best original start otherwise; the
#      agreement statistics against the retained value; the stability of
#      the returned weights over all runs ending within near_tol of the
#      retained value; and the printed numbers recomputed after
#      substituting, in every program at once, the near-optimal run farthest
#      from the retained weights, the best original start, or the polished
#      point of a deterministic local-polish diagnostic (rounds of a
#      pairwise-transfer pattern search and one SLSQP run from every retained
#      solution, repeated until a round gains nothing, drawing nothing), which
#      measures how far SLSQP stalls short of a local optimum and changes no
#      retained solution.
# All programs are solved before any layer is built (the retained solution
# can come from the further starts), with the draws in this order:
# rng_panel (leniency 15, noise 15x22, shape 15x22); rng_kmeans (20 restarts
# x 3 initial centroids); rng_solve (per program, group-major then
# indicator: 300 Dirichlet(1) normalization points; then per delta, group,
# indicator: 12 Dirichlet(1) starts); rng_followup (per flagged program, in
# solve order: 200 Dirichlet(1) starts); rng_probe, seeded seed + 1 = 202604
# (config probe_seed; per delta, group, indicator: 12 Dirichlet(1) starts).
# The local-polish diagnostic draws nothing.
# =====================================================================
print("\n-- generating and solving Case III (fig6-8_bdt_case.xlsx) ...")
from scipy.optimize import minimize as _sp_minimize, linprog as _sp_linprog
from scipy.special import gamma as _sp_gamma, gammainc as _sp_gammainc
from scipy import __version__ as _scipy_version

_BDT = dict(
    seed=NEW_SEED,
    N=15, K=3,
    scale_min=0.0, scale_max=5.0, n_levels=5,
    dim_names=("D1 Instrumented Asset", "D2 BDT Model", "D3 BDT Data",
               "D4 Connection & Interaction", "D5 Function & Service",
               "D6 System Perspective", "D7 Organization Perspective"),
    dim_kind=("tech", "tech", "tech", "tech", "tech", "system", "org"),
    dim_sizes=(3, 3, 4, 3, 3, 3, 3),                     # tab:bdt-instrument
    # generator: stated design inputs
    dim_design_level=(2.75, 2.65, 2.70, 2.55, 2.50, 2.40, 2.80),
    indicator_pattern_halfwidth=0.05,   # within a dimension the indicator targets run
                                        # from level-0.05 to level+0.05 (zero mean)
    stratum_offset=(-0.20, -0.05, +0.20),   # G1 mixed, G2 junior, G3 senior
    leniency_sd=0.10,                   # per-expert leniency, N(0, sd)
    noise_sd=0.08,                      # per-(expert, indicator) noise, N(0, sd)
    shape_range=(4.5, 6.0),             # kappa ~ U(a, b); the resulting member sds are in the config
    target_clip=(1.0, 4.5),             # target means clipped to this range
    # behavioral parameters and the fairness normalization
    rho=0.4, vartheta=0.2, power_kappa=2.0, bonferroni_order=(1, 1),
    # the intra-group program
    lambda_G=0.5,                       # weight-Gini regularizer coefficient
    deltas=(0.0, 0.1, 0.2, 0.3, 0.4, 0.5),
    operating_delta=0.5,
    # numerics
    L=257,                              # quantile grid, midpoints (l-0.5)/L
    L_fig=4001,                         # finer quantile grid on which the Figure 6.8 densities are drawn
    NX=501,                             # x grid on [0,5] for the overlap integrals (trapezoid)
    norm_sample=300,                    # Dirichlet(1) points for the min-max ranges (+ vertices + uniform)
    n_random_starts=12,                 # Dirichlet(1) starts per program (+ uniform + best sampled point,
                                        # the latter dropped when it IS the uniform point)
    slsqp_ftol=1e-10, slsqp_maxiter=300,
    followup_gap=1e-3,                  # a program whose runner-up original start falls short of its
                                        # best original start by more than this gets a follow-up search
    followup_starts=200,                # Dirichlet(1) starts of that follow-up (own generator)
    probe_starts=12,                    # a second, independent set of Dirichlet(1) starts for EVERY
                                        # program, from a generator of its own seeded seed + probe_seed_offset
    probe_seed_offset=1,
    near_tol=1e-5,                      # a run ending no more than this below the retained value counts
                                        # as near-optimal (the order of the solver's terminal resolution);
                                        # the weight stability is reported over those runs
    # the local-polish diagnostic (no draws; it changes no retained solution): rounds of a
    # pairwise-transfer pattern search and one SLSQP run from every retained solution
    polish_step0=2e-2,                  # initial transfer step between two members' weights
    polish_step_min=1e-9,               # the search stops once the halved step falls to this
    polish_max_evals=6000,              # ... or once this many objective evaluations are spent
    polish_accept_tol=1e-14,            # a transfer is accepted when it raises the objective by more
    polish_round_tol=1e-12,             # rounds (pattern search, one SLSQP run) repeat until one gains no more
    polish_max_rounds=10,               # ... or this many rounds have run
    moment_M=4096, moment_s=4,          # reported moments: midpoint rule in t on M cells, with
                                        # p = t^s / (t^s + (1-t)^s) grading both ends of (0,1)
    # k-means and TOPSIS
    kmeans_restarts=20, kmeans_maxiter=100,
    topsis_weights=(1.0 / 3.0, 1.0 / 3.0, 1.0 / 3.0),
    # best-worst method, linear model (Rezaei 2016, Omega 64:126-130): best D1, worst D6
    bwm_best="D1", bwm_worst="D6",
    bwm_best_to_others=(1, 2, 1, 2, 2, 3, 2),
    bwm_others_to_worst=(3, 2, 3, 2, 2, 1, 2),
    # consistency: the input-based ratio of Liang, Brunelli and Rezaei (2020,
    # Omega 96:102175, eqs. 6-7), CR^I = max_j |a_Bj a_jW - a_BW| / (a_BW^2 - a_BW),
    # which does not depend on the weighting model; its threshold for a_BW = 3 is
    # 0.1667 for every number of criteria from 3 to 9 (their Table 3, row "3"),
    # a value that at this scale acceptable (ordinal-consistent) and unacceptable
    # (ordinal-inconsistent) systems both attain (their Section 5.2), so the
    # ordinal-consistency condition of their Definition 4 is tested as well.
    # (The optimum xi^L of the linear model is on the weight scale and is NOT
    # divided by the consistency index that Rezaei 2015 tabulates for the
    # nonlinear model.)
    bwm_CRI_threshold=0.1667,
)
_BDT_HI = _BDT["scale_max"]
_BDT_SIZES = (4, 6, 5)                  # designed strata; k-means must recover them
# stratum descriptions (Table tab:bdt-groups); the rule-12 checks test them
# against the k-means centroids of the attribute matrix below: G3 has the
# longest experience and the highest credentials, G2 the shortest experience,
# the lowest credentials and the most digital-twin projects
_BDT_STRATA = ("mixed experience", "junior, broad but shallow project record",
               "senior, long experience and high credentials")
_BDT_COLORS = ("blue", "amber", "teal")
# the stated 15 x 3 professional-attribute matrix: years of experience,
# credential level (1-5), digital-twin project count (rows e1..e15)
_BDT_ATTR = np.array([
    [12, 4, 3], [8, 3, 3], [14, 4, 4], [10, 4, 4],                        # e1-e4
    [3, 2, 6], [4, 2, 7], [2, 2, 6], [5, 3, 8], [3, 2, 6], [4, 3, 7],     # e5-e10
    [18, 5, 5], [15, 4, 5], [22, 5, 6], [17, 4, 5], [20, 4, 6],           # e11-e15
], float)
_BDT_ATTR_NAMES = ("years_experience", "credential_level", "dt_projects")


# ---- truncated Weibull on [0, 5]  (eq:bdt-weibull-trunc) ---------------------
def _bdt_H(x, k, lam):
    return 1.0 - np.exp(-(x / lam) ** k)


def _bdt_h(x, k, lam):
    x = np.asarray(x, float)
    with np.errstate(divide="ignore", invalid="ignore"):
        v = (k / lam) * (x / lam) ** (k - 1) * np.exp(-(x / lam) ** k)
    return np.where(x > 0, v, 0.0)


def _bdt_pdf(x, k, lam):
    return _bdt_h(x, k, lam) / _bdt_H(_BDT_HI, k, lam)


def _bdt_quantile(p, k, lam):
    return lam * (-np.log1p(-p * _bdt_H(_BDT_HI, k, lam))) ** (1.0 / k)


def _bdt_trunc_mean(k, lam):
    """closed form: lambda Gamma(1+1/k) P(1+1/k, (5/lambda)^k) / H(5)."""
    a = 1.0 + 1.0 / k
    z = (_BDT_HI / lam) ** k
    return lam * _sp_gamma(a) * _sp_gammainc(a, z) / (1.0 - np.exp(-z))


def _bdt_trunc_moment(k, lam, r):
    """closed form r-th raw moment of the truncated Weibull on [0, 5]:
    lambda^r Gamma(1+r/k) P(1+r/k, (5/lambda)^k) / H(5)."""
    a = 1.0 + r / k
    z = (_BDT_HI / lam) ** k
    return lam ** r * _sp_gamma(a) * _sp_gammainc(a, z) / (1.0 - np.exp(-z))


def _bdt_scale_for_mean(k, target):
    """vectorized bisection: the scale at which the TRUNCATED mean equals target."""
    lo = np.full_like(k, 0.5, dtype=float)
    hi = np.full_like(k, 8.0, dtype=float)
    for _ in range(60):
        mid = 0.5 * (lo + hi)
        below = _bdt_trunc_mean(k, mid) < target
        lo = np.where(below, mid, lo)
        hi = np.where(below, hi, mid)
    return 0.5 * (lo + hi)


# ---- 1. the panel ---------------------------------------------------------------
_rng_bdt_panel = np.random.default_rng(_BDT["seed"])
_bdt_N, _bdt_K = _BDT["N"], _BDT["K"]
_bdt_group_of = np.repeat(np.arange(_bdt_K), _BDT_SIZES)
_bdt_dim_of = np.repeat(np.arange(7), _BDT["dim_sizes"])
_bdt_M = len(_bdt_dim_of)
_bdt_ind_target = np.zeros(_bdt_M)
for _b in range(7):
    _idx = np.where(_bdt_dim_of == _b)[0]
    _hw = _BDT["indicator_pattern_halfwidth"]
    _bdt_ind_target[_idx] = _BDT["dim_design_level"][_b] + np.linspace(-_hw, _hw, len(_idx))
_bdt_leniency = _rng_bdt_panel.normal(0.0, _BDT["leniency_sd"], _bdt_N)
_bdt_noise = _rng_bdt_panel.normal(0.0, _BDT["noise_sd"], (_bdt_N, _bdt_M))
_bdt_shape = _rng_bdt_panel.uniform(_BDT["shape_range"][0], _BDT["shape_range"][1],
                                    (_bdt_N, _bdt_M))
_bdt_target = (_bdt_ind_target[None, :]
               + np.array(_BDT["stratum_offset"])[_bdt_group_of][:, None]
               + _bdt_leniency[:, None] + _bdt_noise)
_bdt_target = np.clip(_bdt_target, _BDT["target_clip"][0], _BDT["target_clip"][1])
_bdt_scale = _bdt_scale_for_mean(_bdt_shape, _bdt_target)
_bdt_P = (np.arange(_BDT["L"]) + 0.5) / _BDT["L"]
_bdt_X = np.linspace(0.0, _BDT_HI, _BDT["NX"])
_bdt_DX = float(_bdt_X[1] - _bdt_X[0])
# member standard deviations of the truncated densities, in closed form (the
# 257-level midpoint grid of the solver omits the tails and understates them)
_bdt_member_sd = np.sqrt(_bdt_trunc_moment(_bdt_shape, _bdt_scale, 2)
                         - _bdt_trunc_moment(_bdt_shape, _bdt_scale, 1) ** 2)
assert np.allclose(_bdt_trunc_mean(_bdt_shape, _bdt_scale), _bdt_target, atol=1e-6), \
    "BDT panel: the truncated means do not hit their targets"

# graded midpoint rule for the REPORTED moments of quantile averages: t on M
# midpoints, p = t^s / (t^s + (1-t)^s), weight dp/dt / M.  The grading clusters
# the nodes at both ends of (0,1), where the quantile functions are steep, and
# reproduces the closed-form member moments to 1e-12 (checked here).
_bdt_Mt = (np.arange(_BDT["moment_M"]) + 0.5) / _BDT["moment_M"]
_bdt_Ms = float(_BDT["moment_s"])
_bdt_Mp = _bdt_Mt ** _bdt_Ms / (_bdt_Mt ** _bdt_Ms + (1.0 - _bdt_Mt) ** _bdt_Ms)
_bdt_Mw = (_bdt_Ms * _bdt_Mt ** (_bdt_Ms - 1) * (1.0 - _bdt_Mt) ** (_bdt_Ms - 1)
           / (_bdt_Mt ** _bdt_Ms + (1.0 - _bdt_Mt) ** _bdt_Ms) ** 2 / _BDT["moment_M"])
assert (_bdt_Mp > 0).all() and (_bdt_Mp < 1).all() and abs(_bdt_Mw.sum() - 1.0) < 1e-12
# member quantile functions on the moment grid, (N, M_indicators, moment_M)
_bdt_QM = np.array([[_bdt_quantile(_bdt_Mp, _bdt_shape[_i, _m], _bdt_scale[_i, _m])
                     for _m in range(_bdt_M)] for _i in range(_bdt_N)])
assert np.isfinite(_bdt_QM).all()


def _bdt_mv(QM):
    """exact (graded-rule) mean and variance of quantile functions on the moment grid
    (last axis)."""
    mu = QM @ _bdt_Mw
    return mu, (QM * QM) @ _bdt_Mw - mu ** 2


def _bdt_mass_below(Q, level):
    """P(X < level) for the distribution whose quantile function Q is given on the
    moment grid, by inverting Q there (linear interpolation of p against Q).  Like
    the reported moments it uses the graded levels rather than the 257 solver
    levels, on which the probability could move only in steps of 1/257."""
    assert (np.diff(Q) > 0).all() and Q[0] < level < Q[-1]
    return float(np.interp(level, Q, _bdt_Mp))


_bdt_mm, _bdt_mvv = _bdt_mv(_bdt_QM)
assert np.abs(_bdt_mm - _bdt_trunc_mean(_bdt_shape, _bdt_scale)).max() < 1e-12
assert np.abs(np.sqrt(_bdt_mvv) - _bdt_member_sd).max() < 1e-12, \
    "BDT moment rule does not reproduce the closed-form member sd"
print("   panel: %d experts x %d indicators; member sd %.4f-%.4f (mean %.4f, closed form)"
      % (_bdt_N, _bdt_M, _bdt_member_sd.min(), _bdt_member_sd.max(), _bdt_member_sd.mean()))


# ---- 2. k-means on the attribute matrix, TOPSIS on the centroids ----------------
def _bdt_kmeans(Xr, K, rng, restarts, maxiter):
    """Lloyd's algorithm on standardized columns; `restarts` seeded random
    initializations (K distinct rows), the lowest inertia kept."""
    Z = (Xr - Xr.mean(0)) / Xr.std(0)
    best = None
    for r in range(restarts):
        c = Z[rng.choice(len(Z), K, replace=False)]
        for _ in range(maxiter):
            lab = np.argmin(((Z[:, None, :] - c[None, :, :]) ** 2).sum(-1), axis=1)
            newc = np.array([Z[lab == k].mean(0) if np.any(lab == k) else c[k]
                             for k in range(K)])
            if np.allclose(newc, c):
                break
            c = newc
        inertia = float(sum(((Z[lab == k] - c[k]) ** 2).sum() for k in range(K)))
        if best is None or inertia < best[0] - 1e-12:
            best = (inertia, lab.copy(), r)
    inertia, lab, r = best
    # clusters are numbered in order of first appearance along e1..e15
    order = []
    for l in lab:
        if l not in order:
            order.append(l)
    lab = np.array([order.index(l) for l in lab])
    return inertia, lab, r


_rng_bdt_kmeans = np.random.default_rng(_BDT["seed"])
_bdt_inertia, _bdt_lab, _bdt_best_restart = _bdt_kmeans(
    _BDT_ATTR, _bdt_K, _rng_bdt_kmeans, _BDT["kmeans_restarts"], _BDT["kmeans_maxiter"])
_bdt_sizes = np.bincount(_bdt_lab, minlength=_bdt_K)
assert tuple(int(s) for s in _bdt_sizes) == _BDT_SIZES, \
    "k-means did not recover the designed 4/6/5 strata: %r" % (_bdt_sizes,)
assert np.array_equal(_bdt_lab, _bdt_group_of), "k-means membership differs from the design"
_bdt_cent = np.array([_BDT_ATTR[_bdt_lab == k].mean(0) for k in range(_bdt_K)])
# TOPSIS: vector normalization, stated criterion weights, all benefit criteria
_bdt_Xn = _bdt_cent / np.sqrt((_bdt_cent ** 2).sum(0))
_bdt_Vt = _bdt_Xn * np.array(_BDT["topsis_weights"])
_bdt_ideal, _bdt_anti = _bdt_Vt.max(0), _bdt_Vt.min(0)
_bdt_dpos = np.sqrt(((_bdt_Vt - _bdt_ideal) ** 2).sum(1))
_bdt_dneg = np.sqrt(((_bdt_Vt - _bdt_anti) ** 2).sum(1))
_bdt_close = _bdt_dneg / (_bdt_dpos + _bdt_dneg)
_bdt_Omega = _bdt_close / _bdt_close.sum()
print("   k-means strata %s (restart %d, inertia %.3f); TOPSIS Omega = %s"
      % (tuple(int(s) for s in _bdt_sizes), _bdt_best_restart, _bdt_inertia,
         np.round(_bdt_Omega, 3)))


# ---- 4a. best-worst method, linear model (Rezaei 2016) -------------------------
def _bdt_bwm(aB, aW, best, worst):
    n = len(aB)
    c = np.zeros(n + 1); c[-1] = 1.0
    A_ub, b_ub = [], []
    for j in range(n):
        if j != best:
            for sgn in (1.0, -1.0):
                row = np.zeros(n + 1); row[best] = sgn; row[j] = -sgn * aB[j]; row[-1] = -1.0
                A_ub.append(row); b_ub.append(0.0)
        if j != worst:
            for sgn in (1.0, -1.0):
                row = np.zeros(n + 1); row[j] = sgn; row[worst] = -sgn * aW[j]; row[-1] = -1.0
                A_ub.append(row); b_ub.append(0.0)
    A_eq = np.r_[np.ones(n), 0.0][None, :]
    res = _sp_linprog(c, A_ub=np.array(A_ub), b_ub=np.array(b_ub), A_eq=A_eq, b_eq=[1.0],
                      bounds=[(0.0, None)] * (n + 1), method="highs")
    assert res.success, "BWM linear program failed: %s" % res.message
    return res.x[:n], float(res.x[-1])


_bdt_dim_ids = ["D%d" % (b + 1) for b in range(7)]
_bdt_dw, _bdt_xi = _bdt_bwm(np.array(_BDT["bwm_best_to_others"], float),
                            np.array(_BDT["bwm_others_to_worst"], float),
                            _bdt_dim_ids.index(_BDT["bwm_best"]),
                            _bdt_dim_ids.index(_BDT["bwm_worst"]))
_bdt_dw = _bdt_dw / _bdt_dw.sum()
# input-based consistency ratio (Liang, Brunelli and Rezaei 2020, eqs. 6-7)
_bdt_aB = np.array(_BDT["bwm_best_to_others"], float)
_bdt_aW = np.array(_BDT["bwm_others_to_worst"], float)
_bdt_bi, _bdt_wi = _bdt_dim_ids.index(_BDT["bwm_best"]), _bdt_dim_ids.index(_BDT["bwm_worst"])
assert _bdt_aB[_bdt_bi] == 1 and _bdt_aW[_bdt_wi] == 1 and _bdt_aB[_bdt_wi] == _bdt_aW[_bdt_bi], \
    "BWM vectors: a_BB = a_WW = 1 and a_BW read from both vectors must agree"
_bdt_aBW = float(_bdt_aB[_bdt_wi])
_bdt_prod = _bdt_aB * _bdt_aW                                   # a_Bj * a_jW
_bdt_CRI_j = (np.abs(_bdt_prod - _bdt_aBW) / (_bdt_aBW * _bdt_aBW - _bdt_aBW)
              if _bdt_aBW > 1 else np.zeros(7))
_bdt_CRI = float(_bdt_CRI_j.max())
_bdt_CRI_arg = [_bdt_dim_ids[b] for b in range(7) if _bdt_CRI_j[b] == _bdt_CRI_j.max()]


# ordinal consistency (Liang et al. 2020, Definition 4, eq. 8): the two vectors
# order the criteria alike, i.e. for every pair i, j either
# (a_Bi - a_Bj)(a_jW - a_iW) > 0, or a_Bi = a_Bj and a_iW = a_jW.  At a_BW = 3 the
# tabulated threshold 0.1667 is attained both by ordinal-consistent (acceptable)
# and by ordinal-inconsistent (unacceptable) systems (their Section 5.2), so a
# ratio EQUAL to it settles nothing by itself; acceptability rests on this test.
def _bdt_ordinal_violations(aB, aW):
    """unordered pairs {i, j} that violate the ordinal-consistency condition."""
    bad = []
    for i in range(len(aB)):
        for j in range(i + 1, len(aB)):
            if not ((aB[i] - aB[j]) * (aW[j] - aW[i]) > 0 or (aB[i] == aB[j] and aW[i] == aW[j])):
                bad.append((i, j))
    return bad


_bdt_ord_bad = _bdt_ordinal_violations(_bdt_aB, _bdt_aW)
_bdt_ordinal = int(len(_bdt_ord_bad) == 0)
assert _bdt_ordinal == 1, "BWM comparison vectors are ordinal-inconsistent: %r" % (_bdt_ord_bad,)
print("   BWM weights %s (linear-model xi^L %.4f; input-based CR^I %.4f at %s, threshold %.4f; "
      "ordinal-consistent %d)"
      % (np.round(_bdt_dw, 4), _bdt_xi, _bdt_CRI, ",".join(_bdt_CRI_arg), _BDT["bwm_CRI_threshold"],
         _bdt_ordinal))


# ---- 3. the intra-group program ----------------------------------------------------
_BDT_TIE_TOL = 1e-6      # a start counts as reaching the retained optimum within this objective gap


class _BdtProgram:
    """max_w delta*Phi~(w) - (1-delta)*V~(w) - lambda_G*G(w) on the simplex of one
    subgroup for one indicator (eq:bdt-intra)."""

    def __init__(self, members, ks, lams):
        self.members = members
        self.n = len(ks)
        self.Q = np.array([_bdt_quantile(_bdt_P, k, l) for k, l in zip(ks, lams)])   # n x L
        fq = np.array([_bdt_pdf(self.Q[i], ks[i], lams[i]) for i in range(self.n)])
        self.inv_fq = 1.0 / np.maximum(fq, 1e-300)
        self.fx = np.array([_bdt_pdf(_bdt_X, k, l) for k, l in zip(ks, lams)])     # n x NX

    def qa(self, w):
        return w @ self.Q

    def density_on_x(self, w):
        # f_QA(Q(p)) = 1 / (dQ/dp),  dQ/dp = sum_i w_i / f_i(Q_i(p))
        q = w @ self.Q
        dq = w @ self.inv_fq
        return np.interp(_bdt_X, q, 1.0 / dq, left=0.0, right=0.0)

    def overlaps(self, w):
        f = self.density_on_x(w)
        return _trapz(np.minimum(f[None, :], self.fx), dx=_bdt_DX, axis=1)

    def raw(self, w):
        q = w @ self.Q
        v = float(np.mean((q - q.mean()) ** 2))
        A = self.overlaps(w)
        Phi, _ = _fs_phi(A[None, :], _BDT["rho"], _BDT["vartheta"])
        return float(Phi[0]), v, A

    def set_range(self, rng):
        n = self.n
        S = rng.dirichlet(np.ones(n), _BDT["norm_sample"])
        S = np.vstack([S, np.eye(n), np.ones((1, n)) / n])
        vals = np.array([self.raw(w)[:2] for w in S])
        self.sample, self.sample_vals = S, vals
        self.phi_min, self.phi_max = float(vals[:, 0].min()), float(vals[:, 0].max())
        self.v_min, self.v_max = float(vals[:, 1].min()), float(vals[:, 1].max())

    def tilde(self, phi, v):
        pt = (phi - self.phi_min) / (self.phi_max - self.phi_min) if self.phi_max > self.phi_min else 0.0
        vt = (v - self.v_min) / (self.v_max - self.v_min) if self.v_max > self.v_min else 0.0
        return pt, vt

    def objective(self, w, delta):
        w = np.clip(w, 0.0, None)
        s = w.sum()
        w = w / s if s > 0 else np.ones(self.n) / self.n
        phi, v, _ = self.raw(w)
        pt, vt = self.tilde(phi, v)
        return delta * pt - (1.0 - delta) * vt - _BDT["lambda_G"] * _gini_vec(w)

    def _run(self, delta, w0):
        """one SLSQP run from w0; returns (terminal weights, objective, status, nit)."""
        cons = ({"type": "eq", "fun": lambda w: w.sum() - 1.0},)
        res = _sp_minimize(lambda w: -self.objective(w, delta), w0, method="SLSQP",
                           bounds=[(0.0, 1.0)] * self.n, constraints=cons,
                           options={"ftol": _BDT["slsqp_ftol"],
                                    "maxiter": _BDT["slsqp_maxiter"]})
        w = np.clip(res.x, 0.0, None)
        w /= w.sum()
        return w, self.objective(w, delta), int(res.status), int(res.nit)

    def solve(self, delta, rng):
        n = self.n
        starts = list(rng.dirichlet(np.ones(n), _BDT["n_random_starts"]))
        starts.append(np.ones(n) / n)
        sv = np.array([delta * self.tilde(p, v)[0] - (1.0 - delta) * self.tilde(p, v)[1]
                       - _BDT["lambda_G"] * _gini_vec(w)
                       for (p, v), w in zip(self.sample_vals, self.sample)])
        w_best_sampled = self.sample[int(np.argmax(sv))]
        # the sample contains the center, so the best sampled point can BE the
        # center; a start equal to an earlier one would only repeat that run, so
        # it is dropped and every statistic below counts DISTINCT starts (the
        # retained solution is unaffected: a repeated run returns the same value
        # and the strict comparison below never lets it replace the first)
        self.last_dup_dropped = any(np.array_equal(w_best_sampled, s) for s in starts)
        # (a dropped duplicate is always the center, which the text states)
        assert not self.last_dup_dropped or np.array_equal(w_best_sampled, np.ones(n) / n)
        if not self.last_dup_dropped:
            starts.append(w_best_sampled)
        best_w, best_f, best_s = None, -np.inf, -1
        fvals, stats, wends = [], [], []
        for s_idx, w0 in enumerate(starts):
            w, f, st, nit = self._run(delta, w0)
            fvals.append(f)
            stats.append((st, nit))
            wends.append(w)
            if f > best_f:
                best_f, best_w, best_s = f, w, s_idx
        # multi-start agreement (ch3 rem:nonconvexity-conscience): distinct starts
        # whose terminal value lies within _BDT_TIE_TOL of the best one (section 6
        # takes the same count against the RETAINED value, which a further start can
        # raise; the stability of the returned weights, which the same remark asks
        # for, is recorded there over every run ending within near_tol of the
        # retained value, a tolerance matched to the solver's accuracy)
        self.last_at_best = int(np.sum(np.abs(np.asarray(fvals) - best_f) < _BDT_TIE_TOL))
        # how far the runner-up distinct start falls short of the best one
        self.last_runner_up_gap = float(best_f - np.sort(np.asarray(fvals))[-2])
        # every terminal value and weight vector, for the near-optimal stability record
        # and the substitution checks of section 6
        self.last_fvals, self.last_wends = np.asarray(fvals), np.array(wends)
        # SLSQP exit codes (scipy 1.13): 0 = its convergence test met at accuracy ftol
        # (an objective change or a step below ftol with the constraints satisfied, or
        # its first-order optimality estimate below ftol); 9 = iteration cap
        self.last_n_capped = int(sum(1 for st, nit in stats if st == 9))
        self.last_n_other_exit = int(sum(1 for st, nit in stats if st not in (0, 9)))
        self.last_best_status, self.last_best_nit = stats[best_s]
        return best_w, best_f, best_s, len(starts)

    def followup(self, delta, rng, n_starts):
        """further Dirichlet(1) starts; (objective, weights, status, nit) of every run."""
        out = []
        for w0 in rng.dirichlet(np.ones(self.n), n_starts):
            w, f, st, nit = self._run(delta, w0)
            out.append((f, w, st, nit))
        return out


_rng_bdt_solve = np.random.default_rng(_BDT["seed"])
_bdt_progs = {}
for _k in range(_bdt_K):
    _mem = np.where(_bdt_group_of == _k)[0]
    for _m in range(_bdt_M):
        _Pm = _BdtProgram(_mem, _bdt_shape[_mem, _m], _bdt_scale[_mem, _m])
        _Pm.set_range(_rng_bdt_solve)
        _bdt_progs[(_k, _m)] = _Pm

_bdt_rows_w, _bdt_rows_ind, _bdt_rows_dim, _bdt_rows_sweep = [], [], [], []
_bdt_store = {}
import time as _time
# every program solved from its original distinct starts.  All programs are solved
# before any layer is built, because the solution retained in a program can come
# from the further starts of section 6; rng_solve is drawn in the same order as
# the solves: per delta, group, indicator.
_bdt_orig = {}           # (delta, k, m) -> record of the original distinct starts
for _delta in _BDT["deltas"]:
    _t1 = _time.time()
    for _k in range(_bdt_K):
        for _m in range(_bdt_M):
            _Pm = _bdt_progs[(_k, _m)]
            _w, _f, _s, _ns = _Pm.solve(_delta, _rng_bdt_solve)
            _bdt_orig[(_delta, _k, _m)] = dict(
                w=_w, f=_f, best_start=_s, n_starts=_ns, dup=int(_Pm.last_dup_dropped),
                at_best=_Pm.last_at_best, gap=_Pm.last_runner_up_gap,
                fvals=_Pm.last_fvals, wends=_Pm.last_wends,
                n_capped=_Pm.last_n_capped, n_other=_Pm.last_n_other_exit,
                status=_Pm.last_best_status, nit=_Pm.last_best_nit)
    print("   delta %.1f: %d programs solved from their original starts  [%.0fs]"
          % (_delta, _bdt_K * _bdt_M, _time.time() - _t1))
_bdt_keys = list(_bdt_orig)          # solve order: delta, group, indicator


def _bdt_key_name(key):
    d, k, m = key
    return "delta %.1f, G%d, q%d" % (d, k + 1, m + 1)


# ---- 6. multi-start evidence over distinct starts, and the retained solution ------
# (a) the follow-up search.  Every program whose runner-up original start falls short
# of its best original start by more than followup_gap is searched again from
# followup_starts further Dirichlet(1) starts drawn from a generator of its own (so
# no draw of rng_solve moves).  Against the best original start, the follow-up
# records how many further runs return its value (within tie_tol), how many end
# above it and by how much, and how far their weights lie from its weights, next to
# the distance of the runner-up start's weights (the basin it was left for).
_rng_bdt_follow = np.random.default_rng(_BDT["seed"])
_bdt_rows_follow, _bdt_rows_follow_runs = [], []
_bdt_follow_best = {}          # (delta, k, m) -> (value, weights, status, nit) of the best follow-up run
for _key in _bdt_keys:
    _o = _bdt_orig[_key]
    if _o["gap"] <= _BDT["followup_gap"]:
        continue
    _d, _k, _m = _key
    _Pm = _bdt_progs[(_k, _m)]
    _wr, _f0 = _o["w"], _o["f"]
    _t1 = _time.time()
    _fu = _Pm.followup(_d, _rng_bdt_follow, _BDT["followup_starts"])
    _fv = np.array([f for f, w, st, nit in _fu])
    _fwd = np.array([float(np.abs(w - _wr).max()) for f, w, st, nit in _fu])
    _ib = int(np.argmax(_fv))
    _match = np.abs(_fv - _f0) < _BDT_TIE_TOL
    _exceed = _fv > _f0 + _BDT_TIE_TOL
    _bdt_follow_best[_key] = _fu[_ib]
    _ro = np.argsort(_o["fvals"], kind="stable")[::-1]
    _bdt_rows_follow.append({
        "delta": _d, "group": "G%d" % (_k + 1), "indicator": "q%d" % (_m + 1),
        "original_best_objective": _f0, "original_runner_up_gap": _o["gap"],
        "original_starts_at_best": _o["at_best"],
        "runner_up_weight_diff": float(np.abs(_o["wends"][_ro[1]] - _wr).max()),
        "n_followup_starts": len(_fu),
        "n_match_original_best": int(_match.sum()),
        "n_exceed_original_best": int(_exceed.sum()),
        "exceed_min": float((_fv[_exceed] - _f0).min()) if _exceed.any() else np.nan,
        "exceed_max_weight_diff": float(_fwd[_exceed].max()) if _exceed.any() else np.nan,
        "best_followup_objective": float(_fv[_ib]),
        "best_minus_original_best": float(_fv[_ib] - _f0),
        "max_abs_weight_diff_best_vs_original_best": float(_fwd[_ib]),
        "n_runs_capped": int(sum(1 for f, w, st, nit in _fu if st == 9))})
    for _j, (_fj, _wj, _stj, _nitj) in enumerate(_fu):
        _bdt_rows_follow_runs.append({"delta": _d, "group": "G%d" % (_k + 1),
                                      "indicator": "q%d" % (_m + 1), "start": _j + 1,
                                      "objective": float(_fj), "minus_original_best": float(_fj - _f0),
                                      "max_abs_weight_diff_vs_original_best": float(_fwd[_j]),
                                      "status": int(_stj)})
    print("   follow-up %s: runner-up gap %.2e; of %d further starts %d return the best original value "
          "and %d end above it (by %.2e to %.2e, weights within %.2e of the best original start, "
          "runner-up start %.2f away)  [%.0fs]"
          % (_bdt_key_name(_key), _o["gap"], len(_fu),
             _bdt_rows_follow[-1]["n_match_original_best"], _bdt_rows_follow[-1]["n_exceed_original_best"],
             _bdt_rows_follow[-1]["exceed_min"], _bdt_rows_follow[-1]["best_minus_original_best"],
             _bdt_rows_follow[-1]["exceed_max_weight_diff"], _bdt_rows_follow[-1]["runner_up_weight_diff"],
             _time.time() - _t1))

# (b) the second set of starts and the retained solution.  EVERY program is solved
# again from probe_starts further Dirichlet(1) starts drawn, in solve order, from a
# generator seeded seed + probe_seed_offset, so that the starts are independent of
# the original ones and no draw of rng_solve or rng_followup moves.  How far these
# runs end above the best original start measures the accuracy of a multi-start
# solve of this size.  The RETAINED solution of a program is the best further run
# (second set or follow-up) wherever that run ends above the best original start by
# more than tie_tol, and the best original start otherwise (a further run within
# tie_tol of it counts as reaching it, as in the agreement counts).  The agreement
# statistics are taken against the retained value: starts_at_best counts the
# original distinct starts within tie_tol of it (none, when a further run is
# retained), and runner_up_gap is the retained value minus the best original start
# other than the retained run.  Over the original starts and the second set
# together, every run ending no more than near_tol below the retained value is
# near-optimal, and the weights and the group-opinion mean it returns are compared
# with the retained ones.
_t1 = _time.time()
_rng_bdt_probe = np.random.default_rng(_BDT["seed"] + _BDT["probe_seed_offset"])
_bdt_ret = {}                  # (delta, k, m) -> the retained solution and its multi-start record
_bdt_near_alt = {}             # (delta, k, m) -> the near-optimal weights farthest from the retained ones
for _key in _bdt_keys:
    _d, _k, _m = _key
    _Pm = _bdt_progs[(_k, _m)]
    _o = _bdt_orig[_key]
    _pr = _Pm.followup(_d, _rng_bdt_probe, _BDT["probe_starts"])
    _pv = np.array([f for f, w, st, nit in _pr])
    _pw = np.array([w for f, w, st, nit in _pr])
    _ib = int(np.argmax(_pv))
    # the retained solution
    _cands = [("second set",) + tuple(_pr[_ib])]
    if _key in _bdt_follow_best:
        _cands.append(("follow-up",) + tuple(_bdt_follow_best[_key]))
    _src, _fb, _wb, _stb, _nitb = max(_cands, key=lambda c: c[1])
    if _fb > _o["f"] + _BDT_TIE_TOL:
        _rt = dict(w=_wb, f=float(_fb), source=_src, status=int(_stb), nit=int(_nitb))
    else:
        _rt = dict(w=_o["w"], f=_o["f"], source="original starts", status=_o["status"], nit=_o["nit"])
    _f0, _wr = _rt["f"], _rt["w"]
    # agreement of the original distinct starts with the retained value
    _ov = _o["fvals"]
    _others = np.delete(_ov, _o["best_start"]) if _rt["source"] == "original starts" else _ov
    _rt["at_best"] = int(np.sum(np.abs(_ov - _f0) < _BDT_TIE_TOL))
    _rt["gap"] = float(_f0 - _others.max())
    # near-optimal runs over the original starts and the second set (the retained
    # weights themselves are the reference, at distance zero)
    _av, _aw = np.r_[_ov, _pv], np.vstack([_o["wends"], _pw])
    _near = _av >= _f0 - _BDT["near_tol"]
    _mus = _Pm.Q.mean(axis=1)                       # member means (solution grid)
    _nw = _aw[_near]
    _nwd = np.r_[0.0, np.abs(_nw - _wr).max(axis=1)]
    _nmd = np.r_[0.0, np.abs((_nw - _wr) @ _mus)]
    _ia = int(np.argmax(_nwd))
    _bdt_near_alt[_key] = _wr if _ia == 0 else _nw[_ia - 1]
    _rt["row"] = {
        "retained_source": _rt["source"],
        "original_best_objective": _o["f"],
        "original_runner_up_gap": _o["gap"],
        "further_best_minus_original_best": float(_fb - _o["f"]),
        "probe_n_starts": len(_pr),
        "probe_best_minus_original_best": float(_pv[_ib] - _o["f"]),
        "probe_n_match_original_best": int((np.abs(_pv - _o["f"]) < _BDT_TIE_TOL).sum()),
        "probe_n_exceed_original_best": int((_pv > _o["f"] + _BDT_TIE_TOL).sum()),
        "probe_n_runs_capped": int(sum(1 for f, w, st, nit in _pr if st == 9)),
        "probe_n_runs_other_exit": int(sum(1 for f, w, st, nit in _pr if st not in (0, 9))),
        "near_n_regular": int(_near[:len(_ov)].sum()),
        "near_n_probe": int(_near[len(_ov):].sum()),
        # near-optimal runs other than the retained run itself (a follow-up run, when
        # retained, is not among the runs compared)
        "near_n_other": int(_near.sum()) - int(_rt["source"] != "follow-up"),
        "near_max_weight_diff": float(_nwd.max()),
        "near_max_mean_diff": float(_nmd.max()),
        "near_alt_max_weight": float(_bdt_near_alt[_key].max())}
    _bdt_ret[_key] = _rt
print("   second set of starts: %d programs x %d starts  [%.0fs]; retained from further starts: %s"
      % (len(_bdt_keys), _BDT["probe_starts"], _time.time() - _t1,
         "; ".join("%s (%s, +%.2e)" % (_bdt_key_name(k), r["source"], r["row"]["further_best_minus_original_best"])
                   for k, r in _bdt_ret.items() if r["source"] != "original starts") or "none"))


# (b2) the local-polish diagnostic.  Random starts cannot detect an SLSQP run that
# stalls at a kink of the weight Gini, where the weights of two or three members
# nearly coincide, so from EVERY retained solution a deterministic local search (no
# draws) is run in rounds.  A round is a pairwise-transfer pattern search (a transfer
# of min(step, w_j) from member j to member i over the ordered pairs is accepted when
# it raises the objective by more than polish_accept_tol; the step starts at
# polish_step0 and is halved after every full pass that improves nothing, until it
# falls to polish_step_min or polish_max_evals evaluations are spent), then one SLSQP
# run from its end point, the better of the two end points closing the round.  One
# SLSQP run started at such a kink is sensitive to last-bit changes of its start (at
# delta 0.1, G1, q6, perturbations of 1e-15 move its end value by about 1e-5), so the
# rounds repeat until one gains no more than polish_round_tol (at most
# polish_max_rounds); the first round is recorded as well (polish_first_round_*).  The
# polish is a DIAGNOSTIC, not a retention step: no retained solution and no layer
# changes.  Its gain over the retained value (polish_gain) bounds from below how far a
# retained value lies short of a local optimum, and the substitution checks of (d)
# recompute every printed number at the polished weights.
def _bdt_pattern(P, delta, w, f):
    """the pairwise-transfer pattern search from w (objective f)."""
    w = np.array(w, float)
    n, step, ev = len(w), _BDT["polish_step0"], 0
    while step > _BDT["polish_step_min"] and ev < _BDT["polish_max_evals"]:
        improved = False
        for i in range(n):
            for j in range(n):
                if i == j:
                    continue
                t = min(step, w[j])
                if t <= 0:
                    continue
                w2 = w.copy()
                w2[i] += t
                w2[j] -= t
                f2 = P.objective(w2, delta)
                ev += 1
                if f2 > f + _BDT["polish_accept_tol"]:
                    w, f, improved = w2, f2, True
        if not improved:
            step /= 2
    return w, f, ev


def _bdt_polish(P, delta, w, f):
    """rounds of (pattern search, one SLSQP run from its end, the better end point kept)
    from w (objective f) until a round gains no more than polish_round_tol; returns the
    first round's (weights, value, end point) and the settled (weights, value, pattern
    evaluations in all, rounds, end point of the last round)."""
    w = np.array(w, float)
    first, ev_all = None, 0
    for rnd in range(1, _BDT["polish_max_rounds"] + 1):
        f_start = f
        w, f, ev = _bdt_pattern(P, delta, w, f)
        ev_all += ev
        ws, fs, _st, _nit = P._run(delta, w)
        end = "pattern search"
        if fs >= f:
            w, f, end = ws, fs, "SLSQP"
        if first is None:
            first = (w.copy(), f, end)
        if f - f_start <= _BDT["polish_round_tol"]:
            break
    return first, (w, f, ev_all, rnd, end)


_t1 = _time.time()
_bdt_polished = {}             # (delta, k, m) -> settled polished weights (diagnostic only)
for _key in _bdt_keys:
    _d, _k, _m = _key
    _Pm = _bdt_progs[(_k, _m)]
    _rt = _bdt_ret[_key]
    _fr0 = _Pm.objective(_rt["w"], _d)
    (_w1p, _f1p, _end1), (_wp, _fp, _evp, _rndp, _endp) = _bdt_polish(_Pm, _d, _rt["w"], _fr0)
    _bdt_polished[_key] = _wp
    _mus_p = _Pm.Q.mean(axis=1)
    _rt["row"].update({
        "polish_first_round_gain": float(_f1p - _fr0),
        "polish_first_round_max_weight_diff": float(np.abs(_w1p - _rt["w"]).max()),
        "polish_first_round_end_point": _end1,
        "polish_gain": float(_fp - _fr0),
        "polish_max_weight_diff": float(np.abs(_wp - _rt["w"]).max()),
        "polish_mean_diff": float(abs((_wp - _rt["w"]) @ _mus_p)),
        "polish_rounds": int(_rndp),
        "polish_n_evaluations": int(_evp),
        "polish_end_point": _endp})
_bdt_pg = np.array([_bdt_ret[k]["row"]["polish_gain"] for k in _bdt_keys])
_bdt_pg1 = np.array([_bdt_ret[k]["row"]["polish_first_round_gain"] for k in _bdt_keys])
print("   local polish from every retained solution  [%.0fs]: largest gain %.2e (%s; first round alone %.2e), "
      "%d programs gain more than tie_tol, largest weight move %.2e, at most %d rounds"
      % (_time.time() - _t1, _bdt_pg.max(), _bdt_key_name(_bdt_keys[int(np.argmax(_bdt_pg))]), _bdt_pg1.max(),
         int((_bdt_pg > _BDT_TIE_TOL).sum()),
         max(_bdt_ret[k]["row"]["polish_max_weight_diff"] for k in _bdt_keys),
         max(_bdt_ret[k]["row"]["polish_rounds"] for k in _bdt_keys)))

# (c) the layers, from the retained solutions: inter-group quantile averaging under
# Omega, then the indicator hierarchy (means on the solution grid; the reported
# variances on the moment grid)
for _delta in _BDT["deltas"]:
    _repQ = np.zeros((_bdt_K, _bdt_M, _BDT["L"]))
    _repM = np.zeros((_bdt_K, _bdt_M, _BDT["moment_M"]))           # same, on the moment grid
    _repW = {}
    _phis, _phis_t, _ginis = [], [], []
    for _k in range(_bdt_K):
        for _m in range(_bdt_M):
            _Pm = _bdt_progs[(_k, _m)]
            _o, _rt = _bdt_orig[(_delta, _k, _m)], _bdt_ret[(_delta, _k, _m)]
            _w, _f = _rt["w"], _rt["f"]
            _phi, _v, _A = _Pm.raw(_w)
            _pt, _vt = _Pm.tilde(_phi, _v)
            _g = _gini_vec(_w)
            _phis.append(_phi); _phis_t.append(_pt); _ginis.append(_g)
            _repQ[_k, _m] = _Pm.qa(_w)
            _repM[_k, _m] = _w @ _bdt_QM[_Pm.members, _m]
            _repW[(_k, _m)] = _w
            _row = {"delta": _delta, "group": "G%d" % (_k + 1), "indicator": "q%d" % (_m + 1),
                    "dimension": "D%d" % (_bdt_dim_of[_m] + 1)}
            for _i in range(_bdt_N):
                _row["w_e%d" % (_i + 1)] = (float(_w[list(_Pm.members).index(_i)])
                                           if _i in _Pm.members else np.nan)
            for _j, _i in enumerate(_Pm.members):
                _row["A_e%d" % (_i + 1)] = float(_A[_j])
            _row.update({"Phi": _phi, "Phi_tilde": _pt, "V_group": _v, "V_tilde": _vt,
                         "gini_w": _g, "objective": _f,
                         "Phi_min": _Pm.phi_min, "Phi_max": _Pm.phi_max,
                         "V_min": _Pm.v_min, "V_max": _Pm.v_max,
                         "n_starts": _o["n_starts"], "duplicate_start_dropped": _o["dup"],
                         "best_start": _o["best_start"], "starts_at_best": _rt["at_best"],
                         "runner_up_gap": _rt["gap"],
                         "n_runs_capped": _o["n_capped"], "n_runs_other_exit": _o["n_other"],
                         "best_run_status": _rt["status"], "best_run_nit": _rt["nit"],
                         "group_rep_mean": float(_repQ[_k, _m].mean()),
                         "V_group_exact": float(_bdt_mv(_repM[_k, _m])[1])})
            _row.update(_rt["row"])
            _bdt_rows_w.append(_row)
    _indQ = np.einsum("k,kml->ml", _bdt_Omega, _repQ)
    _dimQ = np.array([_indQ[_bdt_dim_of == b].mean(axis=0) for b in range(7)])   # equal indicator weights
    _globQ = _bdt_dw @ _dimQ
    _gmu = float(_globQ.mean())
    _gvar257 = float(np.mean((_globQ - _gmu) ** 2))
    _dmu = _dimQ.mean(axis=1)
    _imu = _indQ.mean(axis=1)
    _gm = _repQ.mean(axis=2)                                        # K x M group rep means
    _indM = np.einsum("k,kml->ml", _bdt_Omega, _repM)
    _dimM = np.array([_indM[_bdt_dim_of == b].mean(axis=0) for b in range(7)])
    _globM = _bdt_dw @ _dimM
    _gmu_e, _gvar = (float(v) for v in _bdt_mv(_globM))
    _dmu_e, _dvar = _bdt_mv(_dimM)
    _imu_e, _ivar = _bdt_mv(_indM)
    _gm_e, _vgroup = _bdt_mv(_repM)                                 # K x M exact group variances
    _mean_err = float(max(abs(_gmu_e - _gmu), np.abs(_dmu_e - _dmu).max(),
                          np.abs(_imu_e - _imu).max(), np.abs(_gm_e - _gm).max()))
    _mass2 = _bdt_mass_below(_globM, 2.0)                           # P(F_G < 2), moment grid
    _mass2_257 = float(np.mean(_globQ < 2.0))                       # share of the 257 solver levels
    _level = int(min(np.floor(_gmu) + 1, 5))
    _bdt_store[_delta] = dict(repQ=_repQ, repW=_repW, indQ=_indQ, dimQ=_dimQ, globQ=_globQ,
                              gmu=_gmu, gvar=_gvar, dmu=_dmu, dvar=_dvar, imu=_imu, gm=_gm,
                              vgroup=_vgroup)
    _rows_d = [r for r in _bdt_rows_w if r["delta"] == _delta]
    _bdt_rows_sweep.append({"delta": _delta, "global_maturity": _gmu, "aggregate_variance": _gvar,
                            "aggregate_variance_L257": _gvar257,
                            "fairness_Phi": float(np.mean(_phis)),
                            "fairness_Phi_tilde": float(np.mean(_phis_t)),
                            "weight_gini": float(np.mean(_ginis)),
                            "mean_group_variance": float(np.mean(_vgroup)),
                            "max_group_variance": float(np.max(_vgroup)),
                            "max_mean_error_L257": _mean_err,
                            "mass_below_2": _mass2, "mass_below_2_L257": _mass2_257,
                            "level_band": _level,
                            "n_vertex_solutions": int(sum(1 for (k, m), w in _repW.items() if w.max() > 0.999)),
                            "min_starts_at_best": int(min(r["starts_at_best"] for r in _rows_d)),
                            "median_starts_at_best": float(np.median([r["starts_at_best"] for r in _rows_d])),
                            "max_runner_up_gap": float(max(r["runner_up_gap"] for r in _rows_d)),
                            "n_retained_from_further_starts": int(sum(r["retained_source"] != "original starts"
                                                                      for r in _rows_d)),
                            "n_duplicate_start_dropped": int(sum(r["duplicate_start_dropped"] for r in _rows_d)),
                            "n_runs": int(sum(r["n_starts"] for r in _rows_d)),
                            "n_runs_capped": int(sum(r["n_runs_capped"] for r in _rows_d)),
                            "n_runs_other_exit": int(sum(r["n_runs_other_exit"] for r in _rows_d))})
    for _m in range(_bdt_M):
        _bdt_rows_ind.append({"delta": _delta, "indicator": "q%d" % (_m + 1),
                              "dimension": "D%d" % (_bdt_dim_of[_m] + 1),
                              "G1_mean": float(_gm[0, _m]), "G2_mean": float(_gm[1, _m]),
                              "G3_mean": float(_gm[2, _m]),
                              "collective_mean": float(_imu[_m]), "collective_variance": float(_ivar[_m])})
    for _b in range(7):
        _bdt_rows_dim.append({"delta": _delta, "dimension": _bdt_dim_ids[_b], "weight": float(_bdt_dw[_b]),
                              "mean": float(_dmu[_b]), "variance": float(_dvar[_b])})
    print("   delta %.1f  maturity %.4f  V %.4f  Phi %.4f  Gini(w) %.4f  q1 groups %s -> %.4f"
          % (_delta, _gmu, _gvar, np.mean(_phis), np.mean(_ginis), np.round(_gm[:, 0], 3), _imu[0]))

_bdt_level_names = ["Digital Model", "Digital Shadow", "Primary BDT", "Enhanced BDT", "Ideal BDT"]


# (d) substitution checks.  Every layer the text prints is recomputed from one
# weight vector per program exactly as in (c), (i) at the retained weights (which
# must reproduce the recorded values), (ii) with every program's weights replaced at
# once by its near-optimal run farthest from them, (iii) with every program's
# weights replaced by its best ORIGINAL start (which differs from the retained
# solution only where a further run is retained), and (iv) with every program's
# weights replaced by its polished point of (b2).
def _bdt_layers(repW):
    repQ = np.zeros((_bdt_K, _bdt_M, _BDT["L"]))
    repM = np.zeros((_bdt_K, _bdt_M, _BDT["moment_M"]))
    phis, ginis = [], []
    for k in range(_bdt_K):
        for m in range(_bdt_M):
            P, w = _bdt_progs[(k, m)], repW[(k, m)]
            repQ[k, m] = P.qa(w)
            repM[k, m] = w @ _bdt_QM[P.members, m]
            phis.append(P.raw(w)[0])
            ginis.append(_gini_vec(w))
    indQ = np.einsum("k,kml->ml", _bdt_Omega, repQ)
    dimQ = np.array([indQ[_bdt_dim_of == b].mean(axis=0) for b in range(7)])
    globQ = _bdt_dw @ dimQ
    indM = np.einsum("k,kml->ml", _bdt_Omega, repM)
    dimM = np.array([indM[_bdt_dim_of == b].mean(axis=0) for b in range(7)])
    globM = _bdt_dw @ dimM
    return dict(gmu=float(globQ.mean()), gvar=float(_bdt_mv(globM)[1]),
                dmu=dimQ.mean(axis=1), gm=repQ.mean(axis=2),
                vgroup_max=float(np.max(_bdt_mv(repM)[1])),
                phi=float(np.mean(phis)), gini=float(np.mean(ginis)),
                mass2=_bdt_mass_below(globM, 2.0),
                n_vertex=int(sum(1 for w in repW.values() if w.max() > 0.999)))


def _bdt_printed(Ls):
    """the Case III numbers Section 6.4 prints, formatted as printed, from the layers
    at every delta (Ls: delta -> _bdt_layers(...))."""
    d0, dop = _BDT["deltas"][0], _BDT["operating_delta"]
    z, op = Ls[d0], Ls[dop]
    gms = [Ls[d]["gmu"] for d in _BDT["deltas"]]
    gvs = [Ls[d]["gvar"] for d in _BDT["deltas"]]
    out = {
        "global maturity, delta 0": "%.2f" % z["gmu"],
        "global maturity, delta 0.5": "%.2f" % op["gmu"],
        "global maturity to one decimal, delta 0.5": "%.1f" % op["gmu"],
        "global level, delta 0.5": _bdt_level_names[int(min(np.floor(op["gmu"]), 4))],
        "maturity change over the sweep": "%.2f" % (op["gmu"] - z["gmu"]),
        "maturity range over the sweep below 0.1": str(max(gms) - min(gms) < 0.1),
        "collective fairness, delta 0": "%.3f" % z["phi"],
        "collective fairness, delta 0.5": "%.3f" % op["phi"],
        "weight Gini, delta 0": "%.2f" % z["gini"],
        "weight Gini, delta 0.5": "%.2f" % op["gini"],
        "aggregate variance, delta 0": "%.3f" % z["gvar"],
        "aggregate variance, delta 0.5": "%.3f" % op["gvar"],
        "aggregate variance range over the sweep": "%.2f to %.2f" % (min(gvs), max(gvs)),
        "largest subgroup variance, delta 0.5": "%.3f" % op["vgroup_max"],
        "mass below level two, delta 0.5": "%.2f" % op["mass2"],
        "single-member programs, delta 0": "%d" % z["n_vertex"],
        "single-member programs, delta 0.1": "%d" % Ls[0.1]["n_vertex"],
        "single-member programs from delta 0.2 on": "%d" % max(Ls[d]["n_vertex"] for d in _BDT["deltas"] if d >= 0.2),
        "dimensions in increasing order of mean, delta 0.5": ",".join(
            _bdt_dim_ids[b] for b in np.argsort(op["dmu"], kind="stable")),
        "q1 collective mean, delta 0.5": "%.2f" % float(np.dot(_bdt_Omega, op["gm"][:, 0])),
    }
    for b in range(7):
        out["%s mean, delta 0.5" % _bdt_dim_ids[b]] = "%.2f" % op["dmu"][b]
        out["%s level, delta 0.5" % _bdt_dim_ids[b]] = _bdt_level_names[int(min(np.floor(op["dmu"][b]), 4))]
    for k in range(_bdt_K):
        out["q1 G%d mean, delta 0.5" % (k + 1)] = "%.2f" % op["gm"][k, 0]
    return out


_t1 = _time.time()
_bdt_scen_w = {"retained": {}, "near": {}, "original": {}, "polished": {}}
for _d in _BDT["deltas"]:
    _base = _bdt_store[_d]["repW"]
    _bdt_scen_w["retained"][_d] = dict(_base)
    _bdt_scen_w["near"][_d] = {_km: _bdt_near_alt[(_d,) + _km] for _km in _base}
    _bdt_scen_w["original"][_d] = {_km: _bdt_orig[(_d,) + _km]["w"] for _km in _base}
    _bdt_scen_w["polished"][_d] = {_km: _bdt_polished[(_d,) + _km] for _km in _base}
_bdt_scen_L = {s: {d: _bdt_layers(W) for d, W in ws.items()} for s, ws in _bdt_scen_w.items()}
# the recomputation reproduces the recorded layers at the retained weights
for _d in _BDT["deltas"]:
    _Lr, _st = _bdt_scen_L["retained"][_d], _bdt_store[_d]
    assert abs(_Lr["gmu"] - _st["gmu"]) < 1e-12 and abs(_Lr["gvar"] - _st["gvar"]) < 1e-12
    assert np.abs(_Lr["dmu"] - _st["dmu"]).max() < 1e-12 and np.abs(_Lr["gm"] - _st["gm"]).max() < 1e-12
    assert abs(_Lr["vgroup_max"] - float(np.max(_st["vgroup"]))) < 1e-12
    _sw = [r for r in _bdt_rows_sweep if r["delta"] == _d][0]
    assert abs(_Lr["phi"] - _sw["fairness_Phi"]) < 1e-12 and abs(_Lr["gini"] - _sw["weight_gini"]) < 1e-12
    assert abs(_Lr["mass2"] - _sw["mass_below_2"]) < 1e-12 and _Lr["n_vertex"] == _sw["n_vertex_solutions"]
_bdt_pr = {s: _bdt_printed(L) for s, L in _bdt_scen_L.items()}
_bdt_rows_swap = []
for _q in _bdt_pr["retained"]:
    _bdt_rows_swap.append({"quantity": _q, "printed_retained": _bdt_pr["retained"][_q],
                           "printed_near_substituted": _bdt_pr["near"][_q],
                           "printed_original_best": _bdt_pr["original"][_q],
                           "printed_polished": _bdt_pr["polished"][_q],
                           "unchanged_near": int(_bdt_pr["near"][_q] == _bdt_pr["retained"][_q]),
                           "unchanged_original": int(_bdt_pr["original"][_q] == _bdt_pr["retained"][_q]),
                           "unchanged_polished": int(_bdt_pr["polished"][_q] == _bdt_pr["retained"][_q])})
# unrounded effect of each substitution on the plotted sweep and the group-opinion means
_bdt_swap_eff = {}
for _s in ("near", "original", "polished"):
    _L, _R = _bdt_scen_L[_s], _bdt_scen_L["retained"]
    _bdt_swap_eff[_s] = {
        "global_maturity": max(abs(_L[d]["gmu"] - _R[d]["gmu"]) for d in _BDT["deltas"]),
        "aggregate_variance": max(abs(_L[d]["gvar"] - _R[d]["gvar"]) for d in _BDT["deltas"]),
        "fairness_Phi": max(abs(_L[d]["phi"] - _R[d]["phi"]) for d in _BDT["deltas"]),
        "weight_gini": max(abs(_L[d]["gini"] - _R[d]["gini"]) for d in _BDT["deltas"]),
        "largest_group_variance": max(abs(_L[d]["vgroup_max"] - _R[d]["vgroup_max"]) for d in _BDT["deltas"]),
        "dimension_mean": max(float(np.abs(_L[d]["dmu"] - _R[d]["dmu"]).max()) for d in _BDT["deltas"]),
        "group_opinion_mean": max(float(np.abs(_L[d]["gm"] - _R[d]["gm"]).max()) for d in _BDT["deltas"])}
# single-member (vertex) solutions that a near-optimal run does not reproduce: the
# retained weights put more than 0.999 on one member, the substituted ones do not
_bdt_vertex_lost = [(_d,) + _km for _d in _BDT["deltas"] for _km, _w in _bdt_scen_w["retained"][_d].items()
                    if _w.max() > 0.999 and _bdt_scen_w["near"][_d][_km].max() <= 0.999]
print("   substitution checks  [%.0fs]: printed numbers unchanged %d/%d (near-optimal), %d/%d (best original start), "
      "%d/%d (polished)"
      % (_time.time() - _t1, sum(r["unchanged_near"] for r in _bdt_rows_swap), len(_bdt_rows_swap),
         sum(r["unchanged_original"] for r in _bdt_rows_swap), len(_bdt_rows_swap),
         sum(r["unchanged_polished"] for r in _bdt_rows_swap), len(_bdt_rows_swap)))

_bdt_all = pd.DataFrame(_bdt_rows_w)
_bdt_one = _bdt_all[_bdt_all["starts_at_best"] == 1]
_bdt_none = _bdt_all[_bdt_all["starts_at_best"] == 0]
_bdt_further = _bdt_all[_bdt_all["retained_source"] != "original starts"]
_bdt_nwd_arg = _bdt_all.loc[_bdt_all["near_max_weight_diff"].idxmax()]
_bdt_praise_arg = _bdt_all.loc[_bdt_all["probe_best_minus_original_best"].idxmax()]
_bdt_fraise_arg = _bdt_all.loc[_bdt_all["further_best_minus_original_best"].idxmax()]
_bdt_pgain_arg = _bdt_all.loc[_bdt_all["polish_gain"].idxmax()]
_bdt_pwd_arg = _bdt_all.loc[_bdt_all["polish_max_weight_diff"].idxmax()]
_bdt_pmd_arg = _bdt_all.loc[_bdt_all["polish_mean_diff"].idxmax()]
_bdt_ms = {
    "multistart_n_programs": int(len(_bdt_all)),
    "multistart_n_runs": int(_bdt_all["n_starts"].sum()),
    "multistart_n_duplicate_start_dropped": int(_bdt_all["duplicate_start_dropped"].sum()),
    # agreement of the original distinct starts with the retained value
    "multistart_median_starts_at_best": float(_bdt_all["starts_at_best"].median()),
    "multistart_n_two_or_more_at_best": int((_bdt_all["starts_at_best"] >= 2).sum()),
    "multistart_n_one_at_best": int(len(_bdt_one)),
    "multistart_n_none_at_best": int(len(_bdt_none)),
    "multistart_max_gap_one_at_best": float(_bdt_one["runner_up_gap"].max()) if len(_bdt_one) else np.nan,
    "multistart_max_gap_none_at_best": float(_bdt_none["runner_up_gap"].max()) if len(_bdt_none) else np.nan,
    "multistart_n_followed_up": int(len(_bdt_rows_follow)),
    "multistart_n_runs_capped": int(_bdt_all["n_runs_capped"].sum()),
    "multistart_n_runs_other_exit": int(_bdt_all["n_runs_other_exit"].sum()),
    "multistart_n_best_run_nonzero_status": int((_bdt_all["best_run_status"] != 0).sum()),
    # the second set of starts against the best original start
    "probe_n_runs": int(_bdt_all["probe_n_starts"].sum()),
    "probe_n_programs_exceeded": int((_bdt_all["probe_n_exceed_original_best"] > 0).sum()),
    "probe_max_raise": float(_bdt_all["probe_best_minus_original_best"].max()),
    "probe_max_raise_argmax": "delta %.1f, %s, %s" % (
        _bdt_praise_arg["delta"], _bdt_praise_arg["group"], _bdt_praise_arg["indicator"]),
    "probe_n_runs_capped": int(_bdt_all["probe_n_runs_capped"].sum()),
    "probe_n_runs_other_exit": int(_bdt_all["probe_n_runs_other_exit"].sum()),
    # all further starts (second set and follow-up) against the best original start,
    # and the programs whose retained solution is a further run
    "further_n_programs_exceeded": int((_bdt_all["further_best_minus_original_best"] > _BDT_TIE_TOL).sum()),
    "further_max_raise": float(_bdt_all["further_best_minus_original_best"].max()),
    "further_max_raise_argmax": "delta %.1f, %s, %s" % (
        _bdt_fraise_arg["delta"], _bdt_fraise_arg["group"], _bdt_fraise_arg["indicator"]),
    "retained_n_from_further_starts": int(len(_bdt_further)),
    "retained_from_further_programs": ";".join(
        "delta %.1f, %s, %s (%s)" % (r["delta"], r["group"], r["indicator"], r["retained_source"])
        for _, r in _bdt_further.iterrows()) or "none",
    # the stability of the returned weights over the near-optimal runs
    "near_n_programs_with_other_near_run": int((_bdt_all["near_n_other"] >= 1).sum()),
    "near_weight_diff_max": float(_bdt_all["near_max_weight_diff"].max()),
    "near_weight_diff_median": float(_bdt_all["near_max_weight_diff"].median()),
    "near_weight_diff_argmax": "delta %.1f, %s, %s" % (
        _bdt_nwd_arg["delta"], _bdt_nwd_arg["group"], _bdt_nwd_arg["indicator"]),
    "near_n_programs_weight_diff_above_0.01": int((_bdt_all["near_max_weight_diff"] > 0.01).sum()),
    "near_mean_diff_max": float(_bdt_all["near_max_mean_diff"].max()),
    # the substitution checks (sheet multistart_substitution)
    "swap_n_printed_quantities": len(_bdt_rows_swap),
    "swap_near_n_printed_changed": int(sum(1 - r["unchanged_near"] for r in _bdt_rows_swap)),
    "swap_original_n_printed_changed": int(sum(1 - r["unchanged_original"] for r in _bdt_rows_swap)),
    "swap_polished_n_printed_changed": int(sum(1 - r["unchanged_polished"] for r in _bdt_rows_swap)),
    "swap_near_changed_quantities": ";".join(r["quantity"] for r in _bdt_rows_swap if not r["unchanged_near"])
                                    or "none",
    "swap_original_changed_quantities": ";".join(r["quantity"] for r in _bdt_rows_swap
                                                 if not r["unchanged_original"]) or "none",
    "swap_polished_changed_quantities": ";".join(r["quantity"] for r in _bdt_rows_swap
                                                 if not r["unchanged_polished"]) or "none",
    # the local-polish diagnostic (b2), against the retained value: a diagnostic, not a
    # retention step (config polish)
    "polish_n_programs_gain_above_tie_tol": int((_bdt_all["polish_gain"] > _BDT_TIE_TOL).sum()),
    "polish_max_gain": float(_bdt_all["polish_gain"].max()),
    "polish_max_gain_argmax": "delta %.1f, %s, %s" % (
        _bdt_pgain_arg["delta"], _bdt_pgain_arg["group"], _bdt_pgain_arg["indicator"]),
    "polish_max_weight_diff": float(_bdt_all["polish_max_weight_diff"].max()),
    "polish_max_weight_diff_argmax": "delta %.1f, %s, %s" % (
        _bdt_pwd_arg["delta"], _bdt_pwd_arg["group"], _bdt_pwd_arg["indicator"]),
    "polish_max_mean_diff": float(_bdt_all["polish_mean_diff"].max()),
    "polish_max_mean_diff_argmax": "delta %.1f, %s, %s" % (
        _bdt_pmd_arg["delta"], _bdt_pmd_arg["group"], _bdt_pmd_arg["indicator"]),
    "polish_n_programs_at_round_cap": int((_bdt_all["polish_rounds"] >= _BDT["polish_max_rounds"]).sum()),
    "polish_max_rounds_used": int(_bdt_all["polish_rounds"].max()),
    # the first round alone (pattern search and one SLSQP run), for comparison
    "polish_first_round_max_gain": float(_bdt_all["polish_first_round_gain"].max()),
    "polish_first_round_max_gain_argmax": "delta %.1f, %s, %s" % tuple(
        _bdt_all.loc[_bdt_all["polish_first_round_gain"].idxmax(), ["delta", "group", "indicator"]]),
    "polish_first_round_max_weight_diff": float(_bdt_all["polish_first_round_max_weight_diff"].max()),
    "polish_n_programs_gain_above_further_max_raise": int(
        (_bdt_all["polish_gain"] > _bdt_all["further_best_minus_original_best"].max()).sum()),
    # the accuracy of the retained values: the larger of what the further starts and the
    # local polish gain over them (the text states it as of the order of 1e-5)
    "retained_accuracy_bound": float(max(_bdt_all["polish_gain"].max(),
                                         _bdt_all["further_best_minus_original_best"].max())),
    "swap_near_n_vertex_lost": len(_bdt_vertex_lost),
    "swap_near_vertex_lost_programs": ";".join(_bdt_key_name(t) for t in _bdt_vertex_lost) or "none",
    "swap_near_vertex_lost_min_max_weight": float(min(
        (_bdt_near_alt[t].max() for t in _bdt_vertex_lost), default=np.nan)),
    # the normalized objective terms at the retained solutions (the ranges are sampled,
    # so the solver can pass beyond them)
    "retained_n_Phi_tilde_above_1": int((_bdt_all["Phi_tilde"] > 1.0).sum()),
    "retained_max_Phi_tilde": float(_bdt_all["Phi_tilde"].max()),
    "retained_min_V_tilde": float(_bdt_all["V_tilde"].min()),
    "retained_max_V_tilde": float(_bdt_all["V_tilde"].max()),
}
for _s in ("near", "original", "polished"):
    for _q, _v in _bdt_swap_eff[_s].items():
        _bdt_ms["swap_%s_max_abs_change_%s" % (_s, _q)] = float(_v)
print("   multi-start (distinct starts): %s" % _bdt_ms)

_bdt_op = _bdt_store[_BDT["operating_delta"]]
# Figure 6.8 data: q1 group representatives and the inter-group collective at the
# operating point, as quantile functions and as densities 1/(dQ/dp) on the x grid
# (the solved weights are applied on a finer quantile grid so that the drawn
# tails reach the ends of the support instead of stopping at the first and last
# midpoint of the solution grid)
_bdt_Pf = (np.arange(_BDT["L_fig"]) + 0.5) / _BDT["L_fig"]
_bdt_q1_Q = np.zeros((_bdt_K, _BDT["L_fig"]))
_bdt_q1_dQ = np.zeros((_bdt_K, _BDT["L_fig"]))
for _k in range(_bdt_K):
    _Pm = _bdt_progs[(_k, 0)]
    _w = _bdt_op["repW"][(_k, 0)]
    for _j, _i in enumerate(_Pm.members):
        _Qi = _bdt_quantile(_bdt_Pf, _bdt_shape[_i, 0], _bdt_scale[_i, 0])
        _bdt_q1_Q[_k] += _w[_j] * _Qi
        _bdt_q1_dQ[_k] += _w[_j] / np.maximum(_bdt_pdf(_Qi, _bdt_shape[_i, 0], _bdt_scale[_i, 0]), 1e-300)
assert np.allclose(_bdt_q1_Q.mean(axis=1), _bdt_op["repQ"][:, 0, :].mean(axis=1), atol=2e-3)
_bdt_q1_Qc = _bdt_Omega @ _bdt_q1_Q
_bdt_q1_dQc = _bdt_Omega @ _bdt_q1_dQ
_bdt_q1_f = [np.interp(_bdt_X, _bdt_q1_Q[k], 1.0 / _bdt_q1_dQ[k], left=0.0, right=0.0)
             for k in range(_bdt_K)]
_bdt_q1_fc = np.interp(_bdt_X, _bdt_q1_Qc, 1.0 / _bdt_q1_dQc, left=0.0, right=0.0)
_bdt_q1_mu = [float(_bdt_op["repQ"][k, 0, :].mean()) for k in range(_bdt_K)]   # means on the solution grid
_bdt_q1_muc = float(np.dot(_bdt_Omega, _bdt_q1_mu))            # QA mean identity
assert abs(_bdt_q1_muc - float(_bdt_q1_Qc.mean())) < 2e-3
for _k in range(_bdt_K):
    _mass = float(_trapz(_bdt_q1_f[_k], dx=_bdt_DX))
    _mx = float(_trapz(_bdt_X * _bdt_q1_f[_k], dx=_bdt_DX))
    assert abs(_mass - 1.0) < 0.01 and abs(_mx - _bdt_q1_mu[_k]) < 0.01, \
        "q1 density of G%d: mass %.4f mean %.4f vs %.4f" % (_k + 1, _mass, _mx, _bdt_q1_mu[_k])
assert abs(float(_trapz(_bdt_q1_fc, dx=_bdt_DX)) - 1.0) < 0.01
assert abs(float(_trapz(_bdt_X * _bdt_q1_fc, dx=_bdt_DX)) - _bdt_q1_muc) < 0.01

_bdt_glob_level = _bdt_level_names[int(min(np.floor(_bdt_op["gmu"]), 4))]
print("   operating point: global maturity %.4f (%s), V %.4f, Phi %.4f, Gini(w) %.4f, P(F_G < 2) %.3f"
      % (_bdt_op["gmu"], _bdt_glob_level, _bdt_op["gvar"],
         _bdt_rows_sweep[-1]["fairness_Phi"], _bdt_rows_sweep[-1]["weight_gini"],
         _bdt_rows_sweep[-1]["mass_below_2"]))

_bdt_cfg = {
    "case": "III -- building digital twin maturity (Section 6.4)",
    "figures": "fig6-8_bdt_q1agg, fig6-9_bdt_profile, fig6-10_bdt_sensitivity",
    "labels": "fig:bdt-q1agg, fig:bdt-profile, fig:bdt-sensitivity",
    "source_paper": "chen2024bdt",
    "synthetic": 1,
    "seed": NEW_SEED,
    "generator": ("build_workbooks.py Case III block: generated panel, k-means + TOPSIS, "
                  "BWM, multi-start SLSQP intra-group solves, quantile-averaged hierarchy"),
    "rng_order": ("rng_panel(seed): leniency[15], noise[15x22], shape[15x22]; "
                  "rng_kmeans(seed): 20 restarts x choice of 3 rows; "
                  "rng_solve(seed): per program (group-major, indicator) 300 Dirichlet(1) "
                  "normalization points, then per delta/group/indicator 12 Dirichlet(1) starts; "
                  "rng_followup(seed): per program with original_runner_up_gap > followup_gap, in "
                  "solve order, followup_starts Dirichlet(1) starts; "
                  "rng_probe(seed + probe_seed_offset = probe_seed = %d): per delta/group/indicator, "
                  "probe_starts Dirichlet(1) starts; the local polish draws nothing"
                  % (_BDT["seed"] + _BDT["probe_seed_offset"])),
    "n_experts": _bdt_N, "n_indicators": _bdt_M, "n_dimensions": 7, "n_groups": _bdt_K,
    "scale_min": _BDT["scale_min"], "scale_max": _BDT["scale_max"], "n_levels": _BDT["n_levels"],
    "density_family": "two-parameter Weibull truncated to [0,5], eq:bdt-weibull-trunc",
    "dim_design_level": ",".join("%.2f" % v for v in _BDT["dim_design_level"]),
    "indicator_pattern": "linspace(level-%.2f, level+%.2f, n_b) inside dimension b" % (
        _BDT["indicator_pattern_halfwidth"], _BDT["indicator_pattern_halfwidth"]),
    "stratum_offset_G1_G2_G3": ",".join("%+.2f" % v for v in _BDT["stratum_offset"]),
    "leniency_sd": _BDT["leniency_sd"], "noise_sd": _BDT["noise_sd"],
    "shape_range": "U(%.1f, %.1f)" % _BDT["shape_range"],
    "target_clip": "[%.1f, %.1f]" % _BDT["target_clip"],
    "scale_rule": "lambda solved by bisection so that the truncated mean equals the target",
    "member_sd_min": float(_bdt_member_sd.min()), "member_sd_max": float(_bdt_member_sd.max()),
    "member_sd_mean": float(_bdt_member_sd.mean()),
    "member_sd_rule": "closed form sqrt(m2 - m1^2), m_r = lambda^r Gamma(1+r/k) P(1+r/k, (5/lambda)^k) / H(5)",
    "reported_moments": ("variances and sds of every reported distribution (V_group_exact, "
                         "aggregate_variance, mean/max_group_variance, collective and dimension "
                         "variances, q1_variance) by the graded midpoint rule on the moment grid, "
                         "which reproduces the closed-form member moments to 1e-12; reported means "
                         "on the solution grid (their gap to the moment-grid means is "
                         "max_mean_error_L257 in the sensitivity sheet); V_group and "
                         "aggregate_variance_L257 are the 257-level values the solver uses"),
    "mass_below_2_rule": ("P(F_G < 2) (sensitivity sheet mass_below_2, printed in Section 6.4 as about "
                          "twelve percent at the operating point) by inverting the global quantile "
                          "function on the moment grid, linear interpolation of p against Q, the grid of "
                          "the reported moments; mass_below_2_L257 is the share of the 257 solver levels "
                          "below 2, which moves only in steps of 1/257"),
    "moment_grid":"p = t^s/(t^s+(1-t)^s), t = (j-1/2)/M, M = %d, s = %d, weight (dp/dt)/M" % (
        _BDT["moment_M"], _BDT["moment_s"]),
    "program": ("max_w delta*Phi~(w) - (1-delta)*V~(w) - lambda_G*G(w), w in the subgroup simplex; "
                "no consensus term (alpha = 0, beta = 1 - delta in eq:fair-scalarized)"),
    "Phi": "Bonferroni mean of order (1,1) of the power-normalized (kappa=2) Fehr-Schmidt utilities of the overlaps",
    "V": "variance of the quantile-averaged aggregate (quantile form, midpoint rule on the solution p grid)",
    "G": "Gini coefficient of the weight vector (influence-dispersion regularizer, Chapter 4)",
    "rho": _BDT["rho"], "vartheta": _BDT["vartheta"], "power_kappa": _BDT["power_kappa"],
    "bonferroni_order": "(1,1)",
    "lambda_G": _BDT["lambda_G"],
    "deltas": ",".join("%.1f" % d for d in _BDT["deltas"]),
    "operating_delta": _BDT["operating_delta"],
    "normalization": ("min-max per program over 300 Dirichlet(1) points + the vertices + the uniform "
                      "point; ranges in intra_group_solutions (Phi_min, Phi_max, V_min, V_max)"),
    "solver": ("scipy %s SLSQP, ftol %g, maxiter %d, forward-difference gradients; a run ends when "
               "SLSQP's convergence test is met at accuracy ftol (status 0: with the constraints "
               "satisfied, an objective change or a step below ftol, or the first-order change it "
               "predicts for the next step below ftol) or at the iteration cap (status 9); per program "
               "n_runs_capped, n_runs_other_exit, best_run_status" % (
                   _scipy_version, _BDT["slsqp_ftol"], _BDT["slsqp_maxiter"])),
    "starts": ("%d Dirichlet(1) random starts + the uniform point + the best sampled point, the last "
               "dropped when it equals the uniform point (distinct starts only, the ORIGINAL starts); "
               "the best original start is the solution unless a further start ends above it by more "
               "than tie_tol (see retention)" % (_BDT["n_random_starts"])),
    "n_random_starts": _BDT["n_random_starts"],
    "slsqp_ftol": _BDT["slsqp_ftol"], "slsqp_maxiter": _BDT["slsqp_maxiter"],
    "tie_tol": _BDT_TIE_TOL,
    "retention": ("retained solution = the best run of the further starts (second set, or the "
                  "follow-up where there is one) wherever it ends above the best original start by "
                  "more than tie_tol, else the best original start; retained_source in "
                  "intra_group_solutions ('original starts', 'second set', 'follow-up'), "
                  "original_best_objective and further_best_minus_original_best beside it; the local "
                  "polish (see polish) is a diagnostic, not a retention step, and changes no retained "
                  "solution"),
    "agreement": ("against the retained value: starts_at_best = original distinct starts whose "
                  "terminal value is within tie_tol of it (0 where a further run is retained); "
                  "runner_up_gap = retained value minus the best original start other than the "
                  "retained run; original_runner_up_gap = best minus runner-up original start"),
    "followup_gap": _BDT["followup_gap"], "followup_starts": _BDT["followup_starts"],
    "followup": ("programs with original_runner_up_gap > followup_gap re-searched from "
                 "followup_starts Dirichlet(1) starts (rng_followup); sheet multistart_followup, "
                 "against the best original start: n_match_original_best = runs within tie_tol of "
                 "its value, n_exceed_original_best = runs above it by more than tie_tol "
                 "(exceed_min, best_minus_original_best, exceed_max_weight_diff = largest "
                 "max|w - w(best original start)| over those runs), runner_up_weight_diff = the same "
                 "distance for the runner-up original start; the best follow-up run is retained when "
                 "it is the best further run and ends above the best original start by more than "
                 "tie_tol (see retention)"),
    "probe_starts": _BDT["probe_starts"], "probe_seed_offset": _BDT["probe_seed_offset"],
    # the seed of the second set of starts, which Sections 6.1 and 6.4 print
    "probe_seed": _BDT["seed"] + _BDT["probe_seed_offset"],
    "probe": ("a second set of probe_starts Dirichlet(1) starts for every program (rng_probe); "
              "probe_best_minus_original_best, probe_n_match_original_best, "
              "probe_n_exceed_original_best in intra_group_solutions; probe_max_raise = the largest "
              "amount by which a run of this set ends above the best original start; "
              "further_max_raise = the same over the second set and the follow-up together (the "
              "accuracy of a multi-start solve of this size)"),
    "polish_step0": _BDT["polish_step0"], "polish_step_min": _BDT["polish_step_min"],
    "polish_max_evals": _BDT["polish_max_evals"], "polish_accept_tol": _BDT["polish_accept_tol"],
    "polish_round_tol": _BDT["polish_round_tol"], "polish_max_rounds": _BDT["polish_max_rounds"],
    "polish": ("local-polish diagnostic from every retained solution, drawing nothing, in rounds: a "
               "pairwise-transfer pattern search (a transfer of min(step, w_j) from member j to member i "
               "over the ordered pairs, accepted when it raises the objective by more than "
               "polish_accept_tol; the step starts at polish_step0 and is halved after every full pass "
               "that improves nothing, until it falls to polish_step_min or polish_max_evals evaluations "
               "are spent), then one SLSQP run from its end point, the better end point closing the "
               "round; rounds repeat until one gains no more than polish_round_tol (at most "
               "polish_max_rounds), because one SLSQP run started where member weights nearly coincide "
               "is sensitive to last-bit changes of its start; per program in intra_group_solutions "
               "polish_gain = settled polished minus retained value, polish_max_weight_diff = "
               "max|w - w(retained)|, polish_mean_diff = |difference of the group-opinion mean| "
               "(solution grid), polish_rounds, polish_n_evaluations (pattern search, all rounds), "
               "polish_end_point, and the first round alone as polish_first_round_*; maxima and their "
               "programs as polish_max_* and polish_first_round_max_*; a diagnostic, not a retention "
               "step (the retained solutions are unchanged; substitution scenario 'polished', at the "
               "settled polished weights)"),
    "near_tol": _BDT["near_tol"],
    "near_optimal": ("over the original distinct starts and the second set together, the runs ending "
                     "no more than near_tol below the retained value; near_max_weight_diff = largest "
                     "|w_i - w_i(retained)| over them (0 when none differs), near_max_mean_diff = "
                     "largest difference of the group-opinion mean (solution grid), near_n_other = "
                     "those runs other than the retained run itself, near_alt_max_weight = largest "
                     "weight of the near-optimal run farthest from the retained weights (the one "
                     "substituted)"),
    "substitution": ("every printed Case III number recomputed after replacing, in all programs at "
                     "once, the retained weights by (near) the near-optimal run farthest from them, "
                     "(original) the best original start, which differs from the retained solution "
                     "only where a further run is retained, or (polished) the polished point of the "
                     "local-polish diagnostic; sheet multistart_substitution"),
    "quantile_grid": "L = %d midpoints (l-1/2)/L" % _BDT["L"],
    "figure_quantile_grid": "L_fig = %d midpoints for the q1_quantiles and q1_densities sheets (Figure 6.8)" % _BDT["L_fig"],
    "x_grid": "%d equispaced points on [0,5]; overlaps by the trapezoid rule" % _BDT["NX"],
    "qa_density": "f(Q(p)) = 1 / sum_i w_i / f_i(Q_i(p)), interpolated onto the x grid",
    "inter_group": "quantile averaging under Omega (TOPSIS on the k-means centroids)",
    "indicator_weights": "equal inside a dimension",
    "dimension_weights": "best-worst method, linear model (Rezaei 2016), best %s, worst %s" % (
        _BDT["bwm_best"], _BDT["bwm_worst"]),
    "bwm_best": _BDT["bwm_best"], "bwm_worst": _BDT["bwm_worst"],
    "bwm_best_to_others": ",".join(str(v) for v in _BDT["bwm_best_to_others"]),
    "bwm_others_to_worst": ",".join(str(v) for v in _BDT["bwm_others_to_worst"]),
    "bwm_xi": _bdt_xi,
    "bwm_xi_meaning": ("optimum xi^L of the linear model (Rezaei 2016): the largest deviation "
                       "|w_B - a_Bj w_j|, |w_j - a_jW w_W| on the weight scale; not a consistency ratio"),
    "bwm_a_BW": _bdt_aBW,
    "bwm_CRI": _bdt_CRI,
    "bwm_CRI_argmax": ",".join(_bdt_CRI_arg),
    "bwm_CRI_rule": ("input-based consistency ratio of Liang, Brunelli and Rezaei (2020), "
                     "max_j |a_Bj a_jW - a_BW| / (a_BW^2 - a_BW); per dimension in the dimensions sheet"),
    "bwm_CRI_threshold": _BDT["bwm_CRI_threshold"],
    "bwm_CRI_threshold_source": ("Liang, Brunelli and Rezaei (2020), Omega 96:102175, Table 3: "
                                 "0.1667 for a_BW = 3 at every number of criteria from 3 to 9. "
                                 "At a_BW = 3 this value is attained both by acceptable "
                                 "(ordinal-consistent) and by unacceptable (ordinal-inconsistent) "
                                 "systems (their Section 5.2), so a ratio equal to it does not by "
                                 "itself decide acceptability; see bwm_ordinal_consistent"),
    "bwm_ordinal_consistent": _bdt_ordinal,
    "bwm_ordinal_rule": ("Liang, Brunelli and Rezaei (2020), Definition 4, eq. 8: for all i, j, "
                         "(a_Bi - a_Bj)(a_jW - a_iW) > 0, or a_Bi = a_Bj and a_iW = a_jW; "
                         "bwm_ordinal_n_violations = pairs violating it"),
    "bwm_ordinal_n_violations": len(_bdt_ord_bad),
    "kmeans": ("Lloyd's algorithm on standardized attribute columns (population sd), %d seeded restarts, each "
               "initialized at %d distinct rows and run for at most %d Lloyd iterations, stopping when the "
               "centroids are unchanged (numpy allclose, default tolerances rtol 1e-5, atol 1e-8), an empty "
               "cluster keeping its previous centroid; the restart of minimum inertia is kept, a later restart "
               "replacing an earlier one only when it lowers the inertia by more than 1e-12") % (
        _BDT["kmeans_restarts"], _bdt_K, _BDT["kmeans_maxiter"]),
    "kmeans_restarts": _BDT["kmeans_restarts"],
    "kmeans_maxiter": _BDT["kmeans_maxiter"],
    "kmeans_convergence": "numpy.allclose(new centroids, old centroids) with default tolerances",
    "kmeans_best_restart": _bdt_best_restart, "kmeans_inertia": _bdt_inertia,
    "kmeans_labeling": "clusters numbered in order of first appearance along e1..e15",
    "topsis": "vector normalization, criterion weights %s, all benefit criteria, Omega = closeness / sum" % (
        ",".join("%.4f" % v for v in _BDT["topsis_weights"])),
    "attribute_columns": ",".join(_BDT_ATTR_NAMES),
    "global_maturity": _bdt_op["gmu"], "global_level": _bdt_glob_level,
    "aggregate_variance": _bdt_op["gvar"],
    "fairness_Phi": _bdt_rows_sweep[-1]["fairness_Phi"],
    "weight_gini": _bdt_rows_sweep[-1]["weight_gini"],
    "mass_below_2": _bdt_rows_sweep[-1]["mass_below_2"],
    "indicator": "q1 -- Response and Control",
    "band_lo": 2.0, "band_hi": 3.0,
    "x_label_q1": "maturity level, Response and Control ($q_1$)",
    "x_label_profile": "maturity level",
    "y_label_q1": "probability density",
    "x_label_sensitivity": "fairness weight $\\delta$ in the intra-group program",
    "y_label_sensitivity": "value",
    "profile_offset": 2.0,
    "grid_points": _BDT["NX"],
}
_bdt_cfg.update(_bdt_ms)
_bdt_panel_rows = []
for _i in range(_bdt_N):
    for _m in range(_bdt_M):
        _bdt_panel_rows.append({"expert": "e%d" % (_i + 1), "group": "G%d" % (_bdt_group_of[_i] + 1),
                                "indicator": "q%d" % (_m + 1), "dimension": "D%d" % (_bdt_dim_of[_m] + 1),
                                "indicator_design_target": float(_bdt_ind_target[_m]),
                                "target_mean": float(_bdt_target[_i, _m]),
                                "kappa": float(_bdt_shape[_i, _m]), "lambda": float(_bdt_scale[_i, _m]),
                                "truncated_mean": float(_bdt_trunc_mean(_bdt_shape[_i, _m], _bdt_scale[_i, _m])),
                                "truncated_sd": float(_bdt_member_sd[_i, _m])})
_bdt_expert_rows = [{"expert": "e%d" % (_i + 1),
                     _BDT_ATTR_NAMES[0]: _BDT_ATTR[_i, 0], _BDT_ATTR_NAMES[1]: _BDT_ATTR[_i, 1],
                     _BDT_ATTR_NAMES[2]: _BDT_ATTR[_i, 2],
                     "group_design": "G%d" % (_bdt_group_of[_i] + 1),
                     "group_kmeans": "G%d" % (_bdt_lab[_i] + 1),
                     "leniency": float(_bdt_leniency[_i])} for _i in range(_bdt_N)]
_bdt_topsis_rows = []
for _k in range(_bdt_K):
    _bdt_topsis_rows.append({"group": "G%d" % (_k + 1),
                             "centroid_" + _BDT_ATTR_NAMES[0]: float(_bdt_cent[_k, 0]),
                             "centroid_" + _BDT_ATTR_NAMES[1]: float(_bdt_cent[_k, 1]),
                             "centroid_" + _BDT_ATTR_NAMES[2]: float(_bdt_cent[_k, 2]),
                             "norm_" + _BDT_ATTR_NAMES[0]: float(_bdt_Xn[_k, 0]),
                             "norm_" + _BDT_ATTR_NAMES[1]: float(_bdt_Xn[_k, 1]),
                             "norm_" + _BDT_ATTR_NAMES[2]: float(_bdt_Xn[_k, 2]),
                             "d_plus": float(_bdt_dpos[_k]), "d_minus": float(_bdt_dneg[_k]),
                             "closeness": float(_bdt_close[_k]), "Omega": float(_bdt_Omega[_k])})
_bdt_groups_rows = [{"group": "G%d" % (_k + 1), "stratum": _BDT_STRATA[_k], "size": int(_bdt_sizes[_k]),
                     "Omega": float(_bdt_Omega[_k]), "q1_mu": _bdt_q1_mu[_k],
                     "q1_variance": float(_bdt_op["vgroup"][_k, 0]),
                     "color": _BDT_COLORS[_k]} for _k in range(_bdt_K)]
_bdt_dims_rows = [{"dimension": _BDT["dim_names"][_b], "id": _bdt_dim_ids[_b], "kind": _BDT["dim_kind"][_b],
                   "n_indicators": int(_BDT["dim_sizes"][_b]),
                   "mu": float(_bdt_op["dmu"][_b]), "variance": float(_bdt_op["dvar"][_b]),
                   "weight": float(_bdt_dw[_b]),
                   "level": _bdt_level_names[int(min(np.floor(_bdt_op["dmu"][_b]), 4))],
                   "bwm_best_to_others": _BDT["bwm_best_to_others"][_b],
                   "bwm_others_to_worst": _BDT["bwm_others_to_worst"][_b],
                   "bwm_product_aB_aW": float(_bdt_prod[_b]),
                   "bwm_CRI_j": float(_bdt_CRI_j[_b])} for _b in range(7)]
_bdt_q1_quant = pd.DataFrame({"p": _bdt_Pf, "Q_G1": _bdt_q1_Q[0], "Q_G2": _bdt_q1_Q[1], "Q_G3": _bdt_q1_Q[2],
                              "Q_collective": _bdt_q1_Qc})
_bdt_q1_dens = pd.DataFrame({"x": _bdt_X, "f_G1": _bdt_q1_f[0], "f_G2": _bdt_q1_f[1], "f_G3": _bdt_q1_f[2],
                             "f_collective": _bdt_q1_fc})
_bdt_glob_rows = [{"delta": d, "p": float(p), "Q_global": float(q)}
                  for d in _BDT["deltas"] for p, q in zip(_bdt_P, _bdt_store[d]["globQ"])]
write("fig6-8_bdt_case.xlsx", {
    "config": cfg(_bdt_cfg),
    "groups": _bdt_groups_rows,
    "dimensions": _bdt_dims_rows,
    "levels": [{"level": 1, "name": "Digital\nModel"}, {"level": 2, "name": "Digital\nShadow"},
               {"level": 3, "name": "Primary\nBDT"}, {"level": 4, "name": "Enhanced\nBDT"},
               {"level": 5, "name": "Ideal\nBDT"}],
    "sensitivity": _bdt_rows_sweep,
    "expert_attributes": _bdt_expert_rows,
    "kmeans_topsis": _bdt_topsis_rows,
    "panel_weibull": _bdt_panel_rows,
    "intra_group_solutions": _bdt_rows_w,
    "indicator_collectives": _bdt_rows_ind,
    "dimension_collectives": _bdt_rows_dim,
    "q1_quantiles": _bdt_q1_quant,
    "q1_densities": _bdt_q1_dens,
    "global_quantiles": _bdt_glob_rows,
    "multistart_followup": _bdt_rows_follow,
    "multistart_followup_runs": _bdt_rows_follow_runs,
    "multistart_substitution": _bdt_rows_swap,
})

print("\nAll workbooks written to", DATA)
if _environment is not None:
    _environment.workbooks_written()   # from here on every statement is a self-check

# =====================================================================
# SELF-CHECKS (audit rule 12) -- run on the workbooks just written, so
# they hold for the artifacts themselves, not only for in-memory values.
# Dimensions, weight normalizations, probability normalizations, and the
# invariants of the solved outputs.  Any failure aborts the build.
# =====================================================================
print("\n-- rule-12 self-checks on the written workbooks ...")


def _wb(name):
    return pd.read_excel(os.path.join(DATA, name), sheet_name=None)


def _cfgd(sheets, sheet="config"):
    c = sheets[sheet]
    return dict(zip(c["key"].astype(str), c["value"]))


def _sums_to_one(v, what, tol=1e-9):
    s = float(np.asarray(v, float).sum())
    assert abs(s - 1.0) < tol, "%s sums to %.9f, not 1" % (what, s)


# (0) every workbook carries a config or provenance sheet with generator+seed
for _f in sorted(os.listdir(DATA)):
    if not _f.endswith(".xlsx"):
        continue
    _s = _wb(_f)
    _sheet = "config" if "config" in _s else ("provenance" if "provenance" in _s else None)
    assert _sheet is not None, "%s has neither a config nor a provenance sheet" % _f
    _d = _cfgd(_s, _sheet)
    for _k in ("generator", "seed", "synthetic"):
        assert _k in _d, "%s: %s sheet lacks the %r key" % (_f, _sheet, _k)
    # the one seed of the pack is NEW_SEED, and a workbook without random draws
    # says so in its seed entry ("none -- ..."); every workbook is flagged synthetic
    _seedv = str(_d["seed"]).strip()
    assert _seedv.startswith("none") or _seedv in (str(NEW_SEED), "%d.0" % NEW_SEED), \
        "%s: seed %r is neither %d nor an explicit 'none -- ...' statement" % (_f, _seedv, NEW_SEED)
    assert int(float(_d["synthetic"])) == 1, "%s is not flagged synthetic" % _f
print("   provenance: every workbook names generator, seed (%d or an explicit 'none'), synthetic" % NEW_SEED)

# (0b) Claims that the builder's comments and the book make about the GENERATED
#      panels, tested on the written workbooks.  Each check reads its tolerance
#      or range from the builder constants or the config sheet, never from a
#      printed constant.
# Case II: the equal-weight quantile-averaged mean of every barrier meets the
#   calibration tolerance the config states (and the book prints), within the
#   stated number of refits; the guard inside the secant loop allows 5x.
_dt0 = _wb("fig6-5_dt_case.xlsx")
_dt0b = _dt0["barriers"]
_dt0gap = float(np.abs(_dt0b["equal_weight_mean"].to_numpy(float)
                       - _dt0b["design_mean"].to_numpy(float)).max())
assert _dt0gap < DT_CALIB_TOL, "Case II calibration gap %.2e exceeds the stated %g" % (_dt0gap, DT_CALIB_TOL)
assert (_dt0b["calibration_refits"].astype(int) <= DT_CALIB_MAXIT).all(), "a Case II calibration exceeds the refit cap"
assert ("within %g" % DT_CALIB_TOL) in str(_cfgd(_dt0)["target_calibration"]), \
    "the Case II config does not state the calibration tolerance the builder uses"
# Case III: every drawn Weibull shape lies in the configured range, and the
#   member-sd statistics the config records (and the text quotes) are those of
#   the stored panel.
_bdt0 = _wb("fig6-8_bdt_case.xlsx")
_bdt0c = _cfgd(_bdt0)
_bdt0p = _bdt0["panel_weibull"]
_bdt0lo, _bdt0hi = [float(v) for v in _wb_re.findall(r"\d+(?:\.\d+)?", str(_bdt0c["shape_range"]))]
assert ((_bdt0p["kappa"] >= _bdt0lo) & (_bdt0p["kappa"] <= _bdt0hi)).all(), \
    "a Case III shape lies outside the configured %s" % _bdt0c["shape_range"]
_bdt0sd = _bdt0p["truncated_sd"].to_numpy(float)
for _key, _val in (("member_sd_min", _bdt0sd.min()), ("member_sd_max", _bdt0sd.max()),
                   ("member_sd_mean", _bdt0sd.mean())):
    assert abs(float(_bdt0c[_key]) - float(_val)) < 1e-12, "Case III config %s != panel statistic" % _key
# Vignette V: replaying the draw order the config records under the recorded
#   seed reproduces the stored signal x_t and residual u_t, and shows that
#   forecast i minus its bias is m_t + x_t + tau_i xi_it exactly, xi being the
#   recorded forecast-noise draw, so the unobserved residual u_t (and hence the
#   realized y_t) enters no forecast (the identity self-check (6) describes).
_vg0 = _wb("fig6-12_hm_calsharp.xlsx")
_vg0c = _cfgd(_vg0)
_vg0O, _vg0F = _vg0["outcomes"], _vg0["forecasters"]
_vg0MU = _vg0["forecast_mu"].to_numpy(float)
_vg0T, _vg0N = _vg0MU.shape
_vg0rng = np.random.default_rng(int(float(_vg0c["seed"])))
_vg0rng.normal(0.0, 1.0)                                  # recorded draw (1): the pre-sample value
_vg0xu = _vg0rng.normal(0.0, 1.0, (_vg0T, 2))             # recorded draw (2): the (x_t, u_t) pairs
_vg0xi = _vg0rng.normal(0.0, 1.0, (_vg0T, _vg0N))         # recorded draw (3): the forecast noise
assert np.allclose(float(_vg0c["target_sigma_x"]) * _vg0xu[:, 0], _vg0O["leading_signal_x"], rtol=0, atol=1e-12), \
    "the vignette signal does not replay from the recorded seed and draw order"
assert np.allclose(float(_vg0c["target_sigma_u"]) * _vg0xu[:, 1], _vg0O["unobserved_residual_u"], rtol=0, atol=1e-12), \
    "the vignette residual does not replay from the recorded seed and draw order"
_vg0res = (_vg0MU - _vg0F["bias"].to_numpy(float)[None, :]
           - (_vg0O["conditional_mean_given_previous"].to_numpy(float)
              + _vg0O["leading_signal_x"].to_numpy(float))[:, None])
# (the workbook round trip may move the last bit, hence 1e-12 rather than equality)
assert np.allclose(_vg0res, _vg0F["error_scale_tau"].to_numpy(float)[None, :] * _vg0xi, rtol=0, atol=1e-12), \
    "a vignette forecast is not m_t + x_t + bias_i + tau_i xi_it under the recorded draw order"
print("   generator claims: Case II calibration within %g (largest gap %.1e); Case III shapes in %s,"
      " member-sd statistics as configured; vignette signal, residual and forecast noise replay"
      " from the seed, u_t enters no forecast" % (DT_CALIB_TOL, _dt0gap, _bdt0c["shape_range"]))

# (1) Chapter 2 running example: the written workbooks against the definitions
#     of chapter2.tex and the invariants of Algorithm 2.2.  Nothing here is
#     pinned to a printed constant; every expected value is recomputed from
#     the expert rows the workbooks carry.
def _ch2_overlap_from_rows(w, mus, sds):
    """def:consensus recomputed from workbook rows, each overlap exact (crossing points)."""
    m, s = float(w @ mus), float(w @ sds)
    return float(sum(_ch2_overlap_normal(mi, si, m, s) for mi, si in zip(mus, sds)))


def _ch2_overlap_from_rows_trapz(w, mus, sds):
    """def:consensus recomputed from workbook rows by an independent rule, the trapezoid
    rule on the recorded grid (cross-check of the exact form, tolerance 1e-6)."""
    m, s = float(w @ mus), float(w @ sds)
    fqa = _norm.pdf(CH2_X_GRID, m, s)
    return float(sum(_trapz(np.minimum(fqa, _norm.pdf(CH2_X_GRID, mi, si)), CH2_X_GRID) for mi, si in zip(mus, sds)))


# (1a) Table 2.1: the fits are the least squares of the elicited percentiles on
#      Phi^{-1}(p), rounded to one decimal, and the percentiles are the
#      one-decimal quantiles of the rounded fits (internal consistency)
_t21c = _wb("tab2-1_running_example_percentiles.xlsx")["table_2_1"]
_z5c = _norm.ppf([0.05, 0.25, 0.50, 0.75, 0.95])
for _, _r in _t21c.iterrows():
    _xh = _r[["p05", "p25", "p50", "p75", "p95"]].to_numpy(float)
    _c, *_ = np.linalg.lstsq(np.column_stack([np.ones(5), _z5c]), _xh, rcond=None)
    assert round(float(_c[0]), 1) == float(_r["fitted_mu"]) and round(float(_c[1]), 1) == float(_r["fitted_sigma"]), \
        "Table 2.1 %s: LS fit (%.3f, %.3f) does not round to the recorded (%.1f, %.1f)" % (
            _r["expert"], _c[0], _c[1], _r["fitted_mu"], _r["fitted_sigma"])
    assert abs(float(_c[0]) - float(_r["ls_mu_exact"])) < 1e-9 and abs(float(_c[1]) - float(_r["ls_sigma_exact"])) < 1e-9
    assert np.allclose(np.round(float(_r["fitted_mu"]) + float(_r["fitted_sigma"]) * _z5c, 1), _xh), \
        "Table 2.1 %s: elicited values are not the one-decimal quantiles of the fitted normal" % _r["expert"]
_t21_mu = _t21c["fitted_mu"].to_numpy(float)
_t21_sd = _t21c["fitted_sigma"].to_numpy(float)

# (1b) Figures 2.1 and 2.2: expert weights sum to one; the quantile-average row is
#      (w.mu, w.sigma) of the expert rows (prop:qa-properties(iii)); the
#      logarithmic-pool row is the normalized product of def:log-pool, checked both
#      by its precision-average closed form and by numerical normalization
for _f in ("fig2-1_linear_pool_vs_qa.xlsx", "fig2-2_three_pooling_operators.xlsx"):
    _g = _wb(_f)["gaussians"]
    _e = _g[_g["role"].astype(str).str.startswith("expert")]
    _sums_to_one(_e["weight"], "%s expert weights" % _f)
    _ew = _e["weight"].to_numpy(float); _em = _e["mu"].to_numpy(float); _es = _e["sigma"].to_numpy(float)
    _qa = _g[_g["legend"] == "quantile average"].iloc[0]
    assert abs(float(_qa["mu"]) - float(_ew @ _em)) < 1e-9 and abs(float(_qa["sigma"]) - float(_ew @ _es)) < 1e-9, \
        "%s: quantile-average row is not (w.mu, w.sigma) of the expert rows" % _f
    if "logarithmic pool" in set(_g["legend"].astype(str)):
        _lo = _g[_g["legend"] == "logarithmic pool"].iloc[0]
        _prec = float(np.sum(_ew / _es ** 2))
        assert abs(float(_lo["mu"]) - float(np.sum(_ew * _em / _es ** 2) / _prec)) < 1e-9
        assert abs(float(_lo["sigma"]) - _prec ** -0.5) < 1e-9
        _gx = np.exp(sum(wi * _norm.logpdf(CH2_X_GRID, mi, si) for wi, mi, si in zip(_ew, _em, _es)))
        _gx = _gx / _trapz(_gx, CH2_X_GRID)
        _mnum = float(_trapz(CH2_X_GRID * _gx, CH2_X_GRID))
        _snum = float(np.sqrt(_trapz((CH2_X_GRID - _mnum) ** 2 * _gx, CH2_X_GRID)))
        assert abs(_mnum - float(_lo["mu"])) < 1e-3 and abs(_snum - float(_lo["sigma"])) < 1e-3, \
            "%s: normalized product (%.4f, %.4f) disagrees with the recorded log pool" % (_f, _mnum, _snum)
        assert np.max(np.abs(_gx - _norm.pdf(CH2_X_GRID, float(_lo["mu"]), float(_lo["sigma"])))) < 1e-3
        # the Figure 2.2 panel is the rounded fit of Table 2.1
        assert np.allclose(_em, _t21_mu) and np.allclose(_es, _t21_sd), "Figure 2.2 experts differ from Table 2.1 fits"
    # the shapes the text describes, recomputed from the expert rows on a grid of its own:
    # the Figure 2.1 linear pool has two separated peaks with an almost empty trough, the
    # Figure 2.2 linear pool is a single hump, wider than every expert (by the location
    # term of its variance) and lower and broader than the two normal aggregates
    _c = _cfgd(_wb(_f))
    _x = np.linspace(float(_c["domain_min"]), float(_c["domain_max"]), 30001)
    _lp = sum(wi * _norm.pdf(_x, mi, si) for wi, mi, si in zip(_ew, _em, _es))
    _dl = np.diff(_lp)
    _nmodes = int(np.sum((_dl[:-1] > 0) & (_dl[1:] <= 0)))
    assert _nmodes == int(_c["linear_pool_modes"]), "%s: recorded linear-pool mode count is wrong" % _f
    if _f.startswith("fig2-1"):
        _modes = _x[1:-1][(_dl[:-1] > 0) & (_dl[1:] <= 0)]
        _trough = float(np.min(_lp[(_x >= _modes[0]) & (_x <= _modes[-1])]))
        assert _nmodes == 2 and _trough < 0.1 * float(np.max(_lp)), "Figure 2.1: the linear pool is not two separated peaks"
    else:
        _agg = _g[_g["role"] == "aggregate"]["sigma"].to_numpy(float)
        _lpsd = float(np.sqrt(_ew @ _es ** 2 + _ew @ (_em - _ew @ _em) ** 2))
        assert abs(_lpsd - float(_c["linear_pool_sd"])) < 1e-12
        assert _nmodes == 1, "Figure 2.2: the linear pool is not a single hump"
        assert _lpsd > max(float(np.max(_es)), float(np.max(_agg))), "Figure 2.2: the linear pool is not the widest density"
        assert float(np.max(_lp)) < 1.0 / (math.sqrt(2.0 * math.pi) * float(np.max(_agg))), "Figure 2.2: the linear pool is not the lowest aggregate"
        # wider than every expert because the spread of the locations adds to the average
        # variance: mixture variance = average expert variance + location variance, and the
        # average variance alone does not exceed the largest expert variance
        _avar, _lvar = float(_ew @ _es ** 2), float(_ew @ (_em - _ew @ _em) ** 2)
        assert abs(_avar - float(_c["linear_pool_average_expert_variance"])) < 1e-12 and abs(_lvar - float(_c["linear_pool_location_variance"])) < 1e-12
        assert abs(_avar + _lvar - _lpsd ** 2) < 1e-12, "Figure 2.2: mixture variance is not average variance plus location variance"
        assert _avar <= float(np.max(_es)) ** 2 < _avar + _lvar, "Figure 2.2: the location term is not what makes the pool wider than every expert"
print("   Chapter 2 figures: QA and log-pool rows agree with the definitions applied to the expert rows; linear-pool shapes as described")

# (1b') Example ex:heterogeneous: the moments recomputed from the stated inputs of the sheet
#       by an independent route (trapezoid rule in z = Phi^{-1}(p) on [-12, 12]), their
#       printed roundings, the strict variance bound, and the symmetry of the aggregate
_hx = _wb("fig2-2_three_pooling_operators.xlsx")["ex_heterogeneous"].set_index("quantity")
_hv = lambda q: float(_hx.loc[q, "value"])
_hlm, _hls, _hc = _hv("log-normal expert, mean of log x"), _hv("log-normal expert, sd of log x"), _hv("reflection point")
_hsm, _hss, _hw = _hv("sharp expert, mean"), _hv("sharp expert, sd"), _hv("weight of each expert")
assert abs(3.0 * _hw - 1.0) < 1e-12
_hz = np.linspace(-12.0, 12.0, 240001); _hpz = _norm.pdf(_hz)
_hQ = [np.exp(_hlm + _hls * _hz), 2.0 * _hc - np.exp(_hlm - _hls * _hz), _hsm + _hss * _hz]


def _ch2_zmoments(Q):
    m = float(_trapz(Q * _hpz, _hz))
    return m, float(_trapz((Q - m) ** 2 * _hpz, _hz))


_hm = [_ch2_zmoments(Q) for Q in _hQ]
_hqbar = _hw * sum(_hQ)
_hmu_qa, _hV_qa = _ch2_zmoments(_hqbar)
_hls_val = (_hw * sum(math.sqrt(v) for _, v in _hm)) ** 2
for _q, _val in (("log-normal expert, mean", _hm[0][0]), ("log-normal expert, sd", math.sqrt(_hm[0][1])),
                 ("reflected expert, mean", _hm[1][0]), ("reflected expert, sd", math.sqrt(_hm[1][1])),
                 ("mu_QA", _hmu_qa), ("V_QA", _hV_qa), ("(sum_i w_i sigma_i)^2", _hls_val)):
    assert abs(_hv(_q) - _val) < 1e-9, "ex:heterogeneous %s: recorded %.12f, recomputed %.12f" % (_q, _hv(_q), _val)
    # (the text cell is read back as a number by pandas, so compare at two decimals)
    assert "%.2f" % float(_hx.loc[_q, "printed"]) == _ch2_printed(_hv(_q)), "ex:heterogeneous %s: printed value is not the rounding" % _q
assert abs(_hv("log-normal expert, sd") - _hv("reflected expert, sd")) < 1e-12
assert _hV_qa < _hls_val - 1e-3, "ex:heterogeneous: the variance bound of thm:qa-moments is not strict"
assert np.max(np.abs(_hqbar + _hqbar[::-1] - 2.0 * _hc)) < 1e-9, "ex:heterogeneous: the aggregate is not symmetric about the reflection point"
assert abs(_hsm - _hc) < 1e-12 and abs(_hmu_qa - _hc) < 1e-9
print("   ex:heterogeneous: moments, strict variance bound and symmetry recomputed from the stated inputs")

# (1c) Table 2.5: the solve of Algorithm 2.2 read back from the workbook
_t25s = _wb("tab2-5_biobjective_example.xlsx")
_t25c = _t25s["table_2_5"]; _bnd = _t25s["bounds"].set_index("bound"); _lvl = _t25s["level_set_alpha_star"].set_index("point")
_swp = _t25s["alpha_sweep"].sort_values("alpha").reset_index(drop=True); _cfg25 = _cfgd(_t25s)
_alpha_s = float(_cfg25["alpha_star"])
_W25 = _t25c[["w1", "w2", "w3", "w4"]].to_numpy(float)
_Amin_c, _Amax_c = float(_bnd.loc["A_min", "value"]), float(_bnd.loc["A_max", "value"])
_Vmin_c, _Vmax_c = float(_bnd.loc["V_min", "value"]), float(_bnd.loc["V_max", "value"])


def _ch2_obj_from_rows(w, alpha=_alpha_s):
    _a = _ch2_overlap_from_rows(w, _t21_mu, _t21_sd); _v = float(w @ _t21_sd) ** 2
    return alpha * (_a - _Amin_c) / (_Amax_c - _Amin_c) - (1 - alpha) * (_v - _Vmin_c) / (_Vmax_c - _Vmin_c)


for _k, _r in _t25c.iterrows():
    _w = _W25[_k]
    _sums_to_one(_w, "Table 2.5 row %d weights" % _k)
    assert np.all(_w >= -1e-12)
    assert abs(float(_r["A"]) - _ch2_overlap_from_rows(_w, _t21_mu, _t21_sd)) < 1e-9, "Table 2.5 row %d: A is not the overlap of def:consensus" % _k
    assert abs(float(_r["A"]) - _ch2_overlap_from_rows_trapz(_w, _t21_mu, _t21_sd)) < 1e-6, "Table 2.5 row %d: A disagrees with the independent trapezoid rule" % _k
    assert abs(float(_r["V"]) - float(_w @ _t21_sd) ** 2) < 1e-12, "Table 2.5 row %d: V is not (w.sigma)^2" % _k
    assert abs(float(_r["fQA_mu"]) - float(_w @ _t21_mu)) < 1e-12 and abs(float(_r["fQA_sigma"]) - float(_w @ _t21_sd)) < 1e-12
    assert _Amin_c - 1e-9 <= float(_r["A"]) <= _Amax_c + 1e-9 and _Vmin_c - 1e-9 <= float(_r["V"]) <= _Vmax_c + 1e-9
# the confidence bounds are attained at the vertices of the sharpest and the most diffuse expert (V = (w.sigma)^2 with w.sigma linear)
assert abs(_Vmin_c - float(np.min(_t21_sd)) ** 2) < 1e-9 and abs(_Vmax_c - float(np.max(_t21_sd)) ** 2) < 1e-9
assert int(np.argmax(_bnd.loc["V_min", ["w1", "w2", "w3", "w4"]].to_numpy(float))) == int(np.argmin(_t21_sd))
# the consensus bounds are attained at vertices and bracket the vertex overlaps
_vertA = [_ch2_overlap_from_rows(np.eye(4)[i], _t21_mu, _t21_sd) for i in range(4)]
assert abs(_Amax_c - max(_vertA)) < 1e-9 and abs(_Amin_c - min(_vertA)) < 1e-9, "consensus bounds are not attained at the vertices found"
assert np.max(_bnd.loc["A_max", ["w1", "w2", "w3", "w4"]].to_numpy(float)) > 0.999
_emph = list(_t25c["emphasis"].astype(str))
assert _emph[1] == "maximum consensus" and abs(float(_t25c.loc[1, "A"]) - _Amax_c) < 1e-9
assert _emph[2] == "minimum variance" and abs(float(_t25c.loc[2, "V"]) - _Vmin_c) < 1e-9
# the alpha* row: its objective dominates every other row, every start, 2000 seeded
# simplex probes, and the endpoints of its level set; the level set is flat and its
# endpoints induce the same aggregate; the reported vector is their midpoint
_wstar = _W25[3]; _fstar = _ch2_obj_from_rows(_wstar)
assert abs(_fstar - float(_cfg25["alpha_star_objective"])) < 1e-9
for _k in range(3):
    assert _fstar >= _ch2_obj_from_rows(_W25[_k]) - 1e-12, "Table 2.5 row %d beats the alpha* row" % _k
_starts_c = _t25s["starts"][["w1", "w2", "w3", "w4"]].to_numpy(float)
assert len(_starts_c) == int(str(_cfg25["starts"]).split("plus")[0].split()[0]) + 5
assert len(CH2_PROBES) == int(str(_cfg25["probes"]).split()[0])
for _w0 in np.vstack([_starts_c, CH2_PROBES]):
    assert _fstar >= _ch2_obj_from_rows(_w0) - 1e-12, "a start or probe point beats the alpha* row"
_e1 = _lvl.loc["endpoint 1", ["w1", "w2", "w3", "w4"]].to_numpy(float)
_e2 = _lvl.loc["endpoint 2", ["w1", "w2", "w3", "w4"]].to_numpy(float)
assert np.min(_e1) < 1e-9 and np.min(_e2) < 1e-9, "level-set endpoints must lie on the boundary of the simplex"
assert abs(float(_e1 @ _t21_mu) - float(_e2 @ _t21_mu)) < 1e-9 and abs(float(_e1 @ _t21_sd) - float(_e2 @ _t21_sd)) < 1e-9
assert np.allclose(_wstar, 0.5 * (_e1 + _e2), atol=1e-9), "the reported alpha* vector is not the segment midpoint"
for _t in np.linspace(0, 1, 11):
    assert abs(_ch2_obj_from_rows((1 - _t) * _e1 + _t * _e2) - _fstar) < 1e-9, "the objective is not flat along the optimal segment"
assert np.max(_wstar) < 0.999, "the alpha* optimizer is not interior"
# the alpha sweep: the optimal value g(alpha) = max_w [alpha (A~ + V~) - V~] is a
# maximum of affine functions with nonnegative slopes, hence nondecreasing and
# convex in alpha, with g(0) = 0 and g(1) = 1 under min-max normalization
_ga = _swp["alpha"].to_numpy(float); _gv = _swp["objective"].to_numpy(float)
assert len(np.unique(np.round(_ga, 6))) == len(_ga), "alpha sweep has duplicate levels"
assert abs(_gv[0]) < 1e-9 and abs(_gv[-1] - 1.0) < 1e-9 and _ga[0] == 0.0 and _ga[-1] == 1.0
assert np.all(np.diff(_gv) >= -1e-9), "g(alpha) must be nondecreasing"
for _i in range(1, len(_ga) - 1):
    if _ga[_i + 1] > _ga[_i] > _ga[_i - 1]:
        _lin = _gv[_i - 1] + (_gv[_i + 1] - _gv[_i - 1]) * (_ga[_i] - _ga[_i - 1]) / (_ga[_i + 1] - _ga[_i - 1])
        assert _gv[_i] <= _lin + 1e-9, "g(alpha) must be convex (violated at alpha=%.2f)" % _ga[_i]
for _, _r in _swp.iterrows():
    _w = _r[["w1", "w2", "w3", "w4"]].to_numpy(float)
    _sums_to_one(_w, "alpha sweep row alpha=%.2f" % _r["alpha"])
    assert abs(float(_r["objective"]) - _ch2_obj_from_rows(_w, float(_r["alpha"]))) < 1e-9
_int = _swp[_swp["interior"] == 1]["alpha"].to_numpy(float)
assert len(_int) > 0 and _alpha_s in set(np.round(_int, 2)), "alpha* must lie in the interior window"
assert str(_cfg25["interior_alpha_window_on_grid"]) == "[%.2f, %.2f]" % (_int.min(), _int.max())
# the edges of the interior-compromise window: the recorded exact edges bracket the grid
# window and exclude the grid levels outside it, those levels return the two vertices the
# text names (minimum variance below, maximum consensus above), and the printed window is
# the exact one rounded to two decimals.  Each edge is re-derived by an independent route,
# one-sided difference quotients of the objective recomputed from the rows along the
# simplex directions that leave the vertex: their best value must be positive on the
# interior side of the edge and nonpositive on the vertex side.
_elo, _ehi = float(_cfg25["interior_alpha_lower_edge"]), float(_cfg25["interior_alpha_upper_edge"])
assert _elo < _int.min() and _int.max() < _ehi, "the grid window is not inside the exact window"
assert _ga[_ga < _int.min()].max() < _elo and _ga[_ga > _int.max()].min() > _ehi, "a vertex level of the grid lies inside the exact window"
assert str(_cfg25["interior_alpha_window_printed"]) == "[%s, %s]" % (_ch2_printed(_elo), _ch2_printed(_ehi)), \
    "the printed window is not the exact one rounded to two decimals"
_kmv = int(np.argmin(_t21_sd))
_kmc = int(np.argmax(_bnd.loc["A_max", ["w1", "w2", "w3", "w4"]].to_numpy(float)))
for _, _r in _swp.iterrows():
    _w = _r[["w1", "w2", "w3", "w4"]].to_numpy(float)
    if _r["alpha"] < _elo:
        assert _w[_kmv] > 0.999, "alpha=%.2f lies below the window but is not the minimum-variance vertex" % _r["alpha"]
    elif _r["alpha"] > _ehi:
        assert _w[_kmc] > 0.999, "alpha=%.2f lies above the window but is not the maximum-consensus vertex" % _r["alpha"]
    else:
        assert int(_r["interior"]) == 1 and np.max(_w) < 0.999, "alpha=%.2f lies inside the window but is a vertex" % _r["alpha"]
_edg = _t25s["window_edges"].set_index("edge")
assert float(_edg.loc["lower", "alpha_edge"]) == _elo and float(_edg.loc["upper", "alpha_edge"]) == _ehi
assert int(_edg.loc["lower", "vertex_expert"]) == _kmv + 1 and int(_edg.loc["upper", "vertex_expert"]) == _kmc + 1
for _side in ("lower", "upper"):
    assert abs(float(_edg.loc[_side, "alpha_edge"]) - float(_edg.loc[_side, "alpha_edge_ratio_cross_check"])) < 1e-8
    assert float(_edg.loc[_side, "bracket_lo"]) < float(_edg.loc[_side, "alpha_edge"]) < float(_edg.loc[_side, "bracket_hi"])


def _ch2_vertex_dq(k, alpha, n=80, h=1e-7):
    """largest one-sided difference quotient of the objective recomputed from the rows at the
    vertex e_k along the simplex directions lam - e_k, lam on a grid of the opposite face"""
    ek = np.eye(4)[k]; oth = [j for j in range(4) if j != k]
    f0 = _ch2_obj_from_rows(ek, alpha); best = -np.inf
    for i in range(n + 1):
        for j in range(n + 1 - i):
            lam = np.zeros(4); lam[oth[0]] = i / n; lam[oth[1]] = j / n; lam[oth[2]] = (n - i - j) / n
            best = max(best, (_ch2_obj_from_rows(ek + h * (lam - ek), alpha) - f0) / h)
    return best


_dq_off = 5e-4
assert _ch2_vertex_dq(_kmv, _elo - _dq_off) <= 0.0 < _ch2_vertex_dq(_kmv, _elo + _dq_off), "the lower edge is misplaced"
assert _ch2_vertex_dq(_kmc, _ehi - _dq_off) > 0.0 >= _ch2_vertex_dq(_kmc, _ehi + _dq_off), "the upper edge is misplaced"
# the recorded global checks at edge -/+ delta: the vertex outside the window, a strictly
# better interior optimizer inside it (objectives recomputed from the rows); and on the
# vertex side at the closer offset delta_vertex, the vertex again.  Since g(alpha) -
# f_vertex(alpha) is convex and nonnegative and the vertex is not locally optimal past the
# bisection edge (the difference quotients above), the global edge lies between that
# vertex-side level and the bisection edge, so both must print as the text prints the edge.
_edelta = float(_cfg25["edge_check_delta"]); _edv = float(_cfg25["edge_check_delta_vertex"])
assert 0.0 < _edv < _edelta
_ech = _t25s["window_edge_checks"]
assert len(_ech) == 6
for _, _r in _ech.iterrows():
    _w = _r[["w1", "w2", "w3", "w4"]].to_numpy(float); _al = float(_r["alpha"]); _off = float(_r["offset"])
    _sums_to_one(_w, "window edge check alpha=%.5f" % _al)
    _k = _kmv if _r["edge"] == "lower" else _kmc
    _e = _elo if _r["edge"] == "lower" else _ehi
    assert min(abs(_off - _edelta), abs(_off - _edv)) < 1e-15 and abs(abs(_al - _e) - _off) < 1e-12
    assert "%.2f" % float(_r["alpha_printed"]) == _ch2_printed(_al)
    _gap = _ch2_obj_from_rows(_w, _al) - _ch2_obj_from_rows(np.eye(4)[_k], _al)
    assert abs(_gap - float(_r["gap_to_vertex"])) < 1e-9
    if _elo < _al < _ehi:
        assert str(_r["expected"]) == "interior" and np.max(_w) < 0.999 and _gap > 1e-9, "alpha=%.5f: no interior optimizer beats the vertex" % _al
    else:
        assert str(_r["expected"]).startswith("vertex") and _w[_k] > 0.999 and abs(_gap) < 1e-9, "alpha=%.5f: the optimizer is not the vertex" % _al
_pw = str(_cfg25["interior_alpha_window_printed"]).strip("[]").split(", ")
for _side, _e, _sg, _p in (("lower", _elo, -1.0, _pw[0]), ("upper", _ehi, 1.0, _pw[1])):
    _er = _ech[_ech["edge"] == _side]
    _vt = _er[_er["expected"].astype(str).str.startswith("vertex")]
    assert len(_er) == 3 and len(_vt) == 2, "%s edge: expected one interior and two vertex-side global solves" % _side
    assert np.all(np.sign(_vt["alpha"].to_numpy(float) - _e) == _sg), "%s edge: a vertex-side solve lies on the interior side" % _side
    _avc = float(_vt["alpha"].to_numpy(float)[np.argmin(_vt["offset"].to_numpy(float))])
    assert abs(abs(_avc - _e) - _edv) < 1e-12
    assert _ch2_printed(_avc) == _ch2_printed(_e) == _p, \
        "%s edge: the global edge lies in [%.6f, %.6f], which does not print as %s throughout" % (_side, min(_avc, _e), max(_avc, _e), _p)
print("   Table 2.5: bounds at vertices, alpha* row dominates rows, starts, probes, and its level set; g(alpha) monotone convex;"
      " window edges (%.4f, %.4f) confirmed by difference quotients and the global solves, printed %s" % (
          _elo, _ehi, _cfg25["interior_alpha_window_printed"]))

# (2) BIM full assessment (Case I, sec:bim-elicit to sec:bim-results): the
#     GENERATED and SOLVED pipeline read back from fig6-1_bim_case.xlsx and
#     checked against the definitions it implements and against the claims
#     the text makes about it.  Expected values are recomputed from the rows
#     the workbook carries (attribute matrix, latent judgments, profiles,
#     comparison vectors), never pinned to printed constants; the overlaps are
#     recomputed by the independent Chapter 2 routine _ch2_overlap_normal.
_bim = _wb("fig6-1_bim_case.xlsx")
_bcf = _cfgd(_bim)
_bgr = _bim["groups"].reset_index(drop=True)
_bdm = _bim["dimensions"].reset_index(drop=True)
_bea = _bim["expert_attributes"]
_bpn = _bim["panel"]
_bsol = _bim["intra_group_solutions"]
_bic = _bim["indicator_collectives"].set_index("indicator")
_bbw = _bim["bwm"]
_bbs = _bim["bwm_summary"].set_index("problem")
_bth = _bim["bwm_cri_thresholds"].set_index("scale_a_BW")
_bder = _bim["derived"].set_index("quantity")
_bG = ["G1", "G2", "G3"]
# (2a) k-means reproduced the designed 6/7/7 strata; TOPSIS on the centroids of
#      the recorded attribute matrix returns the recorded Omega; the credibility
#      order is junior < mid-career < senior and no group leads on every
#      attribute; the profile ranges of tab:bim-groups are those of the matrix
_bX = _bea[list(_C1_ATTR_NAMES)].to_numpy(float)
assert np.array_equal(_bX, _C1_ATTR), "expert_attributes differ from the stated attribute matrix"
assert (_bea["group_kmeans"] == _bea["group_design"]).all(), "k-means groups differ from the design"
_bsz = _bea["group_kmeans"].value_counts().reindex(_bG).tolist()
assert _bsz == _bgr["size"].tolist() == list(_C1["sizes"]), "group sizes %r" % (_bsz,)
_bcent = np.array([_bX[(_bea["group_kmeans"] == g).to_numpy()].mean(0) for g in _bG])
_bVt = _bcent / np.sqrt((_bcent ** 2).sum(0)) * np.array(_C1["topsis_weights"])
_bdp = np.sqrt(((_bVt - _bVt.max(0)) ** 2).sum(1))
_bdn = np.sqrt(((_bVt - _bVt.min(0)) ** 2).sum(1))
_bOm = (_bdn / (_bdp + _bdn)) / (_bdn / (_bdp + _bdn)).sum()
assert np.allclose(_bOm, _bgr["Omega"].to_numpy(float), atol=1e-12), "Omega is not TOPSIS on the recorded centroids"
_sums_to_one(_bgr["Omega"], "BIM group Omega")
assert np.all(np.diff(_bgr["Omega"].to_numpy(float)) > 0), "Omega must increase from junior to senior"
# the pattern the text states: seniors lead on experience, types and
# certifications, mid-career members on projects while trailing on types and
# certifications, and juniors hold more certifications than mid-career members
_bY, _bPj, _bT, _bC = (_C1_ATTR_NAMES.index(n) for n in _C1_ATTR_NAMES)
assert all(int(np.argmax(_bcent[:, a])) == 2 for a in (_bY, _bT, _bC)) and int(np.argmax(_bcent[:, _bPj])) == 1
assert all(int(np.argmin(_bcent[:, a])) == 1 for a in (_bT, _bC)) and _bcent[0, _bC] > _bcent[1, _bC]
for _k in range(3):
    assert not np.all(_bcent[_k] >= _bcent.max(0) - 1e-12), "G%d leads on every attribute" % (_k + 1)
    _rows = _bX[(_bea["group_kmeans"] == _bG[_k]).to_numpy()]
    for _a, _nm in enumerate(_C1_ATTR_NAMES):
        _lo, _hi = int(_rows[:, _a].min()), int(_rows[:, _a].max())
        assert str(_bgr.loc[_k, "range_" + _nm]) == ("%d" % _lo if _lo == _hi else "%d-%d" % (_lo, _hi))
# (2b) every reported profile is the hundredths-rounded band distribution of its
#      latent normal (largest remainder, tails to the end bands), sums to one,
#      and its fit is the midpoint moment match; the recorded extreme tail mass
#      outside [0,10] is that of the fits
_bP = _bpn[["p_L%d" % (_l + 1) for _l in range(10)]].to_numpy(float)
assert len(_bpn) == _C1["N"] * 24
assert np.allclose(_bP.sum(1), 1.0, atol=1e-12) and np.allclose(_bP * 100, np.round(_bP * 100), atol=1e-9)
for _r, (_, _row) in enumerate(_bpn.iterrows()):
    _cnt = _c1_report(float(_row["latent_mu"]), float(_row["latent_sd"]))
    assert np.array_equal(_cnt, np.round(_bP[_r] * 100).astype(int)), "profile %s/%s is not the rounded latent normal" % (
        _row["expert"], _row["indicator"])
    _pm = float(_bP[_r] @ _c1_mid)
    _ps = math.sqrt(float(_bP[_r] @ _c1_mid ** 2) - _pm ** 2)         # E[x^2] - mu^2, a second route
    assert abs(_pm - float(_row["fit_mu"])) < 1e-12 and abs(_ps - float(_row["fit_sigma"])) < 1e-9
_bfm, _bfs = _bpn["fit_mu"].to_numpy(float), _bpn["fit_sigma"].to_numpy(float)
_btail = np.asarray(_norm.cdf(-_bfm / _bfs) + _norm.sf((10.0 - _bfm) / _bfs))
assert abs(_btail.max() - float(_bcf["max_mass_outside_0_10_fits"])) < 1e-12
assert float(_bcf["max_mass_outside_0_10_aggregates"]) <= _btail.max(), "an aggregate has more tail mass than the fits"
# (2c) each intra-group program: the recorded w lies on the group simplex and
#      induces the recorded group opinion (w.mu, w.sigma); A and V are the
#      def:consensus and def:confidence values (overlaps recomputed by the
#      Chapter 2 routine); the bounds bracket every vertex, V's bounds are the
#      vertex extremes; the retained objective equals its recomputation and is
#      not beaten by any vertex, the centroid, or 200 seeded probe points; all
#      starts agree (the text reports this) and the bound agreement matches
#      the config; the member-coincidence count matches the config
_balpha = float(_bcf["alpha"])
_bprobe_rng = np.random.default_rng(int(_bcf["seed"]))
_bn_member = 0
_bn_member_inside = 0
_bn_member_w1 = 0
_bmtol = float(_bcf["member_coincidence_tol"])


def _b_strictly_inside_hull(P, j):
    """Is point j strictly inside the convex hull of the rows of P (2-D)?
    Andrew's monotone chain on all points, then a strict left-turn test of
    point j against every hull edge (a route independent of Qhull)."""
    pts = sorted(map(tuple, P))
    cross = lambda o, a, b: (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0])
    lower, upper = [], []
    for p in pts:
        while len(lower) >= 2 and cross(lower[-2], lower[-1], p) <= 0:
            lower.pop()
        lower.append(p)
    for p in reversed(pts):
        while len(upper) >= 2 and cross(upper[-2], upper[-1], p) <= 0:
            upper.pop()
        upper.append(p)
    hull = lower[:-1] + upper[:-1]                      # counter-clockwise
    x = tuple(P[j])
    return all(cross(hull[i], hull[(i + 1) % len(hull)], x) > 1e-12 for i in range(len(hull)))


for _, _r in _bsol.iterrows():
    _g, _q = str(_r["group"]), str(_r["indicator"])
    _mem = _bea.index[(_bea["group_kmeans"] == _g).to_numpy()].to_numpy()
    _sub = _bpn[(_bpn["group"] == _g) & (_bpn["indicator"] == _q)]
    assert list(_sub["expert"]) == ["e%d" % (i + 1) for i in _mem]
    _mu, _sd = _sub["fit_mu"].to_numpy(float), _sub["fit_sigma"].to_numpy(float)
    _w = np.array([float(_r["w_e%d" % (i + 1)]) for i in _mem])
    assert np.all(_w >= -1e-12)
    _sums_to_one(_w, "BIM w %s/%s" % (_g, _q), tol=1e-9)
    _others = [c for c in _bsol.columns if c.startswith("w_e") and int(c[3:]) - 1 not in _mem]
    assert _r[_others].isna().all(), "non-members carry weight in %s/%s" % (_g, _q)
    assert abs(float(_r["m"]) - float(_w @ _mu)) < 1e-9 and abs(float(_r["s"]) - float(_w @ _sd)) < 1e-9
    _Af = lambda w: sum(_ch2_overlap_normal(mi, si, float(w @ _mu), float(w @ _sd)) for mi, si in zip(_mu, _sd))
    _Vf = lambda w: float(w @ _sd) ** 2
    assert abs(_Af(_w) - float(_r["A"])) < 1e-9 and abs(_Vf(_w) - float(_r["V"])) < 1e-12
    _Amin, _Amax, _Vmin, _Vmax = (float(_r[c]) for c in ("A_min", "A_max", "V_min", "V_max"))
    assert abs(_Vmin - _sd.min() ** 2) < 1e-9 and abs(_Vmax - _sd.max() ** 2) < 1e-9, "V bounds are not the vertex extremes"
    _E = np.eye(len(_mem))
    _vA = [_Af(e) for e in _E]
    assert min(_vA) >= _Amin - 1e-9 and max(_vA) <= _Amax + 1e-9, "A bounds do not bracket the vertices"
    _obj = lambda w: _balpha * (_Af(w) - _Amin) / (_Amax - _Amin) - (1 - _balpha) * (_Vf(w) - _Vmin) / (_Vmax - _Vmin)
    _f = _obj(_w)
    assert abs(_f - float(_r["objective"])) < 1e-9, "objective of %s/%s not reproduced" % (_g, _q)
    for _w0 in np.vstack([_E, np.full((1, len(_mem)), 1.0 / len(_mem)),
                          _bprobe_rng.dirichlet(np.ones(len(_mem)), 200)]):
        assert _obj(_w0) <= _f + 1e-9, "a vertex, the centroid or a probe beats the solution of %s/%s" % (_g, _q)
        assert _Af(_w0) <= _Amax + 1e-9 and _Af(_w0) >= _Amin - 1e-9, "A bounds exceeded at a probe of %s/%s" % (_g, _q)
    assert int(_r["starts_at_optimum"]) == int(_r["n_starts"]), "not every start agrees in %s/%s" % (_g, _q)
    assert float(_r["ms_spread_at_optimum"]) < float(_bcf["group_opinion_agreement_bound"])
    _bdist = np.abs(_mu - float(_r["m"])) + np.abs(_sd - float(_r["s"]))
    _bat = int(np.min(_bdist) < _bmtol)
    assert _bat == int(_r["coincides_with_member"]), "member coincidence flag of %s/%s" % (_g, _q)
    if _bat:
        # the coinciding member's overlap is one, the peak of that term
        _jm = int(np.argmin(_bdist))
        assert abs(_ch2_overlap_normal(_mu[_jm], _sd[_jm], float(_r["m"]), float(_r["s"])) - 1.0) < 1e-4
        _bn_member_inside += int(_b_strictly_inside_hull(np.column_stack([_mu, _sd]), _jm))
        _bn_member_w1 += int(float(_w.max()) > 1.0 - _bmtol)
    _bn_member += _bat
assert int(_bcf["all_starts_agree_in_every_program"]) == 1 and int(_bcf["n_programs"]) == len(_bsol) == 72
assert int(_bcf["min_bound_starts_at_optimum"]) == min(
    min(int(v) for v in str(s_).split(",")) for s_ in _bsol["bound_starts_at_optimum"])
assert _bn_member == int(_bcf["n_group_opinions_equal_to_one_member"])
# the text's reading of those coincidences: the member's point lies strictly
# inside the group's hull in all but a few programs (a kink of the overlap,
# not a vertex of the feasible set), and the recorded weights are
# concentrated on that member in only a few of them
assert _bn_member_inside == int(_bcf["n_coinciding_members_strictly_inside_group_hull"])
assert _bn_member_w1 == int(_bcf["n_coinciding_opinions_with_recorded_max_weight_one"])
assert _bn_member_inside > _bn_member // 2 and _bn_member_w1 < _bn_member // 2, \
    "the coincidences are no longer mainly interior kinks with spread recorded weights"
assert int(_bcf["kmeans_maxiter"]) == _C1["kmeans_maxiter"] and ("at most %d Lloyd iterations" % _C1["kmeans_maxiter"]) in str(_bcf["kmeans"])
assert "ties assigned to the lower band" in str(_bcf["report_rule"])
# the weight vector need not be unique, but every start returns the same
# group opinion: the recorded spread is the largest per-program spread and
# lies within the stated agreement bound
assert abs(float(_bcf["max_group_opinion_spread_among_agreeing_starts"])
           - max(float(v) for v in _bsol["ms_spread_at_optimum"])) < 1e-15
assert float(_bcf["max_group_opinion_spread_among_agreeing_starts"]) < float(_bcf["group_opinion_agreement_bound"])
assert "all agreeing starts return the same group opinion" in str(_bcf["weight_uniqueness_note"])
# (2d) inter-group quantile averaging under Omega: every indicator collective
#      satisfies the quantile identity sum_k Omega_k F_k^{-1}(p) at three levels
_bOmv = _bgr["Omega"].to_numpy(float)
for _q in _bic.index:
    _gm = np.array([float(_bic.loc[_q, "%s_mu" % g]) for g in _bG])
    _gs = np.array([float(_bic.loc[_q, "%s_sigma" % g]) for g in _bG])
    for _g, _m_, _s_ in zip(_bG, _gm, _gs):
        _rr = _bsol[(_bsol["group"] == _g) & (_bsol["indicator"] == _q)].iloc[0]
        assert abs(float(_rr["m"]) - _m_) < 1e-12 and abs(float(_rr["s"]) - _s_) < 1e-12
    for _p in (0.05, 0.5, 0.95):
        _z = float(_norm.ppf(_p))
        assert abs(float(_bOmv @ (_gm + _gs * _z)) - (float(_bic.loc[_q, "mu"]) + float(_bic.loc[_q, "sigma"]) * _z)) < 1e-9
# (2e) best-worst weights: the linear model re-solved from the recorded
#      comparison vectors returns the recorded weights, uniquely; the
#      input-based ratio CR^I recomputed from the vectors matches, and every
#      ratio lies below the smallest Liang et al. (2020) threshold for its
#      scale (the text's claim); the threshold sheet is Table 3 of that paper
for _s_, _row in _C1_CRI_THRESH.items():
    assert np.allclose(_bth.loc[_s_].to_numpy(float), np.array(_row)), "CR^I threshold table altered"
for _P in _C1_BWM:
    _sb = _bbw[_bbw["problem"] == _P["problem"]]
    _aB = _sb["best_to_others"].to_numpy(float); _aW = _sb["others_to_worst"].to_numpy(float)
    assert tuple(_aB) == tuple(float(v) for v in _P["best_to_others"]) and tuple(_aW) == tuple(float(v) for v in _P["others_to_worst"])
    _w2, _xi2, _spr2 = _c1_bwm(_aB, _aW, _P["best"] - 1, _P["worst"] - 1)
    assert np.allclose(_w2, _sb["weight"].to_numpy(float), atol=1e-9) and _spr2 < 1e-7
    _sums_to_one(_sb["weight"], "BWM %s" % _P["problem"])
    _cr = _c1_cri(_aB, _aW, _P["worst"] - 1)
    assert abs(_cr - float(_bbs.loc[_P["problem"], "CR_I"])) < 1e-12 and abs(_xi2 - float(_bbs.loc[_P["problem"], "xi_L"])) < 1e-9
    assert _cr < float(_bth.loc[int(_aB[_P["worst"] - 1])].min()), "CR^I of %s exceeds the smallest threshold for its scale" % _P["problem"]
assert abs(float(_bbs.loc["Dimensions", "CR_I"])) < 1e-12, "the dimension comparisons are stated as fully consistent"
_bwi = np.concatenate([_bbw[_bbw["problem"] == p]["weight"].to_numpy(float) for p in
                       ("Technology indicators", "Organization indicators", "Policy indicators")])
assert np.allclose(_bwi, _bic["bwm_weight"].to_numpy(float), atol=1e-12)
assert np.allclose(_bbw[_bbw["problem"] == "Dimensions"]["weight"].to_numpy(float), _bdm["weight"].to_numpy(float), atol=1e-12)
# (2f) the hierarchy: each dimension is the BWM-weighted quantile average of its
#      indicator collectives and the global maturity that of the dimensions
#      (quantile identity at three levels); the text's claims: organization
#      strongest, technology weakest, the global mean between the managed and
#      defined bands and within one sd of their boundary; the printed strings
_bdim_of = _bic["dimension"].astype(str).to_numpy()
for _c, _dn in enumerate(_bdm["dimension"].astype(str)):
    _sel = _bdim_of == _dn
    _wv, _mv, _sv = (_bic.loc[_sel, "bwm_weight"].to_numpy(float), _bic.loc[_sel, "mu"].to_numpy(float),
                     _bic.loc[_sel, "sigma"].to_numpy(float))
    for _p in (0.05, 0.5, 0.95):
        _z = float(_norm.ppf(_p))
        assert abs(float(_wv @ (_mv + _sv * _z)) - (float(_bdm.loc[_c, "mu"]) + float(_bdm.loc[_c, "sigma"]) * _z)) < 1e-9
    assert str(_bdm.loc[_c, "printed"]) == _c1_printed(float(_bdm.loc[_c, "mu"]), float(_bdm.loc[_c, "sigma"]))
    _lo_q = _bic.index[_sel][int(np.argmin(_mv))]
    assert str(_bdm.loc[_c, "lowest_indicator"]).split()[0] == _lo_q
_bgrow = "global maturity (QA of dimensions under BWM weights)"
_bdw, _bdmu, _bdsd = (_bdm["weight"].to_numpy(float), _bdm["mu"].to_numpy(float), _bdm["sigma"].to_numpy(float))
for _p in (0.05, 0.5, 0.95):
    _z = float(_norm.ppf(_p))
    assert abs(float(_bdw @ (_bdmu + _bdsd * _z)) - (float(_bder.loc[_bgrow, "mu"]) + float(_bder.loc[_bgrow, "sigma"]) * _z)) < 1e-9
_bdmi = dict(zip(_bdm["dimension"].astype(str), _bdmu))
assert _bdmi["Organization"] > _bdmi["Policy"] > _bdmi["Technology"], "dimension order is not organization > policy > technology"
assert abs(float(_bcf["margin_organization_over_policy"]) - (_bdmi["Organization"] - _bdmi["Policy"])) < 1e-12
assert abs(float(_bcf["margin_policy_over_technology"]) - (_bdmi["Policy"] - _bdmi["Technology"])) < 1e-12
_bgm, _bgs = float(_bder.loc[_bgrow, "mu"]), float(_bder.loc[_bgrow, "sigma"])
# the equal-dimension-weight global (the published application's rule, recorded for
# comparison) is the plain mean of the dimension collectives and differs from the case's
_beqrow = "equal-dimension-weight global (the published rule, for comparison)"
assert abs(float(_bder.loc[_beqrow, "mu"]) - float(_bdmu.mean())) < 1e-12
assert abs(float(_bder.loc[_beqrow, "sigma"]) - float(_bdsd.mean())) < 1e-12
assert float(_bcf["global_mu_equal_dimension_weights"]) == float(_bder.loc[_beqrow, "mu"])
assert float(_bcf["global_sigma_equal_dimension_weights"]) == float(_bder.loc[_beqrow, "sigma"])
assert str(_bcf["global_equal_dimension_weights_printed"]) == str(_bder.loc[_beqrow, "printed"]) \
    == _c1_printed(float(_bdmu.mean()), float(_bdsd.mean()))
assert abs(float(_bder.loc[_beqrow, "mu"]) - _bgm) > 0.05, "the dimension weighting no longer moves the global"
assert "global maturity" not in _beqrow   # the manifest finds the case's global by that phrase
assert 4.0 < _bgm < 6.0 and abs(_bgm - 5.0) < _bgs, "global maturity is not between the managed and defined levels"
# the results paragraph: the global lies in the defined band [5, 6) just above its
# boundary with the managed band; under the published equal dimension weights it lies
# further inside that band, because the best-worst weights give the weakest dimension,
# technology, the largest weight; the Case I robustness paragraph and tab:robustness
# say that this comparison keeps the global mean in the defined band
_beqm = float(_bder.loc[_beqrow, "mu"])
assert 5.0 <= _bgm < 6.0 and 5.0 <= _beqm < 6.0 and _beqm > _bgm, \
    "the global maturity (%.4f) or its equal-weight counterpart (%.4f) left the defined band" % (_bgm, _beqm)
_bdwi = dict(zip(_bdm["dimension"].astype(str), _bdw))
assert min(_bdmi, key=_bdmi.get) == "Technology" and max(_bdwi, key=_bdwi.get) == "Technology", \
    "the best-worst weights no longer give the weakest dimension, technology, the largest weight"
# the design sentence of sec:bim-data: the indicator offsets place q1, q2, q3, q5, q6
# (the published recommendation) and q7 (data richness) below the technology level,
# and within the organization dimension senior leadership, role and competency
# (q15, q16, q17) well above its level and quality and safety management (q14)
# marginally above it, every other organization indicator below it ("marginally":
# less than half the smallest of the three larger offsets)
_boffs = [float(v) for v in str(_bcf["indicator_offset"]).split(",")]
assert len(_boffs) == len(_bdim_of) == 24
_boff = dict(("q%d" % (_k + 1), _v) for _k, _v in enumerate(_boffs))
_bdimq = dict(("q%d" % (_k + 1), _d) for _k, _d in enumerate(_bdim_of))
assert list(_bic.index) == ["q%d" % (_k + 1) for _k in range(24)]
assert {q for q in _boff if _bdimq[q] == "Technology" and _boff[q] < 0} == {"q1", "q2", "q3", "q5", "q6", "q7"}
_borg_pos = {q for q in _boff if _bdimq[q] == "Organization" and _boff[q] > 0}
assert _borg_pos == {"q14", "q15", "q16", "q17"}, "organization indicators above the level %r" % sorted(_borg_pos)
assert 0 < _boff["q14"] < 0.5 * min(_boff["q15"], _boff["q16"], _boff["q17"]), \
    "the q14 offset is no longer marginal beside the q15, q16, q17 offsets"
# the indicator-level pattern the results paragraph names: the four weakest
# technology indicators, and the strongest organization and policy indicators
_btech = _bic[_bdim_of == "Technology"]["mu"].sort_values()
assert set(_btech.index[:4]) == {"q1", "q2", "q3", "q6"}, "weakest technology indicators %r" % list(_btech.index[:4])
assert _btech.index[0] == "q3" and _btech.index[-1] == "q10"
_borg = _bic[_bdim_of == "Organization"]["mu"].sort_values(ascending=False)
assert list(_borg.index[:3]) == ["q15", "q17", "q14"], "strongest organization indicators %r" % list(_borg.index[:3])
# (2g) q4 (Figure 6.1): the group opinions rise from junior to senior, the
#      collective is their quantile average under Omega, its sd is the Omega
#      average of the group sds and so below the linear pool's
_b4 = _bgr[["mu", "sigma"]].to_numpy(float)
assert np.all(np.diff(_b4[:, 0]) > 0), "q4 group opinions are not ordered junior < mid-career < senior"
_b4row = "q4 collective (inter-group QA under Omega)"
assert abs(float(_bder.loc[_b4row, "mu"]) - float(_bOmv @ _b4[:, 0])) < 1e-12
assert abs(float(_bder.loc[_b4row, "sigma"]) - float(_bOmv @ _b4[:, 1])) < 1e-12
_b4ic = np.array([[float(_bic.loc["q4", "%s_mu" % g]), float(_bic.loc["q4", "%s_sigma" % g])] for g in _bG])
assert np.allclose(_b4, _b4ic, atol=1e-12), "the groups sheet does not hold the q4 group opinions"
_blp = math.sqrt(float(_bOmv @ (_b4[:, 1] ** 2 + (_b4[:, 0] - float(_bder.loc[_b4row, "mu"])) ** 2)))
assert abs(_blp - float(_bcf["q4_linear_pool_sd"])) < 1e-12 and float(_bder.loc[_b4row, "sigma"]) < _blp
assert float(_bder.loc[_b4row, "sigma"]) < float(_b4[:, 1].max()), "the collective is not narrower than the widest group"
# (2h) tab:bim-elicit: the profile is the panel row of the stated expert for q4,
#      its fit is that row's moment match, and the fit row of the table holds
#      the band probabilities of the fitted normal
_bel = _bim["elicitation_q4"]
_brow = _bpn[(_bpn["expert"] == str(_bcf["elicitation_expert"])) & (_bpn["indicator"] == "q4")].iloc[0]
assert np.allclose(_bel["probability"].to_numpy(float), [float(_brow["p_L%d" % (_l + 1)]) for _l in range(10)], atol=1e-12)
_sums_to_one(_bel["probability"], "BIM q4 elicited profile")
assert len(_bel) == 10 and np.allclose(_bel["level_mid"].to_numpy(float), _c1_mid)
_frow = "tab:bim-elicit normal fit of the q4 profile"
assert abs(float(_bder.loc[_frow, "mu"]) - float(_brow["fit_mu"])) < 1e-12 and \
    abs(float(_bder.loc[_frow, "sigma"]) - float(_brow["fit_sigma"])) < 1e-12, "BIM elicitation fit is not the midpoint moment match"
_bfb = np.diff(np.concatenate([[0.0], _norm.cdf((np.arange(1.0, 10.0) - float(_brow["fit_mu"])) / float(_brow["fit_sigma"])), [1.0]]))
assert np.allclose(_bfb, _bel["fit_probability"].to_numpy(float), atol=1e-12)
assert abs(float(_bel["probability"].iloc[4] + _bel["probability"].iloc[5]) - float(_bcf["elicitation_mass_managed_defined"])) < 1e-12
assert 4.0 <= float(_brow["fit_mu"]) <= 6.0, "BIM elicitation fit mean must lie in the managed-defined bands"
print("   BIM full assessment: k-means strata and TOPSIS Omega, 480 profiles and fits, 72 intra-group optima (all starts agree,"
      " none beaten by vertices or 14,400 probes), BWM weights and CR^I, the QA hierarchy and the text's orderings verified")

# (3) BIM fairness panel: the SOLVED sweep of fig6-4_bim_tradeoff.xlsx is
#     checked against invariants of the program eq:bim-triobjective, never
#     against typed result constants.
_bsw = _wb("fig6-4_bim_tradeoff.xlsx")
_bcfg = _cfgd(_bsw)
_bser = _bsw["series"].set_index("delta")
_bwt = _bsw["weights"].set_index("delta")
_bov = _bsw["overlaps"].set_index("delta")
_bbd = _bsw["bounds"].set_index("objective")
_bimg = _wb("fig6-3_bim_densities.xlsx")
_bexp = _bimg["gaussians"][_bimg["gaussians"]["role"].astype(str).str.startswith("expert")]
assert len(_bexp) == 7, "BIM densities: expected the 7-expert panel"
_bmu = _bexp["mu"].to_numpy(float)
_bsd = _bexp["sigma"].to_numpy(float)
assert np.allclose(_bmu, _BIM_MU) and np.allclose(_bsd, _BIM_SD), "fig6-3 expert rows differ from eq:bim-panel-mu/sigma"
_bW = np.column_stack([_bwt["w%d" % (i + 1)].to_numpy(float) for i in range(7)])
_bdel = _bwt.index.to_numpy(float)
assert list(np.round(_bdel, 6)) == [round(d, 6) for d in _BIM_DELTAS], "fig6-4 sweep grid is not {0,0.1,...,0.6}"
for _k, _d in enumerate(_bdel):
    _sums_to_one(_bW[_k], "BIM w(delta=%.1f)" % _d)
    assert (_bW[_k] >= -1e-12).all(), "BIM w(delta=%.1f) has a negative weight" % _d
# (3a) the recorded bounds are attained/bracketed: the confidence V = (w.sigma)^2
#      is a square of a linear form, so its extremes over the simplex are
#      vertex values; the mean-overlap and fairness bounds must bracket every
#      vertex and every sweep solution.
assert abs(float(_bbd.loc["V", "min"]) - _BIM_SD.min() ** 2) < 1e-6 and \
    abs(float(_bbd.loc["V", "max"]) - _BIM_SD.max() ** 2) < 1e-6, "solved V bounds are not the vertex extremes"
_bcand = np.vstack([np.eye(7), _bW])
for _name, _f in (("Abar", _bim_Abar), ("Phi", _bim_Phi), ("V", _bim_V)):
    _lo, _hi = float(_bbd.loc[_name, "min"]), float(_bbd.loc[_name, "max"])
    _vals = np.array([_f(w) for w in _bcand])
    assert _vals.min() >= _lo - 1e-6 and _vals.max() <= _hi + 1e-6, \
        "BIM %s bounds [%.6f, %.6f] do not bracket the vertices and sweep solutions" % (_name, _lo, _hi)
    assert _hi - _lo > 1e-3, "BIM %s is constant over the simplex; min-max normalization undefined" % _name
    assert abs(float(_bcfg["%s_min" % _name]) - _lo) < 1e-12 and abs(float(_bcfg["%s_max" % _name]) - _hi) < 1e-12, \
        "config bounds for %s disagree with the bounds sheet" % _name
# (3a') the start agreement of the bound programs as the text reports it: all
#       starts reach the three upper bounds and the lower variance bound, fewer
#       reach the lower mean-overlap and fairness bounds, and both of those
#       lower bounds sit at the dissenter's vertex.
for _name in ("Abar", "V", "Phi"):
    assert int(_bbd.loc[_name, "starts_at_max"]) == _BIM_N_STARTS, "BIM %s max: not every start agrees" % _name
assert int(_bbd.loc["V", "starts_at_min"]) == _BIM_N_STARTS, "BIM V min: not every start agrees"
assert int(_bbd.loc["Phi", "starts_at_min"]) < int(_bbd.loc["Abar", "starts_at_min"]) < _BIM_N_STARTS
assert int(_bbd.loc["Phi", "starts_at_min"]) < int(_bser["starts_at_best"].min()), \
    "the lower fairness bound should be the least-agreed program (the text says so)"
for _name in ("Abar", "Phi"):
    _wa = np.array([float(x) for x in str(_bbd.loc[_name, "argmin_w"]).split(",")])
    assert _wa[_BIM_DISSENTER] > 1 - 1e-4, "BIM %s min is not attained at the dissenter's vertex" % _name
# (3a'') the certificate, recomputed here from the workbook's recorded bounds
#        and sweep optima: no point of the dense hull grid improves on any of
#        them, the grid comes within its resolution of each, and each sweep
#        optimum's grid argmax lies next to the recorded aggregate (same branch).
_bscan_rows, _bscan_n = _bim_hull_scan(
    {_n: (float(_bbd.loc[_n, "min"]), float(_bbd.loc[_n, "max"])) for _n in ("Abar", "V", "Phi")},
    {float(_d): (float(_bser.loc[_d, "objective"]), float(_bser.loc[_d, "agg_mu"]), float(_bser.loc[_d, "agg_sd"]))
     for _d in _bdel})
_bscan_sheet = _bsw["hull_scan"].set_index("program")
assert int(_bcfg["hull_scan_points"]) == _bscan_n
for _r in _bscan_rows:
    assert _r["grid_improvement"] <= _BIM_SCAN_IMPROVE_TOL, \
        "BIM hull scan: a grid point improves on %s by %.2e" % (_r["program"], _r["grid_improvement"])
    assert _r["grid_gap"] <= _BIM_SCAN_GAP_TOL, "BIM hull scan: grid too coarse at %s" % _r["program"]
    assert abs(float(_bscan_sheet.loc[_r["program"], "grid_best"]) - _r["grid_best"]) < 1e-12
    if _r["program"].startswith("sweep"):
        assert abs(_r["grid_m"] - _r["recorded_m"]) <= _BIM_SCAN_LOC_TOL[0] and \
            abs(_r["grid_s"] - _r["recorded_s"]) <= _BIM_SCAN_LOC_TOL[1], \
            "BIM hull scan: grid optimum of %s lies on another branch" % _r["program"]
assert float(_bcfg["hull_scan_max_improvement"]) <= _BIM_SCAN_IMPROVE_TOL
# (3b) the recorded diagnostics are recomputed from the recorded weights
#      (mean overlap, minimum overlap, fairness, aggregate, objective).
for _k, _d in enumerate(_bdel):
    _w = _bW[_k]
    _A = _bim_overlaps(_w)
    _P = _bim_Phi(_w)
    assert abs(float(_bser.loc[_d, "consensus_Abar"]) - _A.mean()) < 1e-9
    assert abs(float(_bser.loc[_d, "minimum_overlap"]) - _A.min()) < 1e-9
    assert abs(float(_bser.loc[_d, "fairness_Phibar"]) - _P) < 1e-9
    assert abs(float(_bser.loc[_d, "agg_mu"]) - float(_w @ _bmu)) < 1e-9
    assert abs(float(_bser.loc[_d, "agg_sd"]) - float(_w @ _bsd)) < 1e-9
    assert abs(float(_bser.loc[_d, "objective"]) - _bim_scalarized(_w, _d)) < 1e-9
    assert np.allclose(_bov.loc[_d, ["A%d" % (i + 1) for i in range(7)]].to_numpy(float), _A, atol=1e-9)
    # closed-form overlap of the aggregate with every expert vs the trapezoid rule
    for _i in range(7):
        assert abs(_A[_i] - _bim_overlap_trapz(_bmu[_i], _bsd[_i], float(_w @ _bmu), float(_w @ _bsd))) < 1e-7
# (3c) optimality against a candidate set: at every delta the recorded
#      solution scores at least as high as every vertex and every other
#      sweep solution under that delta's scalarization.
for _k, _d in enumerate(_bdel):
    _best = _bim_scalarized(_bW[_k], _d)
    for _c in _bcand:
        assert _bim_scalarized(_c, _d) <= _best + 1e-9, \
            "BIM delta=%.1f: candidate %s beats the recorded solution" % (_d, np.round(_c, 3))
# (3d) the qualitative finding the text asserts, checked on the solved
#      outputs: the delta = 0 solution is a vertex (all weight on one expert),
#      and from delta = 0 to delta = 0.5 the dissenter's weight, the minimum
#      overlap and the collective fairness rise while the mean overlap and
#      the aggregate mean fall.
_i0 = list(_bdel).index(0.0)
_i5 = list(_bdel).index(0.5)
assert _bW[_i0].max() > 1 - 1e-6, "BIM delta=0 solution is not a vertex"
assert float(_bser.loc[0.5, "weight_on_dissenter"]) > float(_bser.loc[0.0, "weight_on_dissenter"]) + 0.05
assert float(_bser.loc[0.5, "minimum_overlap"]) > float(_bser.loc[0.0, "minimum_overlap"])
assert float(_bser.loc[0.5, "fairness_Phibar"]) > float(_bser.loc[0.0, "fairness_Phibar"])
assert float(_bser.loc[0.5, "consensus_Abar"]) < float(_bser.loc[0.0, "consensus_Abar"])
assert float(_bser.loc[0.5, "agg_mu"]) < float(_bser.loc[0.0, "agg_mu"])
# (3d') the finer claims of the results and robustness paragraphs:
#   * delta = 0: the vertex is expert 3, whose overlap with the aggregate is
#     exactly one; the dissenter's overlap is the lowest and lies below that
#     of every other majority expert; the solution stays at that vertex
#     through delta = 0.3;
#   * delta = 0.4: the solution leaves the vertex toward a majority expert,
#     not the dissenter, raising consensus and fairness by less than 0.001
#     while the aggregate mean rises and the dissenter's (the minimum)
#     overlap falls, each by less than 0.01 (the opposite direction of the
#     finding);
#   * delta = 0.5: the dissenter and expert 2 tie as least represented;
#   * delta = 0.6: fairness and the dissenter's overlap rise further, the
#     mean and minimum overlaps fall, expert 2 becomes the least represented,
#     and the minimum overlap stays above its delta = 0 value;
#   * every grid point from the threshold on has a lower aggregate mean and a
#     higher minimum overlap than delta = 0 (the robustness finding).
_bA = {float(_d): _bov.loc[_d, ["A%d" % (i + 1) for i in range(7)]].to_numpy(float) for _d in _bdel}
_bvx = int(np.argmax(_bW[_i0]))
assert _bvx == 2, "BIM delta=0 vertex is not expert 3 (the text names expert 3)"
assert abs(_bA[0.0][_bvx] - 1.0) < 1e-12, "the delta=0 aggregate is not expert 3's own distribution"
assert abs(float(_bser.loc[0.0, "agg_mu"]) - _bmu[_bvx]) < 1e-12 and abs(float(_bser.loc[0.0, "agg_sd"]) - _bsd[_bvx]) < 1e-12
assert int(np.argmin(_bA[0.0])) == _BIM_DISSENTER
assert _bA[0.0][_BIM_DISSENTER] < min(_bA[0.0][i] for i in range(7) if i not in (_bvx, _BIM_DISSENTER))
for _d in (0.1, 0.2, 0.3):
    assert float(_bwt.loc[_d, "w3"]) > 1 - 1e-6, "BIM delta=%.1f solution has left the expert-3 vertex" % _d
_bw4 = _bwt.loc[0.4].to_numpy(float)
assert 1e-3 < 1 - _bw4[_bvx] < 0.05 and _bw4[_BIM_DISSENTER] < 1e-9, "BIM delta=0.4 departure is not a small move off the vertex"
assert int(np.argmax(np.where(np.arange(7) == _bvx, -1, _bw4))) != _BIM_DISSENTER
for _col in ("consensus_Abar", "fairness_Phibar"):
    assert 0 < float(_bser.loc[0.4, _col]) - float(_bser.loc[0.0, _col]) < 1e-3, "delta=0.4 %s change" % _col
assert 0 < float(_bser.loc[0.4, "agg_mu"]) - float(_bser.loc[0.0, "agg_mu"]) < 0.01
assert -0.01 < float(_bser.loc[0.4, "minimum_overlap"]) - float(_bser.loc[0.0, "minimum_overlap"]) < 0
assert _bA[0.4][_BIM_DISSENTER] < _bA[0.0][_BIM_DISSENTER] and int(np.argmin(_bA[0.4])) == _BIM_DISSENTER
_b05 = np.argsort(_bA[0.5])
assert set(_b05[:2]) == {1, _BIM_DISSENTER} and abs(_bA[0.5][1] - _bA[0.5][_BIM_DISSENTER]) < 1e-9, \
    "BIM delta=0.5: the dissenter and expert 2 are not tied as least represented"
assert float(_bser.loc[0.6, "fairness_Phibar"]) > float(_bser.loc[0.5, "fairness_Phibar"])
assert _bA[0.6][_BIM_DISSENTER] > _bA[0.5][_BIM_DISSENTER]
assert float(_bser.loc[0.6, "minimum_overlap"]) < float(_bser.loc[0.5, "minimum_overlap"])
assert float(_bser.loc[0.6, "consensus_Abar"]) < float(_bser.loc[0.5, "consensus_Abar"])
assert int(_bser.loc[0.6, "argmin_overlap_expert"]) == 2 and int(np.argmin(_bA[0.6])) == 1
assert float(_bser.loc[0.6, "weight_on_dissenter"]) > float(_bser.loc[0.5, "weight_on_dissenter"])
for _d in _bdel:
    if _d >= 0.5 - 1e-9:
        assert float(_bser.loc[_d, "agg_mu"]) < float(_bser.loc[0.0, "agg_mu"])
        assert float(_bser.loc[_d, "minimum_overlap"]) > float(_bser.loc[0.0, "minimum_overlap"])
# (3d'') the derived quantities of sheet text_values, recomputed from the sheets
_btv = _bsw["text_values"].set_index("quantity")
_bmaj = [i for i in range(7) if i not in (_bvx, _BIM_DISSENTER)]
_blo = min(_bmaj, key=lambda i: _bA[0.0][i])
_bhi = max(_bmaj, key=lambda i: _bA[0.0][i])
_boff = int(np.argmax(np.where(np.arange(7) == _bvx, -1, _bw4)))
_bm06 = int(np.argmin(_bA[0.6]))
for _q, _v, _e in (
        ("delta=0: overlap of the vertex expert with the aggregate", _bA[0.0][_bvx], _bvx),
        ("delta=0: lowest overlap among the other majority experts", _bA[0.0][_blo], _blo),
        ("delta=0: highest overlap among the other majority experts", _bA[0.0][_bhi], _bhi),
        ("delta=0.4: weight moved off the vertex expert, onto this expert", _bw4[_boff], _boff),
        ("delta=0.4: weight on the dissenter", _bw4[_BIM_DISSENTER], _BIM_DISSENTER),
        ("delta=0.4 minus delta=0: mean overlap (consensus)", float(_bser.loc[0.4, "consensus_Abar"]) - float(_bser.loc[0.0, "consensus_Abar"]), None),
        ("delta=0.4 minus delta=0: collective fairness", float(_bser.loc[0.4, "fairness_Phibar"]) - float(_bser.loc[0.0, "fairness_Phibar"]), None),
        ("delta=0.4 minus delta=0: aggregate mean", float(_bser.loc[0.4, "agg_mu"]) - float(_bser.loc[0.0, "agg_mu"]), None),
        ("delta=0.4 minus delta=0: minimum overlap", float(_bser.loc[0.4, "minimum_overlap"]) - float(_bser.loc[0.0, "minimum_overlap"]), None),
        ("delta=0.4: dissenter's overlap", _bA[0.4][_BIM_DISSENTER], _BIM_DISSENTER),
        ("delta=0.5: overlap of the least-represented expert", _bA[0.5][_b05[0]], _b05[0]),
        ("delta=0.5: overlap of the second least-represented expert", _bA[0.5][_b05[1]], _b05[1]),
        ("delta=0.6: dissenter's overlap", _bA[0.6][_BIM_DISSENTER], _BIM_DISSENTER),
        ("delta=0.6: minimum overlap, attained by this expert", _bA[0.6][_bm06], _bm06),
        ("delta=0.6 minus delta=0.5: minimum overlap", float(_bser.loc[0.6, "minimum_overlap"]) - float(_bser.loc[0.5, "minimum_overlap"]), None),
        ("delta=0.6 minus delta=0.5: mean overlap (consensus)", float(_bser.loc[0.6, "consensus_Abar"]) - float(_bser.loc[0.5, "consensus_Abar"]), None),
        ("delta=0.6 minus delta=0.5: collective fairness", float(_bser.loc[0.6, "fairness_Phibar"]) - float(_bser.loc[0.5, "fairness_Phibar"]), None)):
    assert abs(float(_btv.loc[_q, "value"]) - _v) < 1e-9, "text_values %r disagrees with the sweep sheets" % _q
    if _e is not None:
        assert int(_btv.loc[_q, "expert"]) == int(_e) + 1, "text_values %r names the wrong expert" % _q
# (3e) fig6-3 aggregates and tab6-9 are the solved delta = 0 / 0.5 rows
_bagg = _bimg["gaussians"][_bimg["gaussians"]["role"] == "aggregate"].reset_index(drop=True)
for _j, _d in ((0, 0.0), (1, 0.5)):
    assert abs(float(_bagg.loc[_j, "mu"]) - float(_bser.loc[_d, "agg_mu"])) < 1e-12
    assert abs(float(_bagg.loc[_j, "sigma"]) - float(_bser.loc[_d, "agg_sd"])) < 1e-12
# the delta = 0 legend names the vertex expert whose dotted curve it hides
assert str(_bagg.loc[0, "legend"]).endswith("(= expert %d)" % (_bvx + 1)), "fig6-3 delta=0 legend does not name expert 3"
assert int(_cfgd(_bimg)["delta0_vertex_expert"]) == _bvx + 1
# the tail-mass remark of the panel paragraph: the mass each untruncated
# judgment places outside [0,10], recomputed from the expert rows; the text
# names expert 2 as the largest, all of it above the top of the scale, and
# prints the largest share among the other experts.
_btail = _bimg["tail_mass"].set_index("expert")
_bimgc = _cfgd(_bimg)
assert list(_btail.index) == list(range(1, 8))
for _i in range(7):
    _plo = float(_ndtr((0.0 - _bmu[_i]) / _bsd[_i]))
    _phi = float(_ndtr((_bmu[_i] - 10.0) / _bsd[_i]))
    assert abs(float(_btail.loc[_i + 1, "P_below_0"]) - _plo) < 1e-15
    assert abs(float(_btail.loc[_i + 1, "P_above_10"]) - _phi) < 1e-15
    assert abs(float(_btail.loc[_i + 1, "P_outside_0_10"]) - (_plo + _phi)) < 1e-15
    assert float(_btail.loc[_i + 1, "P_outside_printed"]) == round(_plo + _phi, 3)
_bto = _btail["P_outside_0_10"].to_numpy(float)
assert int(np.argmax(_bto)) + 1 == 2 == int(_bimgc["tail_largest_expert"]), "the text names expert 2 as the largest tail"
assert float(_btail.loc[2, "P_above_10"]) > 0.999 * float(_btail.loc[2, "P_outside_0_10"]), "expert 2's tail is not above the scale"
assert abs(float(_bimgc["tail_largest_outside"]) - _bto.max()) < 1e-15
assert abs(float(_bimgc["tail_largest_among_others"]) - np.sort(_bto)[-2]) < 1e-15
assert float(_bimgc["tail_largest_outside_printed"]) == round(_bto.max(), 3)
assert float(_bimgc["tail_largest_among_others_printed"]) == round(np.sort(_bto)[-2], 3)
# the text names expert 4 as the next largest, printing its share as a
# rounded value (it is 0.0122 > 0.012, so the text must not state a bound)
assert int(np.argsort(_bto)[-2]) + 1 == 4 == int(_bimgc["tail_second_expert"]), "the text names expert 4 as the next largest tail"
assert np.sort(_bto)[-3] < np.sort(_bto)[-2] < np.sort(_bto)[-1], "the two largest tails must be strict for the text's ordering"
# the Figure 6.3 legends write delta as the mathtext symbol, as the caption does
assert all("$\\delta=" in str(_l) and "delta=" not in str(_l).replace("$\\delta=", "") for _l in _bagg["legend"]), "fig6-3 aggregate legends must use the symbol delta"
# Figure 6.4 plots diagnostics, not the scalarized objective
assert "objective" not in str(_bcfg["y_label"]).lower(), "fig6-4 y-axis must not be labelled as the objective"
_t69 = _wb("tab6-9_bim_results.xlsx")
_t69v = _t69["table_6_9"].set_index("quantity")
_t69p = _t69["table_6_9_printed"].set_index("quantity")
for _q, _col in (("minimum overlap", "minimum_overlap"),
                 ("mean overlap (consensus)", "consensus_Abar"),
                 ("collective fairness", "fairness_Phibar"),
                 ("weight on dissenter (exp 5)", "weight_on_dissenter")):
    for _tc, _d in (("consensus_confidence_d0", 0.0), ("fairness_weighted_d0.5", 0.5)):
        assert abs(float(_t69v.loc[_q, _tc]) - float(_bser.loc[_d, _col])) < 1e-12, \
            "tab6-9 %r disagrees with the fig6-4 sweep at delta=%.1f" % (_q, _d)
        assert float(_t69p.loc[_q, _tc]) == round(float(_bser.loc[_d, _col]), 3)
for _tc, _d in (("consensus_confidence_d0", 0.0), ("fairness_weighted_d0.5", 0.5)):
    assert str(_t69p.loc["aggregate distribution", _tc]) == "N(%.2f, %.2f^2)" % (
        round(float(_bser.loc[_d, "agg_mu"]), 2), round(float(_bser.loc[_d, "agg_sd"]), 2))
# (3e') the managerial paragraph: both aggregates lie in the standard level
#       [6,7), the base one in its top fifth and the fair one in its bottom
#       fifth, and the probability below the standard level (x < 6) and below
#       the defined level (x < 5) rises from the base to the fair solution;
#       the band probabilities are the normal cdf of the recorded aggregates.
_bband = _t69["band_probabilities"].set_index("solution")
for _tc, _d in (("consensus_confidence_d0", 0.0), ("fairness_weighted_d0.5", 0.5)):
    _mu_d, _sd_d = float(_bser.loc[_d, "agg_mu"]), float(_bser.loc[_d, "agg_sd"])
    assert 6.0 <= _mu_d < 7.0, "BIM delta=%.1f aggregate mean left the standard level" % _d
    for _col, _edge in (("P_below_standard_x_lt_6", 6.0), ("P_below_defined_x_lt_5", 5.0)):
        assert abs(float(_bband.loc[_tc, _col]) - float(_ndtr((_edge - _mu_d) / _sd_d))) < 1e-12
    assert float(_bband.loc[_tc, "P_below_standard_printed"]) == round(float(_bband.loc[_tc, "P_below_standard_x_lt_6"]), 2)
    assert float(_bband.loc[_tc, "P_below_defined_printed"]) == round(float(_bband.loc[_tc, "P_below_defined_x_lt_5"]), 2)
assert float(_bser.loc[0.0, "agg_mu"]) >= 6.8 - 1e-9 and float(_bser.loc[0.5, "agg_mu"]) <= 6.2 + 1e-9, \
    "the text places the base aggregate at the top and the fair one at the foot of the standard level"
for _col in ("P_below_standard_x_lt_6", "P_below_defined_x_lt_5"):
    assert float(_bband.loc["fairness_weighted_d0.5", _col]) > float(_bband.loc["consensus_confidence_d0", _col])
# (3f) solver settings are recorded
for _key in ("seed", "n_starts", "start_law", "ftol", "tie_rule", "overlap_rule", "bounds_method", "delta_grid", "draw_order",
             "certificate", "hull_scan_grid", "hull_scan_points", "hull_scan_improve_tol", "hull_scan_gap_tol"):
    assert _key in _bcfg, "fig6-4 config lacks the solver setting %r" % _key
assert int(_bcfg["seed"]) == NEW_SEED and int(_bcfg["n_starts"]) == _BIM_N_STARTS
print("   BIM fairness panel: solved sweep verified (bounds bracket vertices, bound-program agreement as printed,"
      " hull-scan certificate of every bound and sweep optimum, diagnostics recomputed from the weights,"
      " closed form == trapezoid, optimal over vertices and sweep solutions, qualitative and sweep-detail claims hold,"
      " text_values, band probabilities and tail masses recomputed, fig6-3 aggregates and tab6-9 == solved rows, settings recorded)")

# (4) DT case (Case II): the generated panel, the 70 solved programs, and
#     the 14 x 5 matrix are checked against invariants only (simplex,
#     probability normalization, the identities the book states, bounds
#     of the min-max normalization, and internal consistency between the
#     sheets); nothing is compared with a typed printed constant.  Since
#     R24 the checks also cover the candidate search (bounds contain every
#     evaluated start, the retained solution beats every candidate and
#     every pairwise transfer), the recorded restart statistics, the
#     leveling-down and active-set diagnostics, and every number of the
#     sheet text_numbers, each recomputed here from the other sheets.
_dt = _wb("fig6-5_dt_case.xlsx")
_dtc = _cfgd(_dt)
_dtN, _dtM = int(_dtc["N_experts"]), int(_dtc["M_barriers"])
_dtb = _dt["barriers"]
_dtids = list(_dtb["id"].astype(str))
_dteq = list(_dt["equity_objectives"]["eq"].astype(str))
assert _dteq == ["EQ1", "EQ2", "EQ3", "EQ4", "EQ5"], _dteq
_dtpg = (np.arange(int(_dtc["p_grid"])) + 0.5) / int(_dtc["p_grid"])


def _dt_chk_quantiles(k, lam, H5):
    return lam * (-np.log(1.0 - _dtpg * H5)) ** (1.0 / k)


def _dt_chk_ineq(u, s):
    """The five indices of tab:swf, written out afresh for the checks."""
    u = np.asarray(u, float)
    m = u.mean()
    if m <= 0:
        return 0.0
    if s == 1:
        return (u.max() - u.min()) / m
    if s == 2:
        return np.abs(u - m).mean() / m
    if s == 3:
        return u.std() / m
    if s == 4:
        return np.abs(u[:, None] - u[None, :]).sum() / (2.0 * u.size ** 2 * m)
    return np.abs(u - m).sum() / (2.0 * u.size * m)


# the panel: 420 fitted truncated Weibulls, proper profiles, positive parameters
_dtp = _dt["panel"]
assert len(_dtp) == _dtN * _dtM, len(_dtp)
_dtprof = _dtp[["p1", "p2", "p3", "p4", "p5"]].to_numpy(float)
assert np.allclose(_dtprof.sum(axis=1), 1.0, atol=1e-9) and (_dtprof >= 0).all()
assert (_dtp["kappa"] > 0).all() and (_dtp["lambda"] > 0).all()
assert ((_dtp["H5"] > 0) & (_dtp["H5"] <= 1)).all()
assert ((_dtp["rating"] >= 0) & (_dtp["rating"] <= 5)).all()
# generator identities: rating x_ik = clip(tau + leniency + bias + noise), the
# profile is the stated Gaussian kernel of the rating, and the equal-weight
# quantile-averaged mean equals the average of the fitted expert means (QA
# means are linear, thm:qa-moments) and the design mean
_dtr = np.clip(_dtp["tau"] + _dtp["leniency"] + _dtp["bias"] + _dtp["noise"], 0.0, 5.0)
assert np.allclose(_dtr, _dtp["rating"], atol=1e-12)
_dtker = np.exp(-((np.arange(5) + 0.5)[None, :] - _dtp["rating"].to_numpy(float)[:, None]) ** 2
                / (2.0 * float(_dtc["profile_kernel_sd"]) ** 2))
assert np.allclose(_dtker / _dtker.sum(axis=1, keepdims=True), _dtprof, atol=1e-12)
assert "x_ik" in str(_dtc["rating_rule"]) and "r_ik" in str(_dtc["background_bias_rule"])
_dtfm = _dtp.pivot(index="expert", columns="barrier", values="fitted_mean")[_dtids].to_numpy(float)
assert np.allclose(_dtfm.mean(axis=0), _dtb["equal_weight_mean"], atol=1e-9)
assert np.allclose(_dtb["equal_weight_mean"], _dtb["design_mean"], atol=1e-3), \
    "DT calibration off by more than 1e-3"
# tab:dt-elicit is a row of the generated panel, chosen by the stated rule
# (fitted mean closest to the barrier's equal-weight mean), and its fit is
# that row's; the text's reading of it (most mass on levels 4-5, first level
# below 0.001, mean in the fourth band) holds
_dtel = _dt["elicitation_b1"]
_sums_to_one(_dtel["probability"], "DT tab:dt-elicit profile")
_elb, _eli = str(_dtc["elicit_barrier"]), int(_dtc["elicit_expert"])
assert (_dtel["barrier"].astype(str) == _elb).all() and (_dtel["expert"] == _eli).all()
_elrow = _dtp[(_dtp["barrier"].astype(str) == _elb) & (_dtp["expert"] == _eli)].iloc[0]
assert np.allclose(_dtel["probability"].to_numpy(float), _elrow[["p1", "p2", "p3", "p4", "p5"]].to_numpy(float), atol=0)
_elcol = _dtfm[:, _dtids.index(_elb)]
assert int(np.argmin(np.abs(_elcol - _elcol.mean()))) + 1 == _eli
for _key, _col in (("kappa", "kappa"), ("lambda", "lambda"), ("mean", "fitted_mean"), ("H5", "H5")):
    assert float(_dtc["elicit_fit_" + _key]) == float(_elrow[_col]), _key
_b1k, _b1l = float(_dtc["elicit_fit_kappa"]), float(_dtc["elicit_fit_lambda"])
_b1H5 = 1.0 - math.exp(-(5.0 / _b1l) ** _b1k)
assert abs(_b1H5 - float(_dtc["elicit_fit_H5"])) < 1e-9
_b1mean = float(np.mean(_dt_chk_quantiles(_b1k, _b1l, _b1H5)))
assert abs(_b1mean - float(_dtc["elicit_fit_mean"])) < 1e-6, (_b1mean, _dtc["elicit_fit_mean"])
_elp = _dtel["probability"].to_numpy(float)
assert _elp[3] + _elp[4] > 0.5 and _elp[0] < 0.001 and 3.0 <= _b1mean < 4.0
assert str(_dtc["elicit_profile"]) == ", ".join("%.3f" % p for p in _elp)
# the 70 solved weight vectors lie on the simplex, and the recorded
# collective mean of each solve is the weighted fitted-expert mean
_dtw = _dt["weights"].set_index("expert")
_dts = _dt["solves"]
assert len(_dts) == _dtM * 5 and _dtw.shape == (_dtN, _dtM * 5)
for _, _row in _dts.iterrows():
    _col = "%s_%s" % (_row["barrier"], _row["eq"])
    _w = _dtw[_col].to_numpy(float)
    assert (_w >= 0).all()
    _sums_to_one(_w, "DT weights %s" % _col, tol=1e-8)
    _mu_lin = float(_w @ _dtfm[:, _dtids.index(str(_row["barrier"]))])
    assert abs(_mu_lin - float(_row["mu"])) < 1e-6, (_col, _mu_lin, _row["mu"])
    assert int((_w > float(_dtc["active_threshold"])).sum()) == int(_row["n_active"])
    assert abs(float(_w.max()) - float(_row["w_max"])) < 1e-12
# min-max bounds and normalized terms: V and W within their solved bounds, so
# V~ and W~ lie in [0,1]; the welfare bounds contain the vertex, edge-grid and
# face-sample values and the welfare of every recorded start of the program
# (they were refined until no candidate left them, see the config sheet)
assert (_dts["V_min"] < _dts["V_max"]).all() and (_dts["W_min"] < _dts["W_max"]).all()
assert ((_dts["V"] >= _dts["V_min"] - 1e-9) & (_dts["V"] <= _dts["V_max"] + 1e-9)).all()
assert ((_dts["V_tilde"] >= -1e-9) & (_dts["V_tilde"] <= 1 + 1e-9)).all()
assert ((_dts["W"] >= _dts["W_min"] - 1e-9) & (_dts["W"] <= _dts["W_max"] + 1e-9)).all()
assert ((_dts["W_tilde"] >= -1e-6) & (_dts["W_tilde"] <= 1 + 1e-6)).all(), float(_dts["W_tilde"].max())
assert (_dts["bound_refinements"] <= int(_dtc["bound_refine_max"])).all()
assert (_dts["W_max"] >= _dts[["W_max_cloud", "W_max_search"]].max(axis=1) - 1e-15).all()
assert (_dts["W_min"] <= _dts[["W_min_cloud", "W_min_search", "W_min_vertex"]].min(axis=1) + 1e-15).all()
assert np.allclose(_dts["objective"], _dts["W_tilde"] - _dts["V_tilde"], atol=1e-9)
_dtst = _dt["starts"]
_dtbt = float(_dtc["bound_tol"])
_dtst_b = _dtst.merge(_dts[["barrier", "eq", "W_min", "W_max", "objective", "mu"]], on=["barrier", "eq"],
                      suffixes=("", "_cell"))
assert ((_dtst_b["W"] <= _dtst_b["W_max"] + _dtbt) & (_dtst_b["W"] >= _dtst_b["W_min"] - _dtbt)).all(), \
    "a recorded start of the Case II program lies outside the welfare bounds"
# the retained solution is verified: it is the best recorded start of its
# program, at least as good as every interior start, zoomed edge point, the
# edge grid, and the face sample, and no pairwise transfer gains on it
_dtat, _dtnt = float(_dtc["agreement_tol"]), float(_dtc["near_optimal_tol"])
for (_b, _e), _g in _dtst_b.groupby(["barrier", "eq"]):
    _row = _dts[(_dts["barrier"] == _b) & (_dts["eq"] == _e)].iloc[0]
    assert abs(float(_g["objective"].max()) - float(_row["objective"])) < 1e-12, (_b, _e)
    assert int(_g["retained"].sum()) == 1 and str(_g.loc[_g["retained"], "start"].iloc[0]) == str(_row["source"])
    assert np.allclose(_g["gap"], _row["objective"] - _g["objective"], atol=1e-12)
    _rnd = _g[_g["kind"] == "interior"]
    assert len(_rnd) == int(_row["n_interior_starts"]) == 1 + int(_dtc["dirichlet_starts_program"])
    assert list(_rnd["start"].astype(str)) == ["barycentre"] + [
        "dirichlet_%d" % _n for _n in range(1, int(_dtc["dirichlet_starts_program"]) + 1)]
    assert int((_rnd["objective"] >= _row["objective"] - _dtat).sum()) == int(_row["n_interior_reach"])
    _rch = _g[_g["objective"] >= _row["objective"] - _dtat]
    assert len(_rch) == int(_row["n_reach"]) and _rch["reaches"].all()
    assert abs(float(_rch["mu"].max() - _rch["mu"].min()) - float(_row["mu_spread_reach"])) < 1e-12
    _nr = _g[_g["objective"] >= _row["objective"] - _dtnt]
    assert abs(float((_nr["mu"] - _row["mu"]).abs().max()) - float(_row["mu_dev_near"])) < 1e-12
for _c in ("objective_interior_best", "objective_edge_zoom_best", "objective_edge_grid_max",
           "objective_face_max", "objective_barycentre"):
    assert (_dts["objective"] >= _dts[_c] - 1e-12).all(), _c
assert (_dts["transfer_gain"] <= float(_dtc["transfer_tol"])).all()
# the welfare identity W = Ubar (1 - I) and the recorded utilities
assert np.allclose(_dts["W"], _dts["Ubar"] * (1.0 - _dts["I"]), atol=1e-12)
_dtU = _dt["utilities"].set_index("expert")
_dtA = _dt["overlaps"].set_index("expert")
_dtAe = _dt["overlaps_equal_weights"].set_index("expert")
assert ((_dtU >= 0) & (_dtU <= 1)).all().all() and ((_dtA >= 0) & (_dtA <= 1)).all().all()
_dtVi = {}
for _b in _dtids:
    _pb = _dtp[_dtp["barrier"].astype(str) == _b].sort_values("expert")
    _Q = np.vstack([_dt_chk_quantiles(k, l, h) for k, l, h in zip(_pb["kappa"], _pb["lambda"], _pb["H5"])])
    _dtVi[_b] = _Q.var(axis=1)
for _, _row in _dts.iterrows():
    _b, _s = str(_row["barrier"]), int(str(_row["eq"])[2:])
    _col = "%s_%s" % (_b, _row["eq"])
    _A, _Ae, _U = _dtA[_col].to_numpy(float), _dtAe[_b].to_numpy(float), _dtU[_col].to_numpy(float)
    assert abs(float(_U.mean()) - float(_row["Ubar"])) < 1e-9
    assert abs(float(_A.mean()) - float(_row["Abar"])) < 1e-9
    assert abs(float(_A.min()) - float(_row["Amin"])) < 1e-9
    assert abs(float(_dt_chk_ineq(_U, _s)) - float(_row["I"])) < 1e-9
    # leveling-down diagnostics against equal weights (Prop. consensus-fairness (iii))
    assert abs(float(_Ae.mean()) - float(_row["Abar_equal_weights"])) < 1e-12
    assert abs(float(_A.mean() - _Ae.mean()) - float(_row["dAbar"])) < 1e-12
    assert int((_A < _Ae).sum()) == int(_row["n_overlap_lower"])
    assert abs(_dt_chk_ineq(_A, _s) - float(_row["I_overlap"])) < 1e-9
    assert abs(_dt_chk_ineq(_Ae, _s) - float(_row["I_overlap_equal_weights"])) < 1e-9
    assert bool(_row["leveling_down"]) == bool(_A.mean() < _Ae.mean() and _dt_chk_ineq(_A, _s) < _dt_chk_ineq(_Ae, _s))
    # the active set against the fitted means and the expert variances
    _w = _dtw[_col].to_numpy(float)
    _act = set(np.where(_w > float(_dtc["active_threshold"]))[0].tolist())
    _fm = _dtfm[:, _dtids.index(_b)]
    assert str(_row["active_experts"]) == ",".join(str(i + 1) for i in sorted(_act))
    assert bool(_row["active_is_mean_extremes"]) == (_act == {int(np.argmax(_fm)), int(np.argmin(_fm))})
    assert bool(_row["active_has_sharpest"]) == (int(np.argmin(_dtVi[_b])) in _act)
    assert bool(_row["active_has_flattest"]) == (int(np.argmax(_dtVi[_b])) in _act)
    _fa = _fm[sorted(_act)]
    assert bool(_row["active_straddle_center"]) == bool(_fa.max() > _fm.mean() > _fa.min())
# the 14 x 5 matrix, its average, and the ranking follow from the solves
_dtMx = _dt["matrix"].set_index("id")
assert list(_dtMx.index.astype(str)) == _dtids
for _, _row in _dts.iterrows():
    assert abs(float(_dtMx.loc[str(_row["barrier"]), str(_row["eq"])]) - float(_row["mu"])) < 1e-12
_dtMv = _dtMx[_dteq].to_numpy(float)
assert np.allclose(_dtMx["avg"], _dtMv.mean(axis=1), atol=1e-12)
assert np.allclose(_dtMx["spread"], _dtMv.max(axis=1) - _dtMv.min(axis=1), atol=1e-12)
assert np.allclose(_dtMx["eq1_departure"], np.abs(_dtMv[:, 0] - _dtMv[:, 1:].mean(axis=1)), atol=1e-12)
assert np.allclose(_dtMx["span_eq2_eq5"], _dtMv[:, 1:].max(axis=1) - _dtMv[:, 1:].min(axis=1), atol=1e-12)
assert np.allclose(_dtMx["abs_eq2_minus_eq5"], np.abs(_dtMv[:, 1] - _dtMv[:, 4]), atol=1e-12)
_dtord = sorted(range(_dtM), key=lambda r: -float(_dtMx["avg"].iloc[r]))
_dtorder = [_dtids[r] for r in _dtord]
assert _dtorder == list(_dt["ranking"]["id"].astype(str)), _dtorder
assert list(_dtMx.sort_values("rank").index.astype(str)) == _dtorder
_dttriad = str(_dtc["triad"]).split(",")
assert _dttriad == _dtorder[:3], (_dttriad, _dtorder[:3])
# the text's reading of the ranking: B14 alone in the fifth level, the
# bottom three alone below the fourth, and the same order under every index
_dtavg = _dtMx["avg"].astype(float)
assert list(_dtavg[_dtavg >= 4.0].index.astype(str)) == _dtorder[:1]
assert sorted(_dtavg[_dtavg < 3.0].index.astype(str)) == sorted(_dtorder[-3:])
for _j, _e in enumerate(_dteq):
    _oe = [_dtids[r] for r in np.argsort(-_dtMv[:, _j], kind="stable")]
    assert ">".join(_oe) == str(_dt["per_objective_order"].set_index("eq").loc[_e, "order"])
    assert _oe == _dtorder, "Case II: the order under %s differs from the average's" % _e
# the adjacent pairs that tie at the table's two decimals, and their margins
_dtties = []
for _j, _e in enumerate(_dteq):
    for _p in range(_dtM - 1):
        _a, _bb = _dtord[_p], _dtord[_p + 1]
        if "%.2f" % _dtMv[_a, _j] == "%.2f" % _dtMv[_bb, _j]:
            _dtties.append((_e, _dtids[_a], _dtids[_bb], _dtMv[_a, _j] - _dtMv[_bb, _j]))
_dtnt_ = _dt["near_ties"]
assert [(str(r.eq), str(r.upper), str(r.lower)) for r in _dtnt_.itertuples()] == [t[:3] for t in _dtties]
assert np.allclose(_dtnt_["margin"].to_numpy(float), [t[3] for t in _dtties], atol=1e-12)
# every number of the sheet text_numbers, recomputed from the sheets above
_dttn = _dt["text_numbers"].set_index("key")
_dtv = {}
_dtv["elicit_expert"] = _eli
_dtv["elicit_p1_max"] = float(_elp[0])
_dtv["elicit_kappa"], _dtv["elicit_lambda"], _dtv["elicit_mean"] = _b1k, _b1l, _b1mean
for _p in (0, 1, 2, 3, _dtM - 1, _dtM - 2, _dtM - 3):
    _dtv["avg_rank%d_%s" % (_p + 1, _dtorder[_p])] = float(_dtavg[_dtorder[_p]])
_dtv["n_orders_equal_to_average"] = 5
_dtv["n_ties_at_2dp"] = len(_dtties)
for _t in _dtties:
    _dtv["tie_%s_%s_%s" % _t[:3]] = _t[3]
_dtsp = _dtMx["spread"].astype(float)
for _b in list(_dtsp.sort_values(kind="stable").index[:2]) + list(_dtsp.sort_values(ascending=False, kind="stable").index[:2]) + _dttriad[1:]:
    _dtv["spread_%s" % _b] = float(_dtsp[_b])
_dtv["spread_max"] = float(_dtsp.max())
_dtv["n_eq1_lowest"] = int((_dtMv[:, 0] < _dtMv[:, 1:].min(axis=1) - float(_dtc["tie_tol"])).sum())
_dtv["n_eq1_tie_lowest"] = int((np.abs(_dtMv[:, 0] - _dtMv[:, 1:].min(axis=1)) <= float(_dtc["tie_tol"])).sum())
_dtv["n_eq1_highest"] = int((_dtMv[:, 0] > _dtMv[:, 1:].max(axis=1) + float(_dtc["tie_tol"])).sum())
_dtdep = _dtMx["eq1_departure"].astype(float).sort_values(ascending=False, kind="stable")
for _b in _dtdep.index[:2]:
    _dtv["eq1_departure_%s" % _b] = float(_dtdep[_b])
_dtsp4 = _dtMx["span_eq2_eq5"].astype(float)
_dtv["span4_max_%s" % _dtsp4.idxmax()] = float(_dtsp4.max())
_dtv["span4_max_other"] = float(_dtsp4.drop(_dtsp4.idxmax()).max())
_dtd25 = _dtMx["abs_eq2_minus_eq5"].astype(float)
_dtv["eq2_eq5_max_%s" % _dtd25.idxmax()] = float(_dtd25.max())
_dt2 = _dts[_dts["n_active"] == 2]
_dtv["n_two_expert"] = len(_dt2)
_dtv["n_active_max"] = int(_dts["n_active"].max())
_dtv["n_pair_mean_extremes"] = int(_dt2["active_is_mean_extremes"].sum())
_dtv["n_pair_straddle"] = int(_dt2["active_straddle_center"].sum())
_dtv["n_sharpest_active"] = int(_dts["active_has_sharpest"].sum())
_dtv["n_flattest_active"] = int(_dts["active_has_flattest"].sum())
_dtv["barriers_flattest_not_active"] = len(set(_dts.loc[~_dts["active_has_flattest"].astype(bool), "barrier"]))
_dtv["Abar_min"], _dtv["Abar_max"] = float(_dts["Abar"].min()), float(_dts["Abar"].max())
_dtv["Abar_eq_min"] = float(_dtAe.mean(axis=0).min())
_dtv["Abar_eq_max"] = float(_dtAe.mean(axis=0).max())
_dtv["n_Abar_lower"] = int((_dts["Abar"] < _dts["Abar_equal_weights"]).sum())
_dtv["Abar_drop_min"] = float((_dts["Abar_equal_weights"] - _dts["Abar"]).min())
_dtv["Abar_drop_max"] = float((_dts["Abar_equal_weights"] - _dts["Abar"]).max())
_dtv["Abar_drop_max_pct"] = float(((_dts["Abar_equal_weights"] - _dts["Abar"]) / _dts["Abar_equal_weights"]).max() * 100)
_dtv["n_all_overlaps_lower"] = int((_dts["n_overlap_lower"] == _dtN).sum())
_dtv["n_leveling_down"] = int(_dts["leveling_down"].sum())
_dtv["Ubar_min"], _dtv["Ubar_max"] = float(_dts["Ubar"].min()), float(_dts["Ubar"].max())
_dtv["V_rel_max_pct"] = float((_dts["V"] / _dts["V_equal_weights"] - 1.0).abs().max() * 100)
_dtv["W_tilde_min"], _dtv["W_tilde_max"] = float(_dts["W_tilde"].min()), float(_dts["W_tilde"].max())
_dtv["expert_mean_sd_min"] = float(_dtfm.std(axis=0).min())
_dtv["expert_mean_sd_max"] = float(_dtfm.std(axis=0).max())
for _b in _dttriad:
    _dtv["triad_%s_%s" % (_b, _dtc["triad_objective"])] = float(_dtMx.loc[_b, str(_dtc["triad_objective"])])
_dtv["interior_reach_median"] = float(_dts["n_interior_reach"].median())
_dtv["n_interior_reach_zero"] = int((_dts["n_interior_reach"] == 0).sum())
_dtv["interior_reach_min_nonzero"] = int(_dts.loc[_dts["n_interior_reach"] > 0, "n_interior_reach"].min())
_dtv["n_bound_refined"] = int((_dts["bound_refinements"] > 0).sum())
_dtv["n_W_min_at_vertex"] = int((np.abs(_dts["W_min"] - _dts["W_min_vertex"]) <= 1e-12).sum())
_dtv["n_retained_from_interior"] = int(_dts["source"].astype(str).str.match(r"^(barycentre|dirichlet_\d+)$").sum())
_dtfla = _dts.groupby("barrier")["active_has_flattest"].agg(["all", "any"])
_dtv["n_barriers_flattest_active"] = int(_dtfla["all"].sum())
_dtv["n_barriers_flattest_mixed"] = int((_dtfla["any"] & ~_dtfla["all"]).sum())
assert _dtv["n_barriers_flattest_mixed"] == 0, "the flattest expert is active under some but not all indices"
_dtv["n_barriers_leveling_down"] = int(_dts.groupby("barrier")["leveling_down"].any().sum())
_dtv["n_Ubar_higher"] = int((_dts["Ubar"] > _dts["Ubar_equal_weights"]).sum())
for _b in _dtids:
    _k = _dtids.index(_b)
    if abs(_dtMv[_k, 0] - _dtMv[_k, 1:].min()) <= float(_dtc["tie_tol"]):
        _jt = 1 + int(np.argmin(_dtMv[_k, 1:]))
        _dtv["eq1_tie_%s_partner" % _b] = _jt + 1
        _dtv["eq1_tie_%s_weight_diff" % _b] = float(np.abs(_dtw["%s_EQ1" % _b] - _dtw["%s_%s" % (_b, _dteq[_jt])]).max())
        assert _dtv["eq1_tie_%s_weight_diff" % _b] < 1e-6, "EQ1 ties another index on %s without the same solution" % _b
# the program no interior start reaches: its interior starts either collapse
# onto one expert or stop between zero_reach_gap_min and _max below the
# retained value with (nearly) its collective mean
_dtzero = _dts[_dts["n_interior_reach"] == 0]
_dtz = _dtst[_dtst["kind"] == "interior"].merge(_dtzero[["barrier", "eq", "mu"]], on=["barrier", "eq"],
                                              suffixes=("", "_cell"))
_dtzs = _dtz[_dtz["n_active"] > 1]
_dtv["zero_reach_n_collapsed"] = int((_dtz["n_active"] == 1).sum())
_dtv["zero_reach_n_spread"] = len(_dtzs)
_dtv["zero_reach_nact_min"] = int(_dtzs["n_active"].min())
_dtv["zero_reach_nact_max"] = int(_dtzs["n_active"].max())
_dtv["zero_reach_gap_min"] = float(_dtzs["gap"].min())
_dtv["zero_reach_gap_max"] = float(_dtzs["gap"].max())
_dtv["zero_reach_mu_dev_max"] = float((_dtzs["mu"] - _dtzs["mu_cell"]).abs().max())
assert (_dtzero["source"].astype(str).str.startswith("edge")).all(), "a program no interior start reaches is not solved by the edge search"
assert _dtv["zero_reach_n_collapsed"] + _dtv["zero_reach_n_spread"] == len(_dtzero) * (1 + int(_dtc["dirichlet_starts_program"]))
# the mean range over the starts that reach the retained value: one program
# returns two means (both printing alike at two decimals), the others agree

for (_b, _e), _g in _dtst_b.groupby(["barrier", "eq"]):
    _row = _dts[(_dts["barrier"] == _b) & (_dts["eq"] == _e)].iloc[0]
    _gr = _g[_g["objective"] >= _row["objective"] - _dtat]
    assert abs(float(_gr["mu"].min()) - float(_row["mu_reach_min"])) < 1e-12
    assert abs(float(_gr["mu"].max()) - float(_row["mu_reach_max"])) < 1e-12
_dtmsr = (_dts["mu_reach_max"] - _dts["mu_reach_min"]).to_numpy(float)
assert np.allclose(_dtmsr, _dts["mu_spread_reach"], atol=1e-12)
_dtkx = int(np.argmax(_dtmsr))
_dtv["mu_reach_split_cell"] = _dtkx
_dtv["mu_reach_split_lo"] = float(_dts["mu_reach_min"].iloc[_dtkx])
_dtv["mu_reach_split_hi"] = float(_dts["mu_reach_max"].iloc[_dtkx])
assert "%.2f" % _dtv["mu_reach_split_lo"] == "%.2f" % _dtv["mu_reach_split_hi"] == "%.2f" % float(_dts["mu"].iloc[_dtkx])
_dtv["mu_reach_spread_max_other"] = float(np.delete(_dtmsr, _dtkx).max())
# the flat programs (near-optimal starts moving the mean by more than flat_tol)
_dtflat_tol = float(_dtc["flat_tol"])
_dtflat = _dts[_dts["mu_dev_near"] > _dtflat_tol]
_dtv["n_flat"] = len(_dtflat)
for _r in _dtflat.itertuples():
    _dtv["mu_dev_near_%s_%s" % (_r.barrier, _r.eq)] = float(_r.mu_dev_near)
_dtv["mu_dev_near_max_other"] = float(_dts.loc[_dts["mu_dev_near"] <= _dtflat_tol, "mu_dev_near"].max())
# the doubled-discretization check: its programs are those of the tied pairs
# and the flat programs; every fine solution lies on the simplex, its mean is
# the weighted fine-grid expert mean recomputed from the panel sheet, it
# keeps the retained pair of experts, and every tied pair keeps its order
_dtff = int(_dtc["fine_factor"])
_dtfc = _dt["fine_check"]
_dtfw = _dt["fine_weights"].set_index("expert")
_dtfcells = []
for _r in _dtnt_.itertuples():
    for _b in (str(_r.upper), str(_r.lower)):
        if (_b, str(_r.eq)) not in _dtfcells:
            _dtfcells.append((_b, str(_r.eq)))
for _r in _dtflat.itertuples():
    if (str(_r.barrier), str(_r.eq)) not in _dtfcells:
        _dtfcells.append((str(_r.barrier), str(_r.eq)))
assert [(str(r.barrier), str(r.eq)) for r in _dtfc.itertuples()] == _dtfcells
assert list(_dtfw.columns) == ["%s_%s" % c for c in _dtfcells]
_dtpgf = (np.arange(_dtff * int(_dtc["p_grid"])) + 0.5) / (_dtff * int(_dtc["p_grid"]))
_dtfmu = {}
for _r in _dtfc.itertuples():
    _b, _e = str(_r.barrier), str(_r.eq)
    _wf = _dtfw["%s_%s" % (_b, _e)].to_numpy(float)
    assert (_wf >= 0).all() and abs(_wf.sum() - 1.0) < 1e-12
    _pb = _dtp[_dtp["barrier"].astype(str) == _b].sort_values("expert")
    _mif = np.array([np.mean(l * (-np.log(1.0 - _dtpgf * h)) ** (1.0 / k))
                     for k, l, h in zip(_pb["kappa"], _pb["lambda"], _pb["H5"])])
    assert abs(float(_wf @ _mif) - float(_r.mu_fine)) < 1e-9, (_b, _e)
    assert abs(float(_r.mu) - float(_dtMx.loc[_b, _e])) < 1e-12
    assert abs(float(_r.mu_fine_minus_mu) - (float(_r.mu_fine) - float(_r.mu))) < 1e-12
    assert str(_r.active_experts_fine) == ",".join(str(i + 1) for i in np.where(_wf > float(_dtc["active_threshold"]))[0])
    assert abs(float(_r.w_l1_fine_vs_coarse) - float(np.abs(_wf - _dtw["%s_%s" % (_b, _e)].to_numpy(float)).sum())) < 1e-12
    assert float(_r.transfer_gain_fine) <= float(_dtc["transfer_tol"])
    _dtfmu[(_b, _e)] = float(_r.mu_fine)
_dtv["fine_pgrid"] = _dtff * int(_dtc["p_grid"])
_dtv["fine_xbins"] = _dtff * int(_dtc["x_bins"])
_dtv["n_fine_tie_cells"] = int(_dtfc["reason"].astype(str).str.contains("near tie").sum())
_dtv["n_fine_same_pair"] = int((_dtfc["active_experts"].astype(str) == _dtfc["active_experts_fine"].astype(str)).sum())
assert _dtv["n_fine_same_pair"] == len(_dtfc), "a program solved again on the doubled discretization changes its pair"
_dtv["fine_mu_shift_max"] = float(_dtfc["mu_fine_minus_mu"].abs().max())
for _r in _dtnt_.itertuples():
    _m = _dtfmu[(str(_r.upper), str(_r.eq))] - _dtfmu[(str(_r.lower), str(_r.eq))]
    assert abs(float(_r.upper_mu_fine) - _dtfmu[(str(_r.upper), str(_r.eq))]) < 1e-12
    assert abs(float(_r.lower_mu_fine) - _dtfmu[(str(_r.lower), str(_r.eq))]) < 1e-12
    assert abs(float(_r.margin_fine) - _m) < 1e-12
    assert _m > 0, "a tied pair reverses on the doubled discretization"
    _dtv["tie_fine_%s_%s_%s" % (_r.eq, _r.upper, _r.lower)] = _m
# near-optimal mean ranges recomputed from the starts, and the adjacent pairs they make overlap
_dtnearr = {}
for (_b, _e), _g in _dtst_b.groupby(["barrier", "eq"]):
    _row = _dts[(_dts["barrier"] == _b) & (_dts["eq"] == _e)].iloc[0]
    _gn = _g[_g["objective"] >= _row["objective"] - _dtnt]
    assert abs(float(_gn["mu"].min()) - float(_row["mu_near_min"])) < 1e-12
    assert abs(float(_gn["mu"].max()) - float(_row["mu_near_max"])) < 1e-12
    _dtnearr[(_b, _e)] = (float(_gn["mu"].min()), float(_gn["mu"].max()))
_dtnov = [(_e, _a, _bb) for _e in _dteq for _a, _bb in zip(_dtorder[:-1], _dtorder[1:])
          if _dtnearr[(_bb, _e)][1] >= _dtnearr[(_a, _e)][0]]
_dtv["n_near_optimal_overlaps"] = len(_dtnov)
for _e, _a, _bb in _dtnov:
    _dtv["near_overlap_%s_%s_%s_lower_max" % (_e, _a, _bb)] = _dtnearr[(_bb, _e)][1]
    assert _dtnearr[(_bb, _e)][1] > float(_dtMx.loc[_a, _e]), \
        "the text says a near-optimal solution of the lower barrier's program alone reverses the pair"
assert [bool(x) for x in _dtnt_["near_optimal_overlap"]] == [(str(r.eq), str(r.upper), str(r.lower)) in _dtnov
                                                              for r in _dtnt_.itertuples()]
assert set(_dtnov) <= {(str(r.eq), str(r.upper), str(r.lower)) for r in _dtnt_.itertuples()}, \
    "a near-optimal overlap outside the pairs the text names as tied"
_dtv["n_retained_from_edges"] = int(_dts["source"].astype(str).str.startswith("edge").sum())
assert set(_dtv) == set(_dttn.index), sorted(set(_dtv) ^ set(_dttn.index))
for _key, _val in _dtv.items():
    assert abs(float(_dttn.loc[_key, "value"]) - float(_val)) < 1e-9, (_key, _dttn.loc[_key, "value"], _val)
# the numbers the text uses as upper bounds ("within", "at most", "less than")
# are not exceeded by the values they bound
for _key in ["span4_max_other", "Abar_drop_max_pct", "V_rel_max_pct", "mu_dev_near_max_other",
             "zero_reach_mu_dev_max", "mu_reach_spread_max_other", "fine_mu_shift_max",
             "elicit_p1_max"] + [k for k in _dttn.index if str(k).startswith("eq2_eq5_max_")]:
    assert float(_dttn.loc[_key, "value"]) <= float(_dttn.loc[_key, "printed"]), (_key, _dttn.loc[_key])
# the structural claims of Sections 6.3.6-6.3.7, asserted on the solved outputs:
# every solution is two-expert, its pair is the pair of extreme fitted means
# (one above the panel centre, one below) and contains the sharpest expert;
# the mean overlap is below its equal-weight value at every solution, and the
# leveling-down move occurs on every barrier; W_min sits at a vertex in every
# program; all interior starts reach the retained value in the median program
assert (_dts["n_active"] == 2).all(), "Case II: a solution is not two-expert"
for _c in ("active_is_mean_extremes", "active_straddle_center", "active_has_sharpest"):
    assert _dts[_c].astype(bool).all(), _c
assert (_dts["dAbar"] < 0).all() and (_dts["Abar"] < _dts["Abar_equal_weights"]).all()
assert _dts.groupby("barrier")["leveling_down"].any().all()
assert (np.abs(_dts["W_min"] - _dts["W_min_vertex"]) <= 1e-12).all()
assert float(_dts["n_interior_reach"].median()) == float(_dts["n_interior_starts"].iloc[0])
# the triad densities of Figure 6.7 integrate to one on [0,5]
_dtd = _dt["triad_density"]
for _b in _dttriad:
    _area = float(_trapz(_dtd["f_" + _b].to_numpy(float), _dtd["x"].to_numpy(float)))
    assert abs(_area - 1.0) < 2e-3, (_b, _area)
print("   DT case: panel identities, tab:dt-elicit a panel row, 70 simplex weight vectors, bounds containing "
      "every candidate, verified optima, restart statistics, leveling-down and active-set diagnostics, "
      "W = Ubar(1 - I), matrix/average/ranking/near-tie consistency, doubled-discretization check, "
      "text numbers, triad densities verified")

# (5) BDT case (generated and solved): invariants of the solved outputs, never
#     typed constants.  Simplex and normalization identities, the quantile-
#     averaging mean identity at every layer and every delta, the hierarchical
#     variance bound (prop:hier-variance), the k-means / TOPSIS / BWM records,
#     the panel targets, and the Figure 6.8 densities.
_bdt = _wb("fig6-8_bdt_case.xlsx")
_bdtc = _cfgd(_bdt)
_sums_to_one(_bdt["groups"]["Omega"], "BDT group Omega")
_sums_to_one(_bdt["dimensions"]["weight"], "BDT dimension weights")
_bgrp = _bdt["groups"]
_bsizes = [int(s) for s in _bgrp["size"]]
assert sum(_bsizes) == int(_bdtc["n_experts"]) == 15
_batt = _bdt["expert_attributes"]
assert len(_batt) == 15 and _batt.shape[1] >= 6
assert list(_batt["group_kmeans"].value_counts().reindex(["G1", "G2", "G3"])) == _bsizes, \
    "k-means membership counts differ from the groups sheet"
assert (_batt["group_kmeans"] == _batt["group_design"]).all()
_bkt = _bdt["kmeans_topsis"]
assert ((_bkt["closeness"] >= 0) & (_bkt["closeness"] <= 1)).all()
assert np.allclose(_bkt["Omega"], _bkt["closeness"] / _bkt["closeness"].sum())
assert np.allclose(_bkt["Omega"], _bgrp["Omega"])
_bOmega = _bgrp["Omega"].to_numpy(float)
assert _bOmega[2] > _bOmega[0] > _bOmega[1], "designed Omega ordering G3 > G1 > G2 lost: %r" % (_bOmega,)
# the stratum descriptions of Table tab:bdt-groups against the k-means centroids:
# G3 "senior, long experience and high credentials" has the largest years and
# credential centroids; G2 "junior, broad but shallow project record" the smallest
# years and credential centroids and the most digital-twin projects
_bcy = _bkt["centroid_years_experience"].to_numpy(float)
_bcc = _bkt["centroid_credential_level"].to_numpy(float)
_bcp = _bkt["centroid_dt_projects"].to_numpy(float)
assert _bcy.argmax() == 2 and _bcc.argmax() == 2, "G3 is not the senior stratum"
assert _bcy.argmin() == 1 and _bcc.argmin() == 1 and _bcp.argmax() == 1, \
    "G2 is not the junior stratum with the broadest project record"
assert "senior" in str(_bgrp.loc[2, "stratum"]) and "junior" in str(_bgrp.loc[1, "stratum"])
# BWM linear model: the recorded xi^L bounds every comparison deviation
_bdim = _bdt["dimensions"]
_bw = _bdim["weight"].to_numpy(float)
_baB = _bdim["bwm_best_to_others"].to_numpy(float)
_baW = _bdim["bwm_others_to_worst"].to_numpy(float)
_bids = [str(v) for v in _bdim["id"]]
_bbi = _bids.index(str(_bdtc["bwm_best"]))
_bwi = _bids.index(str(_bdtc["bwm_worst"]))
assert _bw[_bbi] == _bw.max() and _bw[_bwi] == _bw.min()
_bxi = float(_bdtc["bwm_xi"])
assert (np.abs(_bw[_bbi] - _baB * _bw) <= _bxi + 1e-7).all(), "BWM best-to-others deviations exceed xi"
assert (np.abs(_bw - _baW * _bw[_bwi]) <= _bxi + 1e-7).all(), "BWM others-to-worst deviations exceed xi"
# BWM consistency (Section 6.4 text): the input-based ratio CR^I of Liang et al.
# (2020), recomputed from the comparison vectors, equals the recorded value, is
# attained at the dimensions the text names, and sits AT the threshold of 0.1667
# they tabulate for a_BW = 3.  At that scale the threshold is attained by
# acceptable and unacceptable systems alike, so the acceptability the text states
# rests on the ordinal-consistency condition of their Definition 4, recomputed below
_baBW = float(_baB[_bwi])
assert _baBW == float(_baW[_bbi]) == float(_bdtc["bwm_a_BW"]) == 3.0
_bcri_j = np.abs(_baB * _baW - _baBW) / (_baBW * _baBW - _baBW)
assert np.allclose(_bcri_j, _bdim["bwm_CRI_j"].to_numpy(float), atol=1e-15)
assert np.allclose(_baB * _baW, _bdim["bwm_product_aB_aW"].to_numpy(float))
assert abs(float(_bcri_j.max()) - float(_bdtc["bwm_CRI"])) < 1e-15
assert [_bids[j] for j in range(7) if _bcri_j[j] == _bcri_j.max()] == ["D2", "D4", "D5", "D7"] \
    == str(_bdtc["bwm_CRI_argmax"]).split(","), "the dimensions attaining CR^I differ from the text"
assert (_baB * _baW)[[1, 3, 4, 6]].tolist() == [4.0] * 4
assert float(_bdtc["bwm_CRI"]) <= float(_bdtc["bwm_CRI_threshold"]) == 0.1667, "BWM CR^I above the threshold"
assert float(_bdtc["bwm_CRI_threshold"]) - float(_bdtc["bwm_CRI"]) < 1e-4, "BWM CR^I no longer at the boundary"
_bord_bad = [(i, j) for i in range(7) for j in range(i + 1, 7)
             if not ((_baB[i] - _baB[j]) * (_baW[j] - _baW[i]) > 0
                     or (_baB[i] == _baB[j] and _baW[i] == _baW[j]))]
assert _bord_bad == [] and int(_bdtc["bwm_ordinal_consistent"]) == 1 \
    and int(_bdtc["bwm_ordinal_n_violations"]) == 0, "BWM comparison vectors are ordinal-inconsistent"
# the two vectors order the dimensions identically: D1, D3 above D2, D4, D5, D7 above D6
assert [_bids[j] for j in np.argsort(_baB, kind="stable")] == [_bids[j] for j in np.argsort(-_baW, kind="stable")] \
    == ["D1", "D3", "D2", "D4", "D5", "D7", "D6"]
# panel: every truncated mean equals its target, and every member sd is the
# closed-form sd of its truncated Weibull, recomputed here from kappa and lambda
_bpan = _bdt["panel_weibull"]
assert len(_bpan) == 15 * 22
assert np.allclose(_bpan["truncated_mean"], _bpan["target_mean"], atol=1e-5)
_bk_, _bl_ = _bpan["kappa"].to_numpy(float), _bpan["lambda"].to_numpy(float)
_bsd_cf = np.sqrt(_bdt_trunc_moment(_bk_, _bl_, 2) - _bdt_trunc_moment(_bk_, _bl_, 1) ** 2)
assert np.allclose(_bpan["truncated_sd"].to_numpy(float), _bsd_cf, rtol=0, atol=1e-12)
assert ((_bpan["truncated_sd"] > 0.4) & (_bpan["truncated_sd"] < 0.8)).all()
assert abs(float(_bdtc["member_sd_mean"]) - float(_bsd_cf.mean())) < 1e-12
assert abs(float(_bdtc["member_sd_min"]) - float(_bsd_cf.min())) < 1e-12
assert abs(float(_bdtc["member_sd_max"]) - float(_bsd_cf.max())) < 1e-12
# every solved weight vector is a probability vector over its subgroup only
_bsol = _bdt["intra_group_solutions"]
_wcols = ["w_e%d" % i for i in range(1, 16)]
assert len(_bsol) == 6 * 3 * 22
_tt = float(_bdtc["tie_tol"])
for _, _r in _bsol.iterrows():
    _wv = _r[_wcols].to_numpy(float)
    _mem = ~np.isnan(_wv)
    assert _mem.sum() == _bsizes[int(_r["group"][1]) - 1]
    _wv = _wv[_mem]
    assert (_wv >= -1e-12).all() and abs(_wv.sum() - 1.0) < 1e-8, "BDT weight vector off the simplex"
    assert abs(_gini_vec(np.clip(_wv, 0, None)) - float(_r["gini_w"])) < 1e-8
    assert 0.0 <= float(_r["Phi"]) <= 1.0 + 1e-12
    assert float(_r["Phi_min"]) <= float(_r["Phi_max"]) and float(_r["V_min"]) <= float(_r["V_max"])
    # the normalized terms are the recorded raw values scaled by the recorded ranges
    assert abs((float(_r["Phi"]) - float(_r["Phi_min"])) / (float(_r["Phi_max"]) - float(_r["Phi_min"]))
               - float(_r["Phi_tilde"])) < 1e-9
    assert abs((float(_r["V_group"]) - float(_r["V_min"])) / (float(_r["V_max"]) - float(_r["V_min"]))
               - float(_r["V_tilde"])) < 1e-9
    # distinct starts: 13 when the best sampled point is the center (dropped), else 14
    _nst = int(_r["n_starts"])
    assert _nst == int(_bdtc["n_random_starts"]) + 2 - int(_r["duplicate_start_dropped"])
    assert 0 <= int(_r["best_start"]) < _nst
    assert 0 <= int(_r["n_runs_capped"]) + int(_r["n_runs_other_exit"]) <= _nst
    # the retained solution (config "retention"): the best original start unless the
    # best further run ends above it by more than tie_tol, and then that run, which no
    # original start reaches; the agreement statistics are taken against it
    _src = str(_r["retained_source"])
    _fr_ = float(_r["further_best_minus_original_best"])
    _pr_ = float(_r["probe_best_minus_original_best"])
    assert _fr_ >= _pr_
    if _src == "original starts":
        assert _fr_ <= _tt and float(_r["objective"]) == float(_r["original_best_objective"])
        assert 1 <= int(_r["starts_at_best"]) <= _nst
        assert float(_r["runner_up_gap"]) == float(_r["original_runner_up_gap"])
    else:
        assert _src in ("second set", "follow-up") and _fr_ > _tt
        assert (_src == "second set") == (_pr_ == _fr_)
        assert abs(float(_r["objective"]) - float(_r["original_best_objective"]) - _fr_) < 1e-12
        assert int(_r["starts_at_best"]) == 0
        assert abs(float(_r["runner_up_gap"]) - _fr_) < 1e-12
    assert float(_r["runner_up_gap"]) >= 0.0 and float(_r["original_runner_up_gap"]) >= 0.0
    assert (int(_r["starts_at_best"]) >= 2) == (float(_r["runner_up_gap"]) < _tt)
    # the second set of starts, against the best original start
    assert int(_r["probe_n_starts"]) == int(_bdtc["probe_starts"])
    assert int(_r["probe_n_match_original_best"]) + int(_r["probe_n_exceed_original_best"]) \
        <= int(_r["probe_n_starts"])
    assert (int(_r["probe_n_exceed_original_best"]) > 0) == (_pr_ > _tt)
    # the near-optimal runs (original starts and second set, ending within near_tol below
    # the retained value): every original start at the retained value is one of them,
    # the retained run itself is one unless it is a follow-up run, and a program with
    # no near-optimal run besides the retained one has no weight spread
    _nn = int(_r["near_n_regular"]) + int(_r["near_n_probe"])
    assert int(_r["near_n_regular"]) >= int(_r["starts_at_best"]) and int(_r["near_n_probe"]) >= 0
    assert int(_r["near_n_other"]) == _nn - int(_src != "follow-up") and int(_r["near_n_other"]) >= 0
    if _src == "original starts":
        assert int(_r["near_n_probe"]) >= int(_r["probe_n_match_original_best"]) \
            + int(_r["probe_n_exceed_original_best"])
    assert float(_r["near_max_weight_diff"]) >= 0.0 and float(_r["near_max_mean_diff"]) >= 0.0
    assert int(_r["near_n_other"]) >= 1 or float(_r["near_max_weight_diff"]) == 0.0
    # the local-polish diagnostic (config polish): it only accepts improvements, so
    # its gain over the retained value is never negative
    assert float(_r["polish_gain"]) >= 0.0 and float(_r["polish_max_weight_diff"]) >= 0.0
    assert float(_r["polish_mean_diff"]) >= 0.0 and int(_r["polish_n_evaluations"]) >= 0
    assert str(_r["polish_end_point"]) in ("SLSQP", "pattern search")
    assert str(_r["polish_first_round_end_point"]) in ("SLSQP", "pattern search")
    # rounds only improve, and the first round is one of them
    assert float(_r["polish_first_round_gain"]) >= 0.0 and float(_r["polish_gain"]) >= float(_r["polish_first_round_gain"])
    assert 1 <= int(_r["polish_rounds"]) <= int(_bdtc["polish_max_rounds"])
# the solver record the text cites: runs, capped runs, runs ending otherwise
assert int(_bdtc["multistart_n_runs"]) == int(_bsol["n_starts"].sum()) \
    == (int(_bdtc["n_random_starts"]) + 2) * len(_bsol) - int(_bdtc["multistart_n_duplicate_start_dropped"])
assert int(_bdtc["multistart_n_duplicate_start_dropped"]) == int(_bsol["duplicate_start_dropped"].sum())
assert int(_bdtc["multistart_n_runs_capped"]) == int(_bsol["n_runs_capped"].sum())
assert int(_bdtc["multistart_n_runs_other_exit"]) == int(_bsol["n_runs_other_exit"].sum())
assert int(_bdtc["multistart_n_best_run_nonzero_status"]) == int((_bsol["best_run_status"] != 0).sum())
# the multi-start agreement statistics of Section 6.4 (against the retained value),
# recomputed from the rows
_bsab = _bsol["starts_at_best"].to_numpy(int)
_bgap = _bsol["runner_up_gap"].to_numpy(float)
_bsrc = _bsol["retained_source"].astype(str).to_numpy()
assert float(np.median(_bsab)) == float(_bdtc["multistart_median_starts_at_best"])
assert int((_bsab >= 2).sum()) == int(_bdtc["multistart_n_two_or_more_at_best"])
assert int((_bsab == 1).sum()) == int(_bdtc["multistart_n_one_at_best"])
assert int((_bsab == 0).sum()) == int(_bdtc["multistart_n_none_at_best"])
assert int(_bdtc["multistart_n_two_or_more_at_best"]) + int(_bdtc["multistart_n_one_at_best"]) \
    + int(_bdtc["multistart_n_none_at_best"]) == len(_bsol)
assert float(_bgap[_bsab == 1].max()) == float(_bdtc["multistart_max_gap_one_at_best"])
if (_bsab == 0).any():
    assert float(_bgap[_bsab == 0].max()) == float(_bdtc["multistart_max_gap_none_at_best"])
# the programs whose retained solution is a further run are exactly those that no
# original start reaches, and exactly those in which a further run ends above the
# best original start by more than tie_tol
_bfur = _bsol["further_best_minus_original_best"].to_numpy(float)
assert ((_bsrc != "original starts") == (_bsab == 0)).all()
assert ((_bsrc != "original starts") == (_bfur > _tt)).all()
assert int(_bdtc["retained_n_from_further_starts"]) == int((_bsrc != "original starts").sum()) \
    == int(_bdtc["further_n_programs_exceeded"])
assert float(_bdtc["further_max_raise"]) == float(_bfur.max())
_bfa = _bsol.loc[_bsol["further_best_minus_original_best"].idxmax()]
assert str(_bdtc["further_max_raise_argmax"]) == "delta %.1f, %s, %s" % (_bfa["delta"], _bfa["group"], _bfa["indicator"])
assert str(_bdtc["retained_from_further_programs"]) == (";".join(
    "delta %.1f, %s, %s (%s)" % (r["delta"], r["group"], r["indicator"], r["retained_source"])
    for _, r in _bsol[_bsrc != "original starts"].iterrows()) or "none")
# the second set of starts (against the best original start) and the stability of the
# returned weights over the near-optimal runs, recomputed from the rows
assert int(_bdtc["probe_n_runs"]) == int(_bsol["probe_n_starts"].sum()) == int(_bdtc["probe_starts"]) * len(_bsol)
assert int(_bdtc["probe_n_programs_exceeded"]) == int((_bsol["probe_n_exceed_original_best"] > 0).sum())
assert float(_bdtc["probe_max_raise"]) == float(_bsol["probe_best_minus_original_best"].max())
_bpa = _bsol.loc[_bsol["probe_best_minus_original_best"].idxmax()]
assert str(_bdtc["probe_max_raise_argmax"]) == "delta %.1f, %s, %s" % (_bpa["delta"], _bpa["group"], _bpa["indicator"])
assert int(_bdtc["probe_n_runs_capped"]) == int(_bsol["probe_n_runs_capped"].sum())
assert int(_bdtc["probe_n_runs_other_exit"]) == int(_bsol["probe_n_runs_other_exit"].sum())
assert int(_bdtc["near_n_programs_with_other_near_run"]) == int((_bsol["near_n_other"] >= 1).sum())
assert float(_bdtc["near_weight_diff_max"]) == float(_bsol["near_max_weight_diff"].max())
assert abs(float(_bdtc["near_weight_diff_median"]) - float(_bsol["near_max_weight_diff"].median())) \
    <= 1e-12 * float(_bsol["near_max_weight_diff"].median())
_bnam = _bsol["near_alt_max_weight"].to_numpy(float)
_bret_max = _bsol[_wcols].max(axis=1).to_numpy(float)
_bvl = _bsol[(_bret_max > 0.999) & (_bnam <= 0.999)]
assert int(_bdtc["swap_near_n_vertex_lost"]) == len(_bvl)
assert str(_bdtc["swap_near_vertex_lost_programs"]) == (";".join(
    "delta %.1f, %s, %s" % (r["delta"], r["group"], r["indicator"]) for _, r in _bvl.iterrows()) or "none")
if len(_bvl):
    assert abs(float(_bdtc["swap_near_vertex_lost_min_max_weight"]) - float(_bvl["near_alt_max_weight"].min())) < 1e-15
_bna = _bsol.loc[_bsol["near_max_weight_diff"].idxmax()]
assert str(_bdtc["near_weight_diff_argmax"]) == "delta %.1f, %s, %s" % (_bna["delta"], _bna["group"], _bna["indicator"])
assert int(_bdtc["near_n_programs_weight_diff_above_0.01"]) == int((_bsol["near_max_weight_diff"] > 0.01).sum())
assert float(_bdtc["near_mean_diff_max"]) == float(_bsol["near_max_mean_diff"].max())
# the substitution sheet: every printed quantity once, the flags consistent with the
# printed forms, and the counts in config
_bsw = _bdt["multistart_substitution"]
assert len(_bsw) == int(_bdtc["swap_n_printed_quantities"]) == _bsw["quantity"].nunique()
for _, _q in _bsw.iterrows():
    assert int(_q["unchanged_near"]) == int(str(_q["printed_near_substituted"]) == str(_q["printed_retained"]))
    assert int(_q["unchanged_original"]) == int(str(_q["printed_original_best"]) == str(_q["printed_retained"]))
    assert int(_q["unchanged_polished"]) == int(str(_q["printed_polished"]) == str(_q["printed_retained"]))
assert int(_bdtc["swap_near_n_printed_changed"]) == int((_bsw["unchanged_near"] == 0).sum())
assert int(_bdtc["swap_original_n_printed_changed"]) == int((_bsw["unchanged_original"] == 0).sum())
assert int(_bdtc["swap_polished_n_printed_changed"]) == int((_bsw["unchanged_polished"] == 0).sum())
for _s in ("near", "original", "polished"):
    assert str(_bdtc["swap_%s_changed_quantities" % _s]) == (";".join(
        str(q) for q in _bsw.loc[_bsw["unchanged_%s" % _s] == 0, "quantity"]) or "none")
# the retained column of the sheet is the printed form of the recorded outputs
_bswd = dict(zip(_bsw["quantity"].astype(str), _bsw["printed_retained"].astype(str)))
_bsen_ = _bdt["sensitivity"].set_index("delta")
assert _bswd["global maturity, delta 0.5"] == "%.2f" % float(_bsen_.loc[0.5, "global_maturity"])
assert _bswd["global maturity, delta 0"] == "%.2f" % float(_bsen_.loc[0.0, "global_maturity"])
assert _bswd["aggregate variance, delta 0.5"] == "%.3f" % float(_bsen_.loc[0.5, "aggregate_variance"])
assert _bswd["aggregate variance, delta 0"] == "%.3f" % float(_bsen_.loc[0.0, "aggregate_variance"])
assert _bswd["largest subgroup variance, delta 0.5"] == "%.3f" % float(_bsen_.loc[0.5, "max_group_variance"])
assert _bswd["collective fairness, delta 0.5"] == "%.3f" % float(_bsen_.loc[0.5, "fairness_Phi"])
assert _bswd["weight Gini, delta 0"] == "%.2f" % float(_bsen_.loc[0.0, "weight_gini"])
assert _bswd["single-member programs, delta 0"] == "%d" % int(_bsen_.loc[0.0, "n_vertex_solutions"])
assert _bswd["mass below level two, delta 0.5"] == "%.2f" % float(_bsen_.loc[0.5, "mass_below_2"])
# Section 6.4 prints "about twelve percent of the mass of F_G lying below level two"
# at the operating point; the config value is that sweep entry
assert "%.2f" % float(_bsen_.loc[float(_bdtc["operating_delta"]), "mass_below_2"]) == "0.12", \
    "P(F_G < 2) at the operating point no longer rounds to the printed twelve percent"
assert float(_bdtc["mass_below_2"]) == float(_bsen_.loc[float(_bdtc["operating_delta"]), "mass_below_2"])
for _b, _id in enumerate(_bids):
    assert _bswd["%s mean, delta 0.5" % _id] == "%.2f" % float(_bdim.loc[_b, "mu"])
for _k in range(3):
    assert _bswd["q1 G%d mean, delta 0.5" % (_k + 1)] == "%.2f" % float(_bgrp.loc[_k, "q1_mu"])
assert int(_bsen_["n_retained_from_further_starts"].sum()) == int(_bdtc["retained_n_from_further_starts"])
# the follow-up search covers exactly the programs whose runner-up original start falls
# short of the best by more than followup_gap, and its records are against the best
# original start
_bfg = float(_bdtc["followup_gap"])
_bfu = _bdt["multistart_followup"]
_bfr = _bdt["multistart_followup_runs"]
_bflag = _bsol[_bsol["original_runner_up_gap"] > _bfg]
assert len(_bfu) == len(_bflag) == int(_bdtc["multistart_n_followed_up"])
for (_, _fr), (_, _sr) in zip(_bfu.iterrows(), _bflag.iterrows()):
    assert (_fr["delta"], _fr["group"], _fr["indicator"]) == (_sr["delta"], _sr["group"], _sr["indicator"])
    assert float(_fr["original_best_objective"]) == float(_sr["original_best_objective"])
    assert float(_fr["original_runner_up_gap"]) == float(_sr["original_runner_up_gap"])
    _runs = _bfr[(_bfr["delta"] == _fr["delta"]) & (_bfr["group"] == _fr["group"])
                 & (_bfr["indicator"] == _fr["indicator"])]
    assert len(_runs) == int(_fr["n_followup_starts"]) == int(_bdtc["followup_starts"])
    # runs that return the best original value and runs that end above it, counted apart
    _mr = _runs["minus_original_best"].to_numpy(float)
    assert int((np.abs(_mr) < _tt).sum()) == int(_fr["n_match_original_best"])
    assert int((_mr > _tt).sum()) == int(_fr["n_exceed_original_best"])
    assert abs(float(_mr.max()) - float(_fr["best_minus_original_best"])) < 1e-15
    if int(_fr["n_exceed_original_best"]) > 0:
        assert abs(float(_mr[_mr > _tt].min()) - float(_fr["exceed_min"])) < 1e-15
        assert abs(float(_runs.loc[_mr > _tt, "max_abs_weight_diff_vs_original_best"].max())
                   - float(_fr["exceed_max_weight_diff"])) < 1e-15
    _bmx = _runs["max_abs_weight_diff_vs_original_best"].to_numpy(float)
    assert abs(float(_bmx[int(np.argmax(_mr))]) - float(_fr["max_abs_weight_diff_best_vs_original_best"])) < 1e-15
    # the best follow-up run is the retained solution exactly when it is the best
    # further run and ends above the best original start by more than tie_tol
    if str(_sr["retained_source"]) == "follow-up":
        assert float(_sr["objective"]) == float(_fr["best_followup_objective"])
        assert float(_sr["further_best_minus_original_best"]) == float(_fr["best_minus_original_best"]) > _tt
    else:
        assert float(_fr["best_minus_original_best"]) <= float(_sr["further_best_minus_original_best"])
# the normalized terms at the retained solutions (the ranges are sampled)
assert int((_bsol["Phi_tilde"] > 1.0).sum()) == int(_bdtc["retained_n_Phi_tilde_above_1"])
assert float(_bsol["Phi_tilde"].max()) == float(_bdtc["retained_max_Phi_tilde"])
assert float(_bsol["V_tilde"].min()) == float(_bdtc["retained_min_V_tilde"])
assert float(_bsol["V_tilde"].max()) == float(_bdtc["retained_max_V_tilde"])
# quantile-averaging mean identity at the inter-group, dimension and global layers
_bind = _bdt["indicator_collectives"]
_bdimc = _bdt["dimension_collectives"]
_bsen = _bdt["sensitivity"].set_index("delta")
for _d in sorted(_bsen.index):
    _ii = _bind[_bind["delta"] == _d].sort_values("indicator", key=lambda s: s.str[1:].astype(int))
    _pred = _bOmega[0] * _ii["G1_mean"].to_numpy() + _bOmega[1] * _ii["G2_mean"].to_numpy() \
        + _bOmega[2] * _ii["G3_mean"].to_numpy()
    assert np.allclose(_pred, _ii["collective_mean"].to_numpy(), atol=1e-9), \
        "BDT inter-group mean identity fails at delta %.1f" % _d
    _dd = _bdimc[_bdimc["delta"] == _d].sort_values("dimension")
    for _b, (_, _rr) in enumerate(_dd.iterrows()):
        _mm = _ii[_ii["dimension"] == _rr["dimension"]]["collective_mean"].to_numpy()
        assert abs(float(_mm.mean()) - float(_rr["mean"])) < 1e-9, "BDT dimension mean identity fails"
    _gm = float(np.dot(_dd["weight"], _dd["mean"]))
    assert abs(_gm - float(_bsen.loc[_d, "global_maturity"])) < 1e-9, \
        "BDT global maturity at delta %.1f is not the weighted mean of the dimension means" % _d
    # prop:hier-variance: aggregate variance <= largest subgroup variance, and at the
    # indicator layer each collective's variance <= the largest of its group variances
    assert float(_bsen.loc[_d, "aggregate_variance"]) <= float(_bsen.loc[_d, "max_group_variance"]) + 1e-12
    _sd = _bsol[_bsol["delta"] == _d]
    for _, _rr in _ii.iterrows():
        _vg = _sd[_sd["indicator"] == _rr["indicator"]]["V_group_exact"].to_numpy(float)
        assert float(_rr["collective_variance"]) <= _vg.max() + 1e-12, "prop:hier-variance fails at the indicator layer"
    assert abs(float(_sd["V_group_exact"].max()) - float(_bsen.loc[_d, "max_group_variance"])) < 1e-12
    assert abs(float(_sd["V_group_exact"].mean()) - float(_bsen.loc[_d, "mean_group_variance"])) < 1e-12
    # the solution-grid means (printed) agree with the moment-grid means to well below
    # the printed precision, and the 257-level variance understates the exact one
    assert float(_bsen.loc[_d, "max_mean_error_L257"]) < 1e-3
    assert float(_bsen.loc[_d, "aggregate_variance_L257"]) < float(_bsen.loc[_d, "aggregate_variance"])
    assert abs(float(_sd["Phi"].mean()) - float(_bsen.loc[_d, "fairness_Phi"])) < 1e-9
    assert abs(float(_sd["gini_w"].mean()) - float(_bsen.loc[_d, "weight_gini"])) < 1e-9
# the dimensions sheet is the delta = 0.5 row of dimension_collectives
_d5 = _bdimc[_bdimc["delta"] == 0.5].sort_values("dimension")
assert np.allclose(_d5["mean"].to_numpy(), _bdim["mu"].to_numpy(float))
assert abs(float(_bdtc["global_maturity"]) - float(_bsen.loc[0.5, "global_maturity"])) < 1e-12
assert abs(float(_bdtc["aggregate_variance"]) - float(_bsen.loc[0.5, "aggregate_variance"])) < 1e-15
# the printed variances (Section 6.4: the aggregate variance at delta = 0 and 0.5, the
# largest subgroup variance at 0.5) recomputed by adaptive quadrature from the recorded
# weights and Weibull parameters alone, an evaluation independent of the moment grid
from scipy.integrate import quad as _sp_quad
_bkap = {(r["expert"], r["indicator"]): (float(r["kappa"]), float(r["lambda"])) for _, r in _bpan.iterrows()}
_bnb = dict(zip(_bdim["id"].astype(str), _bdim["n_indicators"].astype(int)))
_bdwd = dict(zip(_bdim["id"].astype(str), _bw))


def _bdt_quad_var(terms):
    """variance of the quantile function sum_t c_t Q_t(p) by adaptive quadrature."""
    ks = np.array([t[1] for t in terms]); ls = np.array([t[2] for t in terms])
    cs = np.array([t[0] for t in terms]); h5 = 1.0 - np.exp(-(5.0 / ls) ** ks)

    def q(p):
        return float(cs @ (ls * (-np.log1p(-p * h5)) ** (1.0 / ks)))
    opts = dict(limit=400, epsabs=1e-13, epsrel=1e-13)
    m1 = _sp_quad(q, 0.0, 1.0, **opts)[0]
    m2 = _sp_quad(lambda p: q(p) ** 2, 0.0, 1.0, **opts)[0]
    return m2 - m1 * m1


def _bdt_row_terms(r, scale):
    return [(scale * float(r["w_e%d" % i]),) + _bkap[("e%d" % i, r["indicator"])]
            for i in range(1, 16) if not np.isnan(float(r["w_e%d" % i]))]


for _d in (0.0, 0.5):
    _terms = []
    for _, _r in _bsol[_bsol["delta"] == _d].iterrows():
        _sc = _bdwd[_r["dimension"]] / _bnb[_r["dimension"]] * _bOmega[int(_r["group"][1]) - 1]
        _terms += _bdt_row_terms(_r, _sc)
    _vq = _bdt_quad_var(_terms)
    assert abs(_vq - float(_bsen.loc[_d, "aggregate_variance"])) < 1e-9, \
        "aggregate variance at delta %.1f: quadrature %.10f vs recorded %.10f" % (
            _d, _vq, float(_bsen.loc[_d, "aggregate_variance"]))
_r5 = _bsol[_bsol["delta"] == 0.5]
_r5 = _r5.loc[_r5["V_group_exact"].idxmax()]
_vq = _bdt_quad_var(_bdt_row_terms(_r5, 1.0))
assert abs(_vq - float(_bsen.loc[0.5, "max_group_variance"])) < 1e-9, \
    "largest subgroup variance: quadrature %.10f vs recorded %.10f" % (_vq, float(_bsen.loc[0.5, "max_group_variance"]))
# that value lies within 1e-5 of a three-decimal rounding boundary, so its printed
# rounding is settled only because both evaluations agree far more closely than that
_v3 = float(_bsen.loc[0.5, "max_group_variance"]) * 1e3
assert abs((_v3 - np.floor(_v3)) - 0.5) * 1e-3 > 1e-8
# P(F_G < 2) at every delta, recomputed from the recorded weights and Weibull
# parameters alone by solving Q_G(p) = 2 for p, an evaluation independent of the
# moment grid on which the sensitivity sheet inverts Q_G
from scipy.optimize import brentq as _sp_brentq
for _d in sorted(_bsen.index):
    _terms = []
    for _, _r in _bsol[_bsol["delta"] == _d].iterrows():
        _sc = _bdwd[_r["dimension"]] / _bnb[_r["dimension"]] * _bOmega[int(_r["group"][1]) - 1]
        _terms += _bdt_row_terms(_r, _sc)
    _mks = np.array([t[1] for t in _terms]); _mls = np.array([t[2] for t in _terms])
    _mcs = np.array([t[0] for t in _terms]); _mh5 = 1.0 - np.exp(-(5.0 / _mls) ** _mks)
    _pq = _sp_brentq(lambda p: float(_mcs @ (_mls * (-np.log1p(-p * _mh5)) ** (1.0 / _mks))) - 2.0,
                     0.0, 1.0 - 1e-12, xtol=1e-14, rtol=1e-14)
    assert abs(_pq - float(_bsen.loc[_d, "mass_below_2"])) < 1e-6, \
        "P(F_G < 2) at delta %.1f: root %.8f vs recorded %.8f" % (_d, _pq, float(_bsen.loc[_d, "mass_below_2"]))
# Section 6.4 multi-start claims, each held to the bound the text prints.  The further
# starts end more than tie_tol above the best original start in the programs whose
# retained solution is a further run, by at most 1.2e-5, the accuracy the text states
# (of the order of 1e-5).  Of the programs without a second original start at the
# retained value, those with one fall short by at most 3e-4, those with none by at
# most that 1.2e-5.
assert 1e-6 < float(_bdtc["further_max_raise"]) <= 1.2e-5, "further starts beyond the printed 1.2e-5"
# the local-polish diagnostic, recomputed from the rows, and the accuracy it supports:
# no retained value lies more than 5e-5 below its settled polished point and no weight
# moves by more than 2e-3, so the retained values are accurate to the order of 1e-5;
# the polish changes no retained solution and, substituted everywhere at once, no
# printed number
_bpg = _bsol["polish_gain"].to_numpy(float)
assert float(_bdtc["polish_max_gain"]) == float(_bpg.max())
assert int(_bdtc["polish_n_programs_gain_above_tie_tol"]) == int((_bpg > _tt).sum())
assert int(_bdtc["polish_n_programs_gain_above_further_max_raise"]) == \
    int((_bpg > float(_bdtc["further_max_raise"])).sum())
assert int(_bdtc["polish_n_programs_at_round_cap"]) == \
    int((_bsol["polish_rounds"] >= int(_bdtc["polish_max_rounds"])).sum())
assert int(_bdtc["polish_max_rounds_used"]) == int(_bsol["polish_rounds"].max())
assert float(_bdtc["polish_first_round_max_gain"]) == float(_bsol["polish_first_round_gain"].max()) \
    <= float(_bdtc["polish_max_gain"])
assert float(_bdtc["polish_first_round_max_weight_diff"]) == float(_bsol["polish_first_round_max_weight_diff"].max())
_pa1 = _bsol.loc[_bsol["polish_first_round_gain"].idxmax()]
assert str(_bdtc["polish_first_round_max_gain_argmax"]) == "delta %.1f, %s, %s" % (_pa1["delta"], _pa1["group"], _pa1["indicator"])
assert float(_bdtc["polish_max_weight_diff"]) == float(_bsol["polish_max_weight_diff"].max())
assert float(_bdtc["polish_max_mean_diff"]) == float(_bsol["polish_mean_diff"].max())
for _nm, _col in (("polish_max_gain", "polish_gain"), ("polish_max_weight_diff", "polish_max_weight_diff"),
                  ("polish_max_mean_diff", "polish_mean_diff")):
    _pa = _bsol.loc[_bsol[_col].idxmax()]
    assert str(_bdtc[_nm + "_argmax"]) == "delta %.1f, %s, %s" % (_pa["delta"], _pa["group"], _pa["indicator"])
assert abs(float(_bdtc["retained_accuracy_bound"])
           - max(float(_bpg.max()), float(_bdtc["further_max_raise"]))) < 1e-18
assert 1e-6 < float(_bdtc["polish_max_gain"]) <= 5e-5, "local polish beyond the printed 5e-5"
assert float(_bdtc["polish_max_weight_diff"]) <= 2e-3, "polished weights beyond the printed 2e-3"
assert int(_bdtc["polish_n_programs_at_round_cap"]) == 0, "a local polish did not settle within its rounds"
# the group-opinion means move far less than the half unit (0.005) of their printed two decimals
assert float(_bdtc["polish_max_mean_diff"]) < 5e-4
assert 1e-6 < float(_bdtc["retained_accuracy_bound"]) < 1e-4, "retained values no longer accurate to about 1e-5"
assert int(_bdtc["swap_polished_n_printed_changed"]) == 0, "a printed number changes at the polished weights"
assert "diagnostic, not a retention step" in str(_bdtc["retention"])
assert float(_bdtc["polish_step0"]) == _BDT["polish_step0"] and float(_bdtc["polish_step_min"]) == _BDT["polish_step_min"]
assert int(_bdtc["polish_max_evals"]) == _BDT["polish_max_evals"]
# the seed of the second set of starts, which Sections 6.1 and 6.4 print, is recorded
# and is the seed the rng order names
assert int(float(_bdtc["probe_seed"])) == int(float(_bdtc["seed"])) + int(float(_bdtc["probe_seed_offset"])) \
    == _BDT["seed"] + _BDT["probe_seed_offset"]
assert int(float(_bdtc["probe_seed"])) != int(float(_bdtc["seed"]))
assert ("probe_seed = %d" % int(float(_bdtc["probe_seed"]))) in str(_bdtc["rng_order"])
# the k-means settings behind the strata are recorded
assert int(_bdtc["kmeans_maxiter"]) == _BDT["kmeans_maxiter"] and int(_bdtc["kmeans_restarts"]) == _BDT["kmeans_restarts"]
assert ("at most %d Lloyd iterations" % _BDT["kmeans_maxiter"]) in str(_bdtc["kmeans"])
assert float(_bdtc["probe_max_raise"]) <= float(_bdtc["further_max_raise"])
assert float(_bdtc["multistart_max_gap_one_at_best"]) <= 3e-4, "shortfall beyond the printed 3e-4"
assert float(_bdtc["multistart_max_gap_none_at_best"]) <= 1.2e-5
# the text names the one program followed up, which is among the programs whose
# retained solution is a further run, and there it is the best follow-up run
assert [(float(_fr["delta"]), str(_fr["group"]), str(_fr["indicator"])) for _, _fr in _bfu.iterrows()] \
    == [(0.1, "G3", "q9")], "the followed-up programs differ from the one Section 6.4 names"
_bq9 = _bsol[(_bsol["delta"] == 0.1) & (_bsol["group"] == "G3") & (_bsol["indicator"] == "q9")].iloc[0]
assert str(_bq9["retained_source"]) == "follow-up", "the q9 program does not retain its best follow-up run"
for _, _fr in _bfu.iterrows():
    # no further run returns the value of the best original start and some end above
    # it, all in its basin (their weights at least ten times closer to its weights than
    # the runner-up start's) and within the printed 1.6e-3 of its weights
    assert int(_fr["n_match_original_best"]) == 0 and int(_fr["n_exceed_original_best"]) >= 1
    assert float(_fr["exceed_max_weight_diff"]) < 0.1 * float(_fr["runner_up_weight_diff"])
    assert float(_fr["exceed_max_weight_diff"]) <= 1.6e-3, "follow-up weights beyond the printed 1.6e-3"
    assert float(_fr["best_minus_original_best"]) <= float(_bdtc["further_max_raise"])
# the near-optimal weights, held to the printed bounds 0.072 and 0.011
assert float(_bdtc["near_weight_diff_max"]) <= 0.072, "near-optimal weights beyond the printed 0.072"
assert float(_bdtc["near_mean_diff_max"]) <= 0.011, "group-opinion means beyond the printed 0.011"
# the substitutions: the best original starts change no printed number; the
# near-optimal runs change only the count of single-member programs at delta = 0, and
# only by the vertex solutions they fail to reproduce (the text states this exception)
assert int(_bdtc["swap_original_n_printed_changed"]) == 0, "a printed number changes under the best original starts"
assert str(_bdtc["swap_near_changed_quantities"]) == "single-member programs, delta 0", \
    "the printed numbers that change under the near-optimal weights differ from the text"
_bsq = _bsw.set_index("quantity").loc["single-member programs, delta 0"]
assert int(_bsq["printed_retained"]) - int(_bsq["printed_near_substituted"]) == int(_bdtc["swap_near_n_vertex_lost"])
assert all(s.startswith("delta 0.0,") for s in str(_bdtc["swap_near_vertex_lost_programs"]).split(";"))
# every run of the second set ended on SLSQP's convergence test or at the cap
assert int(_bdtc["probe_n_runs_other_exit"]) == 0
# every retained run ended on SLSQP's convergence test, not at the iteration cap
assert (_bsol["best_run_status"] == 0).all()
assert int(_bdtc["multistart_n_runs_other_exit"]) == 0
# the normalized confidence term stays in the unit interval at every retained solution
assert -1e-12 <= float(_bdtc["retained_min_V_tilde"]) and float(_bdtc["retained_max_V_tilde"]) <= 1.0
# Figure 6.8: densities integrate to one and their means equal the legend means
_bq = _bdt["q1_densities"]
_bx = _bq["x"].to_numpy(float)
_bdx = float(_bx[1] - _bx[0])
for _k in range(3):
    _f = _bq["f_G%d" % (_k + 1)].to_numpy(float)
    assert abs(_trapz(_f, dx=_bdx) - 1.0) < 0.01
    assert abs(_trapz(_bx * _f, dx=_bdx) - float(_bgrp.loc[_k, "q1_mu"])) < 0.01
_fc = _bq["f_collective"].to_numpy(float)
assert abs(_trapz(_fc, dx=_bdx) - 1.0) < 0.01
assert abs(_trapz(_bx * _fc, dx=_bdx) - float(np.dot(_bOmega, _bgrp["q1_mu"]))) < 0.01
_bqq = _bdt["q1_quantiles"]
assert np.allclose(_bqq["Q_collective"], _bOmega[0] * _bqq["Q_G1"] + _bOmega[1] * _bqq["Q_G2"]
                   + _bOmega[2] * _bqq["Q_G3"])
for _c in ("Q_G1", "Q_G2", "Q_G3", "Q_collective"):
    assert (np.diff(_bqq[_c].to_numpy(float)) > 0).all(), "BDT quantile function %s not increasing" % _c
print("   BDT case: simplex/identity/variance-bound invariants hold at every delta;"
      " P(F_G < 2) recomputed by root finding at every delta (printed 0.12 at the operating point);"
      " k-means, TOPSIS, BWM and the Figure 6.8 densities verified")

# (6) vignette V: the solved outputs written to the workbook satisfy the
#     invariants the book states and the claims Section 6.6 makes, re-read
#     from the workbook (nothing here is compared with a typed printed
#     constant): the record regenerates from the recorded seed and draw
#     order, panel dimensions, weight vectors on the simplex, W_H + W_M = 1,
#     the leave-one-out identities (exact and step score, delta = d_cal +
#     d_res), the Hersbach identities, the spread rules, the issue-time
#     information structure, the solver objective is the exact-cdf one,
#     every solve certified, the in-sample and new-quarter scores recomputed
#     from the stored weights, the design-level solution and its closed-form
#     scores, the claims C1-C4 in the draw and at the design level, the
#     single corrections of the resolution and the counterfactual channels of
#     M4 recomputed at the stored weights, the weight of M4 re-solved at its
#     error (draw, design, re-seeds), the resolution order on every grid
#     of sheet resolution_by_L, and the re-seeding frequencies and percentiles
#     recomputed from the stored rows (printed values typical, plotted values
#     typical or disclosed).
_hmc = _wb("fig6-12_hm_calsharp.xlsx")
_hmd = _cfgd(_hmc)
_hmF = _hmc["forecasters"]
_hmN = int(_hmd["n_human"]) + int(_hmd["n_machine"])
_hmT = int(_hmd["horizon_T"])
assert len(_hmF) == _hmN and _hmc["forecast_mu"].shape == (_hmT, _hmN)
assert len(_hmc["outcomes"]) == _hmT
_hmids = list(_hmF["id"].astype(str))
_hmH = [i for i, k in enumerate(_hmids) if k.startswith("H")]
_hmMc = [i for i, k in enumerate(_hmids) if k.startswith("M")]
_hmCal = [i for i, g in enumerate(_hmF["group"].astype(str)) if g != "overconfident machines"]
_hm4, _hm3 = _hmids.index("M4"), _hmids.index("M3")
# (6a) the record regenerates from the recorded seed, draw order and design constants
assert int(_hmd["seed"]) == NEW_SEED and _hmT == _VG_T
assert np.allclose(_hmF["bias"], _VG_BIAS, atol=0) and np.allclose(_hmF["error_scale_tau"], _VG_TAU, atol=0)
assert abs(float(_hmd["target_sigma_u"]) - _VG_SIGNAL["sigma_u"]) < 1e-15
_hmy0, _hmMU0, _, _, _, _hmu0, _ = _vg_generate(int(_hmd["seed"]))
assert float(np.abs(_hmy0 - _hmc["outcomes"]["y"].to_numpy(float)).max()) < 1e-12, "vignette record does not regenerate"
assert float(np.abs(_hmMU0 - _hmc["forecast_mu"].to_numpy(float)).max()) < 1e-12
assert float(np.abs(_hmu0 - _hmc["outcomes"]["unobserved_residual_u"].to_numpy(float)).max()) < 1e-12
# (6b) weights
_hmW = _hmc["weights"]
for _col in ("weight_full", "weight_human_only", "weight_design", "weight_design_human_only",
             "weight_design_machine_only", "weight_design_M4_spread_at_error", "weight_full_M4_spread_at_error"):
    _sums_to_one(_hmW[_col], "vignette %s" % _col, tol=1e-6)
    assert (_hmW[_col] >= -1e-12).all()
_sums_to_one(_hmW["weight_equal"], "vignette equal weights", tol=1e-9)
for _col in ("weight_human_only", "weight_design_human_only"):
    assert (_hmW.loc[_hmW["id"].astype(str).str.startswith("M"), _col] == 0).all()
assert (_hmW.loc[_hmW["id"].astype(str).str.startswith("H"), "weight_design_machine_only"] == 0).all()
for _sheet in ("loo_weights", "design_loo_weights"):
    for _k, _row in _hmc[_sheet].iterrows():
        _wv = _row.drop("withheld").astype(float)
        assert np.isnan(_wv[_row["withheld"]]), "%s row %s keeps the withheld expert" % (_sheet, _row["withheld"])
        _sums_to_one(_wv.dropna(), "vignette %s without %s" % (_sheet, _row["withheld"]), tol=1e-6)
_hmA = _hmc["aggregates"].set_index("quantity")["value"]
assert abs(float(_hmA["human weight share W_H"]) + float(_hmA["machine weight share W_M"]) - 1.0) < 1e-9
assert abs(float(_hmA["human weight share W_H"]) - float(_hmW["weight_full"].to_numpy(float)[_hmH].sum())) < 1e-9
assert abs(float(_hmA["design-level machine weight share W_M"]) - float(_hmW["weight_design"].to_numpy(float)[_hmMc].sum())) < 1e-9
# (6c) leave-one-out and Hersbach identities, per forecaster
_hm_full = float(_hmA["optimized human-machine mean CRPS"])
_hm_ch_full = float(_hmA["full-pool Hersbach CRPS"])
_hm_unc = float(_hmA["uncertainty UNC (sample climatology)"])
assert abs(float(_hmA["full-pool Hersbach REL"]) + float(_hmA["full-pool Hersbach POT"]) - _hm_ch_full) < 1e-9
assert abs(float(_hmd["uncertainty_UNC"]) - _hm_unc) < 1e-9
_hm_y = _hmc["outcomes"]["y"].to_numpy(float)
assert abs(_hm_unc - 0.5 * float(np.abs(_hm_y[:, None] - _hm_y[None, :]).mean())) < 1e-12, "UNC is not 0.5 mean|y - y'|"
_hmsu = float(_hmd["target_sigma_u"])
for _k, _r in _hmF.iterrows():
    assert abs(float(_r["delta_loo"]) - (float(_r["crps_loo_resolve"]) - _hm_full)) < 1e-9
    assert abs(float(_r["delta_cal"]) + float(_r["delta_res"]) - float(_r["delta_loo_hersbach"])) < 1e-12
    assert abs(float(_r["delta_cal_design"]) + float(_r["delta_res_design"]) - float(_r["delta_design"])) < 1e-12
    assert abs(float(_r["REL_i"]) + float(_r["POT_i"]) - float(_r["crps_i_hersbach"])) < 1e-9
    assert abs(float(_r["resolution"]) - (_hm_unc - float(_r["POT_i"]))) < 1e-9
    assert abs(float(_r["reliability_error"]) - float(_r["REL_i"])) < 1e-12
    assert int(_r["count_below_median"]) == int(round(float(_r["frac_below_median"]) * _hmT))
    assert abs(float(_r["below_median_binomial_p"]) - _vg_binom_p(int(_r["count_below_median"]), _hmT)) < 1e-15
    _rmse = math.sqrt(_hmsu ** 2 + float(_r["bias"]) ** 2 + float(_r["error_scale_tau"]) ** 2)
    assert abs(float(_r["design_rmse_of_mean"]) - _rmse) < 1e-12
    assert abs(float(_r["p_below_median_design"]) - float(_ndtr(float(_r["bias"]) / math.sqrt(_hmsu ** 2 + float(_r["error_scale_tau"]) ** 2)))) < 1e-12
    if str(_r["group"]) == "overconfident machines":
        assert float(_r["spread"]) < _rmse, "%s is not overconfident" % _r["id"]
    else:
        assert abs(float(_r["spread"]) - _rmse) < 1e-12, "%s spread not commensurate" % _r["id"]
# (6c') the vignette forms RES_i against the continuous UNC, whereas Section 4 and
#       ex:hmpanel form UNC on the same L-member grid.  The text says the departure
#       raises the resolution of every member by one common amount, so that no
#       comparison between members depends on it: the grid UNC is recomputed from the
#       stored outcomes; the shift is positive and the same for every member; and the
#       order of the members' resolutions and membership of the well-calibrated band
#       (REL_i below band x UNC) are the same under either UNC
_hm_Lg = int(_hmd["hersbach_levels_L"])
_hm_unc_grid = float(_hersbach_X(np.tile(np.quantile(_hm_y, np.arange(1, _hm_Lg + 1) / (_hm_Lg + 1.0)),
                                         (len(_hm_y), 1)), _hm_y)[2])
assert abs(float(_hmd["uncertainty_UNC_grid"]) - _hm_unc_grid) < 1e-12, "uncertainty_UNC_grid does not recompute"
_hm_shift = _hm_unc - _hm_unc_grid
assert _hm_shift > 0, "the continuous UNC no longer exceeds the grid UNC"
_hm_res_c = _hmF["resolution"].to_numpy(float)
_hm_res_g = _hm_unc_grid - _hmF["POT_i"].to_numpy(float)
assert np.allclose(_hm_res_c - _hm_res_g, _hm_shift, rtol=0, atol=1e-12), "the UNC shift is not common to every RES_i"
assert list(np.argsort(-_hm_res_c, kind="stable")) == list(np.argsort(-_hm_res_g, kind="stable"))
_hm_rel_i = _hmF["reliability_error"].to_numpy(float)
assert abs(float(_hmd["well_calibrated_cut"]) - _VG_BAND * _hm_unc) < 1e-12
assert ((_hm_rel_i < _VG_BAND * _hm_unc) == (_hm_rel_i < _VG_BAND * _hm_unc_grid)).all(), \
    "membership of the well-calibrated band depends on the UNC convention"
# issue-time information only: the target obeys y_t = m_t + x_t + u_t with m_t the
# AR(1) conditional mean given y_{t-1}, the innovation split preserves the AR(1) law,
# and forecast i minus its bias is m_t + x_t + tau_i xi_it, so the residual u_t (and
# hence the realized y_t) never enters a forecast
_hmO = _hmc["outcomes"]
_hm_c, _hm_phi = float(_hmd["target_c"]), float(_hmd["target_phi"])
assert abs(float(_hmd["target_sigma_x"]) ** 2 + float(_hmd["target_sigma_u"]) ** 2
           - float(_hmd["target_sigma"]) ** 2) < 1e-12
_hm_prev = np.concatenate([[float(_hmd["presample_y0"])], _hmO["y"].to_numpy(float)[:-1]])
assert np.allclose(_hmO["conditional_mean_given_previous"], _hm_c + _hm_phi * (_hm_prev - _hm_c), atol=1e-12)
assert np.allclose(_hmO["y"], _hmO["conditional_mean_given_previous"] + _hmO["leading_signal_x"]
                   + _hmO["unobserved_residual_u"], atol=1e-12)
# (6d) the in-sample scores of the stored weights, their expected scores on new
#      quarters (closed form from the stored design), and the design level
_hm_mu = _hmc["forecast_mu"].to_numpy(float)
_hm_sd = _hmF["spread"].to_numpy(float)
_hmb, _hmtau = _hmF["bias"].to_numpy(float), _hmF["error_scale_tau"].to_numpy(float)


def _hm_ecrps(w, sd=None):
    """E CRPS on a new quarter, from the stored design (independent restatement)."""
    sd = _hm_sd if sd is None else sd
    S, m = float(w @ sd), float(w @ _hmb)
    R = math.sqrt(S * S + _hmsu ** 2 + float((w * w) @ (_hmtau ** 2)))
    return R * math.sqrt(2.0 / math.pi) * math.exp(-m * m / (2 * R * R)) + m * (2 * float(_ndtr(m / R)) - 1) - S / math.sqrt(math.pi)


_hmS = {}
for _col, _key, _nkey in (("weight_full", "optimized human-machine mean CRPS", "optimized human-machine expected CRPS on new quarters"),
                          ("weight_human_only", "optimized human-only mean CRPS", "optimized human-only expected CRPS on new quarters"),
                          ("weight_equal", "equal-weight human-machine mean CRPS", "equal-weight human-machine expected CRPS on new quarters")):
    _wv = _hmW[_col].to_numpy(float)
    _c = float(_crps_gauss(_hm_mu @ _wv, float(_hm_sd @ _wv), _hm_y).mean())
    assert abs(_c - float(_hmA[_key])) < 1e-9, "vignette %s CRPS %.6f != stored %.6f" % (_col, _c, float(_hmA[_key]))
    assert abs(_hm_ecrps(_wv) - float(_hmA[_nkey])) < 1e-12, "vignette %s new-quarter score" % _col
    _hmS[_col] = (_c, _hm_ecrps(_wv))
for _col, _key in (("weight_design", "design-level optimized human-machine expected CRPS"),
                   ("weight_design_human_only", "design-level optimized human-only expected CRPS"),
                   ("weight_design_machine_only", "design-level optimized machine-only expected CRPS"),
                   ("weight_equal", "design-level equal-weight expected CRPS")):
    assert abs(_hm_ecrps(_hmW[_col].to_numpy(float)) - float(_hmA[_key])) < 1e-12
_hm_cf_sd = _hm_sd.copy()
_hm_cf_sd[_hm4] = float(_hmF.loc[_hm4, "design_rmse_of_mean"])
assert abs(float(_hmW.loc[_hm4, "weight_design_M4_spread_at_error"])
           - float(_hmA["design-level weight of M4 with its spread set to its error"])) < 1e-12
# the design-level leave-one-out pools (exact expected score) reproduce from the stored weights
for _k, _row in _hmc["design_loo_weights"].iterrows():
    _i = _hmids.index(str(_row["withheld"]))
    _wv = np.nan_to_num(_row.drop("withheld").astype(float).to_numpy())
    assert abs(_hm_ecrps(_wv) - float(_hmF.loc[_i, "expected_CRPS_loo_design"])) < 1e-12
# (6e) the solver objective is the exact-cdf objective the cross-check evaluates,
#      every solve is certified, and the full-pool optimum is no worse than its start
_hmSv = _hmc["solves"]
assert len(_hmSv) == 3 + _hmN
_wf = _hmW["weight_full"].to_numpy(float)
assert abs(_vg_objective(_wf, _hm_mu, _hm_sd, _hm_y, float(_hmd["lambda_G"]), float(_hmd["eps_ridge"]))
           - float(_hmSv.loc[_hmSv["solve"] == "full panel", "objective_best_iterate"].iloc[0])) < 1e-12, \
    "the stored objective of the full solve is not the exact-cdf objective"
_wcf = _hmW["weight_full_M4_spread_at_error"].to_numpy(float)
assert abs(_vg_objective(_wcf, _hm_mu, _hm_cf_sd, _hm_y, float(_hmd["lambda_G"]), float(_hmd["eps_ridge"]))
           - float(_hmSv.loc[_hmSv["solve"] == "full panel, M4 spread set to its error", "objective_best_iterate"].iloc[0])) < 1e-12, \
    "the stored objective of the draw counterfactual is not the exact-cdf objective with M4's spread at its error"
assert abs(float(_wcf[_hm4]) - float(_hmA["weight of M4 with its spread set to its error (draw)"])) < 1e-12
assert (_hmSv["objective_gap"].abs() < float(_hmd["certificate_gap_bound"])).all(), "a vignette solve is not certified"
assert (_hmSv["tolerance_met"] == 1).all(), "a vignette solve hit the iteration guard"
assert (_hmc["design_solves"]["objective_spread_over_starts"] < float(_hmd["certificate_gap_bound"])).all()
_hm_eq_obj = float(_hmA["equal-weight human-machine mean CRPS"]) + 0.5 * float(_hmd["eps_ridge"]) / _hmN
assert float(_hmSv.loc[_hmSv["solve"] == "full panel", "objective_best_iterate"].iloc[0]) <= _hm_eq_obj + 1e-12
# (6f) the claims of Section 6.6 in the reported draw and at the design level
_c_in, _c_new = _hmS["weight_full"]
_h_in, _h_new = _hmS["weight_human_only"]
_e_in, _e_new = _hmS["weight_equal"]
assert _c_in < _h_in and _c_in < _e_in, "C1 in-sample fails"
assert _c_new < _h_new and _c_new < _e_new, "C1 on new quarters fails"
assert (float(_hmA["design-level optimized human-machine expected CRPS"]) < float(_hmA["design-level optimized human-only expected CRPS"])
        and float(_hmA["design-level optimized human-machine expected CRPS"]) < float(_hmA["design-level equal-weight expected CRPS"]))
# the machine-only pool at the design level lies between the equal-weight pool and the optimized
# human-machine pool (the sentence of Section 6.6.4 that places it there)
assert (float(_hmA["design-level optimized human-machine expected CRPS"]) < float(_hmA["design-level optimized machine-only expected CRPS"])
        < float(_hmA["design-level equal-weight expected CRPS"])), "the design-level machine-only pool is not between the two pools"
# ... and it prints strictly between them at the three decimals of the text
assert (round(float(_hmA["design-level optimized human-machine expected CRPS"]), 3)
        < round(float(_hmA["design-level optimized machine-only expected CRPS"]), 3)
        < round(float(_hmA["design-level equal-weight expected CRPS"]), 3)), \
    "the design-level machine-only score does not print strictly between the two pools at three decimals"
# the design-level solves are the full panel, the humans alone, the machines alone, the
# leave-one-out panels, and the M4 counterfactual, and the machine-only row is the printed score
_hmDS = _hmc["design_solves"].set_index("solve")
assert len(_hmDS) == 4 + _hmN
assert abs(float(_hmDS.loc["design machine-only", "expected_CRPS"])
           - float(_hmA["design-level optimized machine-only expected CRPS"])) < 1e-12
assert float(_hmA["machine weight share W_M"]) > 0.5 and float(_hmA["design-level machine weight share W_M"]) > 0.5, "C2 fails"
_f4 = _hmF.iloc[_hm4]
_f3 = _hmF.iloc[_hm3]
assert float(_f4["weight_full"]) > 1.0 / _hmN and float(_f4["weight_design"]) > 1.0 / _hmN, "C3: M4 weight not substantial"
assert float(_f4["delta_loo_hersbach"]) > 0 and float(_f4["delta_res"]) > 0 and float(_f4["delta_cal"]) < 0, "C3: M4 split"
assert float(_f4["delta_design"]) > 0 and float(_f4["delta_res_design"]) > 0 and float(_f4["delta_cal_design"]) < 0, "C3: M4 design split"
assert float(_f4["below_median_binomial_p"]) < float(_hmd["bias_test_alpha"]) and float(_f4["p_below_median_design"]) > 0.9, "C3: M4 bias hidden"
# the text: M3, equally biased and narrow, gets almost no weight and contributes nothing,
# its bias is flagged too, and about half its apparent resolution comes from its
# miscalibration, chiefly from its narrow spread (the spread correction alone removes
# most of the drop, and more than removing its mean error alone), recalibration leaving
# it below every human forecaster (in the draw and at the design level)
assert float(_f3["weight_full"]) < 0.5 / _hmN and float(_f3["weight_design"]) < 0.5 / _hmN
assert abs(float(_f3["delta_loo_hersbach"])) < 1e-3 and float(_f3["below_median_binomial_p"]) < float(_hmd["bias_test_alpha"])
for _sfx in ("", "_design"):
    _c, _cr = "resolution" + _sfx, "resolution_recalibrated" + _sfx
    _cm, _cs = "resolution_mean_error_removed" + _sfx, "resolution_spread_corrected" + _sfx
    assert 0.4 <= 1.0 - float(_f3[_cr]) / float(_f3[_c]) <= 0.6, "M3: the miscalibration share of its resolution is not about half"
    assert float(_f3[_cr]) < float(_hmF[_cr].to_numpy(float)[_hmH].min()), "M3 recalibrated not below every human"
    _drop, _drop_s, _drop_m = (float(_f3[_c]) - float(_f3[_cr]), float(_f3[_c]) - float(_f3[_cs]),
                               float(_f3[_c]) - float(_f3[_cm]))
    assert _drop_s > 0.5 * _drop and _drop_s > _drop_m, "M3: the spread is not the chief part of the drop (%s)" % _sfx
# the single corrections restate from the stored record (draw) and the stored design (design level)
_hm_zl = _norm_ppf(np.arange(1, int(_hmd["hersbach_levels_L"]) + 1) / (int(_hmd["hersbach_levels_L"]) + 1.0))
_hm_uncd = float(_hmd["uncertainty_UNC_design"])
for _j in range(_hmN):
    _e = _hm_y - _hm_mu[:, _j]
    _sdj = float(_hmF.loc[_j, "spread"])
    assert abs(_hm_unc - _hersbach(_hm_mu[:, _j] + _e.mean(), np.full(_hmT, _sdj), _hm_y, _hm_zl)[2]
               - float(_hmF.loc[_j, "resolution_mean_error_removed"])) < 1e-12
    assert abs(_hm_unc - _hersbach(_hm_mu[:, _j], np.full(_hmT, float(_e.std())), _hm_y, _hm_zl)[2]
               - float(_hmF.loc[_j, "resolution_spread_corrected"])) < 1e-12
    _esd = math.sqrt(_hmsu ** 2 + float(_hmtau[_j]) ** 2)
    assert abs(_hm_uncd - _vg_pop_hersbach(0.0, _esd, _sdj, _hm_zl)[2]
               - float(_hmF.loc[_j, "resolution_mean_error_removed_design"])) < 1e-12
    assert abs(_hm_uncd - _vg_pop_hersbach(-float(_hmb[_j]), _esd, _esd, _hm_zl)[2]
               - float(_hmF.loc[_j, "resolution_spread_corrected_design"])) < 1e-12
# the text: the resolution channel of M4 comes from its reading of the signal and the
# calibration channel nets the narrowing gain against the bias cost (sheet
# channel_counterfactuals, recomputed here at the stored weights held fixed)
_hmCF = _hmc["channel_counterfactuals"]
_hm_cftol = float(_hmd["channel_counterfactual_resolution_tolerance"])
_hm_wl4 = np.nan_to_num(_hmc["loo_weights"].set_index("withheld").loc["M4", _hmids].to_numpy(float))
_hm_dwl4 = np.nan_to_num(_hmc["design_loo_weights"].set_index("withheld").loc["M4", _hmids].to_numpy(float))
_hm_dwf = _hmW["weight_design"].to_numpy(float)
assert len(_hmCF) == 6 and set(_hmCF["level"]) == {"draw", "design"}
for _k, _row in _hmCF.iterrows():
    _MUc = _hm_mu.copy()
    _MUc[:, _hm4] += float(_row["M4_bias"]) - float(_hmb[_hm4])
    _SDc = _hm_sd.copy()
    _SDc[_hm4] = float(_row["M4_spread"])
    _bc = _hmb.copy()
    _bc[_hm4] = float(_row["M4_bias"])
    if str(_row["level"]) == "draw":
        _a = _hersbach(_MUc @ _wf, np.full(_hmT, float(_SDc @ _wf)), _hm_y, _hm_zl)
        _l = _hersbach(_MUc @ _hm_wl4, np.full(_hmT, float(_SDc @ _hm_wl4)), _hm_y, _hm_zl)
    else:
        _a, _l = (_vg_pop_hersbach(-float(_w @ _bc), math.sqrt(_hmsu ** 2 + float((_w * _w) @ (_hmtau ** 2))),
                                   float(_w @ _SDc), _hm_zl) for _w in (_hm_dwf, _hm_dwl4))
    assert abs((_l[0] - _a[0]) - float(_row["delta"])) < 1e-12
    assert abs((_l[1] - _a[1]) - float(_row["delta_cal"])) < 1e-12
    assert abs((_l[2] - _a[2]) - float(_row["delta_res"])) < 1e-12
for _lev, _sfx in (("draw", ""), ("design", "_design")):
    _q = _hmCF[_hmCF["level"] == _lev].set_index("setting")
    _g = _q.loc["as generated"]
    assert abs(float(_g["delta_cal"]) - float(_f4["delta_cal" + _sfx])) < 1e-12
    assert abs(float(_g["delta_res"]) - float(_f4["delta_res" + _sfx])) < 1e-12
    _nb, _ws = _q.loc["M4 bias removed"], _q.loc["M4 spread set to its design error"]
    assert abs(float(_ws["M4_spread"]) - float(_f4["design_rmse_of_mean"])) < 1e-12 and float(_nb["M4_bias"]) == 0.0
    for _cf in (_nb, _ws):
        assert abs(float(_cf["delta_res"]) - float(_g["delta_res"])) < _hm_cftol * abs(float(_g["delta_res"])), \
            "M4 (%s): the resolution channel is not its reading of the signal" % _lev
    assert float(_nb["delta_cal"]) > 0, "M4 (%s): narrowing gains nothing in calibration once the bias is removed" % _lev
    assert float(_ws["delta_cal"]) < float(_g["delta_cal"]), "M4 (%s): narrowness does not help its calibration channel" % _lev
# the text: M4 has the highest plotted and the highest recalibrated resolution, on the
# plotted L = 19 representation and on every finer one, in the draw and at the design level
assert int(np.argmax(_hmF["resolution"].to_numpy(float))) == _hm4
assert int(np.argmax(_hmF["resolution_recalibrated"].to_numpy(float))) == _hm4
_hmRL = _hmc["resolution_by_L"]
assert sorted(set(_hmRL["L"].astype(int))) == sorted(int(v) for v in str(_hmd["resolution_by_L_levels"]).split(","))
assert int(_hmd["hersbach_levels_L"]) in set(_hmRL["L"].astype(int))
for _L, _q in _hmRL.groupby("L"):
    _q = _q.set_index("id").loc[_hmids]
    for _c in ("resolution", "resolution_recalibrated", "resolution_design", "resolution_recalibrated_design"):
        assert int(np.argmax(_q[_c].to_numpy(float))) == _hm4, "M4 not the most resolving at L = %d (%s)" % (_L, _c)
    for _c in ("resolution_recalibrated", "resolution_recalibrated_design"):
        assert float(_q[_c].iloc[_hm3]) < float(_q[_c].to_numpy(float)[_hmH].min())
    # the plotted excess of the narrow biased M4 over its recalibrated value persists on every grid
    assert float(_q["resolution"].iloc[_hm4]) > float(_q["resolution_recalibrated"].iloc[_hm4])
    assert float(_q["resolution_design"].iloc[_hm4]) > float(_q["resolution_recalibrated_design"].iloc[_hm4])
    if int(_L) == int(_hmd["hersbach_levels_L"]):
        for _c in ("resolution", "resolution_recalibrated", "resolution_design", "resolution_recalibrated_design"):
            assert np.allclose(_q[_c].to_numpy(float), _hmF[_c].to_numpy(float), atol=1e-12, rtol=0), _c
# ... and it shrinks as the representation is refined (in part the narrow-spread grid
# dependence of Chapter 5): the plotted resolution of M3 and M4 at the finest L lies below
# its L = 19 value, in the draw and at the design level; the finest value of M4 is printed
_hmLs = sorted(set(_hmRL["L"].astype(int)))
_hmq0 = _hmRL[_hmRL["L"] == _hmLs[0]].set_index("id")
_hmq1 = _hmRL[_hmRL["L"] == _hmLs[-1]].set_index("id")
for _k in ("M3", "M4"):
    for _c in ("resolution", "resolution_design"):
        assert float(_hmq1.loc[_k, _c]) < float(_hmq0.loc[_k, _c]), "%s: the excess does not shrink (%s)" % (_k, _c)
assert abs(float(_hmq1.loc["M4", "resolution"])
           - float(_hmA["M4 plotted resolution on the finest representation (L = %d)" % _hmLs[-1]])) < 1e-12
# the text: narrowness raises M4's design-level weight, and the pooled calibrated members are over-dispersed
assert float(_hmA["design-level weight of M4 with its spread set to its error"]) < float(_f4["weight_design"])
assert float(_hmA["equal-weight pool of the six calibrated members, spread"]) > \
    float(_hmA["equal-weight pool of the six calibrated members, design root-mean-square error of the mean"])
# the wording of the text: M1, M2 and M4 hold the three largest weights, in the draw and
# at the design level (its frequency across re-seeds is printed), just under a quarter
# each in this draw, M4 at nearly twice its equal share, with an individual score inside
# the human range; narrowness about doubles its design-level weight; H3 holds no weight yet
# records a (penalty-driven) contribution; in this draw the three machines with the largest
# weights contribute most, then H2, and M3 contributes nothing measurable; the others'
# largest plotted and recalibrated resolutions print alike; the six-member pool's design
# and realized errors print alike
assert 1.9 <= float(_f4["weight_full"]) * _hmN < 2.0
_hmwf = _hmW.set_index("id")["weight_full"]
_hmwd = _hmW.set_index("id")["weight_design"]
assert max(abs(_hmwf["M1"] - _hmwf["M4"]), abs(_hmwf["M2"] - _hmwf["M4"])) < float(_hmd["tie_tolerance"])
assert all(0.24 <= _hmwf[k] < 0.25 for k in ("M1", "M2", "M4")), "not just under a quarter each"
for _wser in (_hmwf, _hmwd):
    assert min(_wser[k] for k in ("M1", "M2", "M4")) > max(_wser[k] for k in ("H1", "H2", "H3", "H4", "M3"))
_hmci = _hmF["mean_CRPS_individual"].to_numpy(float)
assert _hmci[_hmH].min() <= _hmci[_hm4] <= _hmci[_hmH].max()
# the text (Chapter 7): M4 held, as its design dictates, about twice the weight it would
# have held at its error; the ratio at the design level lies in the stated band (its
# across-seed median and the draw's percentile are checked in (6g)), and the weight falls
# in the draw and at the design level
_hm_rb = tuple(float(v) for v in str(_hmd["counterfactual_ratio_band"]).split(","))
_hm_rdes = float(_vg_cf_ratio(float(_f4["weight_design"]), float(_hmA["design-level weight of M4 with its spread set to its error"])))
_hm_rdraw = float(_vg_cf_ratio(float(_f4["weight_full"]), float(_wcf[_hm4])))
assert abs(_hm_rdes - float(_hmA["ratio of M4's weight to its weight at its error (design level)"])) < 1e-12
assert abs(_hm_rdraw - float(_hmA["ratio of M4's weight to its weight at its error (draw)"])) < 1e-12
assert _hm_rb[0] <= _hm_rdes <= _hm_rb[1], "M4: the design-level ratio is not about two (%.3f)" % _hm_rdes
assert _hm_rdraw > 1.0 and _hm_rdes > 1.0, "M4: narrowness does not raise its weight"
_hm_h3 = _hmids.index("H3")
assert (float(_hmW.loc[_hm_h3, "weight_full"]) < 1e-9 and float(_hmF.loc[_hm_h3, "delta_loo_hersbach"]) > 0
        and float(_hmF.loc[_hm_h3, "delta_loo"]) > 0), "H3: zero weight with a nonzero contribution"
_hmdl = _hmF.set_index("id")["delta_loo_hersbach"]
assert _hmdl["M2"] > _hmdl["M1"] > _hmdl["M4"] > _hmdl["H2"] > max(_hmdl[k] for k in ("H1", "H3", "H4", "M3"))
assert abs(_hmdl["M3"]) < 1e-3
_hm_oth = [i for i in range(_hmN) if i not in (_hm3, _hm4)]
assert (round(float(_hmF["resolution"].to_numpy(float)[[i for i in range(_hmN) if i != _hm4]].max()), 2)
        == round(float(_hmF["resolution_recalibrated"].to_numpy(float)[_hm_oth].max()), 2))
assert (round(float(_hmA["equal-weight pool of the six calibrated members, design root-mean-square error of the mean"]), 2)
        == round(float(_hmA["equal-weight pool of the six calibrated members, realized root-mean-square error of the mean"]), 2))
# the text: the calibrated machines contribute through both channels, in the draw and at the design level
for _k in ("M1", "M2"):
    _r = _hmF.set_index("id").loc[_k]
    assert min(float(_r["delta_cal"]), float(_r["delta_res"]), float(_r["delta_cal_design"]), float(_r["delta_res_design"])) > 0, _k
_band = float(_hmd["well_calibrated_cut"])
assert abs(_band - 0.2 * _hm_unc) < 1e-12
assert (_hmF["REL_i"].to_numpy(float)[_hmCal] < _band).all(), "C4: a calibrated member leaves the band"
assert (_hmF["REL_i"].to_numpy(float)[[_hm3, _hm4]] > _band).all()
_ud = float(_hmd["uncertainty_UNC_design"])
assert abs(_ud - float(_hmA["uncertainty UNC (design)"])) < 1e-12
assert (_hmF["REL_design"].to_numpy(float)[_hmCal] < 0.2 * _ud).all() and \
    (_hmF["REL_design"].to_numpy(float)[[_hm3, _hm4]] > 0.2 * _ud).all(), "C4 fails at the design level"
# leave-one-out robustness (tab:robustness): every re-solve beats human-only, in-sample and on new quarters
assert (_hmF["crps_loo_resolve"].to_numpy(float) < _h_in).all() and int(_hmA["leave-one-out re-solves scoring below the human-only aggregate"]) == _hmN
assert (_hmF["expected_CRPS_loo_new_quarters"].to_numpy(float) < _h_new).all()
assert int(_hmA["leave-one-out pools scoring below the human-only pool on new quarters"]) == _hmN
assert abs(float(_hmA["equal-weight pool of the six calibrated members, realized root-mean-square error of the mean"])
           - float(np.sqrt(np.mean((_hm_y - _hm_mu[:, _hmCal].mean(axis=1)) ** 2)))) < 1e-12
# the draw conforms to its design (config design_conformity_rule): realized errors, the
# independence of the reading noise from the unobserved residual (recomputed from the
# stored record), and the realized residual scale (percentile checked in (6g))
_hm_rlo, _hm_rhi = (float(v) for v in str(_hmd["design_conformity_rmse_ratio"]).split(","))
_hmO_base = (_hmO["conditional_mean_given_previous"] + _hmO["leading_signal_x"]).to_numpy(float)
_hmO_u = _hmO["unobserved_residual_u"].to_numpy(float)
for _j, _r in _hmF.iterrows():
    _ratio = float(_r["realized_rmse_of_mean"]) / float(_r["design_rmse_of_mean"])
    assert abs(_ratio - float(_r["realized_over_design_rmse"])) < 1e-12
    assert _hm_rlo <= _ratio <= _hm_rhi, "%s: realized error departs from its design" % _r["id"]
    _xi = (_hm_mu[:, _j] - _hmO_base - float(_r["bias"])) / float(_r["error_scale_tau"])
    _cr = float(np.corrcoef(_hmO_u, _xi)[0, 1])
    assert abs(_cr - float(_r["realized_corr_residual_noise"])) < 1e-12
    assert abs(_cr) < float(_hmd["design_conformity_corr_z"]) / math.sqrt(_hmT), "%s reads the residual" % _r["id"]
# (6g) the re-seeding study: rows, seeds, claims and frequencies recomputed
_hmR = _hmc["reseed"]
_hmK = int(_hmd["reseed_count"])
assert len(_hmR) == 1 + _hmK and int(_hmR["seed"].iloc[0]) == NEW_SEED
assert list(_hmR["seed"].iloc[1:]) == list(range(int(_hmd["reseed_first"]), int(_hmd["reseed_first"]) + _hmK))
_hmR1 = _hmR.iloc[1:].reset_index(drop=True)
_hmSum = _hmc["reseed_summary"].set_index("claim")
for _k, _dsc, _pred in _VG_CLAIMS:
    _col = np.array([int(bool(_pred(r))) for r in _hmR1.to_dict("records")])
    assert (_col == _hmR1[_k].to_numpy(int)).all(), "re-seed claim column %s is not its predicate" % _k
    assert abs(float(_hmSum.loc[_k, "frequency"]) - _col.mean()) < 1e-12
    assert int(_hmA["re-seeding count: " + _k]) == int(_col.sum()) == int(_hmSum.loc[_k, "count"])
    assert int(_hmSum.loc[_k, "holds_in_seed_202603"]) == int(bool(_pred(_hmR.iloc[0].to_dict())))
assert int(_hmR["epigraph_failures"].sum()) == 0, "an SLSQP run of the re-seeding study ended without success"
_hmall = _hmR1[_VG_CORE].min(axis=1).to_numpy(int)
assert abs(float(_hmSum.loc["all_core_claims", "frequency"]) - _hmall.mean()) < 1e-12
assert int(_hmSum.loc["all_core_claims", "holds_in_seed_202603"]) == 1, "the reported draw fails a claim"
# every core claim is a property of the design: it holds in at least nine records in ten
assert all(float(_hmSum.loc[_k, "frequency"]) >= 0.9 for _k in _VG_CORE), _hmSum["frequency"]
# the seed-202603 row (epigraph solve) reproduces the Algorithm 4.2 values of the draw
_hmR0 = _hmR.iloc[0]
assert abs(float(_hmR0["ins_hm"]) - _c_in) < 1e-6 and abs(float(_hmR0["ins_h"]) - _h_in) < 1e-6
assert max(abs(float(_hmR0["w_" + k]) - float(_wf[j])) for j, k in enumerate(_hmids)) < 1e-3
assert max(abs(float(_hmR0["delta_" + k]) - float(_hmF.loc[j, "delta_loo_hersbach"])) for j, k in enumerate(_hmids)) < 1e-5
assert abs(float(_hmR0["rmse_cal6"]) - float(_hmA["equal-weight pool of the six calibrated members, realized root-mean-square error of the mean"])) < 1e-12
assert abs(float(_hmR0["UNC"]) - _hm_unc) < 1e-12 and abs(float(_hmR0["sd_u"]) - float(_hmA["realized rms of the unobserved residual u"])) < 1e-12
assert all(abs(float(_hmR0["crps_" + k]) - float(_hmF.loc[j, "mean_CRPS_individual"])) < 1e-12 for j, k in enumerate(_hmids))
assert all(abs(float(_hmR0["RESrec_" + k]) - float(_hmF.loc[j, "resolution_recalibrated"])) < 1e-12 for j, k in enumerate(_hmids))
assert all(int(_hmR0["below_" + k]) == int(_hmF.loc[j, "count_below_median"]) for j, k in enumerate(_hmids))
# the percentiles of the typicality sheet, recomputed; every statistic the text reports
# for the draw lies inside the stated range of the across-seed distribution
_hmTy = _hmc["typicality"].set_index("statistic")
for _k, _row in _hmTy.iterrows():
    assert abs(float(_row["seed_202603"]) - float(_hmR0[_k])) < 1e-12
    assert abs(float(_row["percentile_of_seed_202603"]) - 100.0 * float((_hmR1[_k] <= float(_hmR0[_k])).mean())) < 1e-9
_hm_lo, _hm_hi = (float(v) for v in str(_hmd["typical_range_percent"]).split(","))
_hm_disc = dict(kv.split(":") for kv in str(_hmd["typical_statistics_disclosed"]).split(","))
_hm_printed = [k.strip() for k in str(_hmd["typical_statistics_printed"]).split(",")]
assert not set(_hm_printed) & set(_hm_disc), "a disclosed atypical statistic is also listed as typical"
for _k in _hm_printed + ["sd_u"]:
    _p = float(_hmTy.loc[_k, "percentile_of_seed_202603"])
    assert _hm_lo <= _p <= _hm_hi, "printed statistic %s of the draw is atypical: percentile %.1f" % (_k, _p)
for _k in [k.strip() for k in str(_hmd["typical_statistics_plotted"]).split(",")]:
    _p = float(_hmTy.loc[_k, "percentile_of_seed_202603"])
    if _k not in _hm_disc:
        assert _hm_lo <= _p <= _hm_hi, "plotted statistic %s of the draw is atypical and undisclosed: percentile %.1f" % (_k, _p)
assert set(_VG_TYPICAL_PLOTTED) == set(k.strip() for k in str(_hmd["typical_statistics_plotted"]).split(","))
assert set("REL_" + k for k in _hmids) | set("RES_" + k for k in _hmids) == set(_VG_TYPICAL_PLOTTED)
for _k, _side in _hm_disc.items():
    _p = float(_hmTy.loc[_k, "percentile_of_seed_202603"])
    assert (_p < _hm_lo) if _side == "below" else (_p > _hm_hi), "disclosed statistic %s is not atypical on its side" % _k
# the text discloses: the calibrated members' reliability errors are smaller than in most
# records (the largest over UNC, and its across-seed median, are printed), M3's larger
assert _hm_disc["max_REL_calibrated_over_UNC"] == "below" and _hm_disc["REL_M3"] == "above"
assert all(_hm_disc.get(k) in (None, "below") for k in ("REL_" + _hmids[i] for i in _hmCal))
assert abs(float(_hmA["largest reliability error of the six calibrated members over UNC"])
           - float(_hmR0["max_REL_calibrated_over_UNC"])) < 1e-12
assert abs(float(_hmR0["max_REL_calibrated_over_UNC"])
           - float(_hmF["REL_i"].to_numpy(float)[_hmCal].max()) / _hm_unc) < 1e-12
assert abs(float(_hmA["across-seed median of the largest reliability error of the six calibrated members over UNC"])
           - float(_hmR1["max_REL_calibrated_over_UNC"].median())) < 1e-12
assert abs(float(_hmR0["RES_finest_M4"]) - float(_hmq1.loc["M4", "resolution"])) < 1e-12
# "about twice" across the re-seeded records: the median ratio lies in the band, the
# draw's ratio inside the typical range, and the weight falls in nine records in ten;
# the seed-202603 row (epigraph solve) reproduces the Algorithm 4.2 counterfactual
assert abs(float(_hmR0["w_M4_spread_at_error"]) - float(_wcf[_hm4])) < 1e-3
_hm_rr = _vg_cf_ratio(_hmR1["w_M4"].to_numpy(float), _hmR1["w_M4_spread_at_error"].to_numpy(float))
_hm_rmed = float(np.median(_hm_rr))
_hm_rpct = float(100.0 * np.mean(_hm_rr <= float(_vg_cf_ratio(float(_hmR0["w_M4"]), float(_hmR0["w_M4_spread_at_error"])))))
assert abs(_hm_rmed - float(_hmA["across-seed median ratio of M4's weight to its weight at its error"])) < 1e-12
assert abs(_hm_rpct - float(_hmA["percentile of the draw's ratio of M4's weight to its weight at its error"])) < 1e-9
assert _hm_rb[0] <= _hm_rmed <= _hm_rb[1], "M4: the across-seed median ratio is not about two (%.3f)" % _hm_rmed
assert _hm_lo <= _hm_rpct <= _hm_hi, "M4: the draw's ratio is atypical (percentile %.1f)" % _hm_rpct
assert float(_hmSum.loc["S_M4_weight_falls_with_spread_at_error", "frequency"]) >= 0.9
assert int(_hmSum.loc["S_M4_weight_falls_with_spread_at_error", "holds_in_seed_202603"]) == 1
print("   vignette V: record regenerates; simplex, share, leave-one-out, Hersbach, spread and"
      " exact-objective identities hold; %d solves certified; claims C1-C4 hold in the draw and at"
      " the design level; %d re-seeded records, core claims jointly in %.1f%%"
      % (len(_hmSv), _hmK, 100.0 * _hmall.mean()))

# (7) ex:hmpanel study (fig5-3_loo, fig5-2_calibration_sharpness, tab5-3):
#     the WRITTEN workbooks are tested against invariants recomputed from
#     the data they store, never against typed constants.
_loo = _wb("fig5-3_loo.xlsx")
_lcfg = _cfgd(_loo)
_T7 = int(_lcfg["T"])
assert _loo["forecast_mu"].shape == (_T7, 10)
assert len(_loo["outcomes"]) == _T7
_sums_to_one(_loo["weights"]["weight_full_panel"], "loo full-panel weights", tol=1e-9)
_sums_to_one(_loo["weights"]["weight_human_only"], "loo human-only weights", tol=1e-9)
assert (_loo["weights"]["weight_full_panel"] >= 0).all() and (_loo["weights"]["weight_human_only"] >= 0).all()
assert len(_loo["loo"]) == 10 and len(_loo["individual"]) == 10
_y7 = _loo["outcomes"]["y"].to_numpy(float)
_MU7 = _loo["forecast_mu"].to_numpy(float)
_SD7 = _loo["experts"]["reported_sd"].to_numpy(float)
_lamG7, _lamC7 = float(_lcfg["lambda_G"]), float(_lcfg["lambda_C"])
_res7 = dict(zip(_loo["results"]["quantity"].astype(str), _loo["results"]["value"].astype(float)))
_wf7 = _loo["weights"]["weight_full_panel"].to_numpy(float)
_wh7 = _loo["weights"]["weight_human_only"].to_numpy(float)[:5]
_Pf7 = _hm_program(_y7, _MU7, _SD7, list(range(10)), _lamG7, _lamC7)
_Ph7 = _hm_program(_y7, _MU7, _SD7, list(range(5)), _lamG7, _lamC7)
# (7a) the stored objective and mean CRPS are recomputed from the stored data
_Lf7, _cf7, _, _ = _Pf7["terms"](_wf7)
_Lh7, _ch7, _, _ = _Ph7["terms"](_wh7)
assert abs(_Lf7 - _res7["full-panel objective"]) < 1e-9, "full objective not reproducible from the workbook"
assert abs(_cf7 - _res7["collaborative mean CRPS"]) < 1e-9, "collaborative CRPS not reproducible"
assert abs(_Lh7 - _res7["human-only objective"]) < 1e-9, "human-only objective not reproducible"
assert abs(_ch7 - _res7["human-only mean CRPS"]) < 1e-9, "human-only CRPS not reproducible"
assert abs(_res7["human weight share W_H"] - _wf7[:5].sum()) < 1e-12
assert abs(_res7["machine weight share W_M"] - _wf7[5:].sum()) < 1e-12
assert abs(_res7["CRPS gain (collab vs human-only, pct)"] - 100.0 * (_cf7 / _ch7 - 1.0)) < 1e-9
# (7b) the solver log certifies every solve: converged under the recorded
#      settings and matched by the independent SLSQP solve
_slog = _loo["solver_log"]
assert len(_slog) == 12, "expected 12 solves (full, human-only, 10 leave-one-out)"
assert (_slog["objective_gap"] <= float(_HM_VERIFY["obj_gap"])).all(), "a solve is beaten by SLSQP"
assert (_slog["weights_inf_gap_vs_slsqp"] <= float(_HM_VERIFY["w_gap"])).all()
assert (_slog["alg42_last_dw_inf"] < float(_lcfg["solver_tolerance"])).all()
assert (_slog["alg42_iterations"] >= int(_lcfg["solver_r_min"])).all()
assert (_slog["alg42_iterations"] < int(_lcfg["solver_r_max_cap"])).all()
# (7c) optimality invariant of the convex program: the returned full-panel
#      vector is no worse than every other feasible point the study stored
#      (equal weights, the vertices, the embedded human-only solution, the
#      embedded leave-one-out solutions), under the full-panel objective
_cands7 = [np.full(10, 0.1)] + [np.eye(10)[j] for j in range(10)] + [np.concatenate([_wh7, np.zeros(5)])]
for _, _srow in _slog.iterrows():
    if str(_srow["solve"]).startswith("leave-out"):
        _sidx = [int(v) for v in str(_srow["experts"]).split(",")]
        _sw = np.array([float(v) for v in str(_srow["weights"]).split(",")])
        _emb = np.zeros(10)
        _emb[_sidx] = _sw
        _cands7.append(_emb)
for _cw in _cands7:
    assert _Lf7 <= _Pf7["obj"](_cw) + 1e-12, "a stored feasible point beats the returned full-panel optimum"
# (7d) the channel identity is exact, and the full-panel Hersbach terms
#      are reproduced from the stored aggregate on the recorded grid
_lrows = _loo["loo"]
assert np.allclose(_lrows["delta_cal"] + _lrows["delta_res"], _lrows["delta_total"], atol=1e-12)
_L7 = int(_lcfg["hersbach_levels_L"])
_z7 = _norm_ppf(np.arange(1, _L7 + 1) / (_L7 + 1.0))
_chf7, _relf7, _potf7 = _hersbach(_MU7 @ _wf7, np.full(_T7, float(_SD7 @ _wf7)), _y7, _z7)
assert abs(_relf7 - _res7["full-panel Hersbach REL"]) < 1e-12
assert abs(_potf7 - _res7["full-panel Hersbach POT"]) < 1e-12
assert abs(_chf7 - _res7["full-panel Hersbach CRPS"]) < 1e-12
_Xc7 = np.tile(np.quantile(_y7, np.arange(1, _L7 + 1) / (_L7 + 1.0)), (_T7, 1))
assert abs(_hersbach_X(_Xc7, _y7)[2] - float(_lcfg["uncertainty_UNC_grid"])) < 1e-12
_ind7 = _loo["individual"]
assert np.allclose(_ind7["RES_i"] + _ind7["POT_i"], float(_lcfg["uncertainty_UNC_grid"]), atol=1e-12)
assert np.allclose(_ind7["crps_i_hersbach"], _ind7["REL_i"] + _ind7["POT_i"], atol=1e-12)
_bs7 = _loo["binning_sensitivity"]
_bs19 = _bs7[_bs7["L"] == _L7].set_index("id").loc[list(_lrows["id"].astype(str))]
assert np.allclose(_bs19["delta_total"].to_numpy(float), _lrows["delta_total"].to_numpy(float), atol=1e-12)
_unst7 = [k for k in _lrows["id"].astype(str)
          if len(set(np.sign(_bs7.loc[_bs7["id"] == k, "delta_total"].to_numpy(float)))) > 1]
assert ",".join(_unst7) == str(_lcfg["total_sign_unstable_across_L"]) or (not _unst7 and str(_lcfg["total_sign_unstable_across_L"]) == "none")
# (7e) the design invariant the figure is built to show: both overconfident
#      machines in the (-, +) cell on the recorded grid
_over7 = _lrows[_lrows["group"].astype(str) == "machine overconfident"]
assert len(_over7) == 2 and (_over7["delta_cal"] < 0).all() and (_over7["delta_res"] > 0).all()
# (7f) Figure 5.2 is the individual sheet, point for point, with the same UNC
_cs7 = _wb("fig5-2_calibration_sharpness.xlsx")
_cs7c = _cfgd(_cs7)
assert abs(float(_cs7c["uncertainty_UNC_grid"]) - float(_lcfg["uncertainty_UNC_grid"])) < 1e-15
_pts7 = _cs7["points"].set_index("id")
_ind7i = _ind7.set_index("id")
assert list(_pts7.index) == list(_ind7i.index)
assert np.allclose(_pts7["RES"], _ind7i["RES_i"], atol=1e-15) and np.allclose(_pts7["REL"], _ind7i["REL_i"], atol=1e-15)
# the callout of Figure 5.2 names a group of the points sheet, its text names the
# first and last member of that group, and the group's members nearly coincide: the
# RES range and the REL range of the group are each below the stated fraction of the
# axis range
_c52grp = str(_cs7c["callout_group"])
_c52pts = _pts7[_pts7["group"].astype(str) == _c52grp]
assert len(_c52pts) >= 2, "the Figure 5.2 callout group %r is not a group of the points sheet" % _c52grp
assert str(_cs7c["callout_text"]).replace("\\n", "\n").split("\n")[0] == "%s\u2013%s" % (_c52pts.index[0], _c52pts.index[-1])
_c52frac = float(_cs7c["callout_coincident_max_axis_fraction"])
_c52xr = float(_cs7c["xmax"]) - float(_cs7c["xmin"])
_c52yr = float(_cs7c["ymax"]) - float(_cs7c["ymin"])
assert float(_c52pts["RES"].max() - _c52pts["RES"].min()) < _c52frac * _c52xr, "the Figure 5.2 callout group spreads in RES"
assert float(_c52pts["REL"].max() - _c52pts["REL"].min()) < _c52frac * _c52yr, "the Figure 5.2 callout group spreads in REL"
# Figures 5.2 and 6.12 code the same three roles by the same color and marker
_role7 = {"human": ("human experts", "human forecasters"),
          "machine calibrated": ("calibrated machines", "calibrated machines"),
          "machine overconfident": ("overconfident machines", "overconfident machines")}
for _g7, (_n52, _n612) in _role7.items():
    _s52 = _pts7[_ind7i["group"].astype(str) == _g7]
    _s612 = _hmF[_hmF["group"].astype(str) == _n612]
    assert len(_s52) and len(_s612) and set(_s52["group"].astype(str)) == {_n52}
    assert {(str(a), str(b)) for a, b in zip(_s52["color"], _s52["marker"])} == \
        {(str(a), str(b)) for a, b in zip(_s612["color"], _s612["marker"])}, \
        "Figures 5.2 and 6.12 code the role %s differently" % _g7
# (7g) Table 5.3 is the rounded results sheet
_t53 = _wb("tab5-3_hm_panel.xlsx")
_t53r = _t53["table_5_3"]
_t53p = _cfgd(_t53, "provenance")
assert float(_t53r.loc[0, "mean_CRPS"]) == round(_res7["human-only mean CRPS"], 2)
assert float(_t53r.loc[1, "mean_CRPS"]) == round(_res7["collaborative mean CRPS"], 2)
assert float(_t53r.loc[2, "optimized_weight_share"]) == round(_res7["human weight share W_H"], 2)
assert float(_t53r.loc[3, "optimized_weight_share"]) == round(_res7["machine weight share W_M"], 2)
assert int(_t53p["gain_pct_printed"]) == round(_res7["CRPS gain (collab vs human-only, pct)"])
assert int(_t53p["seed"]) == int(_lcfg["seed"])
print("   ex:hmpanel study: objectives and scores recomputed from the stored data; 12 solves"
      " certified by the solver log; optimality against %d stored feasible points;"
      " exact channel identity; UNC on the grid; fig5-2 == individual sheet, its callout group"
      " coincident within the stated axis fraction; tab5-3 == rounded results"
      % len(_cands7))

# (7h) the claims Section 5.5 and Example ex:corr-lineage make about the
#      dependence structure of ex:hmpanel, recomputed from the stored
#      outcomes and forecast medians (instrument of Section 5.6)
_E7 = _MU7 - _y7[:, None]
_Rerr7 = np.corrcoef(_E7, rowvar=False)
_Rfc7 = np.corrcoef(_MU7, rowvar=False)
assert np.allclose(_Rerr7, _loo["error_correlation"][list(_loo["forecast_mu"].columns)].to_numpy(float), atol=1e-12)
assert np.allclose(_Rfc7, _loo["forecast_correlation"][list(_loo["forecast_mu"].columns)].to_numpy(float), atol=1e-12)
_dep7 = _loo["error_dependence"].set_index("quantity")
for _q, _dr in _dep7.iterrows():
    assert float(_dr["printed"]) == round(float(_dr["value"]), 2), "error_dependence %s: printed != rounded" % _q
_H7, _C7, _O7 = [0, 1, 2, 3, 4], [5, 6, 7], [8, 9]
_M7 = _C7 + _O7


def _blk7(Rm, a, b):
    return np.array([Rm[i, j] for i in a for j in b if (a != b or i < j)])


def _dv7(q):
    return float(_dep7.loc[q, "value"])


for _nm, _a, _b in (("human-human", _H7, _H7), ("machine-machine", _M7, _M7),
                    ("calibrated-overconfident", _C7, _O7), ("human-machine", _H7, _M7)):
    assert abs(_blk7(_Rerr7, _a, _b).min() - _dv7("error r %s min" % _nm)) < 1e-12
    assert abs(_blk7(_Rerr7, _a, _b).max() - _dv7("error r %s max" % _nm)) < 1e-12
    assert abs(_blk7(_Rfc7, _a, _b).min() - _dv7("forecast r %s min" % _nm)) < 1e-12
    assert abs(_blk7(_Rfc7, _a, _b).max() - _dv7("forecast r %s max" % _nm)) < 1e-12
for _nm, _idx in (("human", _H7), ("machine", _M7), ("panel", list(range(10)))):
    _we = np.full(len(_idx), 1.0 / len(_idx))
    assert abs(1.0 / (_we @ _Rerr7[np.ix_(_idx, _idx)] @ _we) - _dv7("N_eff equal weights %s" % _nm)) < 1e-12
assert abs(1.0 / (_wf7 @ _Rerr7 @ _wf7) - _dv7("N_eff full-panel optimized weights")) < 1e-12
_r2_7 = np.array([np.corrcoef(_MU7[:, i], _y7)[0, 1] ** 2 for i in range(10)])
assert abs(_r2_7.min() - _dv7("squared corr(median, outcome) min")) < 1e-12
assert abs(_r2_7.max() - _dv7("squared corr(median, outcome) max")) < 1e-12
# the text: each subpanel's errors are highly correlated (every within-subpanel
# pair above every human-machine pair); the human and machine forecasts are
# nearly uncorrelated; the calibrated and overconfident machines read partly
# different signals (forecast r well below the within-block values); every
# median misses most of the outcome variance; each subpanel is no less
# redundant than the ex:corr-lineage lineage, while the full panel counts for
# more independent experts than the humans alone
assert min(_blk7(_Rerr7, _H7, _H7).min(), _blk7(_Rerr7, _M7, _M7).min()) > _blk7(_Rerr7, _H7, _M7).max()
assert np.abs(_blk7(_Rfc7, _H7, _M7)).max() < 0.15
assert _blk7(_Rfc7, _C7, _O7).max() < min(_blk7(_Rfc7, _C7, _C7).min(), _blk7(_Rfc7, _O7, _O7).min())
# members of a subpanel share much of their information (forecast r above 0.75)
assert min(_blk7(_Rfc7, _H7, _H7).min(), _blk7(_Rfc7, _C7, _C7).min(), _blk7(_Rfc7, _O7, _O7).min()) > 0.75
# the machine subpanel is unbiased (by design every component has mean zero;
# in the sample no machine's mean error is two standard errors from zero)
for _j in _M7:
    assert abs(_E7[:, _j].mean()) < 2.0 * _E7[:, _j].std() / math.sqrt(_T7)
assert _r2_7.max() < 0.5
_dcfg7 = _cfgd(_wb("fig5-4_depnet.xlsx"))
assert max(_dv7("N_eff equal weights human"), _dv7("N_eff equal weights machine")) <= float(_dcfg7["neff_lineage_equalweight"])
assert round(_dv7("N_eff equal weights human"), 1) == round(_dv7("N_eff equal weights machine"), 1) == round(float(_dcfg7["neff_lineage_equalweight"]), 1)
assert _dv7("N_eff equal weights panel") > _dv7("N_eff equal weights human")
# (7i) the leave-one-out claims of Section 5.5: M1-M3 totals near zero or
#      negative on every grid (below 1% of the value the machines add
#      together, the human-only minus collaborative score) and negative
#      under the continuous def:loo score, their only positive channel the
#      calibration channel; the resolution channel of M4/M5 positive in the
#      convention-free limit
_gain7 = _res7["human-only mean CRPS"] - _res7["collaborative mean CRPS"]
assert _gain7 > 0
_lr7 = _lrows.set_index("id")
for _k in ("M1", "M2", "M3"):
    _tk = _bs7.loc[_bs7["id"] == _k]
    assert _tk["delta_total"].max() < 0.01 * _gain7, "%s total not near zero" % _k
    assert (_tk["delta_cal"] > 0).all() and (_tk["delta_res"] < 0).all()
    assert float(_lr7.loc[_k, "delta_cont_defloo"]) < 0
for _k in ("M4", "M5"):
    assert float(_lr7.loc[_k, "delta_res_limit"]) > 0
_mua7 = {}
for _, _srow in _slog.iterrows():
    _sidx = [int(v) for v in str(_srow["experts"]).split(",")]
    _sw = np.array([float(v) for v in str(_srow["weights"]).split(",")])
    _mua7[str(_srow["solve"])] = _MU7[:, _sidx] @ _sw


def _gmd7(e):
    return 0.5 * np.abs(e[:, None] - e[None, :]).mean()


for _k in _lr7.index:
    _dl = _gmd7(_y7 - _mua7["leave-out " + _k]) - _gmd7(_y7 - _mua7["full panel"])
    assert abs(_dl - float(_lr7.loc[_k, "delta_res_limit"])) < 1e-12
# (7j) the individual-resolution claims of Section 5.5 (Figure 5.2): on the
#      L = 19 grid M4 and M5 have the two highest RES_i; at L = 99 H3 ranks
#      above both; with the same means at an honest spread, and in the
#      convention-free limit, their RES_i lies below every human expert's and
#      above every calibrated machine's, the two versions agreeing to the
#      printed three decimals; the limit is recomputed, and checked against a
#      fine grid for the experts whose spread is honest
_ind7i = _ind7.set_index("id")
_ids7 = list(_ind7i.index)
for _L in [int(v) for v in str(_lcfg["binning_sensitivity_L"]).split(",")]:
    _pl = np.arange(1, _L + 1) / (_L + 1.0)
    _U7L = _hersbach_X(np.tile(np.quantile(_y7, _pl), (_T7, 1)), _y7)[2]
    for _j, _k in enumerate(_ids7):
        _v = _U7L - _hersbach(_MU7[:, _j], np.full(_T7, _SD7[_j]), _y7, _norm_ppf(_pl))[2]
        assert abs(_v - float(_ind7i.loc[_k, "RES_i_L%d" % _L])) < 1e-12
assert np.allclose(_ind7i["RES_i_L%d" % _L7], _ind7i["RES_i"], atol=1e-15)
_top2 = list(_ind7i["RES_i"].sort_values(ascending=False).index[:2])
assert sorted(_top2) == ["M4", "M5"], _top2
assert float(_ind7i.loc["H3", "RES_i_L99"]) > max(float(_ind7i.loc[k, "RES_i_L99"]) for k in ("M4", "M5"))
_Uc7 = 0.5 * np.abs(_y7[:, None] - _y7[None, :]).mean()
for _j, _k in enumerate(_ids7):
    assert abs(_Uc7 - _gmd7(_y7 - _MU7[:, _j]) - float(_ind7i.loc[_k, "RES_i_limit"])) < 1e-12
_hum7 = ["H1", "H2", "H3", "H4", "H5"]
_cal7 = ["M1", "M2", "M3"]
for _col in ("RES_i_limit", "RES_i_honest_spread"):
    for _k in ("M4", "M5"):
        assert float(_ind7i.loc[_k, _col]) < min(float(_ind7i.loc[h, _col]) for h in _hum7)
        assert float(_ind7i.loc[_k, _col]) > max(float(_ind7i.loc[c, _col]) for c in _cal7)
for _k in ("M4", "M5"):
    assert round(float(_ind7i.loc[_k, "RES_i_limit"]), 3) == round(float(_ind7i.loc[_k, "RES_i_honest_spread"]), 3)
_Lfine7 = 999
_pf7 = np.arange(1, _Lfine7 + 1) / (_Lfine7 + 1.0)
_Uf7 = _hersbach_X(np.tile(np.quantile(_y7, _pf7), (_T7, 1)), _y7)[2]
for _j, _k in enumerate(_ids7):
    if _k in ("M4", "M5"):
        continue
    _vf = _Uf7 - _hersbach(_MU7[:, _j], np.full(_T7, _SD7[_j]), _y7, _norm_ppf(_pf7))[2]
    assert abs(_vf - float(_ind7i.loc[_k, "RES_i_limit"])) < 5e-4, "%s: L=%d resolution %.5f vs limit %.5f" % (
        _k, _Lfine7, _vf, float(_ind7i.loc[_k, "RES_i_limit"]))
# (7k) the reliability floor quoted in Section 5.5: closed form reproduced
#      by the Hersbach routine on a deterministic N(0,1) quasi-sample, most
#      of it from the two outer bins at every L, and the calibrated experts'
#      REL_i scattered about their floors on the L = 19 grid
_fl7 = _loo["reliability_floor"].set_index("L")
_yq7 = _norm_ppf((np.arange(1, 200001) - 0.5) / 200000.0)
for _L, _fr in _fl7.iterrows():
    _v = _hersbach(np.zeros_like(_yq7), np.ones_like(_yq7), _yq7, _norm_ppf(np.arange(1, _L + 1) / (_L + 1.0)))[1]
    assert abs(_v - float(_fr["REL_floor_per_unit_sd"])) < 0.02 * float(_fr["REL_floor_per_unit_sd"]), (_L, _v)
    assert float(_fr["share_from_outer_bins"]) > 0.5
    assert str(_fr["printed"]) == "%.2g" % float(_fr["REL_floor_per_unit_sd"])
assert list(_fl7["REL_floor_per_unit_sd"]) == sorted(_fl7["REL_floor_per_unit_sd"], reverse=True)
_calib7 = _hum7 + _cal7
assert np.allclose(_ind7i.loc[_calib7, "REL_floor_if_calibrated"],
                   _ind7i.loc[_calib7, "reported_sd"] * float(_fl7.loc[_L7, "REL_floor_per_unit_sd"]), atol=1e-15)
assert _ind7i.loc[_calib7, "REL_i"].min() < _ind7i.loc[_calib7, "REL_floor_if_calibrated"].min()
assert _ind7i.loc[_calib7, "REL_i"].max() > _ind7i.loc[_calib7, "REL_floor_if_calibrated"].max()
assert abs(float(_cs7c["rel_floor_per_unit_sd"]) - float(_fl7.loc[_L7, "REL_floor_per_unit_sd"])) < 1e-15
assert np.allclose(_pts7["REL_floor"], _ind7i["REL_floor_if_calibrated"], atol=1e-15)
assert np.allclose(_pts7["RES_limit"], _ind7i["RES_i_limit"], atol=1e-15)
# the text_numbers sheet summarizes the individual sheet at the printed rounding
_tn7 = _loo["text_numbers"].set_index("quantity")
_tn7_expect = {
    "REL_i min, humans and calibrated machines (L = %d)" % _L7: (_ind7i.loc[_calib7, "REL_i"].min(), 4),
    "REL_i max, humans and calibrated machines (L = %d)" % _L7: (_ind7i.loc[_calib7, "REL_i"].max(), 4),
    "REL floor min, humans and calibrated machines (L = %d)" % _L7: (_ind7i.loc[_calib7, "REL_floor_if_calibrated"].min(), 4),
    "REL floor max, humans and calibrated machines (L = %d)" % _L7: (_ind7i.loc[_calib7, "REL_floor_if_calibrated"].max(), 4),
    "RES_i limit M4": (_ind7i.loc["M4", "RES_i_limit"], 3),
    "RES_i limit M5": (_ind7i.loc["M5", "RES_i_limit"], 3),
    "RES_i honest spread M4 (L = %d)" % _L7: (_ind7i.loc["M4", "RES_i_honest_spread"], 3),
    "RES_i honest spread M5 (L = %d)" % _L7: (_ind7i.loc["M5", "RES_i_honest_spread"], 3),
}
assert sorted(_tn7.index) == sorted(_tn7_expect)
for _q, (_v, _dp) in _tn7_expect.items():
    assert abs(float(_tn7.loc[_q, "value"]) - float(_v)) < 1e-15
    assert float(_tn7.loc[_q, "printed"]) == round(float(_v), _dp)
# (7l) the outlier-rule sensitivity of rem:binning-sensitivity: the width-
#      weighted outer-bin reading leaves every total unchanged, keeps M4/M5 in
#      the (-, +) cell at every L, and holds the H3-H5 resolution-dominant
#      split at none of the four L; the stored channels are recomputed
_Ls7 = [int(v) for v in str(_lcfg["binning_sensitivity_L"]).split(",")]
assert np.allclose(_bs7["delta_cal_ww_outer"] + _bs7["delta_res_ww_outer"], _bs7["delta_total"], atol=1e-12)
for _L in _Ls7:
    _zL = _norm_ppf(np.arange(1, _L + 1) / (_L + 1.0))
    _mF, _sF = _mua7["full panel"], float(_SD7 @ _wf7)
    _XF = _mF[:, None] + _sF * _zL[None, :]
    _cF7, _rF7, _pF7 = _hersbach_X(_XF, _y7)
    _oF7 = _hm_outer_terms(_XF, _y7)
    for _, _srow in _slog.iterrows():
        if not str(_srow["solve"]).startswith("leave-out"):
            continue
        _k = str(_srow["solve"]).split()[-1]
        _sidx = [int(v) for v in str(_srow["experts"]).split(",")]
        _sw = np.array([float(v) for v in str(_srow["weights"]).split(",")])
        _XI = _mua7[str(_srow["solve"])][:, None] + float(_SD7[_sidx] @ _sw) * _zL[None, :]
        _cI7, _rI7, _pI7 = _hersbach_X(_XI, _y7)
        _oI7 = _hm_outer_terms(_XI, _y7)
        _row = _bs7[(_bs7["id"] == _k) & (_bs7["L"] == _L)].iloc[0]
        assert abs((_rI7 - _oI7[0] + _oI7[2]) - (_rF7 - _oF7[0] + _oF7[2]) - float(_row["delta_cal_ww_outer"])) < 1e-12
        assert abs((_pI7 - _oI7[1]) - (_pF7 - _oF7[1]) - float(_row["delta_res_ww_outer"])) < 1e-12
        # the outer terms of _hersbach_X are exactly the outlier rule restated
        assert abs(_oI7[0] + _oI7[1] - _oI7[2]) < 1e-12


def _Lset7(col_cal, col_res, ks, cond):
    return [L for L in _Ls7 if all(cond(float(_bs7[(_bs7["id"] == k) & (_bs7["L"] == L)][col_cal].iloc[0]),
                                        float(_bs7[(_bs7["id"] == k) & (_bs7["L"] == L)][col_res].iloc[0]))
                                   for k in ks)]


_ww_h = _Lset7("delta_cal_ww_outer", "delta_res_ww_outer", ("H3", "H4", "H5"), lambda c, r: r > c)
_ww_m = _Lset7("delta_cal_ww_outer", "delta_res_ww_outer", ("M4", "M5"), lambda c, r: c < 0 and r > 0)
assert _ww_h == [] and str(_lcfg["L_with_H3H5_chiefly_resolution_ww_outer"]) == "none"
assert _ww_m == _Ls7 and str(_lcfg["L_with_M4M5_in_minus_plus_cell_ww_outer"]) == ",".join(str(L) for L in _Ls7)
print("   ex:hmpanel claims: error dependence (N_eff humans %.2f, machines %.2f, panel %.2f), LOO of M1-M3 and"
      " M4/M5, individual RES_i on the grid, honest spread and limit, reliability floor, outlier-rule"
      " sensitivity" % (_dv7("N_eff equal weights human"), _dv7("N_eff equal weights machine"),
                        _dv7("N_eff equal weights panel")))

# (7m) Figure 5.4: the stipulated matrix is a valid correlation matrix, the
#      config summaries are the realized block ranges, the lineage block has
#      mean 0.85, the effective numbers follow from the stored matrix, and the
#      correlation among the humans decides which way adding a block moves them
_dep7w = _wb("fig5-4_depnet.xlsx")
_dids = list(_dep7w["nodes"]["id"].astype(str))
_dgrp = list(_dep7w["nodes"]["group"].astype(str))
# the Figure 5.4 node IDs never reuse the ex:hmpanel expert IDs (a different panel)
assert not (set(_dids) & ({"H%d" % _k for _k in range(1, 6)} | {"M%d" % _k for _k in range(1, 6)}))
assert _dids == ["A%d" % _k for _k in range(1, 6)] + ["V%d" % _k for _k in range(1, 6)]
assert list(_dep7w["correlation"]["id"].astype(str)) == _dids
_R7d = _dep7w["correlation"][_dids].to_numpy(float)
assert np.allclose(_R7d, _R7d.T) and np.allclose(np.diag(_R7d), 1.0)
assert np.linalg.eigvalsh(_R7d).min() > 0
assert abs(np.linalg.eigvalsh(_R7d).min() - float(_dcfg7["min_eigenvalue"])) < 1e-12
_dh = [i for i, g in enumerate(_dgrp) if g == "human"]
_dl7 = [i for i, g in enumerate(_dgrp) if g == "lineage"]
_hh7 = np.array([_R7d[i, j] for i in _dh for j in _dh if i < j])
_hl7 = np.array([_R7d[i, j] for i in _dh for j in _dl7])
_ll7 = np.array([_R7d[i, j] for i in _dl7 for j in _dl7 if i < j])
assert abs(_ll7.mean() - float(_dcfg7["lineage_mean_corr"])) < 1e-3
assert (_hh7 >= 0.05).all() and (_hh7 < 0.20).all() and (_hl7 >= 0.02).all() and (_hl7 < 0.10).all()
for _nm, _v in (("human_block", _hh7), ("cross_block", _hl7), ("lineage_block", _ll7)):
    assert abs(float(_dcfg7[_nm + "_min"]) - _v.min()) < 1e-12 and abs(float(_dcfg7[_nm + "_max"]) - _v.max()) < 1e-12
for _nm, _v in (("human_block", _hh7), ("cross_block", _hl7)):
    assert str(_dcfg7[_nm + "_printed"]) == "%.2f to %.2f" % (_v.min(), _v.max())
assert ("%.2f to %.2f" % (_hh7.min(), _hh7.max())) in str(_dcfg7["human_note"])
assert _ll7.min() > _hh7.max() and _hh7.max() < 0.5
assert abs(1.0 / (np.full(10, 0.1) @ _R7d @ np.full(10, 0.1)) - float(_dcfg7["neff_panel_equalweight"])) < 1e-12
assert abs(1.0 / (np.full(5, 0.2) @ _R7d[np.ix_(_dl7, _dl7)] @ np.full(5, 0.2)) - float(_dcfg7["neff_lineage_equalweight"])) < 1e-12
# the comparison Example ex:corr-lineage draws: at equal weights the stipulated
# humans alone count for more independent experts than the humans joined by
# the lineage, whereas in ex:hmpanel adding the machines raises the count
_nh7d = 1.0 / (np.full(5, 0.2) @ _R7d[np.ix_(_dh, _dh)] @ np.full(5, 0.2))
assert abs(_nh7d - float(_dcfg7["neff_human_equalweight"])) < 1e-12
assert float(_dcfg7["neff_human_printed"]) == round(_nh7d, 2)
assert float(_dcfg7["neff_panel_printed"]) == round(float(_dcfg7["neff_panel_equalweight"]), 2)
assert float(_dcfg7["neff_panel_equalweight"]) < _nh7d
assert _dv7("N_eff equal weights panel") > _dv7("N_eff equal weights human")
# the lineage is less correlated with its humans than the ex:hmpanel machines
# are with theirs, so the cross-block level does not separate the examples
assert _hl7.max() < _dv7("error r human-machine min")
# the reason the text gives for the opposite directions of the count: the
# correlation among the humans that each block joins.  Exchanging only the
# within-human blocks of the two matrices (the ex:hmpanel error correlation and
# the stipulated matrix) reverses both directions; exchanging only the added
# blocks (machines, lineage) or only the cross blocks reverses neither
assert _dh == _H7 == list(range(5)) and _dl7 == _M7 == list(range(5, 10)), \
    "the humans are not the first five and the added block the last five of both matrices"


def _neq7m(R, idx):
    _w = np.full(len(idx), 1.0 / len(idx))
    return 1.0 / float(_w @ R[np.ix_(idx, idx)] @ _w)


def _up7m(R):
    """Adding the block to the humans at equal weights raises the effective number."""
    return _neq7m(R, list(range(10))) > _neq7m(R, _H7)


assert _up7m(_Rerr7) and not _up7m(_R7d)
for _blk7m in ((_H7, _H7), (_M7, _M7), (_H7, _M7)):
    _X7m, _Y7m = _Rerr7.copy(), _R7d.copy()
    for _a7m, _b7m in (_blk7m, _blk7m[::-1]):
        _X7m[np.ix_(_a7m, _b7m)], _Y7m[np.ix_(_a7m, _b7m)] = _R7d[np.ix_(_a7m, _b7m)], _Rerr7[np.ix_(_a7m, _b7m)]
    _flip7m = _blk7m == (_H7, _H7)
    assert _up7m(_X7m) == (not _flip7m) and _up7m(_Y7m) == _flip7m, "the human block does not decide the direction"
print("   Figure 5.4: stipulated matrix valid (min eigenvalue %.3f), block ranges and N_eff reproduced;"
      " humans %.2f -> panel %.2f with the lineage; the within-human blocks decide the direction"
      % (float(_dcfg7["min_eigenvalue"]), _nh7d, float(_dcfg7["neff_panel_equalweight"])))

# (7n) the group leave-outs Section 5.5 cites: every re-solve certified and
#      reproduced from its stored weights; the three calibrated machines
#      together cost little, the two overconfident machines more, all five
#      the whole subpanel gain; each group of machines supplies most of what
#      the other offers, and each calibrated machine's leave-one-out total is
#      below half of what it adds to the humans alone; H3, the most
#      informative human, has the largest total of the panel
_gl7 = _loo["group_loo"].set_index("removed_set")
_gnames7 = ["calibrated machines M1-M3", "overconfident machines M4, M5", "all machines M1-M5",
            "all machines but M1 (humans + M1)", "all machines but M2 (humans + M2)",
            "all machines but M3 (humans + M3)"]
assert list(_gl7.index) == _gnames7
_z7g = _norm_ppf(np.arange(1, _L7 + 1) / (_L7 + 1.0))
_chf7g, _relf7g, _potf7g = _hersbach(_mua7["full panel"], np.full(_T7, float(_SD7 @ _wf7)), _y7, _z7g)
for _gn, _gr in _gl7.iterrows():
    _gidx = [int(v) for v in str(_gr["experts"]).split(",")]
    _gw7 = np.array([float(v) for v in str(_gr["weights"]).split(",")])
    _rem7 = [k for k in _ids7 if _ids7.index(k) not in _gidx]
    assert ",".join(_rem7) == str(_gr["removed"]), _gn
    _sums_to_one(_gw7, "group leave-out " + _gn, tol=1e-9)
    assert (_gw7 >= 0).all()
    _Pg7 = _hm_program(_y7, _MU7, _SD7, _gidx, _lamG7, _lamC7)
    _Lg7, _cg7, _, _ = _Pg7["terms"](_gw7)
    assert abs(_Lg7 - float(_gr["objective_alg42"])) < 1e-9 and abs(_cg7 - float(_gr["mean_crps"])) < 1e-9
    assert float(_gr["objective_gap"]) <= float(_HM_VERIFY["obj_gap"])
    assert float(_gr["weights_inf_gap_vs_slsqp"]) <= float(_HM_VERIFY["w_gap"])
    assert float(_gr["alg42_last_dw_inf"]) < float(_lcfg["solver_tolerance"])
    assert int(_lcfg["solver_r_min"]) <= int(_gr["alg42_iterations"]) < int(_lcfg["solver_r_max_cap"])
    assert abs(float(_gr["delta_vs_full"]) - (_cg7 - _cf7)) < 1e-9
    assert abs(float(_gr["gain_vs_human_only"]) - (_ch7 - _cg7)) < 1e-9
    _chg7, _relg7, _potg7 = _hersbach(_MU7[:, _gidx] @ _gw7, np.full(_T7, float(_SD7[_gidx] @ _gw7)), _y7, _z7g)
    assert abs(_chg7 - _chf7g - float(_gr["delta_step_L%d" % _L7])) < 1e-12
    assert abs(_relg7 - _relf7g - float(_gr["delta_cal_L%d" % _L7])) < 1e-12
    assert abs(_potg7 - _potf7g - float(_gr["delta_res_L%d" % _L7])) < 1e-12
    # (read_excel returns the numeric-looking text cell as a float)
    assert float(_gr["printed_delta_vs_full"]) == float(_hm_sig2(float(_gr["delta_vs_full"])))
_gd7 = {k: float(_gl7.loc[k, "delta_vs_full"]) for k in _gnames7}
_gg7 = {k: float(_gl7.loc[k, "gain_vs_human_only"]) for k in _gnames7}
_trio7, _oc7, _allm7 = _gnames7[:3]
# the all-machines row is the human-only solve, and its cost is the subpanel gain
assert str(_gl7.loc[_allm7, "solve"]) == "human-only" and abs(_gd7[_allm7] - _gain7) < 1e-12
assert 0 < _gd7[_trio7] < _gd7[_oc7] < _gd7[_allm7] and _gd7[_trio7] < 0.1 * _gd7[_allm7]
# removing the calibrated trio is the panel "humans + M4, M5", and removing
# M4, M5 is "humans + M1-M3": each group's cost is below half its value to the humans alone
assert _gd7[_trio7] < 0.5 * _gg7[_oc7] and _gd7[_oc7] < 0.5 * _gg7[_trio7]
for _k, _gn in zip(("M1", "M2", "M3"), _gnames7[3:]):
    assert _gg7[_gn] > 0
    assert float(_lr7.loc[_k, "delta_cont_defloo"]) < 0.5 * _gg7[_gn]
    assert _bs7.loc[_bs7["id"] == _k, "delta_total"].max() < 0.5 * _gg7[_gn]
# H3: largest total of the panel (exact score and every grid), most informative human
assert str(_lr7["delta_cont_defloo"].astype(float).idxmax()) == "H3"
for _L in _Ls7:
    _bL = _bs7[_bs7["L"] == _L].set_index("id")["delta_total"].astype(float)
    assert str(_bL.idxmax()) == "H3", (_L, _bL.idxmax())
for _col in ("RES_i", "RES_i_limit"):
    assert str(_ind7i.loc[_hum7, _col].astype(float).idxmax()) == "H3"
print("   group leave-outs: M1-M3 %.4f, M4/M5 %.4f, all machines %.4f; each group substitutes for the"
      " other; H3 the largest total" % (_gd7[_trio7], _gd7[_oc7], _gd7[_allm7]))

# (8) Case IV: invariants of the SOLVED outputs, re-read from the workbook.
#     Nothing here is pinned to a typed result constant: the panel must
#     regenerate from the recorded seed and draw order, the reported weights
#     must lie on the simplex, reproduce the stored means and satisfy the
#     KKT form of the nearest-to-uniform projection, the Gini / advocate-share
#     / Ubar columns must follow from the stored weights and means, the
#     priorities must be the ranks of the means, the swap pair must be the
#     contested station (derived from the stations sheet) and its consensus
#     neighbour, the multi-start SQP must agree with the 1-D certificate,
#     and the robustness sheets must be complete and consistent.  (8g)-(8o)
#     test the statements of Section 6.5 on the written sheets: the level-set
#     ranges of the advocates' share (vertex enumeration), the overlaps that
#     the level sets fix and the overlap Gini beside the mean overlap, the
#     consensus optimum as one member's mean, every near-tied local maximum
#     (relative gap) and every combination of near-tied optima, the
#     leveling-down record, the swap margin against its re-seed spread, the early-warning
#     threshold, and the realized panel behind the largest uncontested gap.
_rtw = _wb("fig6-11_rt_case.xlsx")
_rtd = _cfgd(_rtw)
_rtn, _rtk = int(_rtd["n_stakeholders"]), int(_rtd["n_stations"])
assert _rtw["panel_mu"].shape == (_rtn, _rtk)
assert len(_rtw["panel"]) == _rtn
_rtmu = _rtw["panel_mu"].to_numpy(float)
_rtr = _rtw["results"].reset_index(drop=True)
_rtids = list(_rtr["station"].astype(str))
assert _rtids == list(_rtw["panel_mu"].columns.astype(str)) == list(_rtw["stations"]["id"].astype(str))
_rtadv = (_rtw["panel"]["constituency"].astype(str) == "accessibility advocates").to_numpy()
assert int(_rtadv.sum()) == int(_rtw["constituencies"].set_index("key").loc["adv", "size"])
# (8a) the panel regenerates from the recorded seed under the recorded draw order
_rtchk, _ = _rt_generate_panel(np.random.default_rng(int(_rtd["seed"])))
assert float(np.abs(_rtchk - _rtmu).max()) < 1e-12, "Case IV panel does not regenerate from the seed"
assert (_rtmu >= float(_rtd["scale_min"])).all() and (_rtmu <= float(_rtd["scale_max"])).all()
assert int(_rtd["generator_clip_binding"]) == int((_rtmu <= 0).any() or (_rtmu >= 10).any())
# (8b) the SQP starts are 1 + K points of the simplex, the first the uniform vector
_rts = _rtw["sqp_starts"].to_numpy(float)
assert _rts.shape == (1 + _RT_K_STARTS, _rtn) and (_rts >= 0).all()
assert np.allclose(_rts.sum(axis=1), 1.0, atol=1e-9) and np.allclose(_rts[0], 1.0 / _rtn)
# (8c) per regime: reported weights on the simplex, reproducing the stored
#      mean, with the diagnostics recomputed from them, and the KKT form of
#      the nearest-to-uniform projection (positive entries affine in mu,
#      zero entries where that affine form is nonpositive)
for _reg in ("eff", "eq"):
    _W = _rtw["weights_" + _reg].to_numpy(float)
    assert _W.shape == (_rtn, _rtk) and (_W >= -1e-12).all()
    for _j in range(_rtk):
        _w = np.clip(_W[:, _j], 0.0, None)
        _m = float(_rtr.loc[_j, "mean_" + _reg])
        _mus = _rtmu[:, _j]
        _sums_to_one(_w, "Case IV %s weights at %s" % (_reg, _rtids[_j]), tol=1e-8)
        assert _mus.min() - 1e-9 <= _m <= _mus.max() + 1e-9
        assert abs(float(_w @ _mus) - _m) < 1e-7, "Case IV %s: weights do not reproduce the mean at %s" % (_reg, _rtids[_j])
        assert abs(_gini_vec(_w) - float(_rtr.loc[_j, "weight_gini_" + _reg])) < 1e-8
        assert abs(float(_w[_rtadv].sum()) - float(_rtr.loc[_j, "adv_share_" + _reg])) < 1e-8
        _pos = _w > 1e-9
        assert _pos.sum() >= 2
        _cb = np.linalg.lstsq(np.column_stack([np.ones(_pos.sum()), _mus[_pos]]), _w[_pos], rcond=None)[0]
        assert float(np.abs(_cb[0] + _cb[1] * _mus[_pos] - _w[_pos]).max()) < 1e-6, \
            "Case IV %s at %s: reported weights are not the nearest-to-uniform point of the level set" % (_reg, _rtids[_j])
        assert (_cb[0] + _cb[1] * _mus[~_pos] <= 1e-6).all()
        # utility level and equity from the closed-form overlaps at the stored mean
        _A = _rt_overlap(_mus, _m)[0]
        _, _Ut = _fs_phi(_A[None, :], float(_rtd["rho"]), float(_rtd["theta"]))
        assert abs(float(_Ut[0].mean()) - float(_rtr.loc[_j, "Ubar_" + _reg])) < 1e-9
        assert abs(_gini_vec(_Ut[0]) - float(_rtr.loc[_j, "gini_U_" + _reg])) < 1e-9
        assert abs(float(_A.mean()) - float(_rtr.loc[_j, "Abar_" + _reg])) < 1e-9
        _regime = "eff" if _reg == "eff" else "eq4"
        assert abs(float(_rt_objective([_m], _mus, _regime)[0]) - float(_rtr.loc[_j, "objective_" + _reg])) < 1e-9
    _avg = float(np.mean(_rtr["weight_gini_" + _reg]))
    assert abs(_avg - float(_rtd["weight_gini_%s_avg" % _reg])) < 1e-9
# (8d) priorities are the ranks of the means; the swap pair is the contested
#      station of the stations sheet and its neighbour in the consensus queue
for _reg in ("eff", "eq"):
    _rank = np.argsort(np.argsort(_rtr["mean_" + _reg].to_numpy(float))) + 1
    assert list(_rank) == [int(v) for v in _rtr["priority_" + _reg]], "Case IV %s priorities are not the ranks" % _reg
_rswap = [s for s, a, b in zip(_rtids, _rtr["priority_eff"], _rtr["priority_eq"]) if int(a) != int(b)]
assert _rswap == [t.strip() for t in str(_rtd["swapped"]).split(",")], _rswap
_rtsp = _rtw["stations"]
_rtspread = (_rtsp[["ops", "pas", "adv"]].max(axis=1) - _rtsp[["ops", "pas", "adv"]].min(axis=1)).to_numpy(float)
_rtcont = str(_rtsp.loc[int(np.argmax(_rtspread)), "id"])
assert _rtcont == str(_rtd["contested_station"])
_rtq = list(_rtr.sort_values("priority_eff")["station"].astype(str))
_rtnb = _rtq[_rtq.index(_rtcont) - 1]
assert sorted(_rswap) == sorted([_rtcont, _rtnb]), "swap pair %r is not the contested station and its neighbour" % _rswap
_rt4 = _rtr.set_index("station")
assert float(_rt4.loc[_rtcont, "mean_eff"]) > float(_rt4.loc[_rtnb, "mean_eff"])
assert float(_rt4.loc[_rtcont, "mean_eq"]) < float(_rt4.loc[_rtnb, "mean_eq"])
assert float(_rt4.loc[_rtcont, "adv_share_eq"]) > float(_rt4.loc[_rtcont, "adv_share_eff"])
_rtgaps = {s: abs(float(_rt4.loc[s, "mean_eff"]) - float(_rt4.loc[s, "mean_eq"])) for s in _rtids if s != _rtcont}
assert len(_rtgaps) == _rtk - 1
assert abs(float(_rtd["max_uncontested_mean_gap"]) - max(_rtgaps.values())) < 1e-9
assert str(_rtd["max_uncontested_mean_gap_station"]) == max(_rtgaps, key=_rtgaps.get)
assert str(_rtd["uncontested_gaps_below_threshold"]).startswith(
    "%d/%d" % (sum(1 for g in _rtgaps.values() if g < _RT_GAP_THRESHOLD), len(_rtgaps)))
# (8e) the multi-start SQP agrees with the 1-D certificate at every solve
_rtc = _rtw["certification"]
assert len(_rtc) == 2 * _rtk
assert (np.abs(_rtc["objective_sqp"] - _rtc["objective_scan"]) <= _RT_CERT_OBJ_TOL).all()
assert (np.abs(_rtc["m_sqp"] - _rtc["m_scan"]) <= _RT_CERT_M_TOL).all()
assert (_rtc["m_sqp"].round(_RT_PRINT_DP) == _rtc["m_scan"].round(_RT_PRINT_DP)).all(), \
    "SQP and scan means print differently at %d dp" % _RT_PRINT_DP
assert (_rtc["projection_kkt_maxdiff"] < 1e-6).all()
for _, _cr in _rtc.iterrows():
    _col = "eff" if _cr["regime"] == "eff" else "eq"
    assert abs(float(_cr["m_sqp"]) - float(_rt4.loc[str(_cr["station"]), "mean_" + _col])) < 1e-12
# (8f) robustness sheets are complete and consistent with the config
_rsd = _rtw["robustness_seeds"]
assert len(_rsd) == _RT_RESEED_N and list(_rsd["seed"]) == list(range(int(_rtd["seed"]) + 1, int(_rtd["seed"]) + 1 + _RT_RESEED_N))
assert str(_rtd["robustness_reseed_swap_direction"]) == "%d/%d" % (int(_rsd["swap_direction_preserved"].sum()), _RT_RESEED_N)
assert str(_rtd["robustness_reseed_other_orders_unchanged"]) == "%d/%d" % (int(_rsd["other_orders_unchanged"].sum()), _RT_RESEED_N)
# every re-seeded panel either keeps every other order or moves the contested station
# exactly two places with its two neighbours keeping theirs
assert (((_rsd["other_orders_unchanged"] == 1) & (_rsd["contested_places_moved"] == 1))
        | ((_rsd["other_orders_unchanged"] == 0) & (_rsd["contested_places_moved"] == 2) & (_rsd["S5_S6_order_unchanged"] == 1))).all()
assert str(_rtd["robustness_reseed_contested_moves_two_places"]).startswith(
    "%d/%d" % (int((_rsd["contested_places_moved"] == 2).sum()), _RT_RESEED_N))
# the claims of Section 6.5 and Table tab:robustness on the re-seeded panels: the
# direction of the swap is preserved in every one; where no other pair changes
# order, only the contested station and its consensus neighbour change priority;
# where the contested station moves two places, the stations that change priority
# are exactly it, its neighbour and the station next ahead in the consensus queue,
# the contested mean passes both of theirs and those two keep their order, so the
# uncontested stations keep their order among themselves in every re-seeded panel
_rtnx = _rtq[_rtq.index(_rtcont) - 2]                 # the station next ahead in the consensus queue
assert (_rsd["swap_direction_preserved"] == 1).all(), "a re-seeded panel reverses the swap"
assert ((_rsd["contested_places_moved"] == 1) | (_rsd["contested_places_moved"] == 2)).all()
for _, _sr in _rsd.iterrows():
    _chg = set(str(_sr["stations_reordered"]).split(","))
    if int(_sr["contested_places_moved"]) == 1:
        assert _chg == {_rtcont, _rtnb}, "seed %d: stations %s change priority, not the swap pair alone" % (
            int(_sr["seed"]), sorted(_chg))
        assert float(_sr["S5_mean_eff"]) < float(_sr["S4_mean_eff"]) and float(_sr["S4_mean_eq"]) < float(_sr["S5_mean_eq"])
    else:
        assert _chg == {_rtcont, _rtnb, _rtnx}, "seed %d: the contested station moves two places but %s change priority" % (
            int(_sr["seed"]), sorted(_chg))
        # consensus queue ... S6, S5, S4 ...; equity queue ... S4, S6, S5 ...
        assert float(_sr["S6_mean_eff"]) < float(_sr["S5_mean_eff"]) < float(_sr["S4_mean_eff"]), int(_sr["seed"])
        assert float(_sr["S4_mean_eq"]) < float(_sr["S6_mean_eq"]) < float(_sr["S5_mean_eq"]), int(_sr["seed"])
        assert int(_sr["S5_S6_order_unchanged"]) == 1
_rfs = _rtw["robustness_fs"]
assert len(_rfs) == sum(1 for a in _RT_FS_RHO for b in _RT_FS_THETA if b <= a)
assert (_rfs["theta"] <= _rfs["rho"]).all()
assert str(_rtd["robustness_fs_swap_direction"]) == "%d/%d" % (int(_rfs["swap_direction_preserved"].sum()), len(_rfs))
_rfs0 = _rfs[(np.abs(_rfs["rho"] - float(_rtd["rho"])) < 1e-12) & (np.abs(_rfs["theta"] - float(_rtd["theta"])) < 1e-12)]
assert len(_rfs0) == 1
assert abs(float(_rfs0["S4_mean_eq"].iloc[0]) - float(_rt4.loc[_rtcont, "mean_eq"])) <= _RT_CERT_M_TOL
assert abs(float(_rfs0["S5_mean_eq"].iloc[0]) - float(_rt4.loc[_rtnb, "mean_eq"])) <= _RT_CERT_M_TOL
# the claims on the Fehr-Schmidt grid: the swap is preserved under every setting, and
# only the swap pair changes priority, so the order of the nine uncontested stations
# is unchanged under every one; the advocates' reported share at the contested station
# is the projection of the stored mean, and its range over the grid is the recorded
# (and printed) one
assert (_rfs["swap_direction_preserved"] == 1).all(), "a Fehr-Schmidt setting reverses the swap"
assert (_rfs["other_orders_unchanged"] == 1).all(), "a Fehr-Schmidt setting reorders an uncontested pair"
assert all(set(str(s).split(",")) == {_rtcont, _rtnb} for s in _rfs["stations_reordered"]), \
    "a Fehr-Schmidt setting changes the priority of a station other than the swap pair"
_rtjc = _rtids.index(_rtcont)
for _, _fr in _rfs.iterrows():
    _wfs, _ = _rt_project_uniform(_rtmu[:, _rtjc], float(_fr["S4_mean_eq"]))
    assert abs(float(_wfs[_rtadv].sum()) - float(_fr["S4_adv_share_eq"])) < 1e-8
    assert abs(_gini_vec(_wfs) - float(_fr["S4_weight_gini_eq"])) < 1e-8
assert str(_rtd["robustness_fs_S4_adv_share_range"]) == "%.4f .. %.4f" % (
    float(_rfs["S4_adv_share_eq"].min()), float(_rfs["S4_adv_share_eq"].max())), \
    "the recorded range of the contested advocate share over the Fehr-Schmidt grid is not the sheet's"
# (8g) metadata: the workbook names the case and section of the book
assert str(_rtd["case"]).startswith("IV -- ") and "(Section 6.5)" in str(_rtd["case"])
# (8h) the uncontested-gap statement of the text: the largest gap at one
#      station, every other uncontested gap below the recorded threshold
_rtthr = float(_rtd["gap_threshold"])
assert all(g < _rtthr for s, g in _rtgaps.items() if s != str(_rtd["max_uncontested_mean_gap_station"]))
assert str(_rtd["uncontested_gaps_below_threshold"]).startswith("%d/%d" % (len(_rtgaps) - 1, len(_rtgaps)))
assert abs(float(_rtd["contested_mean_gap"]) - abs(float(_rt4.loc[_rtcont, "mean_eff"]) - float(_rt4.loc[_rtcont, "mean_eq"]))) < 1e-12
# (8i) what the optimal level sets fix: the LP range of the advocates' share,
#      recomputed by enumerating the vertices of {w in simplex : w.mu = m*};
#      the reported share (tie-break representative) lies inside it
_rthead = float(_rtadv.mean())
assert abs(_rthead - float(_rtd["adv_headcount_share"])) < 1e-12
for _reg in ("eff", "eq"):
    for _j in range(_rtk):
        _mus, _m = _rtmu[:, _j], float(_rtr.loc[_j, "mean_" + _reg])
        _vx = [float(_rtadv[i]) for i in range(_rtn) if abs(_mus[i] - _m) <= 1e-9]
        for _a in range(_rtn):
            for _b in range(_rtn):
                if _mus[_a] < _m < _mus[_b]:
                    _t = (_mus[_b] - _m) / (_mus[_b] - _mus[_a])
                    _vx.append(float(_t * _rtadv[_a] + (1.0 - _t) * _rtadv[_b]))
        _lo, _hi = float(_rtr.loc[_j, "adv_share_min_" + _reg]), float(_rtr.loc[_j, "adv_share_max_" + _reg])
        assert abs(min(_vx) - _lo) < 1e-7 and abs(max(_vx) - _hi) < 1e-7, \
            "Case IV %s at %s: stored level-set share range is not the LP range" % (_reg, _rtids[_j])
        assert _lo - 1e-9 <= float(_rtr.loc[_j, "adv_share_" + _reg]) <= _hi + 1e-9
_rtcap = [s for s in _rtids if float(_rt4.loc[s, "adv_share_max_eff"]) < _rthead]
assert _rtcap == [_rtcont] and str(_rtd["consensus_caps_advocates_below_headcount_at"]) == _rtcont
assert str(_rtd["equity_floors_advocates_above_zero_at"]) == ",".join(
    s for s in _rtids if float(_rt4.loc[s, "adv_share_min_eq"]) > 1e-9)
assert float(_rt4.loc[_rtcont, "adv_share_min_eq"]) > 0.0
_rtcr = (max(float(_rt4.loc[_rtcont, "adv_share_min_eff"]), float(_rt4.loc[_rtcont, "adv_share_min_eq"])),
         min(float(_rt4.loc[_rtcont, "adv_share_max_eff"]), float(_rt4.loc[_rtcont, "adv_share_max_eq"])))
assert _rtcr[0] < _rtcr[1] and str(_rtd["contested_adv_share_common_range"]) == "%.4f .. %.4f" % _rtcr
for _reg in ("eff", "eq"):
    assert str(_rtd["contested_adv_share_range_" + _reg]) == "%.4f .. %.4f" % (
        float(_rt4.loc[_rtcont, "adv_share_min_" + _reg]), float(_rt4.loc[_rtcont, "adv_share_max_" + _reg]))
# the representatives' ratio the text calls "more than fifteenfold"
_rtratio = float(_rt4.loc[_rtcont, "adv_share_eq"]) / float(_rt4.loc[_rtcont, "adv_share_eff"])
assert abs(_rtratio - float(_rtd["contested_adv_share_ratio"])) < 1e-9 and _rtratio > 15.0
# (8j) the level-set properties the text states at the contested station, all
#      functions of the stored mean alone: the advocates' mean overlap and the
#      smallest overlap rise, the mean overlap falls, the utility Gini falls
#      and the mean utility rises under the equity regime
for _reg in ("eff", "eq"):
    for _j in range(_rtk):
        _A = _rt_overlap(_rtmu[:, _j], float(_rtr.loc[_j, "mean_" + _reg]))[0]
        _, _Ut = _fs_phi(_A[None, :], float(_rtd["rho"]), float(_rtd["theta"]))
        assert abs(float(_A[_rtadv].mean()) - float(_rtr.loc[_j, "Abar_adv_" + _reg])) < 1e-9
        assert abs(float(_A.min()) - float(_rtr.loc[_j, "Amin_" + _reg])) < 1e-9
        assert abs(float(_Ut[0].min()) - float(_rtr.loc[_j, "Umin_" + _reg])) < 1e-9
        assert abs(_gini_vec(_A) - float(_rtr.loc[_j, "gini_A_" + _reg])) < 1e-9
_rtC = _rt4.loc[_rtcont]
assert _rtC["Abar_adv_eq"] > _rtC["Abar_adv_eff"] and _rtC["Amin_eq"] > _rtC["Amin_eff"]
assert _rtC["Abar_eq"] < _rtC["Abar_eff"] and _rtC["gini_U_eq"] < _rtC["gini_U_eff"] and _rtC["Ubar_eq"] > _rtC["Ubar_eff"]
# in the overlaps: the equity optimum lowers the mean overlap at every station
# (the consensus optimum maximizes it), and at the contested station it also
# lowers the Gini of the overlaps, the leveling-down case of
# prop:consensus-fairness(iii), while the smallest overlap rises
assert (_rtr["Abar_eq"] < _rtr["Abar_eff"]).all()
assert _rtC["gini_A_eq"] < _rtC["gini_A_eff"]
_rtod = [s for s in _rtids if float(_rt4.loc[s, "gini_A_eq"]) < float(_rt4.loc[s, "gini_A_eff"])
         and float(_rt4.loc[s, "Abar_eq"]) < float(_rt4.loc[s, "Abar_eff"])]
assert str(_rtd["overlap_leveling_down_at"]) == ",".join(_rtod) and _rtcont in _rtod
_rtox = [s for s in _rtids if s not in _rtod]
assert len(_rtox) == 1 and str(_rtd["overlap_leveling_down_except"]) == _rtox[0]
_rtdA = {s: float(_rt4.loc[s, "Abar_eff"]) - float(_rt4.loc[s, "Abar_eq"]) for s in _rtids if s != _rtcont}
_s = max(_rtdA, key=_rtdA.get)
assert _s == str(_rtd["uncontested_max_mean_overlap_fall_station"])
assert abs(_rtdA[_s] - float(_rtd["uncontested_max_mean_overlap_fall"])) < 1e-12
# (8k) the consensus optimum is one stakeholder's own mean, the member with
#      the greatest summed overlap with the panel
for _j in range(_rtk):
    _mus = _rtmu[:, _j]
    _k = int(_rtr.loc[_j, "consensus_member"]) - 1
    assert abs(float(_rtr.loc[_j, "mean_eff"]) - _mus[_k]) < 1e-6, "consensus optimum is not a member's mean at %s" % _rtids[_j]
    _tot = _rt_overlap(_mus, _mus).sum(axis=1)          # summed overlap of each member with the panel
    assert int(np.argmax(_tot)) == _k
    assert str(_rtr.loc[_j, "consensus_member_constituency"]) == str(_rtw["panel"]["constituency"].iloc[_k])
# the text: at the contested station that member is an operator
assert str(_rt4.loc[_rtcont, "consensus_member_constituency"]) == "operators"
# (8l) near-ties: every solve re-scanned, EVERY other local maximum compared
#      with the optimum by its relative gap against the recorded tolerance,
#      each near-tied maximum's representative recomputed by the projection,
#      and every combination of reported and near-tied optima re-evaluated
#      for the statements the text makes
_rttol = float(_rtd["near_tie_tolerance"])
_rtnt = _rtw["near_ties"]
_rtexp, _rtnon = [], []                   # expected near_ties rows; non-tie relative gaps
_rtchoices = []                           # per tied solve: the near-tied m and weight Gini
for _, _cr in _rtc.iterrows():
    _j = _rtids.index(str(_cr["station"]))
    _vs = float(_cr["objective_scan"])
    _oth = _rt_scan(_rtmu[:, _j], str(_cr["regime"]), runner_up=True)[3]
    assert [o[1] for o in _oth] == sorted((o[1] for o in _oth), reverse=True)
    if not _oth:
        assert not np.isfinite(float(_cr["runner_up_m"])) and int(_cr["near_tie"]) == 0
        continue
    assert abs(_oth[0][0] - float(_cr["runner_up_m"])) < 1e-9 and abs(_oth[0][1] - float(_cr["runner_up_objective"])) < 1e-12
    assert abs(_vs - _oth[0][1] - float(_cr["runner_up_gap"])) < 1e-12
    assert abs((_vs - _oth[0][1]) / _vs - float(_cr["runner_up_rel_gap"])) < 1e-12
    _tied = [o for o in _oth if (_vs - o[1]) / _vs < _rttol]
    _rtnon += [(_vs - o[1]) / _vs for o in _oth if (_vs - o[1]) / _vs >= _rttol]
    assert int(_cr["near_tie"]) == int(len(_tied) > 0) and int(_cr["n_near_tie_maxima"]) == len(_tied)
    if len(_oth) > len(_tied):
        assert abs((_vs - _oth[len(_tied)][1]) / _vs - float(_cr["first_non_tie_rel_gap"])) < 1e-12
    if _tied:
        _rtchoices.append((str(_cr["station"]), str(_cr["regime"]), []))
    for _rank, (_mt, _vt) in enumerate(_tied, 1):
        _wt, _ = _rt_project_uniform(_rtmu[:, _j], _mt)
        _rtexp.append((str(_cr["station"]), str(_cr["regime"]), _rank, _mt, _vs - _vt, (_vs - _vt) / _vs,
                       float(_wt[_rtadv].sum()), _gini_vec(_wt)))
        _rtchoices[-1][2].append((_mt, _gini_vec(_wt)))
assert len(_rtnt) == len(_rtexp), "near_ties sheet does not list every near-tied local maximum"
for (_, _nr), _e in zip(_rtnt.iterrows(), _rtexp):
    assert (str(_nr["station"]), str(_nr["regime"]), int(_nr["rank"])) == _e[:3]
    _col = "eff" if _e[1] == "eff" else "eq"
    assert abs(float(_nr["m_reported"]) - float(_rt4.loc[_e[0], "mean_" + _col])) < 1e-12
    assert abs(float(_nr["m_local_max"]) - _e[3]) < 1e-9
    assert abs(float(_nr["objective_gap"]) - _e[4]) < 1e-12 and abs(float(_nr["relative_gap"]) - _e[5]) < 1e-12
    assert abs(float(_nr["adv_share_local_max"]) - _e[6]) < 1e-8 and abs(float(_nr["weight_gini_local_max"]) - _e[7]) < 1e-8
    assert abs(float(_nr["weight_gini_reported"]) - float(_rt4.loc[_e[0], "weight_gini_" + _col])) < 1e-12
assert str(_rtd["near_ties"]).startswith("%d/%d solves, %d local maxima" % (len(_rtchoices), len(_rtc), len(_rtexp)))
assert abs(min(_rtnon) - float(_rtd["near_tie_smallest_non_tie_rel_gap"])) < 1e-12 and min(_rtnon) >= _rttol
_rtcb = _rtw["near_tie_combinations"]
assert len(_rtcb) == int(np.prod([1 + len(c[2]) for c in _rtchoices])) == int(_rtd["near_tie_combinations_n"])
_rtwg = {"eff": [], "eq": []}
_rtlv, _rtlo = [], []                     # per combination: leveling-down set, count of mean-overlap falls
for _pick, (_, _cbr) in zip(_rt_itertools.product(*[range(1 + len(c[2])) for c in _rtchoices]), _rtcb.iterrows()):
    assert str(_cbr["pick"]).zfill(len(_pick)) == "".join(str(p) for p in _pick)
    _mm = {k: list(_rtr["mean_" + k].astype(float)) for k in ("eff", "eq")}
    _gg = {k: list(_rtr["weight_gini_" + k].astype(float)) for k in ("eff", "eq")}
    for (_st, _rg, _ch), _p in zip(_rtchoices, _pick):
        if _p:
            _col = "eff" if _rg == "eff" else "eq"
            _j = _rtids.index(_st)
            _mm[_col][_j], _gg[_col][_j] = _ch[_p - 1]
    for _k in ("eff", "eq"):
        assert list(np.argsort(np.argsort(_mm[_k])) + 1) == [int(v) for v in _rtr["priority_" + _k]], \
            "a near-tied local maximum changes the %s priorities" % _k
        _rtwg[_k].append(float(np.mean(_gg[_k])))
    _ug = {s: abs(_mm["eff"][j] - _mm["eq"][j]) for j, s in enumerate(_rtids) if s != _rtcont}
    assert max(_ug, key=_ug.get) == str(_rtd["max_uncontested_mean_gap_station"])
    assert all(g < _rtthr for s, g in _ug.items() if s != str(_rtd["max_uncontested_mean_gap_station"]))
    # the overlap classification of prop:consensus-fairness(iii) (Gini index)
    # on this combination, recomputed from the panel
    _cmb_lo, _cmb_lv = [], []
    for _j, _s in enumerate(_rtids):
        _Ae = _rt_overlap(_rtmu[:, _j], _mm["eff"][_j])[0]
        _Aq = _rt_overlap(_rtmu[:, _j], _mm["eq"][_j])[0]
        if _Aq.mean() < _Ae.mean():
            _cmb_lo.append(_s)
            if _gini_vec(_Aq) < _gini_vec(_Ae):
                _cmb_lv.append(_s)
    assert int(_cbr["n_mean_overlap_lower"]) == len(_cmb_lo) and int(_cbr["n_overlap_leveling_down"]) == len(_cmb_lv)
    assert ("" if pd.isna(_cbr["overlap_leveling_down_at"]) else str(_cbr["overlap_leveling_down_at"])) == ",".join(_cmb_lv)
    _rtlv.append(set(_cmb_lv))
    _rtlo.append(len(_cmb_lo))
assert (_rtcb["priorities_unchanged"] == 1).all() and (_rtcb["others_below_threshold"] == 1).all()
# the reported pick reproduces the classification of the results sheet; the
# text: "seven to nine" stations, the contested one always, the reported
# exception never, and the stations named as turning on the choice
assert _rtlv[0] == set(str(_rtd["overlap_leveling_down_at"]).split(","))
assert str(_rtd["overlap_leveling_down_combo_count_range"]) == "%d .. %d" % (
    min(len(s) for s in _rtlv), max(len(s) for s in _rtlv))
_rtla = [s for s in _rtids if all(s in d for d in _rtlv)]
_rtln = [s for s in _rtids if all(s not in d for d in _rtlv)]
_rtlx = [s for s in _rtids if s not in _rtla and s not in _rtln]
assert str(_rtd["overlap_leveling_down_combo_always"]) == ",".join(_rtla) and _rtcont in _rtla
assert str(_rtd["overlap_leveling_down_combo_never"]) == ",".join(_rtln) == str(_rtd["overlap_leveling_down_except"])
assert str(_rtd["overlap_leveling_down_combo_varies"]) == ",".join(_rtlx)
assert str(_rtd["mean_overlap_lower_combo_count_range"]) == "%d .. %d" % (min(_rtlo), max(_rtlo))
assert max(_rtlo) == _rtk
for _k in ("eff", "eq"):
    assert str(_rtd["near_tie_weight_gini_%s_avg_range" % _k]) == "%.4f .. %.4f" % (min(_rtwg[_k]), max(_rtwg[_k]))
assert str(_rtd["near_tie_priorities_unchanged"]) == "%d/%d" % (len(_rtcb), len(_rtcb))
# the closest near-tie the text prints (smallest relative gap): EVERY near-tied
# local maximum of that solve below the recorded bound, their means as
# recorded, and the runner-up's representative recomputed by the projection
_rtcl = _rtnt.loc[_rtnt["relative_gap"].astype(float).idxmin()]
assert int(_rtcl["rank"]) == 1
assert "%s %s" % (_rtcl["station"], _rtcl["regime"]) == str(_rtd["near_tie_closest"])
_rtcls = _rtnt[(_rtnt["station"].astype(str) == str(_rtcl["station"])) & (_rtnt["regime"].astype(str) == str(_rtcl["regime"]))]
assert len(_rtcls) == int(_rtd["near_tie_closest_n_maxima"]) == max(len(c[2]) for c in _rtchoices) == 2
# the text: the closest near-tie is at "the one solve with two of them"
assert [c[0] + " " + c[1] for c in _rtchoices if len(c[2]) > 1] == [str(_rtd["near_tie_closest"])] \
    == [str(_rtd["near_tie_solves_with_several_maxima"])]
assert (_rtcls["relative_gap"].astype(float) < float(_rtd["near_tie_closest_gap_bound"])).all()
assert abs(float(_rtcls["relative_gap"].max()) - float(_rtd["near_tie_closest_max_rel_gap"])) < 1e-15
assert str(_rtd["near_tie_closest_m_local_maxima"]) == ", ".join("%.6f" % float(v) for v in _rtcls["m_local_max"])
assert abs(float(_rtcl["objective_gap"]) - float(_rtd["near_tie_closest_gap"])) < 1e-15
assert abs(float(_rtcl["relative_gap"]) - float(_rtd["near_tie_closest_rel_gap"])) < 1e-15
_rtj = _rtids.index(str(_rtcl["station"]))
_rtw2, _ = _rt_project_uniform(_rtmu[:, _rtj], float(_rtcl["m_local_max"]))
assert abs(float(_rtw2[_rtadv].sum()) - float(_rtd["near_tie_closest_runner_up_adv_share"])) < 1e-8
assert abs(_gini_vec(_rtw2) - float(_rtd["near_tie_closest_runner_up_weight_gini"])) < 1e-8
assert abs(float(_rtcl["m_reported"]) - float(_rtd["near_tie_closest_m_reported"])) < 1e-12
assert abs(float(_rtcl["m_local_max"]) - float(_rtd["near_tie_closest_m_runner_up"])) < 1e-12
# (8m) leveling-down record: every uncontested mean utility moves by less than
#      0.01, the largest utility-Gini change is the recorded one, and wherever
#      the mean utility falls the smallest utility on the panel rises
_rtdU = {s: float(_rt4.loc[s, "Ubar_eq"]) - float(_rt4.loc[s, "Ubar_eff"]) for s in _rtids if s != _rtcont}
_rtdG = {s: float(_rt4.loc[s, "gini_U_eq"]) - float(_rt4.loc[s, "gini_U_eff"]) for s in _rtids if s != _rtcont}
_s = max(_rtdU, key=lambda s: abs(_rtdU[s]))
assert abs(_rtdU[_s]) < 0.01 and _s == str(_rtd["uncontested_max_abs_mean_utility_change_station"])
assert abs(abs(_rtdU[_s]) - float(_rtd["uncontested_max_abs_mean_utility_change"])) < 1e-12
_s = max(_rtdG, key=lambda s: abs(_rtdG[s]))
assert _s == str(_rtd["uncontested_max_abs_utility_gini_change_station"])
assert abs(abs(_rtdG[_s]) - float(_rtd["uncontested_max_abs_utility_gini_change"])) < 1e-12
_rtfall = [s for s in _rtids if float(_rt4.loc[s, "Ubar_eq"]) < float(_rt4.loc[s, "Ubar_eff"])]
assert ",".join(_rtfall) == str(_rtd["mean_utility_falls_at"]).replace("nan", "")
assert all(float(_rt4.loc[s, "Umin_eq"]) > float(_rt4.loc[s, "Umin_eff"]) for s in _rtfall)
# (8n) the swap margin against its seed-to-seed spread, and the early-warning
#      threshold over the main panel and every re-seeded one
_rtmg = float(_rt4.loc[_rtnb, "mean_eq"]) - float(_rt4.loc[_rtcont, "mean_eq"])
assert abs(_rtmg - float(_rtd["eq_swap_margin"])) < 1e-12
assert np.allclose(_rsd["eq_swap_margin"], _rsd["S5_mean_eq"] - _rsd["S4_mean_eq"], atol=1e-12, rtol=0)
assert np.allclose(_rsd["contested_gap"], np.abs(_rsd["S4_mean_eff"] - _rsd["S4_mean_eq"]), atol=1e-12, rtol=0)
_rtsd = float(np.std(_rsd["eq_swap_margin"].to_numpy(float), ddof=1))
assert abs(_rtsd - float(_rtd["reseed_eq_swap_margin_sd"])) < 1e-12
assert abs(float(_rsd["eq_swap_margin"].min()) - float(_rtd["reseed_eq_swap_margin_min"])) < 1e-12
assert _rtmg > 3.0 * _rtsd, "the swap margin is not more than three re-seed standard deviations"
_rtwt = float(_rtd["warn_threshold"])
_rtwc = [float(_rtd["contested_mean_gap"])] + list(_rsd["contested_gap"].astype(float))
_rtwu = [float(_rtd["max_uncontested_mean_gap"])] + list(_rsd["max_uncontested_gap"].astype(float))
assert all(a > _rtwt > b for a, b in zip(_rtwc, _rtwu)), "the early-warning threshold does not separate every panel"
assert abs(min(_rtwc) - float(_rtd["warn_contested_gap_min"])) < 1e-12
assert abs(max(_rtwu) - float(_rtd["warn_uncontested_gap_max"])) < 1e-12
# (8o) the realized panel behind the largest uncontested gap: recomputed from
#      the panel, the leniencies and the generator means; its two lowest
#      judgments are advocates'
_rtrp = _rtw["realized_panel"].set_index("station")
_rtlen = _rtw["panel"]["leniency"].to_numpy(float)
_rtkey = _rtw["panel"]["constituency"].map(
    dict(zip(_rtw["constituencies"]["constituency"].astype(str), _rtw["constituencies"]["key"].astype(str)))).to_numpy()
for _j, _s in enumerate(_rtids):
    _x = _rtmu[:, _j]
    _o = np.argsort(_x)
    _base = np.array([float(_rtsp.loc[_j, k]) for k in _rtkey])
    assert abs(float(_rtrp.loc[_s, "panel_mean"]) - float(_x.mean())) < 1e-12
    assert abs(float(_rtrp.loc[_s, "regime_gap"]) - abs(float(_rt4.loc[_s, "mean_eff"]) - float(_rt4.loc[_s, "mean_eq"]))) < 1e-12
    for _q in (0, 1):
        _i = int(_o[_q])
        assert int(_rtrp.loc[_s, "low%d_expert" % (_q + 1)]) == _i + 1
        assert abs(float(_rtrp.loc[_s, "low%d_below_panel_mean" % (_q + 1)]) - float(_x.mean() - _x[_i])) < 1e-12
        assert abs(float(_rtrp.loc[_s, "low%d_noise" % (_q + 1)]) - float(_x[_i] - _base[_i] - _rtlen[_i])) < 1e-12
_rtS9 = str(_rtd["max_uncontested_mean_gap_station"])
assert all(_rtadv[int(_rtrp.loc[_rtS9, "low%d_expert" % q]) - 1] for q in (1, 2)), \
    "the two lowest judgments at the largest uncontested gap are not both advocates'"
assert float(_rt4.loc[_rtS9, "mean_eq"]) < float(_rt4.loc[_rtS9, "mean_eff"])      # the equity optimum moves toward them
# the equity near-tie whose local maxima the text prints has a similar low
# tail: its two lowest judgments are advocates' too, each within 0.05 of the
# corresponding distance below the panel mean at the largest-gap station
_rtwide = _rtnt[_rtnt["regime"].astype(str) == "eq4"]
_rtwide = _rtwide.loc[(_rtwide["m_local_max"] - _rtwide["m_reported"]).abs().idxmax()]
assert str(_rtwide["station"]) == str(_rtd["near_tie_equity_widest_station"])
assert abs(float(_rtwide["m_local_max"]) - float(_rtd["near_tie_equity_widest_runner_up_m"])) < 1e-12
_rtW = str(_rtwide["station"])
assert all(_rtadv[int(_rtrp.loc[_rtW, "low%d_expert" % q]) - 1] for q in (1, 2))
assert all(abs(float(_rtrp.loc[_rtW, "low%d_below_panel_mean" % q]) - float(_rtrp.loc[_rtS9, "low%d_below_panel_mean" % q])) < 0.05
           for q in (1, 2))
print("   Case IV: panel regenerates from the seed; reported weights on the simplex, on the"
      " level sets and nearest to uniform (KKT); diagnostics recomputed; priorities are ranks;"
      " swap = contested station %s and neighbour %s; SQP == 1-D certificate at %d solves;"
      " robustness sheets consistent; level-set share ranges = vertex enumeration, consensus"
      " caps the advocates below headcount only at %s; consensus optimum = a member's mean;"
      " %d near-tied solves with %d near-tied local maxima (relative gap), %d combinations keep"
      " every priority; overlap leveling-down (Gini) at %d..%d stations over the combinations;"
      " overlap and utility leveling-down, swap-margin, early-warning and"
      " realized-panel records recomputed"
      % (_rtcont, _rtnb, len(_rtc), _rtcont, len(_rtchoices), len(_rtnt), len(_rtcb),
         min(len(s) for s in _rtlv), max(len(s) for s in _rtlv)))

# (9) Figures 3.1, 4.3 and 4.4, re-read from their workbooks.
# (9a) Figure 3.1: the Lorenz profile is the Case I delta = 0 solution's
#      normalized utilities, and the Gini and Hoover values the text prints
#      are the pairwise/relative-mean-deviation formulas of tab:swf and, the
#      same numbers, twice the Lorenz area and the largest Lorenz gap.
_lzw = _wb("fig3-1_lorenz_gini_hoover.xlsx")
_lzp, _lzl = _lzw["profile"], _lzw["lorenz"]
_lzs = _cfgd(_lzw, "summary")
_lzwt = _lzp["weight_delta0"].to_numpy(float)
assert np.allclose(_lzwt, _bW[_i0], atol=1e-12), "fig3-1 weights are not the Case I delta = 0 solution"
_lzP, _lzU = _bim_Phi(_lzwt, return_util=True)
assert np.allclose(_lzp["utility_U"].to_numpy(float), _lzU, atol=1e-12), "fig3-1 utilities do not recompute"
assert np.allclose(_lzp["overlap_A"].to_numpy(float), _bim_overlaps(_lzwt), atol=1e-12)
assert abs(_lzP - float(_bser.loc[0.0, "fairness_Phibar"])) < 1e-12
_lzx = _lzl["x_share_of_experts"].to_numpy(float)
_lzL = _lzl["L_share_of_utility"].to_numpy(float)
assert np.allclose(_lzx, np.arange(_BIM_N + 1) / _BIM_N, atol=1e-15)
assert np.allclose(_lzL, np.concatenate([[0.0], np.cumsum(np.sort(_lzU)) / np.sum(_lzU)]), atol=1e-12)
assert np.all(np.diff(np.diff(_lzL)) >= -1e-12), "the Lorenz curve is not convex"
_lzG = _gini_vec(_lzU)
_lzH = float(np.abs(_lzU - _lzU.mean()).sum() / (2.0 * _BIM_N * _lzU.mean()))
assert abs(_lzG - (1.0 - np.sum(np.diff(_lzx) * (_lzL[1:] + _lzL[:-1])))) < 1e-12
assert abs(_lzH - np.max(_lzx - _lzL)) < 1e-12
for _key, _v in (("gini_pairwise", _lzG), ("gini_twice_area", _lzG), ("hoover_half_rmd", _lzH), ("hoover_max_gap", _lzH)):
    assert abs(float(_lzs[_key]) - _v) < 1e-12, "fig3-1 summary %s does not recompute" % _key
assert float(_lzs["gini_printed"]) == round(_lzG, 3) and float(_lzs["hoover_printed"]) == round(_lzH, 3)
# the reading the text gives: the least satisfied expert is the dissenter, and
# the largest gap falls after the first 1/N of the panel, the dissenter's share
assert int(_lzs["min_utility_expert"]) == _BIM_DISSENTER + 1 == int(np.argmin(_lzU)) + 1
assert int(_lzs["hoover_at_k"]) == 1 == int(np.argmax(_lzx - _lzL))
assert int(np.sum(_lzU < _lzU.mean())) == 1, "more than the dissenter lies below the mean utility"
assert np.all(np.diff((_lzx - _lzL)[1:]) < 0), "the Lorenz gap does not close steadily after the dissenter"
assert abs(float(_lzs["min_utility"]) - _lzU.min()) < 1e-12 and abs(float(_lzs["second_lowest_utility"]) - np.sort(_lzU)[1]) < 1e-12
assert abs(float(_lzs["max_utility"]) - _lzU.max()) < 1e-12 and abs(float(_lzs["mean_utility"]) - _lzU.mean()) < 1e-12
assert str(_cfgd(_lzw)["seed"]).startswith("none")

# (9b) Figure 4.4: the epsilon-constraint front (grid and refinement floors),
#      the weighted sums, the concave majorant of both, the stretch and its
#      ends.  The stored objective values are first recomputed from the
#      stored weights with the Case I functions; the tolerance is 1e-8
#      because the closed-form overlap _bim_overlap loses about 1e-10 when
#      the aggregate's spread is within ~1e-12 of an expert's (the 2-5
#      mixtures of this front have spread 1.3 up to rounding) and the
#      workbook round trip moves the weights by an ulp.  Everything else is
#      checked on the stored values themselves.
_pw = _wb("fig4-4_pareto_front.xlsx")
_pf, _pws, _pss = _pw["front"], _pw["weighted_sum"], _cfgd(_pw, "summary")
_ppr, _pcfg, _phl = _cfgd(_pw, "printed"), _cfgd(_pw), _pw["hull"]
_pb = _pf["floor_b"].to_numpy(float)
_pkind = _pf["floor_kind"].astype(str).to_numpy()
_pW = _pf[["w%d" % (i + 1) for i in range(_BIM_N)]].to_numpy(float)
_pA = _pf["Abar"].to_numpy(float)
_pP = _pf["Phi"].to_numpy(float)
assert np.allclose([_bim_Abar(w) for w in _pW], _pA, rtol=0, atol=1e-8), "front Abar does not recompute from its weights"
assert np.allclose([_bim_Phi(w) for w in _pW], _pP, rtol=0, atol=1e-8), "front Phi does not recompute from its weights"
# the floors: the equally spaced grid, and the refinement of the two grid intervals next to the deepest grid point
_pgb = _pb[_pkind == "grid"]
assert len(_pgb) == _PAR_N_FLOORS and np.allclose(np.diff(_pgb), _pgb[1] - _pgb[0], rtol=0, atol=1e-12), "grid floors not equally spaced"
assert abs(_pgb[0] - _bim_Abar(np.asarray(_bim_bounds["Phi"]["argmax"]))) < 1e-12 and abs(_pgb[-1] - float(_bbd.loc["Abar", "max"])) < 1e-12
_pk0 = int(_pss["grid_deepest_index_before_refinement"])
assert 0 < _pk0 < _PAR_N_FLOORS - 1
_prb = [_pgb[_lo] + (_pgb[_lo + 1] - _pgb[_lo]) * _t / _PAR_REFINE_SPLIT
        for _lo in (_pk0 - 1, _pk0) for _t in range(1, _PAR_REFINE_SPLIT)]
assert np.allclose(_pb[_pkind == "refinement"], _prb, rtol=0, atol=1e-15), "refinement floors are not the stated split"
assert set(_pkind) == {"grid", "refinement"} and np.all(np.diff(_pb) > 0), "floors not merged in ascending order"
for _k in range(len(_pb)):
    _sums_to_one(_pW[_k], "front w(b=%.5f)" % _pb[_k])
    assert (_pW[_k] >= 0).all()
assert (_pA >= _pb - _PAR_FEAS_TOL).all(), "a front point violates its consensus floor"
assert np.all(np.diff(_pP) <= 0.0), "front fairness rises with the consensus floor"
assert abs(_pP[0] - float(_bbd.loc["Phi", "max"])) < 1e-9 and abs(_pA[-1] - float(_bbd.loc["Abar", "max"])) < 1e-9
assert (_pf["spread_across_starts"].to_numpy(float) < 1e-6).all(), "the SLSQP starts disagree at some floor"
# weighted sums: recomputed from their weights
_plam = _pws["lambda_consensus"].to_numpy(float)
_pwsW = _pws[["w%d" % (i + 1) for i in range(_BIM_N)]].to_numpy(float)
_pwsA = _pws["Abar"].to_numpy(float)
_pwsP = _pws["Phi"].to_numpy(float)
assert np.allclose([_bim_Abar(w) for w in _pwsW], _pwsA, rtol=0, atol=1e-8)
assert np.allclose([_bim_Phi(w) for w in _pwsW], _pwsP, rtol=0, atol=1e-8)
assert len(_plam) == _PAR_N_LAMBDA and np.allclose(_plam, np.linspace(0.0, 1.0, _PAR_N_LAMBDA), rtol=0, atol=1e-15)
# the majorant is the upper hull of ALL the efficient points, epsilon and weighted sum, and every
# weighted-sum solution (a supported point) lies on it, none below it by more than 1e-9
_phv = _par_majorant(_pA, _pP, _pwsA, _pwsP)
assert [(v[0], v[1]) for v in _phv] == list(zip(_phl["source"].astype(str), _phl["row"].astype(int))), \
    "stored hull is not the upper hull of the efficient points"
_phx, _phy = _phl["Abar"].to_numpy(float), _phl["Phi"].to_numpy(float)
assert np.allclose(_phx, [v[2] for v in _phv], rtol=0, atol=0) and np.allclose(_phy, [v[3] for v in _phv], rtol=0, atol=0)
_pg = np.interp(_pA, _phx, _phy) - _pP
_pwg = np.interp(_pwsA, _phx, _phy) - _pwsP
assert (_pg >= -1e-12).all() and np.allclose(_pg, _pf["gap_below_majorant"].to_numpy(float), rtol=0, atol=1e-12)
assert (_pwg >= -1e-12).all() and np.allclose(_pwg, _pws["gap_below_majorant"].to_numpy(float), rtol=0, atol=1e-12)
assert _pwg.max() < 1e-9, "a weighted-sum solution lies below the majorant of the efficient points"
# the stretch: the hull edge over the deepest point, its ends two consecutive hull vertices
_pvlo, _pvhi = int(_pss["unsupported_lo_hull_vertex"]), int(_pss["unsupported_hi_hull_vertex"])
assert _pvhi == _pvlo + 1
_palo, _pahi = _phx[_pvlo], _phx[_pvhi]
_pkd = int(np.argmax(_pg))
assert int(_pss["max_gap_index"]) == _pkd and _palo < _pA[_pkd] < _pahi
assert abs(float(_pss["max_gap"]) - _pg.max()) < 1e-12 and _pg.max() > 100 * _PAR_GAP_TOL
_pin = (_pA > _palo) & (_pA < _pahi)
assert (_pg[_pin] > _PAR_GAP_TOL).all(), "a point inside the stretch touches the majorant"
assert list(_pf["in_main_stretch"].to_numpy(int)) == [int(v) for v in _pin]
assert _pgb[_pk0 - 1] < _pA[_pkd] < _pgb[_pk0 + 1], "the deepest point lies outside the refined interval"
# the bridges: every hull edge that skips an epsilon point more than _PAR_GAP_TOL below it.  The
# stretch is one of them.  The text mentions the others, printing only their count, as further
# dents far shallower than the stretch that adjoin its ends: one shares the stretch's lower hull
# vertex and one its upper, and each is less than _PAR_DENT_RATIO as deep as the stretch
_pbr = []
for _i in range(len(_phx) - 1):
    _ins = (_pA > _phx[_i]) & (_pA < _phx[_i + 1])
    if _ins.any() and _pg[_ins].max() > _PAR_GAP_TOL:
        _pbr.append((_i, float(_pg[_ins].max())))
_pbs = _pw["bridges"]
assert list(_pbs["from_hull_vertex"].astype(int)) == [_i for _i, _ in _pbr], "stored bridges do not recompute"
assert np.allclose(_pbs["max_gap"].to_numpy(float), [_g for _, _g in _pbr], rtol=0, atol=1e-12)
assert list(_pbs["main"].astype(int)) == [int(_i == _pvlo) for _i, _ in _pbr]
_poth = [(_i, _g) for _i, _g in _pbr if _i != _pvlo]
assert sorted(_i for _i, _ in _poth) == [_pvlo - 1, _pvhi], "the other dents do not adjoin the two ends of the stretch"
assert max(_g for _, _g in _poth) < _PAR_DENT_RATIO * _pg.max(), "another dent is not far shallower than the stretch"
assert int(_pss["n_bridging_edges"]) == len(_pbr) and abs(float(_pss["other_dips_max_gap"]) - max(_g for _, _g in _poth)) < 1e-12
assert float(_pcfg["shallow_dent_ratio"]) == _PAR_DENT_RATIO
for _key, _v in (("unsupported_Abar_lo", _palo), ("unsupported_Abar_hi", _pahi),
                 ("unsupported_Phi_at_lo", _phy[_pvlo]), ("unsupported_Phi_at_hi", _phy[_pvhi]),
                 ("front_Abar_min", _pA.min()), ("front_Abar_max", _pA.max()),
                 ("front_Phi_min", _pP.min()), ("front_Phi_max", _pP.max()),
                 ("max_gap_Abar", _pA[_pkd])):
    assert abs(float(_pss[_key]) - _v) < 1e-12, "fig4-4 summary %s does not recompute" % _key
_psn = ((_phy[_pvhi] - _phy[_pvlo]) / (_bim_Pmax - _bim_Pmin)) / ((_pahi - _palo) / (_bim_Amax - _bim_Amin))
_plst = -_psn / (1.0 - _psn)
assert abs(float(_pss["lambda_star"]) - _plst) < 1e-12
# weighted sums: optimal against every front point, never above the epsilon front at
# their own consensus, none inside the stretch, and the ones on either side of lambda*
# are the two ends of the stretch
_pAt = (_pA - _bim_Amin) / (_bim_Amax - _bim_Amin)
_pPt = (_pP - _bim_Pmin) / (_bim_Pmax - _bim_Pmin)
for _j, _lam in enumerate(_plam):
    _v = _lam * (_pwsA[_j] - _bim_Amin) / (_bim_Amax - _bim_Amin) + (1.0 - _lam) * (_pwsP[_j] - _bim_Pmin) / (_bim_Pmax - _bim_Pmin)
    assert abs(_v - float(_pws.loc[_j, "objective"])) < 1e-12
    assert np.max(_lam * _pAt + (1.0 - _lam) * _pPt) <= _v + 1e-9, "a front point beats weighted sum lambda=%.2f" % _lam
    _kf = int(np.searchsorted(_pb, _pwsA[_j] + _PAR_FEAS_TOL)) - 1
    assert _pwsP[_j] <= _pP[_kf] + 1e-9, "weighted sum lambda=%.2f lies above the epsilon front" % _lam
_pleft = _pwsA <= _palo + 1e-7
_pright = _pwsA >= _pahi - 1e-7
assert (_pleft | _pright).all(), "a weighted sum lands in the stretch"
assert list(_pws["inside_unsupported_stretch"].to_numpy(int)) == [0] * len(_plam) and int(_pss["ws_inside_stretch"]) == 0
_pjl = int(np.max(np.where(_pleft)[0]))
_pjr = _pjl + 1
assert _pleft[:_pjr].all() and _pright[_pjr:].all(), "the weighted-sum solutions cross the stretch more than once"
assert _plam[_pjl] < _plst < _plam[_pjr]
assert max(abs(_pwsA[_pjl] - _palo), abs(_pwsP[_pjl] - _phy[_pvlo]),
           abs(_pwsA[_pjr] - _pahi), abs(_pwsP[_pjr] - _phy[_pvhi])) < 1e-7, \
    "the stretch ends are not the weighted sums on either side of lambda*"
assert abs(float(_pss["ws_last_left_lambda"]) - _plam[_pjl]) < 1e-12 and abs(float(_pss["ws_first_right_lambda"]) - _plam[_pjr]) < 1e-12
assert abs(float(_pss["ws_last_left_Abar"]) - _pwsA[_pjl]) < 1e-12 and abs(float(_pss["ws_first_right_Abar"]) - _pwsA[_pjr]) < 1e-12
assert int(_pss["ws_at_fairness_maximizer"]) == int(np.sum(np.abs(_pwsA - _pA[0]) < 1e-6))
# the text says the FIRST weightings (smallest lambda) are the ones returning the fairness maximizer
assert list(np.where(np.abs(_pwsA - _pA[0]) < 1e-6)[0]) == list(range(int(_pss["ws_at_fairness_maximizer"])))
# the Case I delta-sweep solutions lie on or below the consensus-fairness front
for _k, _d in enumerate(_bdel):
    _a, _p = _bim_Abar(_bW[_k]), _bim_Phi(_bW[_k])
    _kf = int(np.searchsorted(_pb, _a + _PAR_FEAS_TOL)) - 1
    assert _p <= (_pP[_kf] if _kf >= 0 else _pP[0]) + 1e-9, "Case I delta=%.1f lies above the front" % _d
# the printed values are the roundings of the recomputed ones
for _key, _v in (("Abar_min_3dp", round(_pA.min(), 3)), ("Abar_max_3dp", round(_pA.max(), 3)),
                 ("Phi_max_3dp", round(_pP.max(), 3)), ("Phi_min_3dp", round(_pP.min(), 3)),
                 ("fairness_range_3dp", round(_pP.max() - _pP.min(), 3)),
                 ("stretch_lo_3dp", round(_palo, 3)), ("stretch_hi_3dp", round(_pahi, 3)),
                 ("max_gap_4dp", round(_pg.max(), 4)), ("max_gap_Abar_3dp", round(_pA[_pkd], 3)),
                 ("other_dents", len(_poth)), ("lambda_star_3dp", round(_plst, 3)),
                 ("ws_last_left_lambda_2dp", round(float(_plam[_pjl]), 2)),
                 ("ws_first_right_lambda_2dp", round(float(_plam[_pjr]), 2)),
                 ("ws_solutions", len(_plam)), ("ws_at_fairness_maximizer", int(_pss["ws_at_fairness_maximizer"])),
                 ("grid_floors", len(_pgb)), ("refinement_floors", len(_prb))):
    assert float(_ppr[_key]) == _v, "fig4-4 printed %s = %r, recomputed %r" % (_key, _ppr[_key], _v)
for _key in ("seed", "rng", "draw_order", "floors", "starts_per_floor", "solver", "unsupported_rule",
             "weighted_sum", "majorant", "method", "refine_split", "confidence_term", "other_dents",
             "shallow_dent_ratio"):
    assert _key in _pcfg, "fig4-4 config lacks %r" % _key
assert int(_pcfg["seed"]) == NEW_SEED and int(_pcfg["refine_split"]) == _PAR_REFINE_SPLIT

# (9c) Figure 4.3: the series are the stated closed-form counts, no draws
_scw = _wb("fig4-3_scalability.xlsx")
# the illustrative epigraph-size count of Section 4: every count follows from the
# stated N, T, L by its formula, and the printed form is the value with thousands commas
_lpw = _scw["lp_reformulation_size"].set_index("quantity")
_lpv = {k: int(v) for k, v in _lpw["value"].items()}
assert (_lpv["N"], _lpv["T"], _lpv["L"], _lpv["N_large"]) == (30, 40, 19, 1000)
assert _lpv["pinball_variables_TL"] == _lpv["T"] * _lpv["L"]
assert _lpv["pair_variables"] == _lpv["N"] * (_lpv["N"] - 1) // 2
assert _lpv["variables"] == _lpv["N"] + _lpv["pinball_variables_TL"] + _lpv["pair_variables"]
assert _lpv["pinball_rows"] == 2 * _lpv["pinball_variables_TL"] and _lpv["pair_rows"] == 2 * _lpv["pair_variables"]
assert _lpv["pinball_nonzeros"] == _lpv["pinball_rows"] * _lpv["nonzeros_per_pinball_row"] \
    and _lpv["nonzeros_per_pinball_row"] == _lpv["N"] + 1
assert _lpv["pair_variables_at_N_large"] == _lpv["N_large"] * (_lpv["N_large"] - 1) // 2

def _lp_txt(x):
    # a comma-free printed form ('30') may come back from read_excel as a number
    return str(int(x)) if isinstance(x, (int, float, np.integer, np.floating)) and float(x) == int(x) else str(x)


assert all(_lp_txt(_lpw.loc[k, "printed"]) == "{:,}".format(v) for k, v in _lpv.items())
_scs = _scw["series"]
assert str(_cfgd(_scw)["seed"]).startswith("none")
for _, _r in _scs.iterrows():
    _n = int(_r["N"])
    _kk, _hh = _hier_count(_n)
    assert int(_r["flat_pairwise"]) == _n * (_n - 1) and int(_r["K_star"]) == _kk == int(round((_n * _n / 2.0) ** (1.0 / 3.0)))
    assert int(_r["hier_pairwise"]) == _hh
assert int(_scs["N"].iloc[0]) == 10 and int(_scs["N"].iloc[-1]) == 10000
# the saving printed for N = 10,000 is the ratio of the last row's counts
_scsum = _cfgd(_scw, "summary")
_scr = int(_scs["flat_pairwise"].iloc[-1]) / float(_scs["hier_pairwise"].iloc[-1])
assert int(_scsum["N_max"]) == int(_scs["N"].iloc[-1])
assert abs(float(_scsum["ratio_flat_over_hier_at_N_max"]) - _scr) < 1e-12
assert int(_scsum["ratio_printed"]) == int(round(_scr))
print("   Figures 3.1, 4.3, 4.4: Lorenz profile = Case I delta=0 utilities, Gini = twice the Lorenz area,"
      " Hoover = largest gap; front recomputed from its weights, nonincreasing, majorant = upper hull of"
      " the epsilon points and weighted sums, every weighted sum on it, none in the unsupported stretch,"
      " whose ends are the weighted sums either side of lambda*; interaction counts = formulas")

# (10) Worked examples of Chapters 2 to 5 (sheet worked_examples of the
#      Figure 2.2, 3.1, 4.3 and 5.4 workbooks).  The sheets are complete and
#      consistent: one row per (example, quantity), every computed row within
#      1e-6 of its independent cross-check and printing the same, every printed
#      string the value at its decimals (half-up) or its exact fraction.  Every
#      computed value is then recomputed here from the stated inputs re-read from
#      the sheets, through the shared helpers where the builder has one (_fs_phi,
#      _gini_vec) and otherwise by a route of its own (numpy.linalg.solve, an
#      eigenvector, the coefficient-of-variation identity eq:bonferroni-cv, the
#      exponential of the mean logarithm, binomial coefficients, a numerical
#      minimization of the fuzzy C-means objective), and the claims the text
#      makes of each example are tested.
import itertools
from scipy.optimize import minimize_scalar as _wx_minsc4
_WX_COLS = ["chapter", "example", "section", "quantity", "role", "value", "cross_check", "decimals", "printed", "rule"]
_wxt = []
for _ch, _f in sorted(_WX_CHAPTER_WORKBOOK.items()):
    _s = _wb(_f)
    assert "worked_examples" in _s, "%s has no worked_examples sheet" % _f
    assert "no random draws" in str(_cfgd(_s).get("worked_examples", "")), "%s config does not describe the sheet" % _f
    # re-read with the text columns kept as text (read_excel would turn a printed '0.600' into 0.6)
    _t = pd.read_excel(os.path.join(DATA, _f), sheet_name="worked_examples", dtype={"decimals": str, "printed": str})
    assert list(_t.columns) == _WX_COLS, "%s: worked_examples columns changed" % _f
    assert (_t["chapter"].astype(int) == _ch).all()
    _wxt.append(_t)
_wxt = pd.concat(_wxt, ignore_index=True)
_wxk = list(zip(_wxt["example"].astype(str), _wxt["quantity"].astype(str)))
assert len(set(_wxk)) == len(_wxk), "a worked-example row is duplicated"
_wxr = {k: r for k, (_, r) in zip(_wxk, _wxt.iterrows())}


def _wx_cell(x):
    """A decimals or printed cell (read as text) as a string; an empty cell reads as NaN."""
    if isinstance(x, float) and math.isnan(x):
        return ""
    assert isinstance(x, str), "a text cell of worked_examples was read as %r" % (x,)
    return x


for _k, _r in _wxr.items():
    _v, _dec, _pr = float(_r["value"]), _wx_cell(_r["decimals"]), _wx_cell(_r["printed"])
    assert str(_r["role"]) in ("stated input", "computed"), _k
    if str(_r["role"]) == "computed":
        _cc = float(_r["cross_check"])
        assert math.isfinite(_cc) and abs(_v - _cc) < 1e-6, "%s: value %.12g, cross-check %.12g" % (_k, _v, _cc)
        if _dec not in ("", "fraction"):
            assert _wx_fmt(_cc, int(_dec)) == _pr, "%s: the cross-check prints differently" % (_k,)
    else:
        assert not math.isfinite(float(_r["cross_check"]))
    if _dec == "":
        assert _pr == "", "%s: printed without decimals" % (_k,)
    elif _dec == "fraction":
        assert abs(float(_WxF(_pr)) - _v) < 1e-12, "%s: the printed fraction %s is not the value" % (_k, _pr)
    else:
        assert _pr == _wx_fmt(_v, int(_dec)), "%s: printed %r, value %.12g at %s decimals" % (_k, _pr, _v, _dec)


def _wv(ex, q):
    return float(_wxr[(ex, q)]["value"])


def _wp(ex, q):
    return _wx_cell(_wxr[(ex, q)]["printed"])


def _wvs(ex, fmt, n):
    return np.array([_wv(ex, fmt % (i + 1)) for i in range(n)])


def _wx_same(ex, q, x, tol=1e-9):
    assert abs(_wv(ex, q) - float(x)) < tol, "%s / %s: stored %.12g, recomputed %.12g" % (ex, q, _wv(ex, q), float(x))


# Chapter 2: the generalized least-squares weights of two correlated experts
_e = "GLS weights of two correlated experts"
_s1, _s2, _rho = _wv(_e, "sigma_1"), _wv(_e, "sigma_2"), _wv(_e, "rho")
_rr = np.array([_wv(_e, "r_1"), _wv(_e, "r_2")])
_Sg = np.array([[_s1 ** 2, _rho * _s1 * _s2], [_rho * _s1 * _s2, _s2 ** 2]])
_x = np.linalg.solve(_Sg, np.ones(2))
_wg = _x / _x.sum()
for _q, _val in (("sigma_1^2", _s1 ** 2), ("sigma_2^2", _s2 ** 2), ("2 rho sigma_1 sigma_2", 2.0 * _Sg[0, 1]),
                 ("denominator", _Sg[0, 0] + _Sg[1, 1] - 2.0 * _Sg[0, 1]),
                 ("numerator of w_1", np.linalg.det(_Sg) * _x[0]), ("numerator of w_2", np.linalg.det(_Sg) * _x[1]),
                 ("w_1", _wg[0]), ("w_2", _wg[1]), ("sigma_1^2 sigma_2^2 (1 - rho^2)", np.linalg.det(_Sg)),
                 ("diffuse-prior posterior variance", 1.0 / _x.sum()), ("combined estimate", _wg @ _rr),
                 ("combined estimate (fraction)", _wg @ _rr), ("correlation at which w_2 turns negative", _s1 / _s2)):
    _wx_same(_e, _q, _val)
# the text: w_2 is negative because rho exceeds sigma_1 / sigma_2, the weights sum to one,
# the posterior is sharper than the better expert alone, the estimate lies below both reports
assert abs(_wg.sum() - 1.0) < 1e-12 and _wg[1] < 0.0 and _rho > _s1 / _s2
assert _wv(_e, "diffuse-prior posterior variance") < min(_s1, _s2) ** 2
assert _wv(_e, "combined estimate") < _rr.min()
# Chapter 2: the DeGroot trust matrix
_e = "DeGroot trust matrix"
_Wt = np.array([[_wv(_e, "W[%d,%d]" % (i + 1, j + 1)) for j in range(3)] for i in range(3)])
assert np.allclose(_Wt.sum(axis=1), 1.0, atol=1e-15) and (_Wt >= 0).all()
assert _Wt[0, 2] == 0.0 == _Wt[2, 0] and (_Wt[1] > 0).all(), "the trust pattern the text describes"
assert np.allclose(np.diag(_Wt), _Wt[1, 1], atol=0), "not every expert places expert 2's weight on their own judgment"
assert (np.linalg.matrix_power(_Wt, 2) > 0).all(), "the chain is not irreducible and aperiodic (W^2 not positive)"
_ev, _evec = np.linalg.eig(_Wt.T)
_pv = np.real(_evec[:, int(np.argmin(np.abs(_ev - 1.0)))])
_pv = _pv / _pv.sum()
_pis = _wvs(_e, "pi_%d", 3)
assert np.allclose(_pis, _pv, atol=1e-12) and np.allclose(_pis @ _Wt, _pis, atol=1e-15)
_wx_same(_e, "pi_2 / pi_1", _pv[1] / _pv[0])
assert abs(_pis[1] / _pis[0] - 2.0) < 1e-12 and abs(_pis[1] / _pis[2] - 2.0) < 1e-12, "the text: twice as heavily"
# Chapter 3: Example ex:bonferroni
_e = "ex:bonferroni"


def _wx_phi_cv(u):
    """eq:bonferroni-cv: Phi_11 = Ubar sqrt(1 - CV^2 / (N - 1)), CV with divisor N."""
    u = np.asarray(u, float)
    return u.mean() * math.sqrt(1.0 - (u.std() / u.mean()) ** 2 / (u.size - 1))


_uA, _uB = _wvs(_e, "U_A[%d]", 4), _wvs(_e, "U_B[%d]", 4)
_wx_same(_e, "Phi_11(A)", _wx_phi_cv(_uA))
_wx_same(_e, "Phi_11(B)", _wx_phi_cv(_uB))
_wx_same(_e, "arithmetic mean of A (and of B)", _uA.mean())
assert abs(_uA.mean() - _uB.mean()) < 1e-12, "the two profiles do not share the mean"
assert _wv(_e, "Phi_11(A)") < _uA.mean() - 1e-3 and abs(_wv(_e, "Phi_11(B)") - _uB.mean()) < 1e-12
_bonf_gap = _uA.mean() - _wv(_e, "Phi_11(A)")
# Chapter 3: Example ex:skewed-panel, recomputed with the builder's _fs_phi (kappa = 2 there)
_e = "ex:skewed-panel"
_A5, _r5, _t5 = _wvs(_e, "A[%d]", 5), _wvs(_e, "rho[%d]", 5), _wvs(_e, "vartheta[%d]", 5)
assert _wv(_e, "kappa") == 2.0, "_fs_phi applies the power normalization at kappa = 2"
_M5 = _A5[None, :] - _A5[:, None]
_U5 = _A5 - _r5 * np.clip(_M5, 0, None).sum(axis=1) / 4.0 - _t5 * np.clip(-_M5, 0, None).sum(axis=1) / 4.0
_P5, _N5 = _fs_phi(_A5[None, :], _r5, _t5)
for _i in range(5):
    _wx_same(_e, "U_FS[%d]" % (_i + 1), _U5[_i])
    _wx_same(_e, "normalized U_FS[%d]" % (_i + 1), _N5[0, _i])
_wx_same(_e, "arithmetic mean of the normalized profile", _N5[0].mean())
_wx_same(_e, "Phi_11", _P5[0])
_wx_same(_e, "Phi_11", _wx_phi_cv(_N5[0]))
# the text: admissible dispositions; the dominant expert proud, above its overlap and
# clipped to 1; the least represented expert (overlap 0.15) negative and clipped to 0;
# Phi well below the mean, by more than the gap of ex:bonferroni
assert (_r5 >= np.maximum(_t5, 0.0)).all() and _t5[0] < 0 and int(np.argmax(_A5)) == 0
# ... and the tail, ordered by falling overlap, increasingly jealous and sympathetic
assert (np.diff(_A5) < 0).all() and (np.diff(_r5) > 0).all() and (np.diff(_t5) > 0).all() and (_t5[1:] > 0).all()
assert _U5[0] > _A5[0] and _N5[0, 0] == 1.0
assert int(np.argmin(_A5)) == 4 and _U5[4] < 0 and _N5[0, 4] == 0.0
assert _N5[0].mean() - _P5[0] > _bonf_gap, "the skewed panel does not widen the gap of ex:bonferroni"
# Chapter 3: the three-expert welfare profile, its equity objectives and Atkinson index
_e = "three-expert welfare profile"
_A3 = _wvs(_e, "A[%d]", 3)
_rh, _th = _wv(_e, "rho (every expert)"), _wv(_e, "vartheta (every expert)")
_M3 = _A3[None, :] - _A3[:, None]
_U3 = _A3 - _rh * np.clip(_M3, 0, None).sum(axis=1) / 2.0 - _th * np.clip(-_M3, 0, None).sum(axis=1) / 2.0
_E3 = np.where(_U3 < 0, 0.0, np.where(_U3 > _A3, 1.0, (np.exp(_U3 / _A3) - 1.0) / (math.e - 1.0)))
_m3 = _E3.mean()
_I3 = {1: np.ptp(_E3) / _m3, 2: np.abs(_E3 - _m3).mean() / _m3, 3: _E3.std() / _m3, 4: _gini_vec(_E3),
       5: np.abs(_E3 - _m3).mean() / (2.0 * _m3)}
for _i in range(3):
    _wx_same(_e, "U_FS[%d]" % (_i + 1), _U3[_i])
    _wx_same(_e, "U_FS[%d] / A[%d]" % (_i + 1, _i + 1), _U3[_i] / _A3[_i])
    _wx_same(_e, "U[%d]" % (_i + 1), _E3[_i])
_wx_same(_e, "mean U", _m3)
_wx_same(_e, "minimum U", _E3.min())
for _k in range(1, 6):
    _wx_same(_e, "EQ_%d" % _k, -_I3[_k])
_wx_same(_e, "I_4", _I3[4])
_wx_same(_e, "W_{1,4}", _m3 - _wv(_e, "lambda") * _I3[4])
# the text: every utility below its overlap, the ratio falling from the best- to the
# least-represented expert, all five objectives negative, EQ_1 the most severe and EQ_4,
# EQ_5 the mildest, EQ_5 half of EQ_2, the minimum the third expert's; and the printed
# arithmetic 0.638 - 0.196 = 0.442 holds on the printed values
assert _rh > 0 and _th > 0 and _rh >= _th, "the text: moderately jealous and moderately sympathetic"
assert (_U3 < _A3).all() and (np.diff(_U3 / _A3) < 0).all()
_EQ = np.array([_wv(_e, "EQ_%d" % k) for k in range(1, 6)])
assert (_EQ < 0).all() and int(np.argmin(_EQ)) == 0 and set(np.argsort(_EQ)[-2:]) == {3, 4}
assert abs(_EQ[4] - _EQ[1] / 2.0) < 1e-12 and int(np.argmin(_E3)) == 2
assert _WxD(_wp(_e, "mean U")) - _WxD(_wp(_e, "I_4")) == _WxD(_wp(_e, "W_{1,4}"))
_e2 = "Atkinson index of the three-expert profile"
_gm = math.exp(float(np.mean(np.log(_E3))))
_wx_same(_e2, "geometric mean", _gm)
_wx_same(_e2, "I_Atk(1)", 1.0 - _gm / _m3)
assert _wv(_e2, "varepsilon_Atk") == 1.0
# the text: 1 - 0.574/0.638 ~ 0.100 on the printed values, "a tenth" of the mean lost
assert _wx_fmt(1.0 - float(_WxD(_wp(_e2, "geometric mean")) / _WxD(_wp(_e, "mean U"))), 3) == _wp(_e2, "I_Atk(1)")
assert abs(_wv(_e2, "I_Atk(1)") - 0.1) < 0.005
# Chapter 3: the leveling-down illustration
_e = "leveling-down four-expert panel"
_Al, _Ap = _wvs(_e, "A[%d]", 4), _wvs(_e, "A'[%d]", 4)
_wx_same(_e, "mean A", _Al.mean())
_wx_same(_e, "mean A'", _Ap.mean())
_wx_same(_e, "I_4(A)", _gini_vec(_Al))
_wx_same(_e, "-I_4(A)", -_gini_vec(_Al))
_wx_same(_e, "I_4(A')", _gini_vec(_Ap))
_wx_same(_e, "drop in mean overlap", _Al.mean() - _Ap.mean())
_lth = (_Al.mean() - _Ap.mean()) / (_gini_vec(_Al) - _gini_vec(_Ap))
_wx_same(_e, "threshold lambda", _lth)
# the text: perfect equity at A' and a lower mean (the leveling-down case (iii)),
# nearly half of the representation discarded, the level-bearing criterion preferring
# A below the threshold and A' above it, and 0.30/0.214 ~ 1.40 on the printed values
assert _gini_vec(_Ap) == 0.0 and _Ap.mean() < _Al.mean() and 0.4 < 1.0 - _Ap.mean() / _Al.mean() < 0.5
for _lam, _pref in ((0.99 * _lth, 1), (1.01 * _lth, -1)):
    assert _pref * ((_Al.mean() - _lam * _gini_vec(_Al)) - (_Ap.mean() - _lam * _gini_vec(_Ap))) > 0
assert _wx_fmt(float(_WxD(_wp(_e, "drop in mean overlap")) / _WxD(_wp(_e, "I_4(A)"))), 2) == _wp(_e, "threshold lambda")
# Chapter 3: the ultimatum-game identification
_e = "ultimatum-game identification"
for _tag in ("thirty percent", "a quarter"):
    _sh = _wv(_e, "s_hat (%s)" % _tag)
    _wx_same(_e, "1 - 2 s_hat (%s)" % _tag, 1.0 - 2.0 * _sh)
    _wx_same(_e, "alpha (%s)" % _tag, _sh / (1.0 - 2.0 * _sh))
    # the responder with this alpha accepts exactly from s_hat on (Fehr and Schmidt)
    _al = _wv(_e, "alpha (%s)" % _tag)
    assert abs(_al / (1.0 + 2.0 * _al) - _sh) < 1e-12
# Chapter 4: Example ex:unbalanced-clustering, every count recomputed as twice a sum of
# binomial coefficients C(n, 2) (the unordered pairs, each counted in both orders)
_e = "ex:unbalanced-clustering"
_Nc4, _Kc4 = int(_wv(_e, "N")), int(_wv(_e, "K"))
_bl4 = [int(_wv(_e, "balanced size n_%d" % (_k + 1))) for _k in range(_Kc4)]
_sk4 = [int(_wv(_e, "skewed size n_%d" % (_k + 1))) for _k in range(_Kc4)]


def _pairs4(sizes):
    return 2 * sum(math.comb(int(n), 2) for n in sizes)


assert sum(_bl4) == _Nc4 == sum(_sk4) and max(_bl4) - min(_bl4) <= 1 and min(_sk4) >= 1
_wx_same(_e, "balanced n_k - 1", _bl4[0] - 1)
_wx_same(_e, "balanced intra-group count", _pairs4(_bl4))
_wx_same(_e, "N - 1", _Nc4 - 1)
_wx_same(_e, "flat count", _pairs4([_Nc4]))
_wx_same(_e, "skewed n_1 - 1", _sk4[0] - 1)
_wx_same(_e, "skewed subgroups of size n_2", _sk4.count(_sk4[1]))
_wx_same(_e, "skewed n_2 - 1", _sk4[1] - 1)
_wx_same(_e, "skewed intra-group count", _pairs4(_sk4))
_rb4 = _pairs4(_bl4) / _pairs4([_Nc4])
_rs4 = 1.0 - _pairs4(_sk4) / _pairs4([_Nc4])
_wx_same(_e, "balanced / flat", _rb4)
_wx_same(_e, "1 - skewed / flat", _rs4)
_wx_same(_e, "skewed / balanced", _pairs4(_sk4) / _pairs4(_bl4))
_wx_same(_e, "share of the largest subgroup in the skewed count", _pairs4(_sk4[:1]) / _pairs4(_sk4))


def _unit4(x):
    """The unit fraction 1/q (q = 2..10) nearest to x."""
    return min(range(2, 11), key=lambda q: abs(x - 1.0 / q))


# the text: the balanced partition cuts the flat count to one third, the skewed one lies
# only a fifth below it (each the nearest unit fraction) and far above the balanced
# figure; the count is governed by the largest subgroup; and the count, Schur-convex in
# the sizes, is smallest at the balanced sizes among all partitions of N into K nonempty
# subgroups (Section 4.2) and larger at the skewed sizes, which majorize them
assert round(3 * _rb4) == 1 and _unit4(_rb4) == 3, "the balanced saving is not a reduction to one third"
assert round(5 * _rs4) == 1 and _unit4(_rs4) == 5, "the skewed count does not lie a fifth below the flat one"
assert _pairs4(_sk4) > 2 * _pairs4(_bl4), "the skewed count is not far above the balanced one"
assert _sk4[0] == max(_sk4) and _pairs4(_sk4[:1]) > 0.9 * _pairs4(_sk4), "the largest subgroup does not govern the count"
_all4 = [np.diff((0,) + _c + (_Nc4,)) for _c in itertools.combinations(range(1, _Nc4), _Kc4 - 1)]
assert min(_pairs4(_p) for _p in _all4) == _pairs4(_bl4), "the balanced sizes do not minimize the count"
_cum4 = [np.cumsum(sorted(_p, reverse=True)) for _p in (_sk4, _bl4)]
assert (_cum4[0] >= _cum4[1]).all() and _pairs4(_sk4) > _pairs4(_bl4), "the skewed sizes do not majorize the balanced"
# Chapter 4: the fuzzy C-means memberships, recomputed as the normalized inverse powers
# d_k^(-2/(m-1)) (the Lagrangian solution) and by minimizing the expert's term of J_m
# numerically (scipy minimize_scalar), with the behavior the text states at the limits of m
_e = "fuzzy C-means memberships"
_mf4 = _wv(_e, "m")
_df4 = np.array([_wv(_e, "distance to the nearer centroid"), _wv(_e, "distance to the farther centroid")])
assert int(_wv(_e, "number of centroids")) == _df4.size == 2 and _df4[0] < _df4[1]


def _memb4(d, m):
    _p = np.asarray(d, float) ** (-2.0 / (m - 1.0))
    return _p / _p.sum()


_opt4 = _wx_minsc4(lambda u: u ** _mf4 * _df4[0] ** 2 + (1.0 - u) ** _mf4 * _df4[1] ** 2, bounds=(0.0, 1.0),
                   method="bounded", options={"xatol": 1e-12})
_wx_same(_e, "exponent 2/(m - 1)", 2.0 / (_mf4 - 1.0))
_wx_same(_e, "ratio d_near / d_far", _df4[0] / _df4[1])
_wx_same(_e, "membership in the nearer subgroup", _memb4(_df4, _mf4)[0])
_wx_same(_e, "membership in the nearer subgroup", _opt4.x, tol=1e-6)
_wx_same(_e, "membership in the farther subgroup", _memb4(_df4, _mf4)[1])
_wx_same(_e, "membership of an equidistant expert", _memb4(np.repeat(_df4[:1], 2), _mf4)[0])
# the text: at m = 2 the exponent equals 2; the expert's two memberships sum to one, the
# nearer the larger, so that maximal membership assigns it to the nearer subgroup; an
# equidistant expert is split evenly (1/K); as m -> 1+ the memberships harden toward the
# nearest-centroid indicator, and large m drives them toward the uniform 1/K
assert _mf4 == 2.0 and _wv(_e, "exponent 2/(m - 1)") == 2.0
assert abs(_wv(_e, "membership in the nearer subgroup") + _wv(_e, "membership in the farther subgroup") - 1.0) < 1e-12
assert int(np.argmax(_memb4(_df4, _mf4))) == int(np.argmin(_df4))
assert abs(_wv(_e, "membership of an equidistant expert") - 1.0 / _df4.size) < 1e-12
assert _memb4(_df4, 1.01)[0] > 1.0 - 1e-9 and np.abs(_memb4(_df4, 1000.0) - 1.0 / _df4.size).max() < 1e-2
# Chapter 5: the equicorrelated ceiling, Example ex:neff, Example ex:corr-lineage
_e = "equicorrelated ceiling"
_wx_same(_e, "ceiling 1/r", 1.0 / _wv(_e, "r"))
_e = "ex:neff"
_Nn, _rn = int(_wv(_e, "N")), _wv(_e, "r")
_wn = np.full(_Nn, 1.0 / _Nn)
assert abs(_wv(_e, "w_i") - 1.0 / _Nn) < 1e-15
_wx_same(_e, "N_eff at R = I", 1.0 / float(_wn @ _wn))
_wx_same(_e, "1 + (N - 1) r", 1.0 + (_Nn - 1) * _rn)
_wx_same(_e, "N_eff, equicorrelated", _Nn / (1.0 + (_Nn - 1) * _rn))
assert _rn == _wv("equicorrelated ceiling", "r") and _wv(_e, "N_eff, equicorrelated") < 1.0 / _rn
_gs = [int(_wv(_e, "group size %d" % (k + 1))) for k in range(3)]
assert sum(_gs) == _Nn and _gs == sorted(_gs, reverse=True)
_lab = np.repeat(np.arange(3), _gs)
_Rn = np.where(_lab[:, None] == _lab[None, :], _wv(_e, "r_w"), _wv(_e, "r_b"))
np.fill_diagonal(_Rn, 1.0)
assert np.linalg.eigvalsh(_Rn).min() > 0
_wx_same(_e, "N_eff, three lineages, equal weights", 1.0 / float(_wn @ _Rn @ _wn))
for _k in range(3):
    _wx_same(_e, "group share %d, equal weights" % (_k + 1), _wn[_lab == _k].sum())
    _wx_same(_e, "group share %d, penalty minimum" % (_k + 1), 1.0 / 3.0)
_wq = np.array([1.0 / (3.0 * _gs[k]) for k in _lab])
_wx_same(_e, "N_eff, equal shares, equal weights within", 1.0 / float(_wq @ _Rn @ _wq))
_wx_same(_e, "N_eff of the five-member lineage alone", 1.0 / float(np.full(_gs[0], 1.0 / _gs[0]) @ _Rn[:_gs[0], :_gs[0]]
                                                                 @ np.full(_gs[0], 1.0 / _gs[0])))
# the text: far below the nominal ten and below even the three lineages, the dominant
# lineage about one expert, the effective number rising at the equal shares, and the
# dominant lineage holding half of the equal weights
_n3 = _wv(_e, "N_eff, three lineages, equal weights")
assert _n3 < 3 < _Nn and round(_wv(_e, "N_eff of the five-member lineage alone")) == 1
assert _wv(_e, "N_eff, equal shares, equal weights within") > _n3
assert abs(_wv(_e, "group share 1, equal weights") - 0.5) < 1e-12
_e = "ex:corr-lineage"
_nl, _rl = int(_wv(_e, "n")), _wv(_e, "r")
_wx_same(_e, "n - 1", _nl - 1)
_wx_same(_e, "N_eff of the lineage", _nl / (1.0 + (_nl - 1) * _rl))
# the stated correlation is the stipulated lineage mean of Figure 5.4, whose equal-weight
# count prints the same "about 1.1" that the example derives
_dpc = _cfgd(_wb(_WX_CHAPTER_WORKBOOK[5]))
assert abs(_rl - float(_dpc["lineage_mean_corr"])) < 1e-12
assert _wx_fmt(float(_dpc["neff_lineage_equalweight"]), 1) == _wp(_e, "N_eff of the lineage")
print("   worked examples: %d rows in %d sheets, every computed value equal to its cross-check and"
      " recomputed from the stated inputs, every printed string its value; the claims of the text on"
      " the GLS weights, the DeGroot limit, ex:bonferroni, ex:skewed-panel, the welfare, Atkinson and"
      " leveling-down illustrations, the ultimatum identification, ex:unbalanced-clustering, the fuzzy"
      " C-means memberships, ex:neff and ex:corr-lineage hold" % (len(_wxt), len(_WX_CHAPTER_WORKBOOK)))

if _environment is not None and _environment.soft_failures():
    print("   rule-12 self-checks finished; %d did not hold on this software stack (listed below)"
          % len(_environment.soft_failures()))
else:
    print("   all rule-12 self-checks passed")
if _environment is not None:
    _environment.reminder()
