#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
gen_case_manifest.py -- writes case_manifest.json (beside the workbooks), the record of the
numbers the book prints from the workbooks.  Every number in it is read from the workbooks
(or computed here from workbook values), so the manifest cannot drift from the pack:
printed_tables and key_numbers, and the summaries of Chapter 2, Case I, Case II and the
vignette, carry the workbook values at full precision, while the readable summaries of
Cases III and IV and the Chapter 5 anchor round them, mostly to the precision the book
prints (a statistic behind a printed bound keeps more digits, and the upper bounds of
the Case III start statistics are rounded up so that the printed bound holds).
The labels (case, title, section) and the descriptive notes (honesty statements, rules,
statements) are text.

It holds
  the Case III and Case IV entries, hm_panel_anchor
                  readable summaries of the Case III and Case IV results and of the Chapter 5
                  panel (Table tab:hmpanel);
  chapter2_running_example, case_i, case_ii, vignette_v
                  readable summaries of the other studies;
  printed_tables  for the tables whose cells are workbook values (by their LaTeX labels:
                  Tables 2.1, 2.5, 5.3, 6.6 to 6.9, 6.11, 6.12, 6.16 to 6.20), the expected
                  content of each data row, cell by cell: a string = text compared
                  case-insensitively, a number or a list of numbers = the numbers of that
                  cell in order, compared at the precision the book prints them
                  ({"value": v, "decimals": d} fixes the rounding explicitly, an integer is
                  compared exactly), null = a text cell that must print no number, and
                  {"free": reason} = a cell whose printed numbers no workbook records, allowed
                  only for the stated reason and reported as a note;
  key_numbers     numbers the prose prints (and the workbook-derived numbers of the
                  robustness table, Table 6.22), each with the section (its label) and the
                  scope in which it appears, the text it is printed as, and its anchor, the
                  normalized LaTeX printed right before it ('#' standing for any number in
                  between), which must occur once in that scope (an anchor given as a list
                  names alternative wordings of one sentence, of which exactly one may occur,
                  and once), and, for a count printed as a word, the words printed after it;
                  the arithmetic of the examples that Chapters 2 to 5 work by hand is
                  read from the sheets worked_examples of the Figure 2.2, 3.1, 4.3 and
                  5.4 workbooks;
  _layout_errors  the studies whose workbook no longer has the layout read here (empty
                  when every reader succeeded).

Only the layout is written here (which sheet and column fills which cell, at what
precision the book prints it, and after which words).  The delivery copy of the book
compares printed_tables and key_numbers with the LaTeX sources
(tools/check_printed_numbers.py); a table or sentence that is reworded, or a workbook
whose layout changes, is then updated here.

Run after build_workbooks.py:  python gen_case_manifest.py
"""
import json
import math
import os
import sys
from fractions import Fraction

# Ask for the numerical kernels of the reference platform before numpy is imported (see
# ../environment.py); the manifest is then computed with the same arithmetic as the workbooks.
sys.path.insert(0, os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")))
try:
    import environment as _environment
    _environment.pin_numerics()
except ImportError:
    pass

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = HERE                                     # simulation_data/
OUTP = os.path.join(HERE, "case_manifest.json")  # written beside the workbooks


_loaded = {}


def load(name):
    """Every sheet of a workbook, read once (the readers below only read them)."""
    if name not in _loaded:
        _loaded[name] = pd.read_excel(os.path.join(DATA, name), sheet_name=None)
    return _loaded[name]


def cfgd(sheets):
    c = sheets["config"]
    return dict(zip(c["key"].astype(str), c["value"]))


def r2(x):
    return round(float(x), 2)


def r3(x):
    return round(float(x), 3)


def up2(x):
    """an upper bound at two significant figures, rounded UP so that the printed
    bound holds for the recorded value (e.g. 0.0015158 -> 0.0016)."""
    v = float(x)
    if v <= 0 or not math.isfinite(v):
        return v
    e = math.floor(math.log10(v)) - 1
    return float("%.2g" % (math.ceil(v / 10 ** e - 1e-9) * 10 ** e))


# ===================================================================== Case IV
rt = load("fig6-11_rt_case.xlsx")
rtc = cfgd(rt)
rr = rt["results"]
st = rt["stations"]
cons = rt["constituencies"]
swapped = [t.strip() for t in str(rtc["swapped"]).split(",")]
contested = str(rtc["contested_station"])
rri = rr.set_index("station")
rp = rt["realized_panel"].set_index("station")
nt = rt["near_ties"]
# "uncontested" as the builder and the book use it: every station but the contested one
unc = rr[rr["station"].astype(str) != contested]


def _max_abs(col_eff, col_eq):
    d = {str(r["station"]): float(r[col_eq]) - float(r[col_eff]) for _, r in unc.iterrows()}
    s = max(d, key=lambda k: abs(d[k]))
    return {"station": s, "value": round(abs(d[s]), 5)}


# the prose bounds, re-derived here from the workbook so the manifest cannot drift
_du = _max_abs("Ubar_eff", "Ubar_eq")
assert _du["value"] < 0.01, "an uncontested mean utility moves by 0.01 or more"
assert str(rtc["case"]).startswith("IV -- ")
# in the overlaps the contested move is the leveling-down case of prop:consensus-fairness(iii)
assert float(rri.loc[contested, "gini_A_eq"]) < float(rri.loc[contested, "gini_A_eff"])
assert float(rri.loc[contested, "Abar_eq"]) < float(rri.loc[contested, "Abar_eff"])
assert (rr["Abar_eq"] < rr["Abar_eff"]).all()
# ... and it stays so under every choice among the nearly tied optima
assert contested in str(rtc["overlap_leveling_down_combo_always"]).split(",")

case_iv = {
    "case": "IV",
    "title": "Urban rail-transit station accessibility with stakeholder equity",
    "section": "6.5",
    "source_paper": "chen2024railtransit",
    "honesty": ("Synthetic instrument and panel under seed %d. Mirrors the "
                "source study at the level of MECHANISM only -- equity "
                "weighting reorders priorities -- and must NOT be presented "
                "as reproducing any substantive finding of the source paper." % int(rtc["seed"])),
    "seed": int(rtc["seed"]),
    "workbook_column_suffixes": ("_eff = consensus-only regime, _eq = equity-aware regime "
                                 "(Sen's abbreviated welfare form)"),
    "instrument": {
        "n_stations": len(st),
        "scale": "continuous accessibility index on [0, 10], Gaussian fits",
        "stations": [{"id": str(r["id"]), "name": str(r["name"]),
                      "generator_means": {"operators": r2(r["ops"]),
                                          "passengers": r2(r["pas"]),
                                          "advocates": r2(r["adv"])}}
                     for _, r in st.iterrows()],
        "contested_station": contested,
    },
    "panel": {
        "N": int(rtc["n_stakeholders"]),
        "constituencies": [{"constituency": str(r["constituency"]),
                            "size": int(r["size"])} for _, r in cons.iterrows()],
        "generator": {
            "leniency_sd": r2(rtc["leniency_sd"]),
            "noise_sd": r2(rtc["noise_sd"]),
            "common_reported_sd": float(rtc["common_sigma"]),
            "draw_order": str(rtc["generator_draw_order"]),
            "clip": str(rtc["generator_clip"]),
            "clip_binding": bool(int(rtc["generator_clip_binding"])),
            "note": ("all stakeholders share a common reported assessment "
                     "precision, so the confidence objective is constant over "
                     "the simplex and non-discriminating; it is DROPPED from "
                     "both regimes, each of which is a single-objective program"),
        },
    },
    "model": {
        "regime_consensus": str(rtc["regime_eff"]),
        "regime_equity": str(rtc["regime_eq"]),
        "overlap": str(rtc["overlap"]),
        "fairness_utilities": ("Fehr-Schmidt utilities of the overlaps "
                               "(rho %.1f, vartheta %.1f), power "
                               "normalization (eq:parabola-norm, kappa %d), as in Cases I and III; "
                               "Case II uses the exponential normalization") % (
                                   float(rtc["rho"]), float(rtc["theta"]), int(rtc["kappa"])),
        "solver": str(rtc["solver"]),
        "solver_starts": str(rtc["solver_starts"]),
        "solver_ftol": float(rtc["solver_ftol"]),
        "certification": str(rtc["certification"]),
        "consensus_optimum": str(rtc["consensus_optimum"]),
        "weight_reporting": ("BOTH regimes report the nearest-to-uniform "
                             "weight vector of their own optimal level set "
                             "(the projection of the uniform vector onto the "
                             "set, a convex QP cross-checked against its KKT "
                             "form), one shared tie-break, so the weight-Gini and "
                             "constituency-share columns are comparable across "
                             "the regimes. The shares are tie-break "
                             "representatives, not properties of the optima; what "
                             "each optimal level set fixes is the LP range of the "
                             "advocates' share (level_set_advocate_share) and the "
                             "overlaps and utilities, which depend on w only "
                             "through the collective mean"),
        "priority_rule": "upgrading priority = ascending collective accessibility (worst station first)",
    },
    "results": {
        "stations": [{"id": str(r["station"]), "name": str(r["name"]),
                      "mean_consensus": r2(r["mean_eff"]),
                      "mean_equity": r2(r["mean_eq"]),
                      "priority_consensus": int(r["priority_eff"]),
                      "priority_equity": int(r["priority_eq"]),
                      "weight_gini_consensus": r2(r["weight_gini_eff"]),
                      "weight_gini_equity": r2(r["weight_gini_eq"]),
                      "advocate_share_consensus": r2(r["adv_share_eff"]),
                      "advocate_share_equity": r2(r["adv_share_eq"]),
                      "level_set_advocate_share_consensus": [r2(r["adv_share_min_eff"]), r2(r["adv_share_max_eff"])],
                      "level_set_advocate_share_equity": [r2(r["adv_share_min_eq"]), r2(r["adv_share_max_eq"])],
                      "advocate_mean_overlap_consensus": r2(r["Abar_adv_eff"]),
                      "advocate_mean_overlap_equity": r2(r["Abar_adv_eq"]),
                      "mean_overlap_consensus": r2(r["Abar_eff"]),
                      "mean_overlap_equity": r2(r["Abar_eq"]),
                      "min_overlap_consensus": r2(r["Amin_eff"]),
                      "min_overlap_equity": r2(r["Amin_eq"]),
                      "mean_utility_consensus": r3(r["Ubar_eff"]),
                      "mean_utility_equity": r3(r["Ubar_eq"]),
                      "utility_gini_consensus": r3(r["gini_U_eff"]),
                      "utility_gini_equity": r3(r["gini_U_eq"]),
                      "consensus_member": "%d (%s)" % (int(r["consensus_member"]), str(r["consensus_member_constituency"]))}
                     for _, r in rr.iterrows()],
        "priority_order_consensus": [str(r["station"]) for _, r in
                                     rr.sort_values("priority_eff").iterrows()],
        "priority_order_equity": [str(r["station"]) for _, r in
                                  rr.sort_values("priority_eq").iterrows()],
        "max_uncontested_mean_gap": r2(rtc["max_uncontested_mean_gap"]),
        "max_uncontested_mean_gap_station": str(rtc["max_uncontested_mean_gap_station"]),
        "uncontested_gaps_below_threshold": str(rtc["uncontested_gaps_below_threshold"]),
        "contested_mean_gap": r2(rtc["contested_mean_gap"]),
        "largest_gap_realized_panel": {
            "station": str(rtc["max_uncontested_mean_gap_station"]),
            "two_lowest_below_panel_mean": [r2(rp.loc[str(rtc["max_uncontested_mean_gap_station"]), "low1_below_panel_mean"]),
                                            r2(rp.loc[str(rtc["max_uncontested_mean_gap_station"]), "low2_below_panel_mean"])],
            "two_lowest_constituencies": [str(rp.loc[str(rtc["max_uncontested_mean_gap_station"]), "low1_constituency"]),
                                          str(rp.loc[str(rtc["max_uncontested_mean_gap_station"]), "low2_constituency"])],
        },
        "swap": {
            "stations": swapped,
            "statement": ("under the equity-aware regime %s moves ahead of %s "
                          "in upgrading priority; every other pairwise order "
                          "is unchanged") % (swapped[0], swapped[1]),
            "contested_means": {
                s: {"consensus": r2(rri.loc[s, "mean_eff"]), "equity": r2(rri.loc[s, "mean_eq"])}
                for s in swapped},
            "equity_margin": r2(rtc["eq_swap_margin"]),
            "reseed_margin_sd": r2(rtc["reseed_eq_swap_margin_sd"]),
            "reseed_margin_min": r3(rtc["reseed_eq_swap_margin_min"]),
            "advocate_share_at_contested": {
                "consensus": r2(rri.loc[contested, "adv_share_eff"]),
                "equity": r2(rri.loc[contested, "adv_share_eq"]),
                "ratio_equity_to_consensus": r2(rtc["contested_adv_share_ratio"]),
                "note": "nearest-to-uniform tie-break representatives",
            },
            "level_set_advocate_share_at_contested": {
                "consensus": str(rtc["contested_adv_share_range_eff"]),
                "equity": str(rtc["contested_adv_share_range_eq"]),
                "common": str(rtc["contested_adv_share_common_range"]),
                "headcount_share": float(rtc["adv_headcount_share"]),
                "consensus_caps_below_headcount_at": str(rtc["consensus_caps_advocates_below_headcount_at"]),
            },
        },
        "weight_gini": {
            "average_consensus": r2(rtc["weight_gini_eff_avg"]),
            "average_equity": r2(rtc["weight_gini_eq_avg"]),
            "statement": ("averaged over the ten stations, and with both "
                          "regimes reporting the nearest-to-uniform "
                          "representative of their optimal level sets, the "
                          "weight Gini stands at %.2f under the "
                          "consensus-only regime and %.2f under the "
                          "equity-aware regime")
                         % (float(rtc["weight_gini_eff_avg"]),
                            float(rtc["weight_gini_eq_avg"])),
            "near_tie_range_consensus": str(rtc["near_tie_weight_gini_eff_avg_range"]),
            "near_tie_range_equity": str(rtc["near_tie_weight_gini_eq_avg_range"]),
        },
        "near_ties": {
            "tolerance_relative": float(rtc["near_tie_tolerance"]),
            "rule": str(rtc["near_tie_rule"]),
            "solves": str(rtc["near_ties"]),
            "smallest_non_tie_relative_gap": float("%.2e" % float(rtc["near_tie_smallest_non_tie_rel_gap"])),
            "local_maxima": [{"station": str(r["station"]), "regime": str(r["regime"]), "rank": int(r["rank"]),
                              "m_reported": r2(r["m_reported"]), "m_local_max": r2(r["m_local_max"]),
                              "objective_gap": float("%.2e" % float(r["objective_gap"])),
                              "relative_gap": float("%.2e" % float(r["relative_gap"]))}
                             for _, r in nt.iterrows()],
            "closest": {"solve": str(rtc["near_tie_closest"]),
                        "n_local_maxima": int(rtc["near_tie_closest_n_maxima"]),
                        "max_relative_gap": float("%.2e" % float(rtc["near_tie_closest_max_rel_gap"])),
                        "relative_gap_bound": float(rtc["near_tie_closest_gap_bound"]),
                        "m_reported": r2(rtc["near_tie_closest_m_reported"]),
                        "m_local_maxima": [r2(v) for v in str(rtc["near_tie_closest_m_local_maxima"]).split(",")],
                        "runner_up_advocate_share": r2(rtc["near_tie_closest_runner_up_adv_share"]),
                        "runner_up_weight_gini": r2(rtc["near_tie_closest_runner_up_weight_gini"])},
            "solves_with_several_local_maxima": str(rtc["near_tie_solves_with_several_maxima"]),
            "combinations": int(rtc["near_tie_combinations_n"]),
            "priorities_unchanged": str(rtc["near_tie_priorities_unchanged"]),
            "gap_statement_holds": str(rtc["near_tie_gap_statement_holds"]),
        },
    },
    "robustness": {
        "reseed": str(rtc["robustness_reseed"]),
        "reseed_swap_direction_preserved": str(rtc["robustness_reseed_swap_direction"]),
        "reseed_other_orders_unchanged": str(rtc["robustness_reseed_other_orders_unchanged"]),
        "reseed_reordered_pairs": sorted(set(str(x) for x in rt["robustness_seeds"]["stations_reordered"])),
        "reseed_contested_moves_two_places": str(rtc["robustness_reseed_contested_moves_two_places"]),
        "fehr_schmidt_grid": str(rtc["robustness_fs_grid"]),
        "fehr_schmidt_swap_direction_preserved": str(rtc["robustness_fs_swap_direction"]),
        "fehr_schmidt_S4_advocate_share_range": str(rtc["robustness_fs_S4_adv_share_range"]),
        "early_warning": {
            "threshold": float(rtc["warn_threshold"]),
            "contested_gap_min": r2(rtc["warn_contested_gap_min"]),
            "uncontested_gap_max": r2(rtc["warn_uncontested_gap_max"]),
            "panels_separated": str(rtc["warn_panels_separated"]),
        },
        "leveling_down": {
            "mean_utility_rises_at": [str(r["station"]) for _, r in rr.iterrows() if float(r["Ubar_eq"]) > float(r["Ubar_eff"])],
            "mean_utility_falls_at": {str(r["station"]): round(float(r["Ubar_eq"]) - float(r["Ubar_eff"]), 5)
                                      for _, r in rr.iterrows() if float(r["Ubar_eq"]) <= float(r["Ubar_eff"])},
            "max_abs_mean_utility_change_uncontested": _du,
            "max_abs_utility_gini_change_uncontested": _max_abs("gini_U_eff", "gini_U_eq"),
            "contested_mean_utility": {"consensus": r3(rri.loc[contested, "Ubar_eff"]),
                                       "equity": r3(rri.loc[contested, "Ubar_eq"])},
            "contested_utility_gini": {"consensus": r3(rri.loc[contested, "gini_U_eff"]),
                                       "equity": r3(rri.loc[contested, "gini_U_eq"])},
            "contested_mean_overlap": {"consensus": r2(rri.loc[contested, "Abar_eff"]),
                                       "equity": r2(rri.loc[contested, "Abar_eq"])},
            "contested_min_overlap": {"consensus": r2(rri.loc[contested, "Amin_eff"]),
                                      "equity": r2(rri.loc[contested, "Amin_eq"])},
            "contested_overlap_gini": {"consensus": r2(rri.loc[contested, "gini_A_eff"]),
                                       "equity": r2(rri.loc[contested, "gini_A_eq"])},
            "mean_overlap_lower_under_equity_at": str(rtc["mean_overlap_lower_under_equity_at"]),
            "overlap_leveling_down_at": str(rtc["overlap_leveling_down_at"]).split(","),
            "overlap_leveling_down_except": str(rtc["overlap_leveling_down_except"]),
            "max_mean_overlap_fall_uncontested": {"station": str(rtc["uncontested_max_mean_overlap_fall_station"]),
                                                  "value": r3(rtc["uncontested_max_mean_overlap_fall"])},
            "overlap_leveling_down_rule": str(rtc["overlap_leveling_down_rule"]),
            "overlap_leveling_down_over_near_tie_combinations": {
                "count_range": str(rtc["overlap_leveling_down_combo_count_range"]),
                "always": str(rtc["overlap_leveling_down_combo_always"]).split(","),
                "never": str(rtc["overlap_leveling_down_combo_never"]).split(","),
                "turns_on_the_choice": str(rtc["overlap_leveling_down_combo_varies"]).split(","),
                "mean_overlap_lower_count_range": str(rtc["mean_overlap_lower_combo_count_range"]),
            },
        },
    },
}

# =================================================================== Case III
bdt = load("fig6-8_bdt_case.xlsx")
bdtc = cfgd(bdt)
bgrp = bdt["groups"]
bdim = bdt["dimensions"]
bsen = bdt["sensitivity"]
bkt = bdt["kmeans_topsis"]
case_iii = {
    "case": "III",
    "title": "Building digital twin maturity via FA-LSCOG",
    "source_paper": "chen2024bdt",
    "honesty": ("Synthetic %d-expert x %d-indicator truncated-Weibull panel under "
                "seed %d, generated and solved in build_workbooks.py. Mirrors "
                "the source study at the level of MECHANISM only and must NOT be "
                "presented as reproducing any elicited value."
                % (int(bdtc["n_experts"]), int(bdtc["n_indicators"]), int(bdtc["seed"]))),
    "seed": int(bdtc["seed"]),
    # the seed of the second, independent set of solver starts (seed + probe_seed_offset)
    "probe_seed": int(float(bdtc["probe_seed"])),
    "program": str(bdtc["program"]),
    "parameters": {"rho": float(bdtc["rho"]), "vartheta": float(bdtc["vartheta"]),
                   "power_kappa": float(bdtc["power_kappa"]), "lambda_G": float(bdtc["lambda_G"]),
                   "operating_delta": float(bdtc["operating_delta"]),
                   "solver": str(bdtc["solver"]), "starts": str(bdtc["starts"]),
                   "normalization": str(bdtc["normalization"])},
    "groups": [{"group": str(r["group"]), "stratum": str(r["stratum"]), "size": int(r["size"]),
                "Omega": r2(r["Omega"]), "Omega_3dp": r3(r["Omega"]),
                "q1_mean": r2(r["q1_mu"])} for _, r in bgrp.iterrows()],
    "q1_collective_mean": r2(sum(float(r["Omega"]) * float(r["q1_mu"]) for _, r in bgrp.iterrows())),
    "topsis_closeness": [r3(v) for v in bkt["closeness"]],
    "dimensions": [{"id": str(r["id"]), "dimension": str(r["dimension"]),
                    "weight": r2(r["weight"]), "weight_3dp": r3(r["weight"]),
                    "mean": r2(r["mu"]), "level": str(r["level"]),
                    "bwm_best_to_others": int(r["bwm_best_to_others"]),
                    "bwm_others_to_worst": int(r["bwm_others_to_worst"])}
                   for _, r in bdim.iterrows()],
    "bwm": {"best": str(bdtc["bwm_best"]), "worst": str(bdtc["bwm_worst"]),
            "model": "linear (rezaei2016linear)",
            "xi_linear": r3(bdtc["bwm_xi"]),
            "a_BW": int(float(bdtc["bwm_a_BW"])),
            "consistency_ratio_input_based": round(float(bdtc["bwm_CRI"]), 4),
            "consistency_ratio_input_based_at": str(bdtc["bwm_CRI_argmax"]),
            "threshold_input_based": round(float(bdtc["bwm_CRI_threshold"]), 4),
            "threshold_source": "liang2020consistency, Table 3",
            "ordinal_consistent": int(bdtc["bwm_ordinal_consistent"]),
            "ordinal_rule": "liang2020consistency, Definition 4 (eq. 8)"},
    "global": {"maturity": r2(bdtc["global_maturity"]), "level": str(bdtc["global_level"]),
               "aggregate_variance": r3(bdtc["aggregate_variance"]),
               "fairness_Phi": r3(bdtc["fairness_Phi"]), "weight_gini": r2(bdtc["weight_gini"]),
               "mass_below_2": r3(bdtc["mass_below_2"])},
    "sweep": [{"delta": r2(r["delta"]), "global_maturity": r2(r["global_maturity"]),
               "fairness_Phi": r3(r["fairness_Phi"]), "aggregate_variance": r3(r["aggregate_variance"]),
               "weight_gini": r2(r["weight_gini"]), "mass_below_2": r3(r["mass_below_2"])}
              for _, r in bsen.iterrows()],
    "moments": {"member_sd_min": r2(bdtc["member_sd_min"]), "member_sd_max": r2(bdtc["member_sd_max"]),
                "member_sd_mean": r2(bdtc["member_sd_mean"]),
                "largest_group_variance_op": r3(bsen.set_index("delta").loc[float(bdtc["operating_delta"]),
                                                                            "max_group_variance"])},
    "multistart": {
        # upper bounds (max_*, *_max_raise, *_diff_max, exceed_max_weight_diff) are
        # rounded UP at two significant figures so that a printed bound holds
        "n_programs": int(bdtc["multistart_n_programs"]), "n_runs": int(bdtc["multistart_n_runs"]),
        "n_best_sampled_is_center": int(bdtc["multistart_n_duplicate_start_dropped"]),
        "n_runs_capped": int(bdtc["multistart_n_runs_capped"]),
        "retention": str(bdtc["retention"]),
        "median_starts_at_best": float(bdtc["multistart_median_starts_at_best"]),
        "n_two_or_more_at_best": int(bdtc["multistart_n_two_or_more_at_best"]),
        "n_one_at_best": int(bdtc["multistart_n_one_at_best"]),
        "n_none_at_best": int(bdtc["multistart_n_none_at_best"]),
        "max_gap_one_at_best": up2(bdtc["multistart_max_gap_one_at_best"]),
        "max_gap_none_at_best": up2(bdtc["multistart_max_gap_none_at_best"]),
        "retained_n_from_further_starts": int(bdtc["retained_n_from_further_starts"]),
        "retained_from_further_programs": str(bdtc["retained_from_further_programs"]),
        "further_n_programs_exceeded": int(bdtc["further_n_programs_exceeded"]),
        "further_max_raise": up2(bdtc["further_max_raise"]),
        "further_max_raise_argmax": str(bdtc["further_max_raise_argmax"]),
        "followup": [{"delta": r2(r["delta"]), "group": str(r["group"]), "indicator": str(r["indicator"]),
                      "original_runner_up_gap": float("%.2g" % float(r["original_runner_up_gap"])),
                      "n_followup_starts": int(r["n_followup_starts"]),
                      "n_match_original_best": int(r["n_match_original_best"]),
                      "n_exceed_original_best": int(r["n_exceed_original_best"]),
                      "exceed_min": float("%.2g" % float(r["exceed_min"])),
                      "best_minus_original_best": float("%.2g" % float(r["best_minus_original_best"])),
                      "exceed_max_weight_diff": up2(r["exceed_max_weight_diff"]),
                      "runner_up_weight_diff": r2(r["runner_up_weight_diff"]),
                      "retained_source": str(bdt["intra_group_solutions"].set_index(
                          ["delta", "group", "indicator"]).loc[(r["delta"], r["group"], r["indicator"]),
                                                               "retained_source"])}
                     for _, r in bdt["multistart_followup"].iterrows()],
        "probe_starts": int(bdtc["probe_starts"]),
        "probe_n_programs_exceeded": int(bdtc["probe_n_programs_exceeded"]),
        "probe_max_raise": up2(bdtc["probe_max_raise"]),
        "probe_max_raise_argmax": str(bdtc["probe_max_raise_argmax"]),
        "near_tol": float(bdtc["near_tol"]),
        "near_weight_diff_max": up2(bdtc["near_weight_diff_max"]),
        "near_weight_diff_argmax": str(bdtc["near_weight_diff_argmax"]),
        "near_n_programs_weight_diff_above_0.01": int(bdtc["near_n_programs_weight_diff_above_0.01"]),
        "near_mean_diff_max": up2(bdtc["near_mean_diff_max"]),
        "swap_n_printed_quantities": int(bdtc["swap_n_printed_quantities"]),
        "swap_near_n_printed_changed": int(bdtc["swap_near_n_printed_changed"]),
        "swap_original_n_printed_changed": int(bdtc["swap_original_n_printed_changed"]),
        "swap_near_changed_quantities": str(bdtc["swap_near_changed_quantities"]),
        "swap_near_single_member_count_delta0": int(bdt["multistart_substitution"].set_index("quantity").loc[
            "single-member programs, delta 0", "printed_near_substituted"]),
        "swap_near_vertex_lost_programs": str(bdtc["swap_near_vertex_lost_programs"]),
        "swap_near_vertex_lost_kept_weight": round(float(bdtc["swap_near_vertex_lost_min_max_weight"]), 4),
        "retained_n_Phi_tilde_above_1": int(bdtc["retained_n_Phi_tilde_above_1"]),
        "retained_max_Phi_tilde": r2(bdtc["retained_max_Phi_tilde"]),
        # the local-polish diagnostic (a diagnostic, not a retention step): how far rounds
        # of a pairwise-transfer pattern search and one SLSQP run raise the retained values
        "polish_n_programs_gain_above_tie_tol": int(bdtc["polish_n_programs_gain_above_tie_tol"]),
        "polish_max_gain": up2(bdtc["polish_max_gain"]),
        "polish_max_gain_argmax": str(bdtc["polish_max_gain_argmax"]),
        "polish_max_weight_diff": up2(bdtc["polish_max_weight_diff"]),
        "polish_max_weight_diff_argmax": str(bdtc["polish_max_weight_diff_argmax"]),
        "polish_max_mean_diff": up2(bdtc["polish_max_mean_diff"]),
        "polish_first_round_max_gain": up2(bdtc["polish_first_round_max_gain"]),
        "swap_polished_n_printed_changed": int(bdtc["swap_polished_n_printed_changed"]),
        "retained_accuracy_bound": up2(bdtc["retained_accuracy_bound"])},
}

# ============================================================ hm-panel anchor
loo = load("fig5-3_loo.xlsx")
res = dict(zip(loo["results"]["quantity"].astype(str), loo["results"]["printed"]))
# the overconfident machines, by their generator group, and the signs of their
# leave-one-out calibration and resolution changes
_oc = loo["loo"][loo["loo"]["group"].astype(str).str.contains("overconfident")]


def _sign(x):
    return "+" if float(x) > 0 else "-" if float(x) < 0 else "0"


hm_anchor = {
    "purpose": ("cross-check only -- these numbers are ALREADY printed in "
                "Chapter 5 (Table tab:hmpanel) and are now backed by the "
                "generated panel of fig5-3_loo.xlsx"),
    "human_only_mean_crps": float(res["human-only mean CRPS"]),
    "collaborative_mean_crps": float(res["collaborative mean CRPS"]),
    "human_weight_share": float(res["human weight share W_H"]),
    "machine_weight_share": float(res["machine weight share W_M"]),
    "overconfident_machines": {str(r["id"]): {"leave_one_out_signs_cal_res": "(%s, %s)" % (_sign(r["delta_cal"]),
                                                                                           _sign(r["delta_res"]))}
                               for _, r in _oc.iterrows()},
}

# ============================================================ printed numbers
# The numbers the book prints from the workbooks: the tables whose cells are
# workbook values, cell by cell, and the key numbers of the prose, each with its
# anchor, read here at full precision.  The layout functions below read the
# workbooks themselves (not the
# summaries above), and a study whose workbook no longer has the layout they
# expect is recorded in _layout_errors instead of stopping the run.
import re as _pn_re
import traceback as _pn_tb


def _pn_f(x):
    return float(x)


def _pn_fmt(x, d):
    """x printed with d decimals, as the book prints it (no negative zero)."""
    s = "%.*f" % (d, float(x))
    return s[1:] if _pn_re.fullmatch(r"-0\.?0*", s) else s


def _pn_vec(v, d=3, zero=5e-4):
    """A weight vector as the book prints it: exact zeros as 0, the rest with d decimals."""
    return "(" + ",".join("0" if abs(float(x)) < zero else _pn_fmt(x, d) for x in v) + ")"


def _pn_normal(mu, sd, d=2):
    """The argument list of N(mu, sd^2) as the book prints it, without spaces."""
    return "(%s,%s^2)" % (_pn_fmt(mu, d), _pn_fmt(sd, d))


def _pn_sci(x):
    """A small positive value as the book prints it in scientific form, m.m \\times 10^{-e}
    (normalized: no braces or spaces), one decimal in the mantissa."""
    v = float(x)
    if not v > 0:
        raise ValueError("%r is not positive" % (x,))
    e = math.floor(math.log10(v))
    m = v / 10.0 ** e
    if "%.1f" % m == "10.0":
        m, e = m / 10.0, e + 1
    return "%.1f\\times10^%d" % (m, e)


def _pn_pow10(x):
    """A power of ten as the book prints it, 10^{-e} (normalized)."""
    e = int(round(math.log10(float(x))))
    if abs(float(x) - 10.0 ** e) > 1e-9 * 10.0 ** e:
        raise ValueError("%r is not a power of ten" % (x,))
    return "10^%d" % e


_PN_ONES = ("zero one two three four five six seven eight nine ten eleven twelve thirteen fourteen "
            "fifteen sixteen seventeen eighteen nineteen").split()
_PN_TENS = "twenty thirty forty fifty sixty seventy eighty ninety".split()


def _pn_words(n):
    """A count below one hundred as the book writes it in words ('thirty-three')."""
    n = int(n)
    if not 0 <= n < 100:
        raise ValueError("%r is not a count the book writes in words" % (n,))
    if n < 20:
        return _PN_ONES[n]
    t, o = divmod(n, 10)
    return _PN_TENS[t - 2] + ("-" + _PN_ONES[o] if o else "")


def _pn_times10(x):
    """A small positive bound as the book prints it, m \\times 10^{-e} (normalized), the
    mantissa integral where the bound's is ('2\\times10^-5')."""
    v = float(x)
    e = math.floor(math.log10(v))
    m = v / 10.0 ** e
    if abs(m - round(m)) < 1e-9 and 1 <= round(m) < 10:
        return "%d\\times10^%d" % (round(m), e)
    return _pn_sci(v)


def _pn_upto(x, d):
    """'at most X': the recorded value rounded up at the d decimals the book prints."""
    return _pn_fmt(math.ceil(float(x) * 10 ** d - 1e-9) / 10.0 ** d, d)


def _pn_below(x, d):
    """'less than X': the smallest number of d decimals above the recorded value."""
    return _pn_fmt((math.floor(float(x) * 10 ** d + 1e-9) + 1) / 10.0 ** d, d)


def _pn_key(file, section, what, anchor, text, value, scope="prose", follow=None):
    """A number the prose prints: in `scope` of the section labelled `section` in `file`
    (scope 'prose' = outside tables and figures, 'table:<label>' = inside that table),
    `anchor` (the LaTeX printed right before the number, without spaces, $ or braces;
    '#' stands for any number printed in between) must occur exactly once, and `text`
    (the number as printed, normalized the same way) must follow it.  A list of anchors
    names alternative wordings of the one sentence (a revision of it, say): exactly one
    occurrence of any of them is allowed.  `follow`, for a count printed as a word, names
    the words printed right after it, so that 'seven' cannot pass for 'seventeen'."""
    key = {"file": file, "section": section, "scope": scope, "what": what,
           "anchor": anchor, "text": text, "value": value}
    if follow:
        key["follow"] = follow
    return key


def _pn_anchor(s):
    """Text read from a workbook (a station name, say) as it reads inside an anchor."""
    return _pn_re.sub(r"\s+", "", str(s))


def _pn_cfg(c, *names):
    """The config value stored under the first of `names` present (one quantity that a
    workbook revision may have renamed)."""
    for n in names:
        if n in c:
            return c[n]
    raise KeyError(" or ".join(names))


def _pn_alt(anchors, suffix=""):
    """Alternative wordings of one anchor, each followed by `suffix`."""
    return [a + suffix for a in anchors]


def _pn_free(reason):
    """A cell whose printed numbers no workbook records: allowed only for `reason`."""
    return {"free": reason}


def _pn_table(label, file, source, rows):
    return {"label": label, "file": file, "source": source, "rows": rows}


def _pn_row(row, cells, key=None):
    """One data row: `cells` as described above; `key` finds the row in the book, either
    {"id": n} (the first cell prints the single number n, as B_3 or Expert 3 do) or
    {"text": s} (the first cell reads s); a table whose rows carry no key is compared
    row by row in order."""
    out = {"row": row, "cells": cells}
    if key is not None:
        out["key"] = key
    return out


def _pn_chapter2():
    t21 = load("tab2-1_running_example_percentiles.xlsx")["table_2_1"]
    t25w = load("tab2-5_biobjective_example.xlsx")
    t25, t25c = t25w["table_2_5"], cfgd(t25w)
    ls = t25w["level_set_alpha_star"].set_index("point")
    bnd = t25w["bounds"].set_index("bound")["value"]
    star = t25.iloc[3]
    alpha_star = float(t25c["alpha_star"])
    w = ["w1", "w2", "w3", "w4"]
    f = _pn_f
    tables = [
        _pn_table("tab:elicitation", "chapter2.tex", "tab2-1_running_example_percentiles.xlsx / table_2_1",
                  [_pn_row(str(r["expert"]),
                           [[int(str(r["expert"]).split()[-1])],
                            f(r["p05"]), f(r["p25"]), f(r["p50"]), f(r["p75"]), f(r["p95"]),
                            [f(r["fitted_mu"]), f(r["fitted_sigma"]), 2]],
                           {"id": int(str(r["expert"]).split()[-1])})
                   for _, r in t21.iterrows()]),
        _pn_table("tab:biobj-example", "chapter2.tex", "tab2-5_biobjective_example.xlsx / table_2_5",
                  [_pn_row(str(r["emphasis"]),
                           [[f(r[c]) for c in w],
                            [alpha_star] if i == 3 else None,
                            f(r["A"]), f(r["V"]),
                            [f(r["fQA_mu"]), f(r["fQA_sigma"]), 2]])
                   for i, (_, r) in enumerate(t25.iterrows())]),
    ]
    sec = "sec:foundations-biobjective"
    ch = "chapter2.tex"
    agg = "interiorcompromisewithaggregate\\mathcalN"
    seg = "alinesegmentofthesimplex,herefrom"
    ext = "normalizationusingthesolvedextrema"
    keys = [
        _pn_key(ch, sec, "maximum-consensus variance", "mostdispersedaggregate(",
                "V=" + _pn_fmt(t25.iloc[1]["V"], 2), f(t25.iloc[1]["V"])),
        _pn_key(ch, sec, "minimum-variance variance", "mostdecisiveaggregate(",
                "V=" + _pn_fmt(t25.iloc[2]["V"], 2), f(t25.iloc[2]["V"])),
        _pn_key(ch, sec, "alpha* aggregate", agg, _pn_normal(star["fQA_mu"], star["fQA_sigma"]),
                [f(star["fQA_mu"]), f(star["fQA_sigma"])]),
        _pn_key(ch, sec, "alpha* consensus", agg + "(#,#^2),", "A=" + _pn_fmt(star["A"], 2), f(star["A"])),
        _pn_key(ch, sec, "alpha* confidence", agg + "(#,#^2),A=#,and", "V=" + _pn_fmt(star["V"], 2), f(star["V"])),
        _pn_key(ch, sec, "alpha* level set, endpoint 1", seg,
                _pn_vec(ls.loc["endpoint 1", w]), [f(x) for x in ls.loc["endpoint 1", w]]),
        _pn_key(ch, sec, "alpha* level set, endpoint 2", seg + "(#,#,#,#)to",
                _pn_vec(ls.loc["endpoint 2", w]), [f(x) for x in ls.loc["endpoint 2", w]]),
        _pn_key(ch, sec, "alpha* reported midpoint", "themidpointofthatsegment,\\mathbfw^*=",
                _pn_vec(ls.loc["midpoint (reported)", w]), [f(x) for x in ls.loc["midpoint (reported)", w]]),
        _pn_key(ch, sec, "solved consensus range", ext,
                "A\\in[%s,%s]" % (_pn_fmt(bnd["A_min"], 2), _pn_fmt(bnd["A_max"], 2)), [f(bnd["A_min"]), f(bnd["A_max"])]),
        _pn_key(ch, sec, "solved confidence range", ext + "A\\in[#,#]and",
                "V\\in[%s,%s]" % (_pn_fmt(bnd["V_min"], 2), _pn_fmt(bnd["V_max"], 2)), [f(bnd["V_min"]), f(bnd["V_max"])]),
    ]
    summary = {"alpha_star": alpha_star,
               "alpha_star_aggregate": {"mu": f(star["fQA_mu"]), "sigma": f(star["fQA_sigma"])},
               "alpha_star_weights_midpoint": [f(x) for x in ls.loc["midpoint (reported)", w]],
               "A_range": [f(bnd["A_min"]), f(bnd["A_max"])], "V_range": [f(bnd["V_min"]), f(bnd["V_max"])]}
    return summary, tables, keys


def _pn_chapter4():
    """The illustrative size of the epigraph reformulation (the paragraph after Remark
    rem:lp-qp in Section sec:scalability-convex): every count is computed by the builder
    from the stated N, T and L (fig4-3_scalability.xlsx, sheet lp_reformulation_size)."""
    lp = load("fig4-3_scalability.xlsx")["lp_reformulation_size"]
    v = dict(zip(lp["quantity"].astype(str), lp["value"]))

    def n(q):
        return int(round(float(v[q])))

    # the prose calls the pair block at N_large 'about half a million auxiliary variables'
    if abs(n("pair_variables_at_N_large") / 1e6 - 0.5) > 0.05:
        raise ValueError("the pair block at N = %d is %d variables, not about half a million"
                         % (n("N_large"), n("pair_variables_at_N_large")))
    sec, ch = "sec:scalability-convex", "chapter4.tex"
    rows = (
        ("panel size N", "ForapanelofN=", "N"),
        ("number of outcomes T", "expertsscoredagainstT=", "T"),
        ("grid levels L", "outcomesonanL=", "L"),
        ("pinball epigraph variables TL", "thereareTL=", "pinball_variables_TL"),
        ("N in the pair count", "pinballvariablesand", "N"),
        ("N - 1 in the pair count", "pinballvariablesand#\\times", None),
        ("pair variables N(N-1)/2", "pinballvariablesand#\\times#/2=", "pair_variables"),
        ("variables N + TL + N(N-1)/2", "hence#+#+#=", "variables"),
        ("pinball rows 2TL", "togetherwith#\\times#=", "pinball_rows"),
        ("pair rows N(N-1)", "pinballrowsand#\\times#=", "pair_rows"),
        ("nonzeros per pinball row N + 1", "containsroughly#,#\\times", "nonzeros_per_pinball_row"),
        ("pinball nonzeros", "containsroughly#,#\\times#=", "pinball_nonzeros"),
        ("large panel size", "halfamillionauxiliaryvariablesatN=", "N_large"),
    )
    keys = []
    for what, anchor, q in rows:
        val = n("N") - 1 if q is None else n(q)
        keys.append(_pn_key(ch, sec, what, anchor, "{:,}".format(val), val))
    return None, [], keys


def _pn_chapter4_figures():
    """The numbers the prose of Chapter 4 prints from its two data figures: the range of
    panel sizes of Figure 4.3 and the saving at the largest (Section
    sec:scalability-hierarchical; fig4-3_scalability.xlsx, sheets series and summary), and
    the solved consensus-fairness front of Figure 4.4 (Section sec:scalability-pareto;
    fig4-4_pareto_front.xlsx, sheets summary and printed, the panel size from the weight
    columns of sheet front).  Each value is taken at full precision from the summary and
    must print as the builder's own rounding in sheet printed."""
    ch = "chapter4.tex"
    # ---- Figure 4.3
    sc = load("fig4-3_scalability.xlsx")
    ser = sc["series"]
    sm = cfgd({"config": sc["summary"]})
    n_lo, n_hi = int(ser["N"].iloc[0]), int(ser["N"].iloc[-1])
    ratio = float(sm["ratio_flat_over_hier_at_N_max"])
    if int(float(sm["N_max"])) != n_hi or \
            abs(ratio - int(ser["flat_pairwise"].iloc[-1]) / float(ser["hier_pairwise"].iloc[-1])) > 1e-9:
        raise ValueError("the saving of fig4-3_scalability.xlsx is not the ratio of the counts at the largest panel")
    if int(round(ratio)) != int(float(sm["ratio_printed"])):
        raise ValueError("ratio_printed %r is not the rounded saving %r" % (sm["ratio_printed"], ratio))
    hs = "sec:scalability-hierarchical"
    keys = [
        _pn_key(ch, hs, "Figure 4.3: smallest panel size", "forpanelsizesfromN=", "{:,}".format(n_lo), n_lo),
        _pn_key(ch, hs, "Figure 4.3: largest panel size", "forpanelsizesfromN=#toN=", "{:,}".format(n_hi), n_hi),
        _pn_key(ch, hs, "Figure 4.3: panel size of the printed saving", "withthepanelsize,andatN=",
                "{:,}".format(n_hi), n_hi),
        _pn_key(ch, hs, "Figure 4.3: flat over two-stage interaction count at the largest panel",
                "therepresentative-levelhierarchyrequiresabout", "%d" % int(round(ratio)), ratio),
    ]
    # ---- Figure 4.4
    pf = load("fig4-4_pareto_front.xlsx")
    ps = cfgd({"config": pf["summary"]})
    pp = cfgd({"config": pf["printed"]})
    pa = "sec:scalability-pareto"

    def dec(skey, pkey, d):
        v = float(ps[skey])
        t = _pn_fmt(v, d)
        if abs(float(t) - float(pp[pkey])) > 1e-12:
            raise ValueError("fig4-4 summary %s = %r prints %s, sheet printed has %s = %r" % (skey, v, t, pkey, pp[pkey]))
        return t, v

    runs, to_ = "Thefrontrunsfromthefairnessmaximizer,withmeanoverlap", "totheconsensusmaximizer,with"
    st, gap = "liesbetweenmeanoverlapsof", "byasmuchas"
    ret, rgt = "returningtheendatmeanoverlap", "andtheendat"
    for what, anchor, skey, pkey, d in (
            ("fairness maximizer, mean overlap", runs, "front_Abar_min", "Abar_min_3dp", 3),
            ("fairness maximizer, fairness", runs + "#andfairness", "front_Phi_max", "Phi_max_3dp", 3),
            ("consensus maximizer, mean overlap", to_, "front_Abar_max", "Abar_max_3dp", 3),
            ("consensus maximizer, fairness", to_ + "#and", "front_Phi_min", "Phi_min_3dp", 3),
            ("unsupported stretch, lower end", st, "unsupported_Abar_lo", "stretch_lo_3dp", 3),
            ("unsupported stretch, upper end", st + "#and", "unsupported_Abar_hi", "stretch_hi_3dp", 3),
            ("largest gap below the concave majorant", gap, "max_gap", "max_gap_4dp", 4),
            ("mean overlap at the largest gap", gap + "#at\\barA=", "max_gap_Abar", "max_gap_Abar_3dp", 3),
            ("fairness range of the front", "againstafairnessrangeof", "fairness_range", "fairness_range_3dp", 3),
            ("last weighted sum below the stretch, mean overlap", ret, "ws_last_left_Abar", "stretch_lo_3dp", 3),
            ("last weighted sum below the stretch, lambda_1", ret + "#at\\lambda_1=", "ws_last_left_lambda",
             "ws_last_left_lambda_2dp", 2),
            ("first weighted sum above the stretch, mean overlap", rgt, "ws_first_right_Abar", "stretch_hi_3dp", 3),
            ("first weighted sum above the stretch, lambda_1", rgt + "#at\\lambda_1=", "ws_first_right_lambda",
             "ws_first_right_lambda_2dp", 2),
            ("weighting at which the weighted sum ties the two ends", "indifferentbetweenthemat\\lambda_1=",
             "lambda_star", "lambda_star_3dp", 3)):
        t, v = dec(skey, pkey, d)
        keys.append(_pn_key(ch, pa, "Figure 4.4: %s" % what, anchor, t, v))
    for what, anchor, key in (("equally spaced floors", "Itissolvedat", "grid_floors"),
                              ("refinement floors", "Itissolvedat#equallyspacedfloorsandat", "refinement_floors"),
                              ("weighted sums", "simplex,andofthe", "ws_solutions")):
        n = int(float(ps[key]))
        if n != int(float(pp[key])):
            raise ValueError("fig4-4 summary %s = %r, printed %r" % (key, ps[key], pp[key]))
        keys.append(_pn_key(ch, pa, "Figure 4.4: %s" % what, anchor, "%d" % n, n))
    # the counts the paragraph writes in words: the panel size, the further dents (at the
    # start of a sentence), the first weightings that return the fairness maximizer
    n_panel = sum(1 for c in pf["front"].columns if _pn_re.fullmatch(r"w\d+", str(c)))
    n_dents = int(float(pp["other_dents"]))
    if n_dents != int(float(ps["n_bridging_edges"])) - 1:
        raise ValueError("fig4-4 printed other_dents %r is not the bridging edges besides the stretch" % pp["other_dents"])
    n_fmax = int(float(ps["ws_at_fairness_maximizer"]))
    if n_fmax != int(float(pp["ws_at_fairness_maximizer"])):
        raise ValueError("fig4-4 summary and printed disagree on the weightings at the fairness maximizer")
    keys += [
        _pn_key(ch, pa, "Figure 4.4: experts of the panel (in words)", "showsthedistinctiononafrontcomputedforthe",
                _pn_words(n_panel), n_panel, follow="-expertfairnesspanel"),
        _pn_key(ch, pa, "Figure 4.4: further dents (in words)", "yeteveryefficientpointinsideitisunsupported.",
                _pn_words(n_dents).capitalize(), n_dents, follow="furtherdents"),
        _pn_key(ch, pa, "Figure 4.4: first weightings at the fairness maximizer (in words)", "Thefirst",
                _pn_words(n_fmax), n_fmax, follow="weightingsallreturnthefairnessmaximizer"),
    ]
    return None, [], keys


def _pn_chapter5():
    lo = load("fig5-3_loo.xlsx")
    rs = dict(zip(lo["results"]["quantity"].astype(str), lo["results"]["value"].astype(float)))
    ho, co = rs["human-only mean CRPS"], rs["collaborative mean CRPS"]
    wh, wm = rs["human weight share W_H"], rs["machine weight share W_M"]
    gain = rs["CRPS gain (collab vs human-only, pct)"]
    lc = cfgd(lo)
    nh, nm = int(lc["n_human"]), int(lc["n_machine"])
    tables = [_pn_table("tab:hmpanel", "chapter5.tex", "fig5-3_loo.xlsx / results, config (rounded copy in tab5-3_hm_panel.xlsx)", [
        _pn_row("human-only aggregate", [None, None, ho], {"text": "human-only aggregate"}),
        _pn_row("collaborative aggregate", [None, None, [co, {"value": gain, "decimals": 0}]],
                {"text": "collaborative aggregate"}),
        _pn_row("human subpanel", [[nh], wh, None], {"text": "human subpanel"}),
        _pn_row("machine subpanel", [[nm], wm, None], {"text": "machine subpanel"}),
    ])]
    sec, ch = "sec:collab-application", "chapter5.tex"
    keys = [
        _pn_key(ch, sec, "human weight share (percent)", "assignsabout", "%d\\%%" % round(100 * wh), wh),
        _pn_key(ch, sec, "machine weight share (percent)", "oftheweighttothehumanexpertsand",
                "%d\\%%" % round(100 * wm), wm),
        _pn_key(ch, sec, "CRPS gain (percent)", "meanCRPSabout",
                "%d\\%%" % round(abs(gain)), gain),
        _pn_key(ch, sec, "human-only mean CRPS", "sothatthehuman-onlyaggregatescores", _pn_fmt(ho, 2), ho),
    ]
    # the SLSQP cross-check of every solve (config entry verification of fig5-3_loo.xlsx)
    ver = _pn_re.search(r"plus (\d+) Dirichlet\(1\) starts under default_rng\((\d+)\).*?"
                        r"value by more than ([0-9.e+-]+)", str(lc["verification"]))
    if ver is None:
        raise ValueError("the verification entry of fig5-3_loo.xlsx no longer states its starts, seed and gap")
    n_ver, s_ver, g_ver = int(ver.group(1)), int(ver.group(2)), float(ver.group(3))
    keys += [
        _pn_key(ch, sec, "random SLSQP starts per solve", "fromtheequal-weightvectorand", _pn_words(n_ver), n_ver,
                follow="randomsimplexstarts"),
        _pn_key(ch, sec, "seed of the SLSQP starts", "randomsimplexstartsunderseed", "%d" % s_ver, s_ver),
        _pn_key(ch, sec, "bound on the SLSQP improvement", "improvesonthereturnedvaluebymorethan",
                _pn_pow10(g_ver), g_ver),
    ]
    # Section sec:collab-loo prints the finite-grid reliability floor of a perfectly
    # calibrated normal forecast (per unit standard deviation) at four grid sizes, at two
    # significant figures, and says that most of it comes from the two outer bins
    rf = lo["reliability_floor"]
    if not (rf["share_from_outer_bins"].astype(float) > 0.5).all():
        raise ValueError("the outer bins carry less than half of the reliability floor at some L")
    fl = "standarddeviationofaperfectlycalibratednormalforecastatL="
    for i, (_, r) in enumerate(rf.iterrows()):
        sep = "" if i == 0 else "#," * i + ("and" if i == len(rf) - 1 else "")
        keys.append(_pn_key(ch, "sec:collab-loo", "reliability floor at L = %d (per unit sd)" % int(r["L"]),
                            "thatfloorisabout" + sep, "%.2g" % float(r["REL_floor_per_unit_sd"]),
                            float(r["REL_floor_per_unit_sd"])))
        keys.append(_pn_key(ch, "sec:collab-loo", "grid size of the reliability floor (%d)" % int(r["L"]),
                            fl + sep, "%d" % int(r["L"]), int(r["L"])))
    return None, tables, keys


def _pn_case_i():
    f = _pn_f
    bim = load("fig6-1_bim_case.xlsx")
    dims, der = bim["dimensions"], bim["derived"]
    glob = der[der["quantity"].astype(str).str.contains("global maturity")].iloc[0]
    t69 = load("tab6-9_bim_results.xlsx")["table_6_9"].set_index("quantity")
    tr = load("fig6-4_bim_tradeoff.xlsx")
    ser = tr["series"].set_index("delta")
    wts = tr["weights"].set_index("delta")

    def normal_of(s):
        m = _pn_re.match(r"\s*N\(\s*([-0-9.eE]+)\s*,\s*([-0-9.eE]+)\s*\^2\s*\)", str(s))
        return [float(m.group(1)), float(m.group(2)), 2]

    # Table 6.9: the dissenter is named in its row label by the id the workbook records
    # ("weight on dissenter (exp N)"), which must be the expert of the fairness panel with
    # the lowest stated mean
    qdis = [q for q in t69.index if str(q).startswith("weight on dissenter")]
    if len(qdis) != 1:
        raise ValueError("tab6-9 names %d dissenter rows" % len(qdis))
    dis = int(_pn_re.search(r"(\d+)\D*$", str(qdis[0])).group(1))
    trc = cfgd(tr)
    if "panel_mu" in trc:
        pmu = [float(x) for x in str(trc["panel_mu"]).split(",")]
        if pmu.index(min(pmu)) + 1 != dis:
            raise ValueError("tab6-9 names expert %d as the dissenter, the lowest stated mean is expert %d"
                             % (dis, pmu.index(min(pmu)) + 1))
    cols = ["consensus_confidence_d0", "fairness_weighted_d0.5"]
    rows69 = [
        _pn_row(lab, [lab] + [get(t69.loc[q, c]) for c in cols], {"text": key})
        for lab, key, q, get in (
            ("aggregate distribution", "aggregate distribution", "aggregate distribution", normal_of),
            ("weight on dissenter (expert %d)" % dis, "weight on dissenter", qdis[0], f),
            ("minimum overlap", "minimum overlap", "minimum overlap", f),
            ("mean overlap (consensus)", "mean overlap", "mean overlap (consensus)", f),
            ("collective fairness", "collective fairness", "collective fairness", f))]
    # Table 6.8: where the workbook records each dimension's maturity band, the band
    # column is compared, the weakest and the strongest dimension (lowest and highest
    # mean) carrying that note
    mus = dims["mu"].astype(float)
    lo8, hi8 = int(mus.idxmin()), int(mus.idxmax())

    def band(i, r):
        if "band" not in dims.columns:
            return None
        return str(r["band"]).lower() + (" (weakest)" if i == lo8 else " (strongest)" if i == hi8 else "")

    rows68 = [_pn_row(str(r["dimension"]),
                      [str(r["dimension"]), f(r["weight"]), [f(r["mu"]), f(r["sigma"]), 2], band(i, r)],
                      {"text": str(r["dimension"]).lower()})
              for i, r in dims.iterrows()]
    rows68.append(_pn_row("global", [None, None, [f(glob["mu"]), f(glob["sigma"]), 2], None], {"text": "global"}))
    # Table 6.6: the elicited profile of q4, and, where the workbook records the band
    # probabilities of its normal fit, the fit row (the fit's N(mu, sigma^2) in its label)
    q4 = bim["elicitation_q4"]
    fit = der[der["quantity"].astype(str).str.contains("tab:bim-elicit")].iloc[0]
    rows66 = [_pn_row("elicited profile", [None] + [f(p) for p in q4["probability"]])]
    if "fit_probability" in q4.columns:
        rows66.append(_pn_row("normal fit", [[f(fit["mu"]), f(fit["sigma"]), 2]]
                              + [f(p) for p in q4["fit_probability"]]))
    # Table 6.7: group, profile, size, Omega, and the total row; the profile ranges and the
    # sizes are compared where the workbook records them
    grp = bim["groups"]
    rng = [c for c in grp.columns if str(c).startswith("range_")]

    def gid(g):
        return int(_pn_re.search(r"G(\d+)", str(g)).group(1))

    # (a workbook without them states the profiles and sizes as labels of the case
    # description, which the check then reports as not backed by the workbook)
    rows67 = [_pn_row(str(r["group"]),
                      [[gid(r["group"])],
                       [int(x) for c in rng for x in str(r[c]).split("-")] if rng else
                       _pn_free("the workbook records no group profiles, which the case description states"),
                       int(r["size"]) if "size" in grp.columns else
                       _pn_free("the workbook records no group sizes, only the panel size (config n_experts)"),
                       f(r["Omega"])],
                      {"id": gid(r["group"])})
              for _, r in grp.iterrows()]
    n_all = int(grp["size"].sum()) if "size" in grp.columns else int(cfgd(bim)["n_experts"])
    rows67.append(_pn_row("total", [None, n_all, f(grp["Omega"].sum())], {"text": "total"}))
    tables = [
        _pn_table("tab:bim-elicit", "chapter6.tex", "fig6-1_bim_case.xlsx / elicitation_q4, derived", rows66),
        _pn_table("tab:bim-groups", "chapter6.tex", "fig6-1_bim_case.xlsx / groups, config", rows67),
        _pn_table("tab:bim-dimensions", "chapter6.tex", "fig6-1_bim_case.xlsx / dimensions, derived", rows68),
        _pn_table("tab:bim", "chapter6.tex", "tab6-9_bim_results.xlsx / table_6_9 (the delta = 0 and 0.5 rows of fig6-4_bim_tradeoff.xlsx)", rows69),
    ]
    sec, ch = "sec:bim-results", "chapter6.tex"
    # Section 6.2.7 compares the global maturity with its value under the equal dimension
    # weights of the published application: further inside the same band, because the
    # best--worst weights give the weakest dimension the largest weight
    bc = cfgd(bim)
    eqmu = f(bc["global_mu_equal_dimension_weights"])
    eqsd = f(bc["global_sigma_equal_dimension_weights"])
    if not (int(eqmu) == int(f(glob["mu"])) and eqmu > f(glob["mu"])):
        raise ValueError("the equal-weight global %.4f is not further inside the band of %.4f"
                         % (eqmu, f(glob["mu"])))
    if int(dims["weight"].astype(float).idxmax()) != int(dims["mu"].astype(float).idxmin()):
        raise ValueError("the dimension with the largest best--worst weight is not the weakest")
    d0, d5, d6 = ser.loc[0.0], ser.loc[0.5], ser.loc[0.6]
    wcols = ["w%d" % i for i in range(1, 8)]
    q4c = der[der["quantity"].astype(str).str.contains("q4 collective")].iloc[0]
    keys = [
        _pn_key(ch, "sec:bim-elicit", "normal fit of the q4 profile", "bandmidpointsreturns\\mathcalN",
                _pn_normal(fit["mu"], fit["sigma"]), [f(fit["mu"]), f(fit["sigma"])]),
        _pn_key(ch, "sec:bim-hierarchy", "inter-group credibility weights Omega",
                "credibilityweights\\boldsymbol\\Omega=", _pn_vec(grp["Omega"], 2), [f(x) for x in grp["Omega"]]),
        _pn_key(ch, "sec:bim-hierarchy", "q4 collective", "returnsthecollective\\mathcalN",
                _pn_normal(q4c["mu"], q4c["sigma"]), [f(q4c["mu"]), f(q4c["sigma"])]),
        _pn_key(ch, sec, "full assessment, global maturity", "theglobalmaturityis\\mathcalN",
                _pn_normal(glob["mu"], glob["sigma"]), [f(glob["mu"]), f(glob["sigma"])]),
        _pn_key(ch, sec, "full assessment, global maturity under the published equal dimension weights",
                "equaldimensionweightsof\\citetchen2023bimdcitwouldbe\\mathcalN",
                _pn_normal(eqmu, eqsd), [eqmu, eqsd]),
        _pn_key(ch, sec, "fairness panel, delta = 0 aggregate", "givingtheaggregate\\mathcalN",
                _pn_normal(d0["agg_mu"], d0["agg_sd"]), [f(d0["agg_mu"]), f(d0["agg_sd"])]),
        _pn_key(ch, sec, "fairness panel, delta = 0 overlap of the dissenter (minimum overlap)",
                ["overlapwiththeaggregateisonly", "thatexpert'soverlapwiththeaggregateis"],
                _pn_fmt(d0["minimum_overlap"], 3), f(d0["minimum_overlap"])),
        _pn_key(ch, sec, "fairness panel, delta = 0.5 weight vector", "thesolutionis\\mathbfw=",
                _pn_vec(wts.loc[0.5, wcols]), [f(x) for x in wts.loc[0.5, wcols]]),
        _pn_key(ch, sec, "fairness panel, delta = 0.5 aggregate",
                "theremainderonexpert#,andtheaggregatemovesto\\mathcalN",
                _pn_normal(d5["agg_mu"], d5["agg_sd"]), [f(d5["agg_mu"]), f(d5["agg_sd"])]),
        _pn_key(ch, sec, "fairness panel, delta = 0.6 weight on the dissenter", "thedissenter'sweightrisesto",
                _pn_fmt(d6["weight_on_dissenter"], 3), f(d6["weight_on_dissenter"])),
        _pn_key(ch, sec, "fairness panel, delta = 0.6 aggregate", "andtheaggregatefallsto\\mathcalN",
                _pn_normal(d6["agg_mu"], d6["agg_sd"]), [f(d6["agg_mu"]), f(d6["agg_sd"])]),
    ]
    # the threshold of the delta sweep: the last grid point at which the dissenter still has
    # no weight and the first at which it has some (Section 6.2.7 twice, Table 6.22)
    wd_ = ser["weight_on_dissenter"].astype(float)
    lo_ = max(float(d) for d in wd_.index if wd_[d] < 1e-9)
    hi_ = min(float(d) for d in wd_.index if wd_[d] >= 1e-9)
    if not (hi_ > lo_ and all(wd_[d] < 1e-9 for d in wd_.index if d <= lo_)
            and all(wd_[d] >= 1e-9 for d in wd_.index if d >= hi_)):
        raise ValueError("the dissenter's weight does not switch on at a single threshold of the delta grid")
    for sc_, sec_, an_ in (("prose", sec, "risespastathresholdbetween"), ("prose", sec, "onlypastathresholdbetween"),
                           ("table:tab:robustness", "sec:app-synthesis", "dissenteronlypastathresholdbetween")):
        keys += [_pn_key(ch, sec_, "fairness panel, delta threshold, last grid point without dissenter weight",
                         an_, _pn_fmt(lo_, 1), lo_, scope=sc_),
                 _pn_key(ch, sec_, "fairness panel, delta threshold, first grid point with dissenter weight",
                         an_ + "#and", _pn_fmt(hi_, 1), hi_, scope=sc_)]
    summary = {
        "case": "I",
        "title": "Building information modeling maturity with a fairness objective",
        "published_application": "chen2023bimdc",
        "full_assessment": {"groups": [{"group": str(r["group"]), "Omega": f(r["Omega"])} for _, r in grp.iterrows()],
                            "q4_collective": {"mu": f(q4c["mu"]), "sigma": f(q4c["sigma"])},
                            "q4_profile_fit": {"mu": f(fit["mu"]), "sigma": f(fit["sigma"])},
                            "dimensions": [{"dimension": str(r["dimension"]), "weight": f(r["weight"]),
                                            "mu": f(r["mu"]), "sigma": f(r["sigma"])} for _, r in dims.iterrows()],
                            "global": {"mu": f(glob["mu"]), "sigma": f(glob["sigma"])}},
        "fairness_panel": {str(d): {"aggregate": [f(ser.loc[d, "agg_mu"]), f(ser.loc[d, "agg_sd"])],
                                    "weight_on_dissenter": f(ser.loc[d, "weight_on_dissenter"]),
                                    "minimum_overlap": f(ser.loc[d, "minimum_overlap"]),
                                    "mean_overlap": f(ser.loc[d, "consensus_Abar"]),
                                    "collective_fairness": f(ser.loc[d, "fairness_Phibar"])}
                           for d in ser.index},
    }
    return summary, tables, keys


def _pn_case_ii():
    f = _pn_f
    dt = load("fig6-5_dt_case.xlsx")
    mx = dt["matrix"].set_index("id")
    order = list(dt["ranking"].sort_values("rank")["id"].astype(str))
    eq = ["EQ1", "EQ2", "EQ3", "EQ4", "EQ5"]
    rows = [_pn_row(b, [[int(b[1:])], str(mx.loc[b, "dimension"])]
                    + [f(mx.loc[b, e]) for e in eq] + [f(mx.loc[b, "avg"]), int(mx.loc[b, "rank"])],
                    {"id": int(b[1:])})
            for b in order]
    # Table 6.11: the elicited profile of B1 and its truncated-Weibull fit (kappa, lambda,
    # mean), stored in the config sheet as elicit_fit_* (or, in older workbooks, b1_fit_*)
    dtc = cfgd(dt)
    fit = [f(_pn_cfg(dtc, "elicit_fit_" + p, "b1_fit_" + p)) for p in ("kappa", "lambda", "mean")]
    tables = [
        _pn_table("tab:dt-elicit", "chapter6.tex", "fig6-5_dt_case.xlsx / elicitation_b1, config elicit_fit_*", [
            _pn_row("elicited profile", [None] + [f(p) for p in dt["elicitation_b1"]["probability"]]),
            _pn_row("fitted truncated Weibull", [None, fit]),
        ]),
        _pn_table("tab:dt-swf-results", "chapter6.tex", "fig6-5_dt_case.xlsx / matrix", rows),
    ]
    sec, ch = "sec:dt-results", "chapter6.tex"
    # the anchors name the barrier the workbook ranks at that place, so a changed ranking
    # is caught as well as a changed mean
    anchors = ["(B_%s)ranksfirstat\\bar\\mu=" % order[0][1:]] + ["(B_%s)at" % b[1:] for b in order[1:4]] \
        + ["(B_%s," % b[1:] for b in order[-3:]]
    keys = [_pn_key(ch, sec, "average collective mean of %s" % b, a, _pn_fmt(mx.loc[b, "avg"], 2), f(mx.loc[b, "avg"]))
            for b, a in zip(order[:4] + order[-3:], anchors)]
    keys.append(_pn_key(ch, "sec:app-synthesis", "largest spread of a barrier mean across the five indices",
                        ["barrierbyatmost", "barrierbyupto"], _pn_fmt(mx["spread"].max(), 2), f(mx["spread"].max()),
                        scope="table:tab:robustness"))
    summary = {
        "case": "II",
        "title": "Digital-transformation barriers in SMCEs under five inequality indices",
        "published_application": "chen2024socialwelfare",
        "ranking": order,
        "average_collective_mean": {b: f(mx.loc[b, "avg"]) for b in order},
        "max_spread_across_indices": {"value": f(mx["spread"].max()), "barrier": str(mx["spread"].idxmax())},
    }
    return summary, tables, keys


def _pn_case_iii():
    f = _pn_f
    b3 = load("fig6-8_bdt_case.xlsx")
    b3c = cfgd(b3)
    grp, dim, sen = b3["groups"], b3["dimensions"], b3["sensitivity"]
    q1c = sum(float(r["Omega"]) * float(r["q1_mu"]) for _, r in grp.iterrows())
    # Table 6.16: the groups and the total row; Table 6.17: the dimensions and the global
    # row, the level column carrying the lowest / highest note of the extreme dimensions
    grp_rows = [_pn_row(str(r["group"]),
                        [[int(str(r["group"])[1:])], str(r["stratum"]), int(r["size"]), f(r["Omega"])],
                        {"id": int(str(r["group"])[1:])})
                for _, r in grp.iterrows()]
    grp_rows.append(_pn_row("total", [None, int(grp["size"].sum()), f(grp["Omega"].sum())], {"text": "total"}))
    lo, hi = int(dim["mu"].astype(float).idxmin()), int(dim["mu"].astype(float).idxmax())
    dim_rows = []
    for i, r in dim.iterrows():
        level = str(r["level"]) + (" (lowest)" if i == lo else " (highest)" if i == hi else "")
        dim_rows.append(_pn_row(str(r["id"]), [[int(str(r["id"])[1:])], f(r["weight"]), f(r["mu"]), level],
                                {"id": int(str(r["id"])[1:])}))
    dim_rows.append(_pn_row("global", [None, None, f(b3c["global_maturity"]), str(b3c["global_level"])],
                            {"text": "global"}))
    tables = [
        _pn_table("tab:bdt-groups", "chapter6.tex", "fig6-8_bdt_case.xlsx / groups", grp_rows),
        _pn_table("tab:bdt-results", "chapter6.tex", "fig6-8_bdt_case.xlsx / dimensions, config", dim_rows),
    ]
    res, dat, ch = "sec:bdt-results", "sec:bdt-data", "chapter6.tex"
    means = "groupopinionshavemeans"
    keys = [_pn_key(ch, res, "q1 mean of %s" % r["group"], a, _pn_fmt(r["q1_mu"], 2), f(r["q1_mu"]))
            for a, (_, r) in zip([means, means + "#,", means + "#,#,and"], grp.iterrows())]
    sd = "memberstandarddeviationsbetween"
    av = "globaldistributionliesbetween"
    keys += [
        _pn_key(ch, res, "q1 collective mean", "collectivedistributionwithmean", _pn_fmt(q1c, 2), q1c),
        _pn_key(ch, res, "highest dimension mean", "standinghighest(", _pn_fmt(dim["mu"].max(), 2), f(dim["mu"].max())),
        _pn_key(ch, res, "lowest dimension mean", "lowest(", _pn_fmt(dim["mu"].min(), 2), f(dim["mu"].min())),
        _pn_key(ch, res, "global maturity", "equal-weightoperatingpoint,is",
                _pn_fmt(b3c["global_maturity"], 2), f(b3c["global_maturity"])),
        _pn_key(ch, res, "collective fairness", "withacollectivefairnessof",
                _pn_fmt(b3c["fairness_Phi"], 3), f(b3c["fairness_Phi"])),
        _pn_key(ch, res, "weight Gini", ",aweightGiniof", _pn_fmt(b3c["weight_gini"], 2), f(b3c["weight_gini"])),
        _pn_key(ch, res, "aggregate variance", "andanaggregatevarianceof",
                _pn_fmt(b3c["aggregate_variance"], 3), f(b3c["aggregate_variance"])),
        _pn_key(ch, dat, "inter-group weights Omega", "clustercentroidsgives\\boldsymbol\\Omega=",
                _pn_vec(grp["Omega"], 2), [f(x) for x in grp["Omega"]]),
        _pn_key(ch, dat, "smallest member sd", sd, _pn_fmt(b3c["member_sd_min"], 2), f(b3c["member_sd_min"])),
        _pn_key(ch, dat, "largest member sd", sd + "#and", _pn_fmt(b3c["member_sd_max"], 2), f(b3c["member_sd_max"])),
        _pn_key(ch, dat, "mean member sd", sd + "#and#,withmean",
                _pn_fmt(b3c["member_sd_mean"], 2), f(b3c["member_sd_mean"])),
        _pn_key(ch, dat, "smallest aggregate variance over the sweep", av,
                _pn_fmt(sen["aggregate_variance"].min(), 2), f(sen["aggregate_variance"].min())),
        _pn_key(ch, dat, "largest aggregate variance over the sweep", av + "#and",
                _pn_fmt(sen["aggregate_variance"].max(), 2), f(sen["aggregate_variance"].max())),
    ]
    # the second set of solver starts is drawn from a generator of its own, seeded with
    # seed + probe_seed_offset, which Section 6.4.6 prints twice
    s2 = int(float(b3c["seed"])) + int(float(b3c["probe_seed_offset"]))
    if int(float(b3c["probe_seed"])) != s2:
        raise ValueError("probe_seed %r is not seed + probe_seed_offset" % (b3c["probe_seed"],))
    keys += [
        _pn_key(ch, dat, "seed of the second set of starts (draw order)", "describedbelow,whichtheseed", "%d" % s2, s2),
        _pn_key(ch, dat, "seed of the second set of starts", "separategeneratorseededwith", "%d" % s2, s2),
        _pn_key(ch, dat, "largest gain of the further starts over the best original start (upper bound)",
                "thisoneamongthem,byatmost", _pn_sci(up2(b3c["further_max_raise"])), up2(b3c["further_max_raise"])),
    ]
    # the accuracy sentence rests on the local-polish diagnostic: its settled largest gain
    # and largest weight move (upper bounds at two significant figures), its round
    # tolerance, the order of magnitude it supports, and the claim that substituting its
    # end points (or the best original starts) changes no printed result
    if int(float(b3c["polish_n_programs_at_round_cap"])) != 0:
        raise ValueError("a local-polish program stopped at the round cap")
    if math.floor(math.log10(float(b3c["retained_accuracy_bound"]))) != -5:
        raise ValueError("the retained accuracy %r is not of the order of 1e-5" % (b3c["retained_accuracy_bound"],))
    for sw in ("swap_polished_n_printed_changed", "swap_original_n_printed_changed"):
        if int(float(b3c[sw])) != 0:
            raise ValueError("%s = %r, but the text says it changes no reported result" % (sw, b3c[sw]))
    keys += [
        _pn_key(ch, dat, "local polish, round tolerance", "untilaroundgainsnomorethan",
                _pn_pow10(b3c["polish_round_tol"]), float(b3c["polish_round_tol"])),
        _pn_key(ch, dat, "local polish, largest gain (upper bound)", "raisesnoretainedvaluebymorethan",
                _pn_sci(up2(b3c["polish_max_gain"])), up2(b3c["polish_max_gain"])),
        _pn_key(ch, dat, "local polish, largest member-weight move (upper bound)", "movesnomemberweightbymorethan",
                _pn_sci(up2(b3c["polish_max_weight_diff"])), up2(b3c["polish_max_weight_diff"])),
    ]
    # the robustness paragraph of Section 6.4.7 and the Case III half of the fairness-weight
    # row of Table 6.22: the global maturity moves by so many hundredths of a level (in words)
    # as delta is swept from 0 to the equal-weight operating point.  The move is the change
    # from delta = 0 to the operating point; it must round to the same hundredths as the range
    # of the global maturity over the sweep, else "moves by" would misstate the sweep
    gm = sen.set_index(sen["delta"].astype(float).round(12))["global_maturity"].astype(float)
    d_op = float(b3c["operating_delta"])
    grid = sorted(float(x) for x in str(b3c["deltas"]).split(","))
    if grid[0] != 0.0 or grid[-1] != d_op or sorted(gm.index) != grid:
        raise ValueError("the Case III sensitivity sweep is not delta = 0 up to the operating point %r" % d_op)
    mv = float(gm[d_op] - gm[0.0])
    rg = float(gm[gm.index <= d_op].max() - gm[gm.index <= d_op].min())
    if round(100 * mv) != round(100 * rg):
        raise ValueError("the global maturity moves by %.4f from delta = 0 to %g but ranges over %.4f on the sweep"
                         % (mv, d_op, rg))
    for sc_, sec_, an_ in (("prose", res, "Theglobalmaturitymovesby"),
                           ("table:tab:robustness", "sec:app-synthesis", "theglobalmaturitymovesby")):
        keys.append(_pn_key(ch, sec_, "global maturity change over the fairness-weight sweep (hundredths, in words)",
                            an_, _pn_words(round(100 * mv)), mv, scope=sc_, follow="hundredthsofalevel"))
    return None, tables, keys


def _pn_case_iv():
    f = _pn_f
    r4 = load("fig6-11_rt_case.xlsx")
    r4c = cfgd(r4)
    rr4, st4 = r4["results"], r4["stations"]
    pair = [t.strip() for t in str(r4c["swapped"]).split(",")]
    cont, nb = pair[0], pair[1]
    by = rr4.set_index("station")
    rows_st = [_pn_row(str(r["id"]), [[int(str(r["id"])[1:])], str(r["name"]),
                                      f(r["ops"]), f(r["pas"]), f(r["adv"])],
                       {"id": int(str(r["id"])[1:])})
               for _, r in st4.iterrows()]
    rows_rs = [_pn_row(str(r["station"]),
                       [[int(str(r["station"])[1:])], str(r["name"]),
                        f(r["mean_eff"]), f(r["mean_eq"]), int(r["priority_eff"]), int(r["priority_eq"]),
                        f(r["adv_share_eff"]), f(r["adv_share_eq"]),
                        f(r["weight_gini_eff"]), f(r["weight_gini_eq"])],
                       {"id": int(str(r["station"])[1:])})
               for _, r in rr4.iterrows()]
    tables = [
        _pn_table("tab:rt-stations", "chapter6.tex", "fig6-11_rt_case.xlsx / stations", rows_st),
        _pn_table("tab:rt-results", "chapter6.tex", "fig6-11_rt_case.xlsx / results (_eff = consensus-only)", rows_rs),
    ]
    res, rob, ch = "sec:rt-results", "sec:rt-robust", "chapter6.tex"

    def k(sec, what, anchor, v, d=2):
        return _pn_key(ch, sec, what, anchor, _pn_fmt(v, d), f(v))

    # the station names and the neighbor's id in the anchors are read from the workbook
    cname, nname = _pn_anchor(by.loc[cont, "name"]), _pn_anchor(by.loc[nb, "name"])
    gini = "selectedweightvectoris"
    swap = "differenceonwhichtheswapdepends,"
    # alternative wordings of one sentence (a revision of Section 6.5 names the shares
    # selected ones and restates the leveling-down sentence)
    util = ["meanutilityrisesatthecontestedstationfrom", "contestedstation,wherethemeanutilityrisesfrom"]
    ugini = "theGinioftheutilitiesfallsfrom"
    keys = [
        k(res, "contested station, consensus-only mean", cname + "aggregatesto", by.loc[cont, "mean_eff"]),
        k(res, "contested station, equity-aware mean", "thecollectivemeanfallsto", by.loc[cont, "mean_eq"]),
        k(res, "swap neighbor, consensus-only mean", "%s(S_%s)at" % (nname, nb[1:]), by.loc[nb, "mean_eff"]),
        k(res, "swap neighbor, equity-aware mean", nname + "'s", by.loc[nb, "mean_eq"]),
        k(res, "contested station, consensus-only advocate share",
          ["theadvocatesreceivingaweightshareof", "theadvocatesreceivingaselectedweightshareof"],
          by.loc[cont, "adv_share_eff"]),
        k(res, "contested station, equity-aware advocate share", ["theirsharerisingto", "theirselectedsharerisingto"],
          by.loc[cont, "adv_share_eq"]),
        k(res, "average weight Gini, consensus-only", gini, r4c["weight_gini_eff_avg"]),
        k(res, "average weight Gini, equity-aware", gini + "#undertheconsensusregimeand", r4c["weight_gini_eq_avg"]),
        k(res, "largest uncontested mean gap", "collectivemeansdifferingbyatmost", r4c["max_uncontested_mean_gap"]),
        k(rob, "swap margin, neighbor equity-aware mean", swap, by.loc[nb, "mean_eq"]),
        k(rob, "swap margin, contested equity-aware mean", swap + "#against", by.loc[cont, "mean_eq"]),
        k(rob, "contested mean utility, consensus-only", util, by.loc[cont, "Ubar_eff"]),
        k(rob, "contested mean utility, equity-aware", _pn_alt(util, "#to"), by.loc[cont, "Ubar_eq"]),
        k(rob, "contested utility Gini, consensus-only", ugini, by.loc[cont, "gini_U_eff"]),
        k(rob, "contested utility Gini, equity-aware", ugini + "#to", by.loc[cont, "gini_U_eq"]),
    ]
    # the rest of Section 6.5.6 (sec:rt-robust), read from the robustness, near-tie,
    # certification and results sheets and the config: the robustness paragraph (the
    # swap margin and its re-seed spread, the re-seeded panels, the Fehr-Schmidt grid and
    # the range over it of the contested advocate share), the near-tie and leveling-down
    # paragraphs, the early-warning threshold and the closing box.  A count the text
    # writes as a word is compared as that word; an "at most" or "less than" bound is the
    # recorded value rounded up at the printed precision.
    rs4, fs4, nt4, ce4 = r4["robustness_seeds"], r4["robustness_fs"], r4["near_ties"], r4["certification"]
    n_re, n_st, n_fs = len(rs4), len(rr4), len(fs4)
    n_dir = int(rs4["swap_direction_preserved"].sum())
    n_oth = int(rs4["other_orders_unchanged"].sum())
    n_two = int((rs4["contested_places_moved"] == 2).sum())
    if n_dir != n_re or int(fs4["swap_direction_preserved"].sum()) != n_fs or n_oth + n_two != n_re:
        raise ValueError("the re-seeded or Fehr-Schmidt panels no longer bear out Section 6.5.6")
    seeds = sorted(int(s) for s in rs4["seed"])
    shr = fs4["S4_adv_share_eq"].astype(float)
    tie_st, tie_rg = str(r4c["near_tie_closest"]).split()
    ml = sorted(float(m) for m in nt4.loc[(nt4["station"].astype(str) == tie_st)
                                            & (nt4["regime"].astype(str) == tie_rg), "m_local_max"])
    if len(ml) != 2:
        raise ValueError("Section 6.5.6 prints two local maxima at the closest near-tie, the workbook has %d" % len(ml))
    tied = int(ce4["near_tie"].astype(int).sum())
    lvl = [int(v) for v in _pn_re.findall(r"\d+", str(r4c["overlap_leveling_down_combo_count_range"]))]
    wgr = {g: [float(v) for v in _pn_re.findall(r"\d+\.\d+", str(r4c["near_tie_weight_gini_%s_avg_range" % g]))]
           for g in ("eff", "eq")}
    wname = _pn_anchor(by.loc[str(r4c["max_uncontested_mean_gap_station"]), "name"])
    n_adv = int(r4["constituencies"].set_index("key").loc["adv", "size"])
    n_sh = int(float(r4c["n_stakeholders"]))

    def grid(col):
        vals = sorted(set(float(x) for x in fs4[col]))
        return ",".join("%g" % v for v in vals), vals

    def fm(x, d=2):
        return _pn_fmt(x, d)

    wd = _pn_words
    cst = by.loc[cont]
    rows6 = [
        # (what, anchor, text as printed, value, the words printed after a count in words)
        ("swap margin", swap + "#against#,is", fm(r4c["eq_swap_margin"]), f(r4c["eq_swap_margin"]), None),
        ("re-seed sd of the swap margin", "morethanthreetimesitsstandarddeviationof",
         fm(r4c["reseed_eq_swap_margin_sd"]), f(r4c["reseed_eq_swap_margin_sd"]), None),
        ("re-seeded panels behind the sd", "morethanthreetimesitsstandarddeviationof#acrossthe",
         wd(n_re), n_re, "re-seededpanels"),
        ("re-seeded panels", "Regeneratingthepanelunder", wd(n_re), n_re, "furtherseeds"),
        ("first re-seed", "furtherseeds,", "%d" % seeds[0], seeds[0], None),
        ("last re-seed", "furtherseeds,#to", "%d" % seeds[-1], seeds[-1], None),
        ("re-seeded panels preserving the swap", "preservesthedirectionoftheswapinall", wd(n_dir), n_dir, ",andin"),
        ("re-seeded panels with no other reordering", "preservesthedirectionoftheswapinall%s,andin" % wd(n_dir),
         wd(n_oth), n_oth, "ofthemnootherpair"),
        ("re-seeded panels where the contested station moves two places", "ratherthanoneintheremaining",
         wd(n_two), n_two, "whiletheotherstations"),
        ("Fehr-Schmidt grid of rho", "Fehr--Schmidtparametersover\\rho_i\\in\\", grid("rho")[0], grid("rho")[1], None),
        ("Fehr-Schmidt grid of vartheta", "\\and\\vartheta_i\\in\\", grid("theta")[0], grid("theta")[1], None),
        ("Fehr-Schmidt settings", "\\vartheta_i\\le\\rho_i,", wd(n_fs), n_fs, "combinations"),
        ("contested advocate share over the Fehr-Schmidt grid, smallest", "contestedstationonlybetween",
         fm(shr.min()), f(shr.min()), None),
        ("contested advocate share over the Fehr-Schmidt grid, largest", "contestedstationonlybetween#and",
         fm(shr.max()), f(shr.max()), None),
        ("uncontested stations (Fehr-Schmidt grid)", "andtheorderofthe", wd(n_st - 1), n_st - 1, "uncontestedstations"),
        ("solves with a near-tied local maximum", "nearlytiedlocalmaximaatsomeuncontestedstations.At",
         wd(tied), tied, "ofthe"),
        ("solves", "nearlytiedlocalmaximaatsomeuncontestedstations.At%softhe" % wd(tied), wd(len(ce4)), len(ce4),
         "solves"),
        ("near-tie tolerance (relative)", "furtherlocalmaximumcomeswithin", _pn_pow10(r4c["near_tie_tolerance"]),
         f(r4c["near_tie_tolerance"]), None),
        ("near-tied local maxima at the closest solve", "theonesolvewith", wd(len(ml)), len(ml), "ofthem"),
        ("closest near-tie, lower local maximum", "objectivehaslocalmaximaat", fm(ml[0]), ml[0], None),
        ("closest near-tie, upper local maximum", "objectivehaslocalmaximaat#and", fm(ml[1]), ml[1], None),
        ("closest near-tie, reported optimum", "fallshortoftheoptimumat", fm(r4c["near_tie_closest_m_reported"]),
         f(r4c["near_tie_closest_m_reported"]), None),
        ("closest near-tie, bound on the relative gaps", "fallshortoftheoptimumat#bylessthan",
         _pn_times10(r4c["near_tie_closest_gap_bound"]), f(r4c["near_tie_closest_gap_bound"]), None),
        ("closest near-tie, runner-up", "nearest-to-uniformrepresentativeat", fm(r4c["near_tie_closest_m_runner_up"]),
         f(r4c["near_tie_closest_m_runner_up"]), None),
        ("closest near-tie, runner-up advocate share", "wouldgivetheadvocatesashareof",
         fm(r4c["near_tie_closest_runner_up_adv_share"]), f(r4c["near_tie_closest_runner_up_adv_share"]), None),
        ("closest near-tie, runner-up weight Gini", "ataweightGiniof",
         fm(r4c["near_tie_closest_runner_up_weight_gini"]), f(r4c["near_tie_closest_runner_up_weight_gini"]), None),
        ("uncontested-difference threshold", "otherthan%s'sbelow" % wname, "%g" % f(r4c["gap_threshold"]),
         f(r4c["gap_threshold"]), None),
        ("average weight Gini over the near-tie combinations, consensus-only, smallest",
         "averageweightGinionlybetween", fm(wgr["eff"][0]), wgr["eff"][0], None),
        ("average weight Gini over the near-tie combinations, consensus-only, largest",
         "averageweightGinionlybetween#and", fm(wgr["eff"][1]), wgr["eff"][1], None),
        ("average weight Gini over the near-tie combinations, equity-aware, smallest",
         "undertheconsensusregimeandbetween", fm(wgr["eq"][0]), wgr["eq"][0], None),
        ("average weight Gini over the near-tie combinations, equity-aware, largest",
         "undertheconsensusregimeandbetween#and", fm(wgr["eq"][1]), wgr["eq"][1], None),
        ("uncontested stations (utility leveling-down)", ugini + "#to#.Attheother", wd(n_st - 1), n_st - 1, "stations"),
        ("bound on the uncontested change of the mean utility", "themeanutilitychangesbylessthan",
         _pn_below(r4c["uncontested_max_abs_mean_utility_change"], 2),
         f(r4c["uncontested_max_abs_mean_utility_change"]), None),
        ("bound on the uncontested change of the utility Gini", "theGinioftheutilitiesbyatmost",
         _pn_upto(r4c["uncontested_max_abs_utility_gini_change"], 3),
         f(r4c["uncontested_max_abs_utility_gini_change"]), None),
        ("stations leveling down over the near-tie combinations, fewest", "acrosstheircombinations",
         wd(lvl[0]), lvl[0], "to"),
        ("stations leveling down over the near-tie combinations, most", "acrosstheircombinations%sto" % wd(lvl[0]),
         wd(lvl[1]), lvl[1], "stationsleveldown"),
        ("bound on the uncontested mean overlap given up", "meanoverlapgivenupisatmost",
         _pn_upto(r4c["uncontested_max_mean_overlap_fall"], 2), f(r4c["uncontested_max_mean_overlap_fall"]), None),
        ("contested overlap Gini, consensus-only", "AtthecontestedstationtheGinioftheoverlapsfallsfrom",
         fm(cst["gini_A_eff"]), f(cst["gini_A_eff"]), None),
        ("contested overlap Gini, equity-aware", "AtthecontestedstationtheGinioftheoverlapsfallsfrom#to",
         fm(cst["gini_A_eq"]), f(cst["gini_A_eq"]), None),
        ("contested mean overlap, consensus-only", "thepanel'smeanoverlapfrom", fm(cst["Abar_eff"]), f(cst["Abar_eff"]), None),
        ("contested mean overlap, equity-aware", "thepanel'smeanoverlapfrom#to", fm(cst["Abar_eq"]), f(cst["Abar_eq"]), None),
        ("contested smallest overlap, consensus-only", "thesmallestoverlaprisesfrom", fm(cst["Amin_eff"]),
         f(cst["Amin_eff"]), None),
        ("contested smallest overlap, equity-aware", "thesmallestoverlaprisesfrom#to", fm(cst["Amin_eq"]),
         f(cst["Amin_eq"]), None),
        ("size of the advocates' constituency", "acriterionthatpreventsa", wd(n_adv), n_adv, "-memberconstituency"),
        ("uncontested stations (managerial)", "theunchangedorderoftheother", wd(n_st - 1), n_st - 1, "stations"),
        ("early-warning threshold", "Onthisdesignathresholdof", "%g" % f(r4c["warn_threshold"]),
         f(r4c["warn_threshold"]), None),
        ("re-seeded panels the threshold separates", "inthemainpanelandinall", wd(n_re), n_re, "re-seededones"),
        ("smallest contested difference", "thesmallestcontesteddifferenceacrossthembeing",
         fm(r4c["warn_contested_gap_min"]), f(r4c["warn_contested_gap_min"]), None),
        ("largest uncontested difference", "thelargestuncontestedone", fm(r4c["warn_uncontested_gap_max"]),
         f(r4c["warn_uncontested_gap_max"]), None),
        ("stations (closing box)", "singlesyntheticaccessibilityindexover", wd(n_st), n_st, "fictitiousstations"),
    ]
    keys += [_pn_key(ch, rob, what, anchor, text, value, follow=follow) for what, anchor, text, value, follow in rows6]
    # the generating parameters of the panel, which Section 6.5.4 (sec:rt-data) states
    rows4 = [
        ("seed", "Thepanelisgeneratedundertheseed", "%d" % int(float(r4c["seed"])), int(float(r4c["seed"])), None),
        ("leniency sd", "shiftedbyapersonalleniencywithstandarddeviation", "%g" % f(r4c["leniency_sd"]),
         f(r4c["leniency_sd"]), None),
        ("noise sd", "andbynoisewithstandarddeviation", "%g" % f(r4c["noise_sd"]), f(r4c["noise_sd"]), None),
        ("common dispersion", "isassignedthecommondispersion", fm(r4c["common_sigma"], 1), f(r4c["common_sigma"]), None),
        ("stakeholders", "Thesolutionproceduretakesthe", wd(n_sh), n_sh, "generatednormaldistributions"),
    ]
    keys += [_pn_key(ch, "sec:rt-data", what, anchor, text, value, follow=follow)
             for what, anchor, text, value, follow in rows4]
    # the rest of Section 6.5.5 (sec:rt-results): the other uncontested differences, the
    # realized judgments behind the largest one, the weight diagnostics of the swap
    # neighbour and of the station with the largest uncontested difference, and the
    # level-set paragraph of the contested station (ranges of the advocates' share, the
    # weight Ginis of the reported representatives, the overlaps the level sets fix)
    rp4 = r4["realized_panel"].set_index("station")
    wst = str(r4c["max_uncontested_mean_gap_station"])
    cr = [float(v) for v in _pn_re.findall(r"\d+\.\d+", str(r4c["contested_adv_share_common_range"]))]
    rows5 = [
        ("threshold of the other uncontested differences", "at%s,andbylessthan" % wname, "%g" % f(r4c["gap_threshold"]),
         f(r4c["gap_threshold"]), None),
        ("uncontested stations below the threshold", "at%s,andbylessthan#ateachoftheother" % wname, wd(n_st - 2),
         n_st - 2, ",andnouncontestedpair"),
        ("lowest judgment below the panel mean at the largest uncontested difference", "bothadvocates',lie",
         fm(rp4.loc[wst, "low1_below_panel_mean"]), f(rp4.loc[wst, "low1_below_panel_mean"]), None),
        ("second-lowest judgment below the panel mean there", "bothadvocates',lie#and",
         fm(rp4.loc[wst, "low2_below_panel_mean"]), f(rp4.loc[wst, "low2_below_panel_mean"]), None),
        ("contested advocate share, consensus-only (reported)", "stationisinadequaterisesfrom",
         fm(cst["adv_share_eff"]), f(cst["adv_share_eff"]), None),
        ("contested advocate share, equity-aware (reported)", "stationisinadequaterisesfrom#to",
         fm(cst["adv_share_eq"]), f(cst["adv_share_eq"]), None),
        ("swap neighbour weight Gini, consensus-only", "residualdifferenceappears,asinthe",
         fm(by.loc[nb, "weight_gini_eff"]), f(by.loc[nb, "weight_gini_eff"]), None),
        ("swap neighbour weight Gini, equity-aware", "residualdifferenceappears,asinthe#against",
         fm(by.loc[nb, "weight_gini_eq"]), f(by.loc[nb, "weight_gini_eq"]), None),
        ("largest-difference station weight Gini, consensus-only", "at%sorthe" % nname,
         fm(by.loc[wst, "weight_gini_eff"]), f(by.loc[wst, "weight_gini_eff"]), None),
        ("largest-difference station weight Gini, equity-aware", "at%sorthe#against" % nname,
         fm(by.loc[wst, "weight_gini_eq"]), f(by.loc[wst, "weight_gini_eq"]), None),
        ("largest-difference station advocate share, equity-aware",
         "At%stheselectedadvocateshareundertheequityregime," % wname,
         fm(by.loc[wst, "adv_share_eq"]), f(by.loc[wst, "adv_share_eq"]), None),
        ("contested mean, consensus-only (again)", "Itscollectivemeanfallsfrom", fm(cst["mean_eff"]), f(cst["mean_eff"]), None),
        ("contested mean, equity-aware (again)", "Itscollectivemeanfallsfrom#to", fm(cst["mean_eq"]), f(cst["mean_eq"]), None),
        ("largest uncontested mean difference (again)", "nearlyafullpointagainst", fm(r4c["max_uncontested_mean_gap"]),
         f(r4c["max_uncontested_mean_gap"]), None),
        ("advocates' headcount share", "belowtheirheadcountshareof", fm(r4c["adv_headcount_share"]),
         f(r4c["adv_headcount_share"]), None),
        ("contested consensus level set, smallest advocate share", "everyconsensusoptimumtheregivingthembetween",
         fm(cst["adv_share_min_eff"], 0) if abs(f(cst["adv_share_min_eff"])) < 5e-3 else fm(cst["adv_share_min_eff"]),
         f(cst["adv_share_min_eff"]), None),
        ("contested consensus level set, largest advocate share", "everyconsensusoptimumtheregivingthembetween#and",
         fm(cst["adv_share_max_eff"]), f(cst["adv_share_max_eff"]), None),
        ("contested equity level set, smallest advocate share", "everyequityoptimumbetween",
         fm(cst["adv_share_min_eq"]), f(cst["adv_share_min_eq"]), None),
        ("contested equity level set, largest advocate share", "everyequityoptimumbetween#and",
         fm(cst["adv_share_max_eq"]), f(cst["adv_share_max_eq"]), None),
        ("advocate shares attainable on both level sets, smallest", "Everyadvocatesharebetween", fm(cr[0]), cr[0], None),
        ("advocate shares attainable on both level sets, largest", "Everyadvocatesharebetween#and", fm(cr[1]), cr[1], None),
        ("contested advocate share, consensus-only (again)", "andneitherthe", fm(cst["adv_share_eff"]),
         f(cst["adv_share_eff"]), None),
        ("contested advocate share, equity-aware (again)", "andneitherthe#northe", fm(cst["adv_share_eq"]),
         f(cst["adv_share_eq"]), None),
        ("contested weight Gini, consensus-only", "withweightGinisof", fm(cst["weight_gini_eff"]),
         f(cst["weight_gini_eff"]), None),
        ("contested weight Gini, equity-aware", "withweightGinisof#and", fm(cst["weight_gini_eq"]),
         f(cst["weight_gini_eq"]), None),
        ("contested advocates' mean overlap, equity-aware", "meanoverlapwiththecollectivedistributionis",
         fm(cst["Abar_adv_eq"]), f(cst["Abar_adv_eq"]), None),
        ("contested advocates' mean overlap, consensus-only", "meanoverlapwiththecollectivedistributionis#,against",
         fm(cst["Abar_adv_eff"]), f(cst["Abar_adv_eff"]), None),
        ("contested smallest overlap, equity-aware (level set)", "thesmallestoverlaponthepanelis",
         fm(cst["Amin_eq"]), f(cst["Amin_eq"]), None),
        ("contested smallest overlap, consensus-only (level set)", "thesmallestoverlaponthepanelis#against",
         fm(cst["Amin_eff"]), f(cst["Amin_eff"]), None),
    ]
    keys += [_pn_key(ch, res, what, anchor, text, value, follow=follow) for what, anchor, text, value, follow in rows5]
    # the Case IV row of the robustness table (Table 6.22) counts the same panels and settings
    keys += [_pn_key(ch, "sec:app-synthesis", "Case IV re-seeded panels (robustness table)",
                     "contestedstationsurvivesall", wd(n_re), n_re, scope="table:tab:robustness",
                     follow="re-seededpanels"),
             _pn_key(ch, "sec:app-synthesis", "Case IV Fehr-Schmidt settings (robustness table)",
                     "contestedstationsurvivesall%sre-seededpanelsandall" % wd(n_re), wd(n_fs), n_fs,
                     scope="table:tab:robustness", follow="Fehr--Schmidtsettings")]
    return None, tables, keys


def _pn_vignette_v():
    hm = load("fig6-12_hm_calsharp.xlsx")
    agg = dict(zip(hm["aggregates"]["quantity"].astype(str), hm["aggregates"]["value"].astype(float)))
    wts = hm["weights"].set_index("id")["weight_full"].astype(float)
    names = ["optimized human-only mean CRPS", "equal-weight human-machine mean CRPS",
             "optimized human-machine mean CRPS"]

    def new_q(n):
        # the expected CRPS on new quarters of the same aggregate, where the workbook has it
        # (the table then prints it after the in-sample score)
        q = n.replace(" mean CRPS", " expected CRPS on new quarters")
        return [agg[q]] if q in agg else []

    tables = [_pn_table("tab:hm-vignette", "chapter6.tex", "fig6-12_hm_calsharp.xlsx / aggregates",
                        [_pn_row(n, [n.replace(" mean CRPS", ""), agg[n]] + new_q(n) + [None],
                                 {"text": n.replace(" mean CRPS", "")}) for n in names])]
    sec, ch = "sec:fc-results", "chapter6.tex"

    def k(what, anchor, q, d=2, scope="prose", section=sec):
        # q may list alternative names of one stored quantity; the first present is read
        q = next(n for n in ([q] if isinstance(q, str) else q) if n in agg)
        return _pn_key(ch, section, what, anchor, _pn_fmt(agg[q], d), agg[q], scope=scope)

    keys = [
        k("machine weight share W_M", "W_\\mathrmM=\\sum_i\\in\\mathrmMw_i=", "machine weight share W_M"),
        k("human weight share W_H", "W_\\mathrmH=\\sum_i\\in\\mathrmHw_i=", "human weight share W_H"),
        k("optimized human-machine mean CRPS", "optimizedhuman-machineaggregatescores", names[2]),
        k("optimized human-only mean CRPS", "optimizedhuman-machineaggregatescores#against", names[0]),
        k("equal-weight human-machine mean CRPS", "fortheoptimizedhuman-onlyaggregateand", names[1]),
        k("largest leave-one-out mean CRPS", ["leave-one-outre-solvesscoring", "leave-one-outre-solvesscores"],
          "largest leave-one-out mean CRPS"),
        k("climatological uncertainty UNC", ["sampleclimatologicaluncertainty", "seventy-twooutcomes,whichis",
                                             "drawnindependentlyfromit,whichis"],
          "uncertainty UNC (sample climatology)"),
        k("equal-weight pool of the calibrated members, spread", "reportsaspreadof",
          "equal-weight pool of the six calibrated members, spread"),
        k("equal-weight pool of the calibrated members, error of the mean", "againstaroot-mean-squareerrorof",
          ["equal-weight pool of the six calibrated members, root-mean-square error of the mean",
           "equal-weight pool of the six calibrated members, realized root-mean-square error of the mean"]),
    ]
    # the design-level scores of Section 6.6.4, printed to three decimals: the optimized
    # human-machine pool, the human-only pool, the equal-weight pool, and the machine-only
    # pool that the next sentence places between the last and the first
    keys += [
        k("design-level optimized human-machine expected CRPS", "theoptimizedhuman-machinepoolscores",
          "design-level optimized human-machine expected CRPS", d=3),
        k("design-level optimized human-only expected CRPS", "theoptimizedhuman-machinepoolscores#against",
          "design-level optimized human-only expected CRPS", d=3),
        k("design-level equal-weight expected CRPS", "forthehuman-onlypooland",
          "design-level equal-weight expected CRPS", d=3),
        k("design-level optimized machine-only expected CRPS", "thesameprogramgivesapoolthatscores",
          "design-level optimized machine-only expected CRPS", d=3),
    ]
    wmax = float(wts.max())
    if "reseed_summary" in hm:
        # a workbook that regenerates the record under further seeds: the prose prints the
        # weight of the other overconfident machine M3, and the robustness table the
        # number of regenerated records and of those in which M4 shows a positive
        # resolution channel with a negative calibration channel
        rs = hm["reseed_summary"].set_index("claim")
        n_re = int(float(cfgd(hm)["reseed_count"]))
        n_m4 = int(rs.loc["C3_M4_resolution_plus_calibration_minus", "count"])
        m3 = float(wts.loc["M3"])
        keys += [
            _pn_key(ch, sec, "weight of the overconfident machine M3", "M3receives", _pn_fmt(m3, 2), m3),
            _pn_key(ch, "sec:app-synthesis", "number of regenerated records (robustness table)", "ineachofthe",
                    "%d" % n_re, n_re, scope="table:tab:robustness"),
            _pn_key(ch, "sec:app-synthesis", "records with M4 in the (-, +) cell (robustness table)",
                    "negativecalibrationchannelin", "%d" % n_m4, n_m4, scope="table:tab:robustness"),
        ]
    else:
        # the overconfident machine M4 is read by its id; the prose also calls its weight
        # the largest single weight, so the largest weight of the pool, whichever member
        # holds it, must print the same at the same place
        m4 = float(wts.loc["M4"])
        held = "M4holdingthelargestsingleweight,"
        still = "M4stillholding"
        keys += [
            _pn_key(ch, sec, "weight of the overconfident machine M4", held, _pn_fmt(m4, 2), m4),
            _pn_key(ch, sec, "largest single weight of the pool", held, _pn_fmt(wmax, 2), wmax),
            _pn_key(ch, "sec:app-synthesis", "weight of the overconfident machine M4 (robustness table)", still,
                    _pn_fmt(m4, 2), m4, scope="table:tab:robustness"),
            _pn_key(ch, "sec:app-synthesis", "largest single weight of the pool (robustness table)", still,
                    _pn_fmt(wmax, 2), wmax, scope="table:tab:robustness"),
        ]
    summary = {
        "vignette": "V",
        "title": "Aggregating human and machine forecasts (exploratory)",
        "mean_crps": {n: agg[n] for n in names},
        "weight_shares": {"human": agg["human weight share W_H"], "machine": agg["machine weight share W_M"]},
        "weights_full_pool": {str(i): float(v) for i, v in wts.items()},
        "largest_weight": {"value": wmax, "held_by": sorted(str(i) for i, v in wts.items() if v >= wmax - 1e-9)},
        "largest_leave_one_out_mean_crps": agg["largest leave-one-out mean CRPS"],
    }
    return summary, tables, keys


def _pn_worked_examples():
    """The examples that Chapters 2 to 5 work by hand: their arithmetic, computed by the
    builder from the stated inputs (sheets worked_examples of the Figure 2.2, 3.1, 4.3 and
    5.4 workbooks), and the other builder-computed numbers printed in the same passages
    (Example ex:heterogeneous, the Figure 3.1 profile, Example ex:corr-lineage and
    Figure 5.4)."""
    W = {}
    for name in ("fig2-2_three_pooling_operators.xlsx", "fig3-1_lorenz_gini_hoover.xlsx", "fig4-3_scalability.xlsx",
                 "fig5-4_depnet.xlsx"):
        # the text columns are read as text (read_excel would turn a printed '0.600' into 0.6)
        t = pd.read_excel(os.path.join(DATA, name), sheet_name="worked_examples",
                          dtype={"decimals": str, "printed": str})
        for _, r in t.iterrows():
            W[(str(r["example"]), str(r["quantity"]))] = r

    def val(ex, q):
        return float(W[(ex, q)]["value"])

    def pr(ex, q, form=None):
        """The printed form of a row: as the sheet prints it, a fraction as \\tfrac (form
        'tfrac'), an integer with thousands commas (form 'comma', as 2{,}610 prints), or the
        value at `form` decimals; the sheet's string must be the value at its recorded
        decimals (or its exact fraction)."""
        r = W[(ex, q)]
        p, d = ("" if pd.isna(r[c]) else str(r[c]) for c in ("printed", "decimals"))
        v = float(r["value"])
        if d == "fraction":
            fr = Fraction(p)
            if abs(float(fr) - v) > 1e-12:
                raise ValueError("%s / %s: the printed fraction %s is not the value %r" % (ex, q, p, v))
            if form == "tfrac" and fr.denominator != 1:
                return ("-" if fr < 0 else "") + "\\tfrac%d%d" % (abs(fr.numerator), fr.denominator)
            return p
        if not p or len(p.partition(".")[2]) != int(d) or \
                abs(float(p) - v) > 0.5 * 10 ** -int(d) * (1 + 1e-9) + 1e-12:
            raise ValueError("%s / %s: printed %r is not the value %r at %s decimals" % (ex, q, p, v, d))
        if form == "comma":
            if int(d) != 0:
                raise ValueError("%s / %s: a count printed with thousands commas has %s decimals" % (ex, q, d))
            return "{:,}".format(int(p))
        return _pn_fmt(v, form) if isinstance(form, int) else p

    def vec(fmt, n):
        return [fmt % (i + 1) for i in range(n)]

    keys = []

    def k(file, sec, ex, q, anchor, form=None, what=None):
        if isinstance(q, str):
            text, value = pr(ex, q, form), val(ex, q)
        else:
            text, value = "(" + ",".join(pr(ex, x, form) for x in q) + ")", [val(ex, x) for x in q]
        name = q if isinstance(q, str) else "%s .. %s" % (q[0], q[-1])
        keys.append(_pn_key(file, sec, "%s: %s%s" % (ex, name, " (%s)" % what if what else ""), anchor, text, value))

    # ---- Chapter 2
    c2, sb, rl, qa = "chapter2.tex", "sec:foundations-suprabayesian", "sec:foundations-related", "sec:foundations-qa"
    g = "GLS weights of two correlated experts"
    post = "posteriorvarianceis\\sigma_1^2\\sigma_2^2(1-\\rho^2)/"
    est = "thecombinedestimateis"
    for q, anchor, what in (
            ("sigma_1", "Take\\sigma_1=", None), ("sigma_2", "Take\\sigma_1=#,\\sigma_2=", None),
            ("rho", "Take\\sigma_1=#,\\sigma_2=#,and\\rho=", None),
            ("sigma_1^2", "Thedenominatoris", None), ("sigma_2^2", "Thedenominatoris#+", None),
            ("2 rho sigma_1 sigma_2", "Thedenominatoris#+#-", None), ("denominator", "Thedenominatoris#+#-#=", None),
            ("numerator of w_1", "sow_1=", None), ("denominator", "sow_1=#/", "under w_1"), ("w_1", "sow_1=#/#=", None),
            ("numerator of w_2", "andw_2=", None), ("denominator", "andw_2=#/", "under w_2"),
            ("w_2", "andw_2=#/#=", None),
            ("denominator", post, "under the posterior variance"),
            ("sigma_1^2 sigma_2^2 (1 - rho^2)", post + "#=", None),
            ("denominator", post + "#=#/", "under the posterior variance, again"),
            ("diffuse-prior posterior variance", post + "#=#/#=", None),
            ("r_1", "Iftheexpertsreportr_1=", None), ("r_2", "Iftheexpertsreportr_1=#andr_2=", None),
            ("r_1", est + "\\tfrac##(", "in the combined estimate"),
            ("r_2", est + "\\tfrac##(#)-\\tfrac##(", "in the combined estimate"),
            ("combined estimate (fraction)", est + "\\tfrac##(#)-\\tfrac##(#)=", None),
            ("combined estimate", est + "\\tfrac##(#)-\\tfrac##(#)=#/#\\approx", None)):
        k(c2, sb, g, q, anchor, what=what)
    k(c2, sb, g, "w_1", est, "tfrac", "in the combined estimate")
    k(c2, sb, g, "w_2", est + "\\tfrac##(#)", "tfrac", "in the combined estimate")
    d = "DeGroot trust matrix"
    wq = [["W[%d,%d]" % (i + 1, j + 1) for j in range(3)] for i in range(3)]
    keys.append(_pn_key(c2, rl, "%s: W" % d, "W=\\beginpmatrix",
                        "\\\\[2pt]".join("&".join(pr(d, q, "tfrac") for q in row) for row in wq),
                        [val(d, q) for row in wq for q in row]))
    k(c2, rl, d, vec("pi_%d", 3), "substituting\\boldsymbol\\pi=", "tfrac")
    # Example ex:heterogeneous (sheet ex_heterogeneous of the Figure 2.2 workbook)
    hx = load("fig2-2_three_pooling_operators.xlsx")["ex_heterogeneous"].set_index("quantity")["value"].astype(float)
    het, refl = "right-skewedlog-normalwithunderlying\\mathcalN", "left-skewedreflectionaboutx="
    for what, anchor, text, value in (
            ("log-normal expert, log x", het, "(%s,%s^2)" % (_pn_fmt(hx["log-normal expert, mean of log x"], 1),
                                                            _pn_fmt(hx["log-normal expert, sd of log x"], 2)),
             [hx["log-normal expert, mean of log x"], hx["log-normal expert, sd of log x"]]),
            ("log-normal expert, mean", het + "(#,#^2)(mean", _pn_fmt(hx["log-normal expert, mean"], 2),
             hx["log-normal expert, mean"]),
            ("log-normal expert, sd", het + "(#,#^2)(mean#,standarddeviation", _pn_fmt(hx["log-normal expert, sd"], 2),
             hx["log-normal expert, sd"]),
            ("reflection point", refl, _pn_fmt(hx["reflection point"], 2), hx["reflection point"]),
            ("reflected expert, mean", refl + "#(mean", _pn_fmt(hx["reflected expert, mean"], 2),
             hx["reflected expert, mean"]),
            ("reflected expert, sd", refl + "#(mean#,standarddeviation", _pn_fmt(hx["reflected expert, sd"], 2),
             hx["reflected expert, sd"]),
            ("sharp expert", "sharpsymmetricexpert\\mathcalN",
             "(%s,%s^2)" % (_pn_fmt(hx["sharp expert, mean"], 2), _pn_fmt(hx["sharp expert, sd"], 2)),
             [hx["sharp expert, mean"], hx["sharp expert, sd"]]),
            ("mu_QA", "theaverageofthemeans,\\mu_\\mathrmQA=", _pn_fmt(hx["mu_QA"], 2), hx["mu_QA"]),
            ("V_QA", "averagedquantilefunctiongivesV_\\mathrmQA=", _pn_fmt(hx["V_QA"], 2), hx["V_QA"]),
            ("(sum_i w_i sigma_i)^2", "strictlybelow\\bigl(\\sum_iw_i\\sigma_i\\bigr)^2=",
             _pn_fmt(hx["(sum_i w_i sigma_i)^2"], 2), hx["(sum_i w_i sigma_i)^2"])):
        keys.append(_pn_key(c2, qa, "ex:heterogeneous: %s" % what, anchor, text,
                            [float(v) for v in value] if isinstance(value, list) else float(value)))

    # the panel of Figure 2.1 and the running-example comparison of Figure 2.2 (sheets
    # gaussians of the two workbooks, the linear pool's sd in the Figure 2.2 config)
    g22 = load("fig2-2_three_pooling_operators.xlsx")
    ga = g22["gaussians"]
    e22 = ga[ga["role"].astype(str) == "expert"]
    qa22 = ga[ga["legend"].astype(str) == "quantile average"].iloc[0]
    lo22 = ga[ga["legend"].astype(str) == "logarithmic pool"].iloc[0]
    lp_sd = float(cfgd(g22)["linear_pool_sd"])
    lp_mu = float((e22["weight"].astype(float) * e22["mu"].astype(float)).sum())
    if abs(lp_mu - float(qa22["mu"])) > 1e-12:
        raise ValueError("the linear pool and the quantile average no longer share the mean")
    wts = [Fraction(float(x)).limit_denominator(100) for x in e22["weight"]]
    e21 = load("fig2-1_linear_pool_vs_qa.xlsx")["gaussians"]
    e21 = e21[e21["role"].astype(str) == "expert"]
    if len(e21) != 2 or abs(float(e21["mu"].sum())) > 1e-12 or e21["sigma"].nunique() != 1:
        raise ValueError("Figure 2.1 is no longer two experts at means -m and +m with a common sd")
    m21, s21 = float(e21["mu"].abs().iloc[0]), float(e21["sigma"].iloc[0])
    srt = "Theirorderingbysharpness,f_\\mathrmLogOP("
    for what, anchor, text, value in (
            ("Figure 2.1, the experts' means (plus or minus)", "normalexpertswithmeans\\pm", _pn_fmt(m21, 1), m21),
            ("Figure 2.1, the common sd", "normalexpertswithmeans\\pm#andcommonstandarddeviation", _pn_fmt(s21, 2), s21),
            ("Figure 2.2, the equal weights", "comparethethreepoolingoperatorsunderequalweights\\mathbfw=",
             "(" + ",".join("\\tfrac%d%d" % (w.numerator, w.denominator) for w in wts) + ")", [float(w) for w in wts]),
            ("Figure 2.2, the quantile average", "f_\\mathrmQA=\\mathcalN",
             _pn_normal(qa22["mu"], qa22["sigma"]), [float(qa22["mu"]), float(qa22["sigma"])]),
            ("Figure 2.2, the logarithmic pool", "f_\\mathrmLogOP=\\mathcalN",
             _pn_normal(lo22["mu"], lo22["sigma"]), [float(lo22["mu"]), float(lo22["sigma"])]),
            ("Figure 2.2, the linear pool's mean", "TheLPsharesthemean", _pn_fmt(lp_mu, 2), lp_mu),
            ("Figure 2.2, the linear pool's sd", "equallyweightedmixture,hasstandarddeviation", _pn_fmt(lp_sd, 2), lp_sd),
            ("Figure 2.2, sharpness order, logarithmic pool", srt, _pn_fmt(lo22["sigma"], 2), float(lo22["sigma"])),
            ("Figure 2.2, sharpness order, quantile average", srt + "#)<f_\\mathrmQA(", _pn_fmt(qa22["sigma"], 2),
             float(qa22["sigma"])),
            ("Figure 2.2, sharpness order, linear pool", srt + "#)<f_\\mathrmQA(#)<f_\\mathrmLP(", _pn_fmt(lp_sd, 2), lp_sd)):
        keys.append(_pn_key(c2, qa, what, anchor, text, value))

    # ---- Chapter 3
    c3, fa, we, el = "chapter3.tex", "sec:conscience-fairness", "sec:conscience-welfare", "sec:conscience-elicitation"
    b, sk = "ex:bonferroni", "ex:skewed-panel"
    k(c3, fa, b, vec("U_A[%d]", 4), "\\widetildeU^\\mathrmFS,A&=")
    k(c3, fa, b, vec("U_B[%d]", 4), "\\widetildeU^\\mathrmFS,B&=")
    k(c3, fa, b, "arithmetic mean of A (and of B)", "whichsharethearithmeticmean")
    k(c3, fa, b, "Phi_11(A)", "gives\\Phi_1,1(\\widetildeU^\\mathrmFS,A)=")
    k(c3, fa, b, "Phi_11(B)", "against\\Phi_1,1(\\widetildeU^\\mathrmFS,B)=")
    k(c3, fa, sk, vec("A[%d]", 5), "Takefiveexpertswithoverlaps\\mathbfA=")
    keys.append(_pn_key(c3, fa, "%s: (rho, vartheta) of the five experts" % sk,
                        "heterogeneousdispositions(\\rho_i,\\vartheta_i)equalto",
                        ",".join("(%s,%s)" % (pr(sk, "rho[%d]" % (i + 1)), pr(sk, "vartheta[%d]" % (i + 1)))
                                 for i in range(5)),
                        [val(sk, q % (i + 1)) for i in range(5) for q in ("rho[%d]", "vartheta[%d]")]))
    k(c3, fa, sk, vec("U_FS[%d]", 5), "areU^\\mathrmFS=")
    k(c3, fa, sk, "normalized U_FS[1]", "andisclippedto", 0, "the clip")
    k(c3, fa, sk, "kappa", "eq:parabola-normat\\kappa=")
    k(c3, fa, sk, "A[5]", "withoverlap")
    k(c3, fa, sk, "normalized U_FS[5]", "clipsthatutilitytoexactly", 0, "the clip")
    k(c3, fa, sk, vec("normalized U_FS[%d]", 5), "Thenormalizedprofileis\\widetildeU^\\mathrmFS=")
    k(c3, fa, sk, "arithmetic mean of the normalized profile", "witharithmeticmean")
    k(c3, fa, sk, "Phi_11", "bytheclosedform\\eqrefeq:bonferroni-closed,is\\Phi_1,1=")
    t3, at, lv = "three-expert welfare profile", "Atkinson index of the three-expert profile", "leveling-down four-expert panel"
    wel = "W_1,4=\\barU-I_4="
    atk = "I_\\mathrmAtk(1)\\approx1-"
    for ex, q, anchor, what in (
            (t3, vec("A[%d]", 3), "overlaptheaggregatetodegrees\\mathbfA=(A_1,A_2,A_3)=", None),
            (t3, "rho (every expert)", "homogeneousdisposition\\rho_i=", None),
            (t3, "vartheta (every expert)", "homogeneousdisposition\\rho_i=#and\\vartheta_i=", None),
            (t3, vec("U_FS[%d]", 3), "thenevaluatetoU^\\mathrmFS=", None),
            (t3, "U_FS[1] / A[1]", "fallsunevenly,from", None),
            (t3, "U_FS[2] / A[2]", "forthebest-representedexpertthrough", None),
            (t3, "U_FS[3] / A[3]", "forthebest-representedexpertthrough#to", None),
            (t3, vec("U[%d]", 3), "yieldstheprofileU=", None),
            (t3, "mean U", "withmean\\barU=", None),
            (t3, "EQ_1", "\\mathrmEQ_1=", None), (t3, "EQ_2", "\\mathrmEQ_2=", None),
            (t3, "EQ_3", "\\mathrmEQ_3=", None), (t3, "EQ_4", "\\mathrmEQ_4=", None),
            (t3, "EQ_5", "\\mathrmEQ_5=", None),
            (t3, "EQ_5", "theHoovervalue", "the Hoover value again"),
            (t3, "mean U", "Theutilitarianmean", "the utilitarian mean"),
            (t3, "minimum U", "andtheegalitarianminimum", None),
            (t3, "lambda", "sothatat\\lambda=", None),
            (t3, "mean U", wel, "in W_{1,4}"), (t3, "I_4", wel + "#-", None), (t3, "W_{1,4}", wel + "#-#=", None),
            (t3, vec("U[%d]", 3), "Continuingtheillustrativethree-expertprofileU=", "Atkinson paragraph"),
            (at, "varepsilon_Atk", "oftheworkedexampleabove,at\\varepsilon_\\mathrmAtk=", None),
            (at, "geometric mean", "equivalentlevelisthegeometricmean(#\\times#\\times#)^1/3\\approx", None),
            (t3, "mean U", "againstthearithmeticmean", "Atkinson paragraph"),
            (at, "geometric mean", atk, "in I_Atk(1)"), (t3, "mean U", atk + "#/", "in I_Atk(1)"),
            (at, "I_Atk(1)", atk + "#/#\\approx", None),
            (lv, vec("A[%d]", 4), "underoneweightingare\\mathbfA=", None),
            (lv, "mean A", "withmeanoverlap\\barA=", None),
            (lv, "I_4(A)", "andGiniinequalityI_4=", None),
            (lv, vec("A'[%d]", 4), "levelingdownto\\mathbfA'=", None),
            (lv, "I_4(A')", "attainsperfectequity,I_4=", None),
            (lv, "-I_4(A)", "bareequityobjective-I_4from", None),
            (lv, "I_4(A')", "bareequityobjective-I_4from#to", "the bare equity objective at A'"),
            (lv, "mean A", "lowersthemeanoverlapfrom", "again"),
            (lv, "mean A'", "lowersthemeanoverlapfrom#to", None),
            (lv, "drop in mean overlap", "forevery\\lambda<", None),
            (lv, "I_4(A)", "forevery\\lambda<#/", "in the threshold"),
            (lv, "threshold lambda", "forevery\\lambda<#/#\\approx", None),
            (lv, "drop in mean overlap", "Belowthisthresholdthedropof", "again"),
            (lv, "I_4(A)", "outweighsthegainof", "in the gain"),
            (lv, "threshold lambda", "forgomorethan", "again")):
        k(c3, we, ex, q, anchor, what=what)
    keys.append(_pn_key(c3, we, "%s: the product under the geometric mean" % at,
                        "equivalentlevelisthegeometricmean(",
                        "\\times".join(pr(t3, q) for q in vec("U[%d]", 3)), [val(t3, q) for q in vec("U[%d]", 3)]))
    # the Figure 3.1 profile (sheet summary of the Figure 3.1 workbook)
    s31 = cfgd({"config": load("fig3-1_lorenz_gini_hoover.xlsx")["summary"]})
    for what, anchor, key, dd in (
            ("second-lowest utility", "Sixexpertsholdutilitiesbetween", "second_lowest_utility", 3),
            ("largest utility", "Sixexpertsholdutilitiesbetween#and", "max_utility", 3),
            ("the dissenting expert", "andthedissentingexpert", "min_utility_expert", 0),
            ("the dissenter's utility", "andthedissentingexpert#only", "min_utility", 3),
            ("Gini coefficient", "OnthisprofiletheGinicoefficientis", "gini_pairwise", 3),
            ("Hoover index", "OnthisprofiletheGinicoefficientis#andtheHooverindex", "hoover_max_gap", 3),
            ("mean utility", "belowthepanelmeanof", "mean_utility", 3)):
        keys.append(_pn_key(c3, we, "Figure 3.1 profile: %s" % what, anchor, _pn_fmt(s31[key], dd), float(s31[key])))
    u = "ultimatum-game identification"
    uq = "ofthepie,\\hats=#,reveals\\alpha_i="
    for q, anchor in (("s_hat (thirty percent)", "ofthepie,\\hats="), ("s_hat (thirty percent)", uq),
                      ("1 - 2 s_hat (thirty percent)", uq + "#/"), ("alpha (thirty percent)", uq + "#/#="),
                      ("s_hat (a quarter)", "indifferentataquarter,\\hats="),
                      ("alpha (a quarter)", "indifferentataquarter,\\hats=#,reveals\\alpha_i=")):
        k(c3, el, u, q, anchor, what="in alpha = s/(1-2s)" if anchor == uq else None)

    # ---- Chapter 4 (sheet worked_examples of the Figure 4.3 workbook)
    c4, hs = "chapter4.tex", "sec:scalability-hierarchical"
    ub, fc = "ex:unbalanced-clustering", "fuzzy C-means memberships"
    bal, flat, skw = "Abalancedpartition(#,#,#)incurs", "againstthe", "dominatesthepanel,incurs"
    mem = "thenreceivesmembership1/\\bigl(1+("
    for ex, q, anchor, form, what in (
            (ub, "N", "TakeapanelofN=", None, None),
            (ub, "K", "TakeapanelofN=#expertsandK=", None, None),
            (ub, "N", "summingto", None, "the sum of the subgroup sizes"),
            (ub, vec("balanced size n_%d", 3), "Abalancedpartition", None, None),
            (ub, "K", bal, None, "in the balanced count"),
            (ub, "balanced size n_1", bal + "#\\times", None, "in the balanced count"),
            (ub, "balanced n_k - 1", bal + "#\\times#\\times", None, None),
            (ub, "balanced intra-group count", bal + "#\\times#\\times#=", "comma", None),
            (ub, "N", flat, None, "in the flat count"),
            (ub, "N - 1", flat + "#\\times", None, None),
            (ub, "flat count", flat + "#\\times#=", "comma", None),
            (ub, vec("skewed size n_%d", 3), "Askewedpartition", None, None),
            (ub, "skewed size n_1", skw, None, "in the skewed count"),
            (ub, "skewed n_1 - 1", skw + "#\\times", None, None),
            (ub, "skewed subgroups of size n_2", skw + "#\\times#+", None, None),
            (ub, "skewed size n_2", skw + "#\\times#+#\\times", None, "in the skewed count"),
            (ub, "skewed n_2 - 1", skw + "#\\times#+#\\times#\\times", None, None),
            (ub, "skewed intra-group count", skw + "#\\times#+#\\times#\\times#=", "comma", None),
            (fc, "m", "Attheconventionalm=", None, None),
            (fc, "exponent 2/(m - 1)", "theexponent2/(m-1)equals", None, None),
            (fc, "distance to the nearer centroid", "Anexpertatdistance", None, None),
            (fc, "distance to the farther centroid", "Anexpertatdistance#fromonecentroidanddistance", None, None),
            (fc, "ratio d_near / d_far", mem, None, None),
            (fc, "exponent 2/(m - 1)", mem + "#/#)^", None, "in the membership"),
            (fc, "membership in the nearer subgroup", mem + "#/#)^#\\bigr)=", None, None),
            (fc, "membership in the farther subgroup", "inthenearersubgroupand", None, None),
            (fc, "membership of an equidistant expert", "equidistantfrombothreceives", None, None)):
        k(c4, hs, ex, q, anchor, form, what)

    # ---- Chapter 5
    c5, de = "chapter5.tex", "sec:collab-dependence"
    ce, n, lin = "equicorrelated ceiling", "ex:neff", "ex:corr-lineage"
    neff = "thetenmodelssupplyonlyN_\\mathrmeff="
    for ex, q, anchor, form, what in (
            (ce, "r", "equicorrelatedpoolatr=", None, None),
            (ce, "ceiling 1/r", "canneverexceedabout", None, None),
            (n, "w_i", "weightedequally,w_i=", None, None),
            (n, "N_eff at R = I", "R=IwouldgiveN_\\mathrmeff=", None, None),
            (n, "r", "soatr=", None, None),
            (n, "N", neff, None, "in the effective number"),
            (n, "1 + (N - 1) r", neff + "#/", None, None),
            (n, "N_eff, equicorrelated", neff + "#/#\\approx", None, None),
            (n, "r_w", "correlatedatr_\\mathrmw=", None, None),
            (n, "r_b", "withinagroupandr_\\mathrmb=", None, None),
            (n, "N_eff, three lineages, equal weights", "effectivenumberisthenN_\\mathrmeff\\approx", None, None),
            (n, vec("group share %d, equal weights", 3), "drawsthegroupweightsharesfrom", None, None),
            (n, vec("group share %d, penalty minimum", 3), "drawsthegroupweightsharesfrom(#,#,#)toward", "tfrac", None),
            (n, "N_eff, equal shares, equal weights within", "withequalweightswithineachgroup,toN_\\mathrmeff\\approx",
             None, None),
            (lin, "r", "within-lineageerrorcorrelationr_ij\\approx", None, None),
            (lin, "n", "thosefivevariantscountasonly", None, None),
            (lin, "n - 1", "thosefivevariantscountasonly#/[1+", None, None),
            (lin, "r", "thosefivevariantscountasonly#/[#+#(", None, "in the count"),
            (lin, "N_eff of the lineage", "thosefivevariantscountasonly#/[#+#(#)]\\approx", None, None)):
        k(c5, de, ex, q, anchor, form, what)
    keys.append(_pn_key(c5, de, "%s: panel size (in words)" % n, "Consider", _pn_words(val(n, "N")), int(val(n, "N")),
                        follow="languagemodels"))
    keys.append(_pn_key(c5, de, "%s: panel size, again (in words)" % n, "farbelowthenominal", _pn_words(val(n, "N")),
                        int(val(n, "N")), follow="andbeloweventhe"))
    gs = [int(val(n, "group size %d" % (i + 1))) for i in range(3)]
    keys.append(_pn_key(c5, de, "%s: group sizes (in words)" % n, "provenancegroupsofsizes",
                        "%s,%s,and%s" % tuple(_pn_words(x) for x in gs), gs, follow=",correlatedat"))
    # Example ex:corr-lineage and the Figure 5.4 paragraph: the dependence instrument on
    # the panel of Example ex:hmpanel (fig5-3_loo.xlsx, sheet error_dependence) and the
    # stipulated matrix of Figure 5.4 (config of fig5-4_depnet.xlsx)
    dep = load("fig5-3_loo.xlsx")["error_dependence"].set_index("quantity")["value"].astype(float)
    d4 = cfgd(load("fig5-4_depnet.xlsx"))
    lo = "tocorrelatewithitshumansatonly"
    for what, anchor, value in (
            ("ex:hmpanel human error correlation, smallest", "ex:hmpanelcorrelateat", dep["error r human-human min"]),
            ("ex:hmpanel human error correlation, largest", "ex:hmpanelcorrelateat#to", dep["error r human-human max"]),
            ("ex:hmpanel machine error correlation, smallest", "amongitshumansandat", dep["error r machine-machine min"]),
            ("ex:hmpanel machine error correlation, largest", "amongitshumansandat#to", dep["error r machine-machine max"]),
            ("ex:hmpanel humans, effective number", "itstwosubpanelscountasonly", dep["N_eff equal weights human"]),
            ("ex:hmpanel machines, effective number", "itstwosubpanelscountasonly#and", dep["N_eff equal weights machine"]),
            ("squared correlation with the outcome, smallest", "withtheoutcomelyingbetween",
             dep["squared corr(median, outcome) min"]),
            ("squared correlation with the outcome, largest", "withtheoutcomelyingbetween#and",
             dep["squared corr(median, outcome) max"]),
            ("lineage-human correlation, smallest", lo, d4["cross_block_min"]),
            ("lineage-human correlation, largest", lo + "#to", d4["cross_block_max"]),
            ("ex:hmpanel human-machine error correlation, smallest", lo + "#to#,belowthe", dep["error r human-machine min"]),
            ("ex:hmpanel human-machine error correlation, largest", lo + "#to#,belowthe#to",
             dep["error r human-machine max"]),
            ("ex:hmpanel human-machine forecast correlation, smallest", "twosubpanelscorrelatingatonly",
             dep["forecast r human-machine min"]),
            ("ex:hmpanel human-machine forecast correlation, largest", "twosubpanelscorrelatingatonly#to",
             dep["forecast r human-machine max"]),
            ("ex:hmpanel effective number, humans", "raisestheeffectivenumberofexpertsfrom", dep["N_eff equal weights human"]),
            ("ex:hmpanel effective number, panel", "raisestheeffectivenumberofexpertsfrom#to", dep["N_eff equal weights panel"]),
            ("stipulated humans, effective number", "lowerstheeffectivenumberfrom", d4["neff_human_equalweight"]),
            ("stipulated humans with the lineage, effective number", "lowerstheeffectivenumberfrom#to", d4["neff_panel_equalweight"]),
            ("Figure 5.4 human block, smallest", "pairwisecorrelationslyingbetween", d4["human_block_min"]),
            ("Figure 5.4 human block, largest", "pairwisecorrelationslyingbetween#and", d4["human_block_max"]),
            ("Figure 5.4 cross block, smallest", "correlationswiththevariantsbetween", d4["cross_block_min"]),
            ("Figure 5.4 cross block, largest", "correlationswiththevariantsbetween#and", d4["cross_block_max"]),
            ("Figure 5.4 lineage block mean", "within-lineagecorrelationof", d4["lineage_mean_corr"])):
        keys.append(_pn_key(c5, de, "ex:corr-lineage: %s" % what, anchor, _pn_fmt(value, 2), float(value)))
    for what, anchor in (("lineage count, detection clause", "inwhichthelineagecountsforonlyabout"),
                         ("Figure 5.4 clique, effective number", "thecliquesuppliesonlyabout")):
        keys.append(_pn_key(c5, de, "ex:corr-lineage: %s" % what, anchor, _pn_fmt(d4["neff_lineage_equalweight"], 1),
                            float(d4["neff_lineage_equalweight"])))
    return None, [], keys


_pn_seed_cache = {}


def _pn_seed_entries():
    """The seed entry of every workbook's config (or provenance) sheet, by workbook."""
    if _pn_seed_cache:
        return _pn_seed_cache
    out = _pn_seed_cache
    for fn in sorted(os.listdir(DATA)):
        if not fn.endswith(".xlsx") or fn.startswith("~$"):
            continue
        for sh in ("config", "provenance"):
            if fn in _loaded:
                c = _loaded[fn].get(sh)
                if c is None:
                    continue
            else:
                try:
                    c = pd.read_excel(os.path.join(DATA, fn), sheet_name=sh)
                except ValueError:
                    continue
            out[fn] = str(dict(zip(c["key"].astype(str), c["value"])).get("seed", "")).strip()
            break
    return out


def _pn_seeds():
    """The one numeric seed that every workbook drawing at random records (a list if they
    differ), read from the config sheets."""
    s = sorted({int(float(v)) for v in _pn_seed_entries().values() if _pn_re.fullmatch(r"\d+(\.0+)?", v)})
    return s[0] if len(s) == 1 else s


def _pn_chapter6_method():
    """The seeds that Section 6.1 states: the seed of every reported panel and solution,
    and, where the Case III workbook records a second set of solver starts drawn from a
    generator of its own, the seed of that generator (the seed plus the recorded offset)."""
    seed = _pn_seeds()
    if not isinstance(seed, int):
        raise ValueError("the workbooks record several seeds: %r" % (seed,))
    sec, ch = "sec:app-practice", "chapter6.tex"
    keys = [_pn_key(ch, sec, "seed of every reported panel and solution", "isfixedbytheseed", "%d" % seed, seed)]
    c3 = cfgd(load("fig6-8_bdt_case.xlsx"))
    # Section 6.1 names the seed of the second set of Case III solver starts, so the
    # expectation is two-sided: a Case III workbook that records no such seed breaks the
    # layout (the checker then FAILs) instead of silently dropping the key
    if "probe_seed_offset" not in c3:
        raise ValueError("Section 6.1 names the seed of a second set of Case III solver starts, but "
                         "fig6-8_bdt_case.xlsx records no probe_seed_offset")
    s2 = int(float(c3["seed"])) + int(float(c3["probe_seed_offset"]))
    if "probe_seed" in c3 and int(float(c3["probe_seed"])) != s2:
        raise ValueError("fig6-8_bdt_case.xlsx records probe_seed %r, not seed + probe_seed_offset = %d"
                         % (c3["probe_seed"], s2))
    if s2 == seed:
        raise ValueError("the second set of Case III solver starts is recorded under the main seed %d" % seed)
    keys.append(_pn_key(ch, sec, "seed of the second set of Case III solver starts",
                        "secondsetofsolverstartsinCaseIII,whichtheseed", "%d" % s2, s2))
    return None, [], keys


printed_tables, key_numbers, layout_errors, summaries = [], [], {}, {}
for _pn_name, _pn_fn in (("chapter2_running_example", _pn_chapter2), ("chapter4_lp_counts", _pn_chapter4),
                         ("chapter4_figures", _pn_chapter4_figures), ("chapter5_hm_panel", _pn_chapter5), ("worked_examples", _pn_worked_examples),
                         ("case_i", _pn_case_i), ("case_ii", _pn_case_ii), ("case_iii", _pn_case_iii),
                         ("case_iv", _pn_case_iv), ("vignette_v", _pn_vignette_v),
                         ("chapter6_method", _pn_chapter6_method)):
    try:
        _pn_s, _pn_t, _pn_k = _pn_fn()
    except Exception as _pn_exc:  # noqa: BLE001 -- recorded in the manifest and reported by the checker
        layout_errors[_pn_name] = "%s: %s" % (type(_pn_exc).__name__, _pn_exc)
        print("   WARNING: the %s layout could not be read (%s)" % (_pn_name, layout_errors[_pn_name]))
        _pn_tb.print_exc()
        continue
    if _pn_s is not None:
        summaries[_pn_name] = _pn_s
    printed_tables += _pn_t
    key_numbers += _pn_k

manifest = {
    "_generated_by": "figure_code/simulation_data/gen_case_manifest.py -- do not edit by hand; regenerate after build_workbooks.py",
    "_seed": _pn_seeds(),
    "_honesty_rule": (
        "Every study below is an ILLUSTRATIVE SYNTHETIC construction. Each case "
        "mirrors its published application at the level of MECHANISM only (for "
        "Case IV, that equity weighting reorders priorities). No number below is "
        "a finding of any cited study, and the prose must not attribute any of "
        "them to one."),
    "case_iii": case_iii,
    "case_iv": case_iv,
    "hm_panel_anchor": hm_anchor,
    **summaries,
    "printed_tables": printed_tables,
    "key_numbers": key_numbers,
    "_layout_errors": layout_errors,
}

if __name__ == "__main__":
    with open(OUTP, "w", encoding="utf-8", newline="\n") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)
    print("wrote", OUTP, "(%d tables, %d key numbers%s)" % (
        len(printed_tables), len(key_numbers),
        "; LAYOUT ERRORS in " + ", ".join(sorted(layout_errors)) if layout_errors else ""))
