# Simulation data and figure code for *The Social Optimization of Collective Judgment*

[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.22764360.svg)](https://doi.org/10.5281/zenodo.22764360)

Data and code accompanying

> Zhen-Song Chen and Xian-Jia Wang, *The Social Optimization of Collective Judgment: A Multiobjective Framework from Consensus to Human-Machine Collaboration*, Elsevier.

The repository holds two things and what they produce. `simulation_data/` is every
synthetic dataset of the book, together with the single deterministic generator that
creates them and solves every optimization the case studies report. `drawing_code/`
is the script behind each of the 20 data-driven figures of the book. Running
`python reproduce.py` rebuilds the datasets and the figures on your machine and
checks that they match the printed figures byte for byte. Nothing in the figures is
hand-edited, and no printed number of the worked cases is a typed-in constant.

Archived version: https://doi.org/10.5281/zenodo.22764360 · Development version: https://github.com/zschenwhu/social-optimization-collective-judgment · Release 1.0.0 (2026-09-15)

## Layout

```
simulation_data/      the synthetic datasets and their generator
  build_workbooks.py    writes every workbook deterministically (seed 202603) and
                        solves every optimization the case studies report
  gen_case_manifest.py  collects the case-study numbers the text prints
  *.xlsx                the 19 workbooks: the authoritative numbers behind every figure
  case_manifest.json    the printed case-study numbers, read from the workbooks
  INDEX.md              what each workbook holds and which figure or table it backs
drawing_code/         the figure scripts (one per data-driven figure)
  fig*.py               read a workbook, compute, draw; no numeric literals beyond styling
  _common.py            palette, fonts, the shared density plotter, the save routine
  run_all.py            rebuild every workbook, then render every figure
  INDEX.md              figure number, label, script, and workbook for each figure
figures_out/          the rendered figures (PDF and PNG)
reference_figures/    the same 20 figures exactly as printed in the book, plus manifest.json
verify_figures.py     the reproducibility gate (see below)
reproduce.py          run_all followed by verify, in one command
```

The conceptual diagrams of the book (flowcharts and schematic diagrams drawn in
TikZ) carry no data and are not part of this repository.

## Quick start

Requirements: Python 3.9 or later with `numpy`, `scipy`, `pandas`, `openpyxl`, and
`matplotlib` (`pip install -r requirements.txt`). The figure text is set in Times New
Roman; without that font matplotlib substitutes a fallback and the byte-level check
of the PDFs reports differences that are font-only.

```bash
python reproduce.py            # rebuild every workbook and figure, then verify
```

A full rebuild takes about fifteen minutes on a laptop, most of it in the seventy
multi-start solves of the Chapter 6 barrier case; the verification step rebuilds
everything a second time in a temporary copy, so `reproduce.py` runs for about
half an hour in total. To render one figure, run its script from `drawing_code/`,
for example `python fig6-5_dt_expansion.py`; each case script also prints the
numbers of its case study so that a printed table can be checked directly.

## What is verified

`verify_figures.py` copies the repository to a temporary folder, deletes the
shipped workbooks there, rebuilds them with `build_workbooks.py`, renders every
figure, and then checks that

1. every data-driven figure of the book (`reference_figures/manifest.json`) has a
   generator in this repository, and no generator or workbook is an orphan;
2. every regenerated PDF equals the printed one, comparing decompressed content
   streams so that creation timestamps cannot mask a difference;
3. the rebuilt workbooks pass the generator's own self-checks (weight vectors on the
   simplex, probability profiles summing to one, solved values within their
   normalization bounds, the identities the book states).

## Figures of the book

The printed number, the label in the LaTeX source, the file that draws it, and the
workbook it reads. Regenerated from the book sources at release time.

| Fig. | Label | Description | Drawn by | Data |
|---|---|---|---|---|
| 2.1 | `fig:lp-vs-qa` | Linear pool versus quantile averaging | `drawing_code/fig2-1_linear_pool_vs_qa.py` | `fig2-1_linear_pool_vs_qa.xlsx` |
| 2.2 | `fig:lp-lop-qa` | Linear, logarithmic, and quantile pools compared | `drawing_code/fig2-2_three_pooling_operators.py` | `fig2-2_three_pooling_operators.xlsx` |
| 3.1 | `fig:lorenz` | The Gini coefficient and Hoover index on the Lorenz curve | `drawing_code/fig3-1_lorenz_gini_hoover.py` | `fig3-1_lorenz_gini_hoover.xlsx` |
| 4.3 | `fig:scalability` | Interaction counts of the flat and hierarchical programs | `drawing_code/fig4-3_scalability.py` | `fig4-3_scalability.xlsx` |
| 4.4 | `fig:pareto` | Weighted-sum scalarization and a nonconvex Pareto front | `drawing_code/fig4-4_pareto_front.py` | `fig4-4_pareto_front.xlsx` |
| 5.2 | `fig:calsharp` | The calibration profile of a mixed panel | `drawing_code/fig5-2_calibration_sharpness.py` | `fig5-2_calibration_sharpness.xlsx`, `fig5-3_loo.xlsx` |
| 5.3 | `fig:loo-chart` | The leave-one-out contribution of each expert | `drawing_code/fig5-2_calibration_sharpness.py` | `fig5-2_calibration_sharpness.xlsx`, `fig5-3_loo.xlsx` |
| 5.4 | `fig:depnet` | The expert-dependence network of a correlated lineage | `drawing_code/fig5-4_depnet.py` | `fig5-4_depnet.xlsx` |
| 6.1 | `fig:bim-groupagg` | Between-group aggregation for the real-time-data indicator | `drawing_code/fig6-1_bim_expansion.py` | `fig6-1_bim_case.xlsx` |
| 6.2 | `fig:bim-dimensions` | The three dimensions and the global maturity distribution | `drawing_code/fig6-1_bim_expansion.py` | `fig6-1_bim_case.xlsx` |
| 6.3 | `fig:bim-densities` | Expert densities and aggregates for the maturity assessment | `drawing_code/fig6-3_bim_densities.py` | `fig6-3_bim_densities.xlsx` |
| 6.4 | `fig:bim-tradeoff` | The fairness-consensus trade-off across the sweep | `drawing_code/fig6-4_bim_tradeoff.py` | `fig6-4_bim_tradeoff.xlsx` |
| 6.5 | `fig:dt-ranking14` | The fourteen barriers ranked | `drawing_code/fig6-5_dt_expansion.py` | `fig6-5_dt_case.xlsx` |
| 6.6 | `fig:dt-crossswf` | Barrier importance across the five equity objectives | `drawing_code/fig6-5_dt_expansion.py` | `fig6-5_dt_case.xlsx` |
| 6.7 | `fig:dt-triad` | Collective distributions of the leading triad | `drawing_code/fig6-5_dt_expansion.py` | `fig6-5_dt_case.xlsx` |
| 6.8 | `fig:bdt-q1agg` | Intra/inter-group aggregation of an indicator | `drawing_code/fig6-8_bdt_expansion.py` | `fig6-8_bdt_case.xlsx` |
| 6.9 | `fig:bdt-profile` | The seven-dimension BDT maturity profile | `drawing_code/fig6-8_bdt_expansion.py` | `fig6-8_bdt_case.xlsx` |
| 6.10 | `fig:bdt-sensitivity` | Sensitivity to the fairness weight | `drawing_code/fig6-8_bdt_expansion.py` | `fig6-8_bdt_case.xlsx` |
| 6.11 | `fig:rt-equity` | Station accessibility under the two regimes | `drawing_code/fig6-11_rt_expansion.py` | `fig6-11_rt_case.xlsx` |
| 6.12 | `fig:hm-calsharp` | Calibration and resolution of the human and machine forecasters | `drawing_code/fig6-12_hm_calsharp.py` | `fig6-12_hm_calsharp.xlsx` |

## Datasets

Every dataset is illustrative synthetic data generated from stated parametric
families; none is taken from a proprietary elicitation. Each workbook carries a
`config` (or `provenance`) sheet naming its generator block, its seed, and the
synthetic flag.

| Workbook | Sheets | Read by | Backs |
|---|---|---|---|
| `fig2-1_linear_pool_vs_qa.xlsx` | `config`, `gaussians` | `fig2-1_linear_pool_vs_qa.py` | `fig2-1_linear_pool_vs_qa` |
| `fig2-2_three_pooling_operators.xlsx` | `config`, `gaussians` | `fig2-2_three_pooling_operators.py` | `fig2-2_three_pooling_operators` |
| `fig3-1_lorenz_gini_hoover.xlsx` | `config`, `annotations` | `fig3-1_lorenz_gini_hoover.py` | `fig3-1_lorenz_gini_hoover` |
| `fig4-3_scalability.xlsx` | `config`, `series` | `fig4-3_scalability.py` | `fig4-3_scalability` |
| `fig4-4_pareto_front.xlsx` | `config`, `front_pieces` | `fig4-4_pareto_front.py` | `fig4-4_pareto_front` |
| `fig5-2_calibration_sharpness.xlsx` | `config`, `points`, `annotations` | `fig5-2_calibration_sharpness.py` | `fig5-2_calibration_sharpness`, `fig5-3_loo` |
| `fig5-3_loo.xlsx` | `config`, `generator_params`, `experts`, `outcomes`, `forecast_mu`, `loo`, `individual`, `binning_sensitivity`, `solver_log`, `results`, `weights` | `fig5-2_calibration_sharpness.py`, `fig5-3_loo.py` | `fig5-2_calibration_sharpness`, `fig5-3_loo` |
| `fig5-4_depnet.xlsx` | `config`, `nodes`, `correlation` | `fig5-4_depnet.py` | `fig5-4_depnet` |
| `fig6-11_rt_case.xlsx` | `config`, `constituencies`, `stations`, `panel`, `panel_mu`, `results`, `weights_eff`, `weights_eq`, `sqp_starts`, `certification`, `robustness_seeds`, `robustness_fs` | `fig6-11_rt_expansion.py` | `fig6-11_rt_equity` |
| `fig6-12_hm_calsharp.xlsx` | `config`, `forecasters`, `outcomes`, `forecast_mu`, `weights`, `loo_weights`, `solves`, `aggregates` | `fig6-12_hm_calsharp.py` | `fig6-12_hm_calsharp` |
| `fig6-1_bim_case.xlsx` | `config`, `groups`, `dimensions`, `elicitation_q4`, `derived` | `fig6-1_bim_expansion.py` | `fig6-1_bim_groupagg`, `fig6-2_bim_dimensions` |
| `fig6-3_bim_densities.xlsx` | `config`, `gaussians`, `aggregates`, `annotations` | `fig6-3_bim_densities.py` | `fig6-3_bim_densities` |
| `fig6-4_bim_tradeoff.xlsx` | `config`, `series`, `weights`, `overlaps`, `bounds` | `fig6-4_bim_tradeoff.py` | `fig6-4_bim_tradeoff` |
| `fig6-5_dt_case.xlsx` | `config`, `barriers`, `dimensions`, `equity_objectives`, `panel_experts`, `panel`, `elicitation_b1`, `solves`, `weights`, `overlaps`, `utilities`, `matrix`, `ranking`, `per_objective_order`, `triad_density` | `fig6-5_dt_expansion.py` | `fig6-5_dt_ranking14`, `fig6-6_dt_crossswf`, `fig6-7_dt_triad` |
| `fig6-8_bdt_case.xlsx` | `config`, `groups`, `dimensions`, `levels`, `sensitivity`, `expert_attributes`, `kmeans_topsis`, `panel_weibull`, `intra_group_solutions`, `indicator_collectives`, `dimension_collectives`, `q1_quantiles`, `q1_densities`, `global_quantiles` | `fig6-8_bdt_expansion.py` | `fig6-10_bdt_sensitivity`, `fig6-8_bdt_q1agg`, `fig6-9_bdt_profile` |
| `tab2-1_running_example_percentiles.xlsx` | `config`, `table_2_1` | (table only) | a table of the book (no plot) |
| `tab2-5_biobjective_example.xlsx` | `config`, `table_2_5`, `bounds`, `alpha_sweep`, `level_set_alpha_star`, `cross_checks`, `starts` | (table only) | a table of the book (no plot) |
| `tab5-3_hm_panel.xlsx` | `table_5_3`, `provenance` | (table only) | a table of the book (no plot) |
| `tab6-9_bim_results.xlsx` | `config`, `table_6_9`, `table_6_9_printed` | (table only) | a table of the book (no plot) |

## How the numbers are produced

Every number the book prints is an output of `simulation_data/build_workbooks.py`
run under `numpy.random.default_rng(202603)`, in a fixed draw order that the
config sheet of each workbook records. The builder solves every optimization the
case studies report; nothing is carried as a typed result. Two consecutive
builds produce byte-identical workbooks (the zip and document timestamps are
pinned), and the rule-12 self-check pass at the end of the builder re-reads the
written workbooks and aborts on any violated invariant (weight vectors on the
simplex, probability profiles summing to one, solved values within their
normalization bounds, identities the book states). The models, in the order of
the book:

* **Chapter 2 running example** (Tables 2.1, 2.5, Figures 2.1, 2.2). The normal
  fits of Table 2.1 are least squares of the elicited quantiles on the standard
  normal quantiles, rounded to one decimal before use. The quantile average of
  normal experts is the closed form of Proposition 2.1(iii), the logarithmic
  pool is the normalized product (closed-form precision average, asserted against
  numerical normalization), and Table 2.5 is solved by Algorithm 2.2: overlap
  consensus in closed form, confidence as the squared weighted standard
  deviation, min-max bounds solved over the simplex, and every scalarized program
  by SLSQP from 64 Dirichlet(1) starts plus the vertices and the centroid. The
  alpha = 0.75 optimum is a line segment of the simplex (both objectives depend
  on the weights only through the aggregate mean and standard deviation); the
  table prints its midpoint, and the level set is stored.
* **Figures 3.1, 4.3, 4.4** are drawn from the formulas the book states: a
  schematic Lorenz curve, the pairwise interaction counts of the flat and the
  balanced two-stage programs, and a stored Pareto curve from which the hull
  chord and the unsupported point are derived in the script and asserted.
* **Chapter 5 human-machine panel** (Example 5.1, Table 5.3, Figures 5.2, 5.3).
  The T = 250 outcomes and the ten normal forecasts are generated from the
  stated signal model; the convex CRPS program with the influence-dispersion and
  representativeness regularizers is solved by Algorithm 4.2 as printed
  (analytic subgradient, diminishing steps, best iterate, tolerance 1e-6), each
  solve certified against multi-start SLSQP; the human-only program and the ten
  leave-one-out programs are re-solved the same way; the calibration and
  resolution channels follow Hersbach's ensemble decomposition (nineteen
  members, outer regions scored), with the uncertainty term formed on the same
  grid. Figure 5.4's effective panel size is computed from the stored
  correlation matrix.
* **Case I, BIM maturity.** The seven-expert fairness panel is solved from the
  stated means, spreads, and behavioral parameters: whole-line overlaps in
  closed form, Fehr-Schmidt utilities, power normalization, Bonferroni collective
  fairness, min-max bounds solved over the simplex, SLSQP from 200 uniform
  simplex starts for every point of the delta grid. The full assessment is worked
  at the group and dimension level from stated inputs (group opinions, the
  credibility weights, the dimension distributions and weights), from which the
  builder computes the quantile averages; no random draws enter Case I.
* **Case II, digital-transformation barriers.** A 30-expert, 14-barrier
  truncated-Weibull panel is generated (leniency, background bias by
  professional domain, noise, five-level profiles, grouped maximum-likelihood
  fits) with targets calibrated so that the equal-weight quantile average
  reproduces the stated design means. The complete-welfare-and-confidence
  program is solved for every barrier under each of the five inequality indices
  by multi-start SLSQP with solved normalization bounds, and the collective
  means, averages, ranking, active-expert counts, and overlap diagnostics are
  stored.
* **Case III, building digital twin maturity.** A 15-expert, 22-indicator
  truncated-Weibull panel is generated from stated stratum levels and spreads;
  k-means on a stated professional-attribute matrix yields the strata, TOPSIS on
  the centroids the inter-group weights, the best-worst linear model the
  dimension weights; the intra-group fairness-confidence program with the
  influence-dispersion term is solved for every indicator, stratum, and delta by
  multi-start SLSQP; quantile averaging under the inter-group and dimension
  weights gives the dimension and global maturities and the delta sweep.
* **Case IV, rail-transit station accessibility.** The 24-stakeholder panel over
  ten stations is generated from the stated constituency means, leniency, and
  precision; the consensus-only and the equity-aware (Sen abbreviated welfare)
  programs are solved by SLSQP from the uniform vector and 40 Dirichlet(1)
  starts with closed-form overlaps, each solve certified by a dense scan of the
  aggregate mean; the reported representative of each optimal level set is the
  nearest-to-uniform weight vector; the re-seeded and Fehr-Schmidt robustness
  sweeps the book reports are computed and stored.
* **Section 6.6 vignette.** An AR(1) target over 24 quarters and eight normal
  forecasters with stated biases and spreads; the CRPS program with the
  influence-dispersion penalty and a ridge term is solved by Algorithm 4.2 and
  certified against SLSQP, together with the human-only, equal-weight, and eight
  leave-one-out comparators, the calibration fractions, and the Hersbach split
  that places the eight points of Figure 6.12.

The `config` sheet of each workbook lists the solver, the number and law of the
starts, the tolerances, the integration rules and grids, the normalization
bounds and how they were obtained, and the seed, so that any printed number can
be traced to its inputs.

## Citation

If you use this data or code, please cite the book and the archived repository
(`CITATION.cff` carries both in machine-readable form):

```
Chen, Z.-S., & Wang, X.-J. Simulation data and figure code for The Social
Optimization of Collective Judgment (version 1.0.0). Zenodo. https://doi.org/10.5281/zenodo.22764360
```

## License

The code is released under the MIT License. The synthetic datasets and the rendered
figures are released under CC BY 4.0. See `LICENSE`.
