# Simulation data and figure code for *The Social Optimization of Collective Judgment*

[![DOI](https://zenodo.org/badge/DOI/10.5281/zenodo.22764359.svg)](https://doi.org/10.5281/zenodo.22764359)

Data and code accompanying

> Zhen-Song Chen and Xian-Jia Wang, *The Social Optimization of Collective Judgment: A Multiobjective Framework from Consensus to Human-Machine Collaboration*, Elsevier.

The repository holds two things and what they produce. `simulation_data/` is every
synthetic dataset of the book, together with the single deterministic generator that
creates them and solves every optimization the case studies report. `drawing_code/`
is the script behind each of the 20 data-driven figures of the book. On the
reference software stack (`requirements.txt`), running
`python reproduce.py` rebuilds the datasets and the figures on your machine and
checks that every rebuilt dataset equals the shipped one and that every regenerated
figure has the same PDF content streams as the printed one (only the embedded
creation date differs). Nothing in the figures is hand-edited, and no printed result
of the worked cases is a typed-in constant: the stated inputs are recorded in the
workbooks, and every result is computed from them.

Archived version: https://doi.org/10.5281/zenodo.22764359 · Development version: https://github.com/zschenwhu/social-optimization-collective-judgment · Release 1.0.0 (2026-10-30)

## Layout

```
simulation_data/      the synthetic datasets and their generator
  build_workbooks.py    writes every workbook deterministically (seed 202603) and
                        solves every optimization the case studies report
  gen_case_manifest.py  collects the numbers of the book's workbook-backed tables and
                        the key numbers of its prose
  *.xlsx                the 19 workbooks: the authoritative numbers behind every figure
  case_manifest.json    those numbers at full precision and a readable summary of each
                        study, all read from the workbooks
  INDEX.md              what each workbook holds and which figure or table it backs
drawing_code/         the figure scripts (one per data-driven figure)
  fig*.py               read a workbook, compute, draw; no numeric literals beyond styling
  _common.py            palette, fonts, the shared density plotter, the save routine
  _scriptio.py          what each script loads and saves, read from its code
  run_all.py            rebuild every workbook, then render every figure
  INDEX.md              figure number, label, script, and workbook for each figure
figures_out/          the rendered figures (PDF and PNG)
reference_figures/    the same 20 figures exactly as printed in the book, plus manifest.json
verify_figures.py     the reproducibility gate (see below)
reproduce.py          run_all followed by verify, in one command
environment.py        the software stack that built the shipped files, and a check of yours
requirements.txt      the exact versions of that stack
```

The conceptual diagrams of the book (flowcharts and schematic diagrams drawn in
Microsoft Visio) carry no data and are not part of this repository.

## Quick start

Requirements: the reference software stack, which is Python 3.9 with `numpy` 1.26.4,
`scipy` 1.13.1, `pandas` 2.3.3, `openpyxl` 3.1.5, and `matplotlib` 3.8.4, together
with `threadpoolctl` (any recent version), through which the pack sets the number of
threads of the numerical library. `pip install -r requirements.txt` installs exactly
these versions (on Python 3.9 to 3.12), and `python environment.py` compares the stack
you are running with them. The
figure text is set in Times New Roman; without that font matplotlib substitutes a
fallback and the content-stream check of the PDFs reports differences that are
font-only.

```bash
pip install -r requirements.txt
```

```bash
python environment.py          # is this the reference stack?
```

```bash
python reproduce.py            # rebuild every workbook and figure, then verify
```

A full rebuild takes between half an hour and an hour on a laptop, most of it in
the searched solves of the seventy programs of the Chapter 6 barrier case and the
multi-start solves of the 396 programs of the digital twin case; the
verification step rebuilds everything a second time in a temporary copy, so
`reproduce.py` runs for about twice as long. To render one figure, run its script from `drawing_code/`,
for example `python fig6-5_dt_expansion.py`; each case script also prints the
numbers of its case study so that a printed table can be checked directly.

## What is verified

`verify_figures.py` copies the repository to a temporary folder, deletes the
shipped workbooks and `case_manifest.json` there, rebuilds them with
`build_workbooks.py` and `gen_case_manifest.py`, renders every figure, and then
checks that

1. every data-driven figure of the book (`reference_figures/manifest.json`) has a
   generator in this repository, and no generator or workbook is an orphan, what a
   script loads and draws being read from its code rather than its documentation;
2. every regenerated PDF equals the printed one, comparing the decompressed content
   streams (the files themselves differ only by their embedded creation dates);
3. every rebuilt workbook, the table-only ones included, has the same content in
   every sheet as the shipped one (the gate also reports which workbooks are
   byte-identical; on the software stack that built the shipped copy all of them
   are), and the rebuilt `case_manifest.json` equals the shipped one;
4. the rebuilt workbooks pass the generator's own self-checks (weight vectors on the
   simplex, probability profiles summing to one, solved values within normalization
   bounds that are themselves solved over the simplex, the identities the book
   states, and the properties the book and the config sheets claim for the
   generated panels); Case III scales its terms by ranges sampled over the simplex,
   which a retained solution may exceed slightly.

The delivery copy of the book runs two further comparisons, which need the LaTeX
sources and are not part of this repository. The first compares the numbers of the
tables listed in `case_manifest.json` and the key numbers of the prose, each found
after the words printed before it, with the manifest at the printed precision. The
second goes figure by figure: every number a figure prints must be the workbook value
at the printed precision and must agree with the table row or the sentence of the
book that prints the same quantity, and every line, bar, and point a figure draws
must be a workbook column or follow from workbook cells by the formula the book
states (a normal density, a pool of densities, a Lorenz curve, a count of pairs).

These checks demand identity, and identity is to be expected only on the reference
stack. The next section says what happens on another one.

## Exact rebuilds and other software stacks

The shipped workbooks and figures were built on the reference stack named in
`environment.py`, on Windows 11 (x86-64). On that platform a rebuild is identical,
every workbook byte for byte and every figure with the same PDF content streams.

The generator is deterministic on every stack, since every draw and every solver start
comes from a fixed seed. But the multi-start solves of the nonconvex programs are
sensitive to round-off, and round-off depends on three things. The first is the
versions of NumPy and SciPy, which `requirements.txt` fixes. The second is the
numerical kernels. NumPy and SciPy each bundle an OpenBLAS library that picks its
kernels from the processor it detects, so the same versions can round differently on
two machines. The third is the number of OpenBLAS threads, which follows the number of
processors, since a matrix product split over another number of threads sums in another
order. The pack therefore requests, before NumPy is imported, the kernels the reference
machine runs (`KERNELS` in `environment.py`), switches off the AVX-512 code paths of
NumPy, which the reference processor does not have, and sets the thread count of the
reference machine (`THREADS`) through the OpenBLAS interface. With the pinned versions
on an x86-64 processor with AVX2, the pack then does the same arithmetic as the
reference machine.

On another stack the collective distributions, rankings, and scores agree with the
shipped ones to rounding, and the following can differ.

* The weight vector recorded for a program whose optimum is not unique (other weights,
  the same aggregate).
* Solver diagnostics: how many starts reach a retained value, which start is retained,
  and the small gains of the further starts and of the local polish in Case III.
* The placement of text in the figures, by about one point, which comes from the font
  metrics of matplotlib and changes the size of the tightly cropped page.

A rebuild was tried on a second stack on the machine that built the shipped files
(Python 3.11, NumPy 1.24.3, SciPy 1.10.1, pandas 2.1.4, openpyxl 3.0.10, matplotlib
3.8.0). Every one of the 14 workbook-backed tables of the book was reproduced at the
precision it prints, and so were all but three of the key numbers of its prose. The three
that differed are solver diagnostics of Case III that the book states as bounds: the
largest gain of the further starts (1.3e-5 there against the printed 1.2e-5) and the
largest gain and weight move of the local polish (2.7e-5 and 1.3e-3 against the printed
4.1e-5 and 1.8e-3). Five self-checks of the generator, all on Case III diagnostics of
this kind, did not hold there. Four of the 19 workbooks were identical in every cell,
the others differed in the ways listed above, and each figure showed the same data
with its text moved by about a point.

The kernels and the thread count were tried separately, with the pinned versions and
the request of the pack switched off. With NumPy's OpenBLAS on its Haswell kernels in
place of the Prescott ones that the reference machine runs, 15 of the 19 workbooks
differed, six self-checks on Case III diagnostics did not hold, and one local polish of
Case III stopped at its round cap, so that the manifest of that case could not be
written in its usual form. With one OpenBLAS thread in place of the 24 of the reference
machine, 18 workbooks were identical and the Case II workbook differed in which start
is retained at near ties, its collective means moving by less than 1e-8, and every
table and key number of the book was still reproduced. Switching off the AVX2 and FMA
code paths of NumPy changed nothing in the 18 workbooks written before that run was
stopped. With the request of the pack in place, a rebuild under outside settings that
ask for other kernels and another thread count returned the shipped files.

Four safeguards follow from this.

1. Every entry point of the pack requests the reference kernels and thread count
   before it imports NumPy, and `python environment.py` prints the versions, the
   kernels, and the thread counts in use.
2. `build_workbooks.py` prints the running stack and, when it is not the reference
   one, stops before writing anything, so that the shipped workbooks cannot be replaced
   by accident with ones that no longer match the numbers printed in the book. A
   package that is not installed is named in a plain message.
3. `verify_figures.py` rebuilds in a temporary copy on any stack, never touches the
   shipped files, and prints for each workbook that differs the largest numerical
   difference.
4. Setting the environment variable `SOCJ_ALLOW_OTHER_STACK=1` lets the generator build
   on another stack. The self-checks that compare a solver diagnostic with the value
   printed in the book are then reported instead of enforced, and the workbooks written
   are those of that stack, not the shipped ones.

The exact rebuild has been verified on the reference platform only. Other operating
systems, processors without AVX2, and processors that are not x86-64 are untested, and
a difference there is round-off in the system libraries, not an error in the shipped
data. NumPy 2 changes how a scalar prints and is outside the pins. The generator has
not been tested on it.

## Figures of the book

The printed number, the label in the LaTeX source, the file that draws it, and the
workbook it reads. Regenerated from the book sources at release time.

| Fig. | Label | Description | Drawn by | Data |
|---|---|---|---|---|
| 2.1 | `fig:lp-vs-qa` | LP and quantile average of two equally weighted normal experts | `drawing_code/fig2-1_linear_pool_vs_qa.py` | `fig2-1_linear_pool_vs_qa.xlsx` |
| 2.2 | `fig:lp-lop-qa` | The three pooling operators applied to the four equally weighted synthetic experts of Example 2.1 | `drawing_code/fig2-2_three_pooling_operators.py` | `fig2-2_three_pooling_operators.xlsx` |
| 3.1 | `fig:lorenz` | Lorenz curve, Gini coefficient, and Hoover index of the Case I fairness panel's utilities at $\delta = 0$ | `drawing_code/fig3-1_lorenz_gini_hoover.py` | `fig3-1_lorenz_gini_hoover.xlsx` |
| 4.3 | `fig:scalability` | Pairwise interaction counts of the flat program and a balanced representative-level hierarchy | `drawing_code/fig4-3_scalability.py` | `fig4-3_scalability.xlsx` |
| 4.4 | `fig:pareto` | The Case I fairness panel's consensus-fairness Pareto front | `drawing_code/fig4-4_pareto_front.py` | `fig4-4_pareto_front.xlsx` |
| 5.2 | `fig:calsharp` | Reliability error against resolution of each expert's forecasts in the panel of Example 5.1 | `drawing_code/fig5-2_calibration_sharpness.py` | `fig5-2_calibration_sharpness.xlsx` |
| 5.3 | `fig:loo-chart` | Calibration and resolution channels of each expert's leave-one-out contribution in Example 5.1 | `drawing_code/fig5-3_loo.py` | `fig5-3_loo.xlsx` |
| 5.4 | `fig:depnet` | The expert-dependence network of Example 5.3, in which five model variants form a dense clique | `drawing_code/fig5-4_depnet.py` | `fig5-4_depnet.xlsx` |
| 6.1 | `fig:bim-groupagg` | The three solved group opinions on the real-time-data indicator $q_4$ and their inter-group QA collective | `drawing_code/fig6-1_bim_expansion.py` | `fig6-1_bim_case.xlsx` |
| 6.2 | `fig:bim-dimensions` | Technology, organization, and policy maturity distributions and the global BIM maturity | `drawing_code/fig6-1_bim_expansion.py` | `fig6-1_bim_case.xlsx` |
| 6.3 | `fig:bim-densities` | Seven expert judgments and the collective aggregates at $\delta=0$ and $\delta=0.5$ for the maturity panel | `drawing_code/fig6-3_bim_densities.py` | `fig6-3_bim_densities.xlsx` |
| 6.4 | `fig:bim-tradeoff` | Consensus $\bar{A}$, collective fairness $\bar{\Phi}$, and minimum overlap across the fairness-weight sweep | `drawing_code/fig6-4_bim_tradeoff.py` | `fig6-4_bim_tradeoff.xlsx` |
| 6.5 | `fig:dt-ranking14` | The fourteen barriers ranked by mean collective importance averaged over the five equity objectives | `drawing_code/fig6-5_dt_expansion.py` | `fig6-5_dt_case.xlsx` |
| 6.6 | `fig:dt-crossswf` | Collective barrier means under the five inequality indices, with barriers in final ranked order | `drawing_code/fig6-5_dt_expansion.py` | `fig6-5_dt_case.xlsx` |
| 6.7 | `fig:dt-triad` | Collective distributions of the leading triad $B_{14}$, $B_6$, and $B_1$ under the Gini criterion $\mathrm{EQ}_4$ | `drawing_code/fig6-5_dt_expansion.py` | `fig6-5_dt_case.xlsx` |
| 6.8 | `fig:bdt-q1agg` | Three subgroup opinions on the response-and-control indicator $q_1$ and their inter-group collective | `drawing_code/fig6-8_bdt_expansion.py` | `fig6-8_bdt_case.xlsx` |
| 6.9 | `fig:bdt-profile` | Maturity of the seven dimensions and the global maturity on the five-level scale | `drawing_code/fig6-8_bdt_expansion.py` | `fig6-8_bdt_case.xlsx` |
| 6.10 | `fig:bdt-sensitivity` | Collective fairness, weight Gini, aggregate variance, and global maturity against the fairness weight $\delta$ | `drawing_code/fig6-8_bdt_expansion.py` | `fig6-8_bdt_case.xlsx` |
| 6.11 | `fig:rt-equity` | The collective accessibility of the ten stations under the consensus-only and the equity-aware regimes | `drawing_code/fig6-11_rt_expansion.py` | `fig6-11_rt_case.xlsx` |
| 6.12 | `fig:hm-calsharp` | Each human and machine forecaster on the reliability–resolution plane | `drawing_code/fig6-12_hm_calsharp.py` | `fig6-12_hm_calsharp.xlsx` |

## Datasets

Every dataset is illustrative synthetic data, either stated directly (for instance
the elicited percentiles of the Chapter 2 running example and the seven-expert
fairness panel of Case I) or generated from stated parametric families under the
seed 202603 (a robustness sweep that regenerates a panel records the further seeds
it uses, and the second set of Case III solver starts is drawn from a generator
seeded by 202603 plus the offset its config sheet records); none is taken from a
proprietary elicitation. Each workbook carries a
`config` (or `provenance`) sheet naming its generator block, its seed (or the
statement that it involves no random draw), and the synthetic flag.

| Workbook | Sheets | Read by | Backs |
|---|---|---|---|
| `fig2-1_linear_pool_vs_qa.xlsx` | `config`, `gaussians` | `fig2-1_linear_pool_vs_qa.py` | `fig2-1_linear_pool_vs_qa` |
| `tab2-1_running_example_percentiles.xlsx` | `config`, `table_2_1` | (table only) | Table 2.1 |
| `fig2-2_three_pooling_operators.xlsx` | `config`, `gaussians`, `ex_heterogeneous`, `worked_examples` | `fig2-2_three_pooling_operators.py` | `fig2-2_three_pooling_operators` |
| `tab2-5_biobjective_example.xlsx` | `config`, `table_2_5`, `bounds`, `alpha_sweep`, `window_edges`, `window_edge_checks`, `level_set_alpha_star`, `cross_checks`, `starts` | (table only) | Table 2.5 |
| `fig3-1_lorenz_gini_hoover.xlsx` | `config`, `profile`, `lorenz`, `summary`, `worked_examples` | `fig3-1_lorenz_gini_hoover.py` | `fig3-1_lorenz_gini_hoover` |
| `fig4-3_scalability.xlsx` | `config`, `series`, `summary`, `lp_reformulation_size`, `worked_examples` | `fig4-3_scalability.py` | `fig4-3_scalability` |
| `fig4-4_pareto_front.xlsx` | `config`, `front`, `hull`, `bridges`, `weighted_sum`, `summary`, `printed` | `fig4-4_pareto_front.py` | `fig4-4_pareto_front` |
| `fig5-2_calibration_sharpness.xlsx` | `config`, `points`, `annotations` | `fig5-2_calibration_sharpness.py` | `fig5-2_calibration_sharpness` |
| `fig5-3_loo.xlsx` | `config`, `generator_params`, `experts`, `outcomes`, `forecast_mu`, `loo`, `group_loo`, `individual`, `binning_sensitivity`, `reliability_floor`, `error_dependence`, `text_numbers`, `error_correlation`, `forecast_correlation`, `solver_log`, `results`, `weights` | `fig5-3_loo.py` | `fig5-3_loo`, Table 5.3 |
| `tab5-3_hm_panel.xlsx` | `table_5_3`, `provenance` | (table only) | Table 5.3 |
| `fig5-4_depnet.xlsx` | `config`, `nodes`, `correlation`, `worked_examples` | `fig5-4_depnet.py` | `fig5-4_depnet` |
| `fig6-1_bim_case.xlsx` | `config`, `groups`, `dimensions`, `elicitation_q4`, `derived`, `expert_attributes`, `panel`, `intra_group_solutions`, `indicator_collectives`, `bwm`, `bwm_summary`, `bwm_cri_thresholds` | `fig6-1_bim_expansion.py` | `fig6-1_bim_groupagg`, `fig6-2_bim_dimensions`, Table 6.6, Table 6.7, Table 6.8 |
| `fig6-3_bim_densities.xlsx` | `config`, `gaussians`, `aggregates`, `annotations`, `tail_mass` | `fig6-3_bim_densities.py` | `fig6-3_bim_densities` |
| `fig6-4_bim_tradeoff.xlsx` | `config`, `series`, `weights`, `overlaps`, `bounds`, `hull_scan`, `text_values` | `fig6-4_bim_tradeoff.py` | `fig6-4_bim_tradeoff`, Table 6.9 |
| `fig6-5_dt_case.xlsx` | `config`, `barriers`, `dimensions`, `equity_objectives`, `panel_experts`, `panel`, `elicitation_b1`, `solves`, `starts`, `overlaps_equal_weights`, `near_ties`, `fine_check`, `fine_weights`, `text_numbers`, `weights`, `overlaps`, `utilities`, `matrix`, `ranking`, `per_objective_order`, `triad_density` | `fig6-5_dt_expansion.py` | `fig6-5_dt_ranking14`, `fig6-6_dt_crossswf`, `fig6-7_dt_triad`, Table 6.11, Table 6.12 |
| `fig6-8_bdt_case.xlsx` | `config`, `groups`, `dimensions`, `levels`, `sensitivity`, `expert_attributes`, `kmeans_topsis`, `panel_weibull`, `intra_group_solutions`, `indicator_collectives`, `dimension_collectives`, `q1_quantiles`, `q1_densities`, `global_quantiles`, `multistart_followup`, `multistart_followup_runs`, `multistart_substitution` | `fig6-8_bdt_expansion.py` | `fig6-8_bdt_q1agg`, `fig6-9_bdt_profile`, `fig6-10_bdt_sensitivity`, Table 6.16, Table 6.17 |
| `tab6-9_bim_results.xlsx` | `config`, `table_6_9`, `table_6_9_printed`, `band_probabilities` | (table only) | Table 6.9 |
| `fig6-11_rt_case.xlsx` | `config`, `constituencies`, `stations`, `panel`, `panel_mu`, `results`, `weights_eff`, `weights_eq`, `sqp_starts`, `certification`, `near_ties`, `near_tie_combinations`, `realized_panel`, `robustness_seeds`, `robustness_fs` | `fig6-11_rt_expansion.py` | `fig6-11_rt_equity`, Table 6.18, Table 6.19 |
| `fig6-12_hm_calsharp.xlsx` | `config`, `forecasters`, `outcomes`, `forecast_mu`, `weights`, `loo_weights`, `design_loo_weights`, `solves`, `design_solves`, `reseed`, `reseed_summary`, `typicality`, `resolution_by_L`, `channel_counterfactuals`, `aggregates` | `fig6-12_hm_calsharp.py` | `fig6-12_hm_calsharp`, Table 6.20 |

## How the numbers are produced

Every number that the book's data figures, its workbook-backed tables, and its five
case studies print is an output of
`simulation_data/build_workbooks.py`. Stated inputs (the elicited percentiles of
Chapter 2, design levels and means, the seven-expert fairness panel of Case I) are
written to the workbooks as they are given. Every random draw and every random
solver start behind a printed result comes from
`numpy.random.default_rng(202603)` or, where the config sheet records an offset of
that seed (the second set of solver starts of Case III), from a generator seeded
by 202603 plus that offset, in a fixed draw order that the config sheet of each
workbook records (a robustness sweep that regenerates a panel records the further
seeds it uses), and a workbook that involves no random draw says so in its `seed`
entry. The arithmetic that the short examples of Chapters 2 to 5 work by hand is
computed from their stated inputs and written, each value with an independent
cross-check and its printed form, to a sheet `worked_examples` of the Figure 2.2,
3.1, 4.3 and 5.4 workbooks. Values that an example only stipulates, such as the weights
and scores that the correlated-lineage example of Chapter 5 supposes, are not
generated (Figure 5.4, which draws that example's correlation structure, is). The
builder solves every optimization the case studies report; nothing is carried as a
typed result. Two consecutive builds produce byte-identical workbooks (the zip and
document timestamps are pinned), and `verify_figures.py` requires every rebuilt
workbook to hold the same content, sheet by sheet, as the shipped one. The rule-12
self-check pass at the end of the builder re-reads the written workbooks and
aborts on any violated invariant (weight vectors on the simplex, probability
profiles summing to one, solved values within normalization bounds that are
themselves solved over the simplex, the identities the book states, and the
properties the book and the config sheets claim for the generated panels). Case
III scales its fairness and confidence terms by ranges sampled over the simplex,
which a retained solution may exceed slightly. The models, in the order of the
book:

* **Chapter 2 running example** (Tables 2.1, 2.5, Figures 2.1, 2.2, Example 2.2).
  The normal fits of Table 2.1 are least squares of the elicited quantiles on the
  standard normal quantiles, rounded to one decimal before use. The quantile
  average of normal experts is the closed form of Proposition 2.1(iii), the
  logarithmic pool is the normalized product (closed-form precision average,
  asserted against numerical normalization), and Table 2.5 is solved by Algorithm
  2.2: overlap consensus in closed form, confidence as the squared weighted
  standard deviation, min-max bounds solved over the simplex, and every scalarized
  program by SLSQP from 64 Dirichlet(1) starts plus the vertices and the centroid.
  The alpha = 0.75 optimum is a line segment of the simplex (both objectives
  depend on the weights only through the aggregate mean and standard deviation);
  the table prints the midpoint of that segment, taken after an interior optimizer
  is polished by Nelder-Mead in the aggregate pair (w.mu, w.sigma) (xatol 1e-9,
  fatol 1e-14), and the level set is stored. The two edges of the
  interior-compromise window are located by bisection on the first-order
  optimality test of the adjacent vertex and confirmed by the full multi-start
  solve on either side of each edge (sheets `window_edges` and
  `window_edge_checks` of the Table 2.5 workbook). The moments of Example 2.2 are
  integrated by scipy quad, each expert's mean and standard deviation from its own
  quantile function and the aggregate mean and variance from the averaged quantile
  function; the location-scale value (sum w sigma)^2 is formed from the experts'
  standard deviations. All are cross-checked by closed forms (sheet
  `ex_heterogeneous` of the Figure 2.2 workbook).
* **Figures 3.1 and 4.4** are computed on the seven-expert Case I fairness panel
  with the functions of the CASE STUDY I, FAIRNESS PANEL block of
  build_workbooks.py. Figure 3.1 is the Lorenz curve of the panel's
  power-normalized fairness utilities at the solved delta = 0 weights; its Gini
  coefficient (pairwise formula and twice the Lorenz area) and Hoover index (half
  the relative mean deviation and the largest Lorenz gap) are computed both ways
  and asserted equal. Figure 4.4 is the consensus-fairness Pareto front of the
  panel, traced by the epsilon-constraint method (maximum fairness subject to a
  floor on the mean overlap, 161 equally spaced floors and 40 refinement floors
  around the deepest point of the dent, SLSQP from the previous floor's solution
  and two Dirichlet(1) starts under seed 202603), with 51 weighted-sum solutions
  (both objectives min-max normalized by their extremes over the whole simplex, as
  in Case I), the concave majorant of all those efficient points (epsilon points
  and weighted sums together), and the deepest unsupported stretch below it, whose
  ends are the weighted sums on either side of the tie weighting and inside which
  no weighted sum lies (two far shallower dents adjoining its ends are recorded in
  the 'bridges' sheet); an inset plots the gap between the majorant and the front
  on its own scale.
* **Figure 4.3** is drawn from the formulas the book states, the pairwise
  interaction counts of the flat program and of the balanced two-stage program
  whose inter-group stage is posed over the subgroup representatives (no random
  draws).
* **Chapter 5 human-machine panel** (Example 5.1, Table 5.3, Figures 5.2, 5.3).
  The T = 250 outcomes and the ten normal forecasts are generated from the stated
  signal model; the convex CRPS program with the influence-dispersion and
  representativeness regularizers is solved by Algorithm 4.2 as printed (analytic
  subgradient, diminishing steps, best iterate, tolerance 1e-6), each solve
  cross-checked against multi-start SLSQP; the human-only program and the ten
  leave-one-out programs are re-solved the same way; the calibration and
  resolution channels follow Hersbach's ensemble decomposition (nineteen members,
  outer regions scored), with the uncertainty term formed on the same grid. Figure
  5.4's effective panel size is computed from the stored correlation matrix.
* **Case I, BIM maturity.** The seven-expert fairness panel is solved from the
  stated means, spreads, and behavioral parameters: whole-line overlaps in closed
  form, Fehr-Schmidt utilities, power normalization, Bonferroni collective
  fairness, min-max bounds solved over the simplex, SLSQP from 200 uniform simplex
  starts for every point of the delta grid; every bound and sweep optimum is then
  checked against a dense grid of the attainable (mean, spread) region, the convex
  hull of the seven experts' (mu, sigma). Within Case I, the Dirichlet starting
  points of these multi-start solves are the fairness panel's only random draws.
  The full assessment generates a 20-expert, 24-indicator panel (latent normal
  judgments from stated dimension design levels, indicator and stratum offsets,
  leniency, noise, and stratum spread ranges), reports each judgment as a
  ten-level profile rounded to hundredths and fits it by its moments; k-means on a
  stated professional-attribute matrix yields the three strata, TOPSIS on the
  centroids the credibility weights; the consensus-confidence base model (alpha =
  0.75, exact overlaps and gradients, solved min-max bounds) is solved for every
  group and indicator by multi-start SLSQP; quantile averaging under the
  credibility weights and the best-worst linear-model weights (input-based
  consistency ratios of Liang et al. 2020 recorded) gives the indicator,
  dimension, and global maturities.
* **Case II, digital-transformation barriers.** A 30-expert, 14-barrier
  truncated-Weibull panel is generated (leniency, background bias by professional
  domain, noise, five-level profiles, grouped maximum-likelihood fits) with
  targets calibrated so that the equal-weight quantile average reproduces the
  stated design means. The complete-welfare-and-confidence program is solved for
  every barrier under each of the five inequality indices by multi-start SLSQP
  together with a search of every vertex, a grid on every two-expert edge refined
  by zooming, and a random sample of the faces of three to six experts; the
  normalization bounds are refined until no solution of the program leaves them,
  and every retained solution is checked against all of these candidates and every
  pairwise weight transfer. The programs behind the adjacent pairs that tie at two
  decimals and the flat programs (whose near-optimal starts move the collective
  mean) are solved again on a doubled discretization (4,000 probabilities, 1,200
  bins). The collective means, averages, ranking, every start of the final solve
  of every program (restart stability), the doubled-discretization solutions, the
  active-expert and overlap (leveling-down) diagnostics, and every number the text
  of Section 6.3 prints (sheet `text_numbers`) are stored. The profile of Table
  6.11 is a member of the generated panel, chosen by a stated rule.
* **Case III, building digital twin maturity.** A 15-expert, 22-indicator
  truncated-Weibull panel is generated from stated stratum levels and spreads, the
  member standard deviations being recorded in closed form; k-means on a stated
  professional-attribute matrix yields the strata, TOPSIS on the centroids the
  inter-group weights, and the best-worst linear model the dimension weights (the
  input-based consistency ratio of Liang et al. 2020 and their ordinal-consistency
  condition recorded); the intra-group fairness-confidence program with the
  influence-dispersion term is solved for every indicator, stratum, and delta by
  multi-start SLSQP from distinct starts, the number of original starts reaching
  the retained value being recorded for every program. A second, independent set
  of starts for every program, drawn under the seed 202603 plus a recorded offset,
  and a follow-up search of every program whose runner-up start falls short of its
  best one measure the accuracy of the solve, and the best further run is retained
  wherever it ends above the best original start by more than 1e-6. The stability
  of the returned weights over the near-optimal runs is recorded; a deterministic
  local polish from every retained solution (rounds of pairwise weight transfers
  and an SLSQP run, repeated until a round gains nothing) measures how far a
  retained value lies below a local optimum without replacing it; and every
  printed number is recomputed with the near-optimal, the best original, and the
  polished weights substituted in all programs at once. Quantile averaging under
  the inter-group and dimension weights gives the dimension and global maturities
  and the delta sweep.
* **Case IV, rail-transit station accessibility.** The 24-stakeholder panel over
  ten stations is generated from the stated constituency means, leniency, and
  precision; the consensus-only and the equity-aware (Sen abbreviated welfare)
  programs are solved by SLSQP from the uniform vector and 40 Dirichlet(1) starts
  with closed-form overlaps, each solve checked against a dense scan of the aggregate
  mean that also records every other local maximum, those within a relative 1e-4
  of the optimum being near-ties and every combination of near-tied optima being
  re-evaluated; the selected representative of each optimal level set is the
  nearest-to-uniform weight vector, and the LP range of the advocates' share over
  each optimal level set records what the set itself fixes; the consensus optimum
  is checked to be one member's own mean; the re-seeded and Fehr-Schmidt
  robustness sweeps, the leveling-down record in the utilities and in the overlaps
  (the latter also over every combination of near-tied optima), the swap margin,
  the early-warning threshold, and the realized panel behind the largest
  uncontested gap are computed and stored.
* **Section 6.6, Case Study V** (`vignette` in the code and the workbooks).
  An AR(1) target over 72 quarters whose innovation is
  split into a leading signal, observable when the forecasts are issued, and a
  residual no forecaster sees; eight normal forecasters built only from that
  issue-time information (the previous value and a noisy reading of the signal)
  with stated biases, error scales, and spreads, so the realized value never
  enters a forecast; the CRPS program with the influence-dispersion penalty and a
  ridge term is solved by Algorithm 4.2 with the exact normal cdf and
  cross-checked by SLSQP on its smooth epigraph form, together with the
  human-only, equal-weight, and eight leave-one-out comparators, the calibration
  fractions, and the Hersbach split that places the eight points of Figure 6.12;
  every fitted pool is scored in-sample and by its closed-form expected CRPS on
  new quarters, the design-level program is solved on that expected score, for the
  full panel, the human forecasters alone, the machine experts alone, and with each
  member withheld, and the record is regenerated and re-solved under 500 further seeds, the frequency of
  each claim of the text and the percentile of each printed or plotted value of
  the draw being stored; the leave-one-out channels of M4 are also recomputed at
  fixed weights with its bias removed and with its spread set to its error, and
  the full panel is re-solved with that spread in the draw, at the design level,
  and in every regenerated record.

The `config` sheet of each workbook lists the solver, the number and law of the
starts, the tolerances, the integration rules and grids, the normalization bounds
and how they were obtained, and the seed, so that any printed number can be traced
to its inputs.

## Citation

If you use this data or code, please cite the book and the archived repository
(`CITATION.cff` carries both in machine-readable form):

```
Chen, Z.-S., & Wang, X.-J. Simulation data and figure code for The Social
Optimization of Collective Judgment (version 1.0.0). Zenodo. https://doi.org/10.5281/zenodo.22764359
```

## License

The code is released under the MIT License. The synthetic datasets and the rendered
figures are released under CC BY 4.0. See `LICENSE`.
