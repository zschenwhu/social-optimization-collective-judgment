# Changelog

## v1.0.0 (2026-10-30)

First release, matching the manuscript of the book as delivered to Elsevier, which names version 1.0.0 as the one that matches its edition. It supersedes the pre-publication draft that carried the tag v1.0.0 on 2026-09-15 and has been withdrawn. For anyone who downloaded that draft, the differences are listed below.

### Differences from the withdrawn draft of 2026-09-15

* `verify_figures.py` also compares every rebuilt workbook, the table-only ones
  included, and the rebuilt `case_manifest.json` with the shipped copies; the
  documentation describes the figure check as identity of the PDF content streams
  (the embedded creation dates differ) rather than of the bytes.
* Indexes: what each script reads and draws is taken from its code
  (`drawing_code/_scriptio.py`), not from its docstring, which had credited
  `fig5-2_calibration_sharpness.py` with Figure 5.3 and with reading
  `fig5-3_loo.xlsx`.
* `case_manifest.json` carries, besides the Case III and Case IV summaries,
  summaries of the Chapter 2 running example, Case I, Case II, and Case V, and
  the full-precision numbers of the tables whose cells are workbook values
  (Tables 2.1, 2.5, 5.3, 6.6 to 6.9, 6.11, 6.12, 6.16 to 6.20) and of the key
  numbers of the prose, each with the words printed before it; the hand-typed and
  outdated figure roster is dropped.
* `build_workbooks.py` checks the seed statement of every workbook and the claims
  made for the generated panels: the Case II calibration tolerance, the Case III
  shape range and member spreads, and the information structure of the Case V
  forecasts.
* Case II: the welfare bounds and the program of every barrier and index are
  now searched over every vertex, a grid on every two-expert edge refined by
  zooming, and a random face sample besides the SLSQP starts; the bounds are
  refined until no solution of the program exceeds them, and each retained
  solution is checked against every candidate and every pairwise transfer.
  The programs behind the three near ties of Table 6.12 and the two flat
  programs are solved again on a doubled discretization. `fig6-5_dt_case.xlsx`
  gains the sheets `starts` (every start of the final solve of every program),
  `overlaps_equal_weights`, `near_ties`, `fine_check`, `fine_weights`, and
  `text_numbers` (every number Section 6.3 prints), the one six-expert solution
  (B8 under EQ1) is replaced by the two-expert optimum it fell short of, Table
  6.11 shows a member of the generated panel, and Figure 6.5 labels its bands
  with the book's level symbols. The barycentre and the Dirichlet(1) starts
  are labeled interior starts, as the book calls them (`kind` 'interior' in
  `starts`, `n_interior_starts` and `n_interior_reach` in `solves`). No printed
  collective mean changes.
* Case I: the full assessment is generated and solved (a 20-expert,
  24-indicator panel from stated design levels, k-means and TOPSIS, the
  consensus-confidence base model by multi-start SLSQP, best-worst dimension
  weights) instead of being worked from stated group-level inputs, and every
  bound and sweep optimum of the fairness panel is checked against a dense grid
  of the attainable (mean, spread) region.
* Figures 3.1 and 4.4 are computed on the Case I fairness panel (the Lorenz curve
  of its utilities at the solved weights; the consensus-fairness Pareto front
  traced by the epsilon-constraint method, with the weighted-sum solutions and
  their concave majorant) instead of being drawn from schematic curves.
* Chapter 5: `fig5-3_loo.xlsx` gains the sheets `group_loo`, `reliability_floor`,
  `error_dependence`, `text_numbers`, `error_correlation`, and
  `forecast_correlation`, and the Hersbach outlier rule is recorded in the config
  sheets that use the decomposition; the nodes of Figure 5.4 are labeled A1 to A5
  and V1 to V5, so that they no longer share the IDs of the experts of Example
  5.1, a different panel.
* Case III: every program is also solved from a second, independent set of
  starts (seed 202603 plus the recorded offset, `probe_seed`), a follow-up
  search covers the programs whose runner-up start falls short, and the best
  further run is retained wherever it ends above the best original start by
  more than 1e-6; the stability of the returned weights over the near-optimal
  runs, the ordinal consistency of the best-worst comparisons, and a
  deterministic local polish from every retained solution (rounds of pairwise
  weight transfers and an SLSQP run until a round gains nothing, a diagnostic
  that replaces no solution) are
  recorded, and every printed number is recomputed with the near-optimal, the
  best original, and the polished weights substituted.
* Case IV: every other local maximum of the certification scan is recorded,
  near-tied optima are re-evaluated in every combination, the LP range of the
  advocates' share over each optimal level set and the leveling-down record in
  the overlaps are stored, and the equity tick of Figure 6.11 reads "equity
  aware (Sen form)".
* Section 6.6 (Case V): the record has 72 quarters; every fitted pool is scored
  in-sample and by its closed-form expected CRPS on new quarters, the design
  level is solved on that expected score for the full panel, the human
  forecasters alone, the machine experts alone, and with each member withheld,
  the record is regenerated and
  re-solved under 500 further seeds, and the leave-one-out channels of M4 are
  recomputed with its bias removed and with its spread set to its error.
* `case_manifest.json` requires the seed of the second set of Case III solver
  starts that Section 6.1 prints (a workbook without it breaks the layout
  instead of dropping the check), and the workbook tables of the indexes list
  the workbooks and the figures each backs in book order.
* The key numbers of `case_manifest.json` also cover the counts of the epigraph
  reformulation in Chapter 4 (new sheet `lp_reformulation_size` of
  `fig4-3_scalability.xlsx`), the finite-grid reliability floor of Chapter 5,
  the Case I global maturity under the published equal dimension weights, and
  the Case III seed of the second set of starts and the bounds of its local
  polish.
* The arithmetic of the short examples that Chapters 2 to 5 work by hand is
  computed by `build_workbooks.py` from the stated inputs and written, each value
  with an independent cross-check and its printed form, to a new sheet
  `worked_examples` of `fig2-2_three_pooling_operators.xlsx`,
  `fig3-1_lorenz_gini_hoover.xlsx`, `fig4-3_scalability.xlsx` and
  `fig5-4_depnet.xlsx`, checked by a new self-check. The Case IV robustness
  claims (which stations change priority in each re-seeded panel and under each
  Fehr-Schmidt setting, and the recorded range of the contested advocate share)
  are asserted, as is the reason Example 5.3 gives for the opposite directions
  of the effective-number change (the correlation among the humans each block
  joins). `case_manifest.json` compares every number of Sections 6.5.5 and 6.5.6,
  the numbers of the worked examples, the numbers the prose of Chapter 4 reads
  from Figures 4.3 and 4.4, the starts, seed and bound of the SLSQP cross-check
  of Example 5.1, the fairness-weight threshold of Case I, and the move of the
  Case III global maturity over the fairness-weight sweep, and
  `check_printed_numbers.py` compares a count printed as a word together with
  the words after it.
* Software stack: `requirements.txt` now pins the exact versions that built the
  shipped workbooks and figures (it gave lower bounds only), and the new
  `environment.py` names that stack and compares the running one with it. The
  versions do not fix the arithmetic by themselves, since NumPy and SciPy pick
  their OpenBLAS kernels from the processor and their thread count from the
  number of processors, so every entry point requests the kernels of the
  reference machine before it imports NumPy, switches off the AVX-512 code paths
  of NumPy, and sets the thread count of the reference machine through
  `threadpoolctl`, which joins the requirements (the drawing scripts import
  `_common.py` first for this reason). The generator prints the stack and, when
  it is not the reference one, stops before writing anything, because a rebuild
  on another stack returns other weight vectors for programs with a non-unique
  optimum and other solver diagnostics; `SOCJ_ALLOW_OTHER_STACK=1` lets it build
  there with the self-checks reported instead of enforced. A package that is not
  installed is named in a plain message. `verify_figures.py` prints the stack, rebuilds in its
  temporary copy on any stack, and prints the largest numerical difference of
  each workbook that differs. `reproduce.py` leaves the shipped files untouched
  on another stack. A scalar formatted through its `repr` in the Case I block,
  which stopped the generator under NumPy 2, is formatted through `float`.
  The case scripts print their numbers under the heading "NUMBERS READ FROM THE
  WORKBOOK" (it read "LOCKED"). No workbook and no figure changed.
* Naming: the book now presents the former exploratory vignette of Section 6.6
  as the fifth of its five case studies, "Case Study V" (Case V). The code, the
  workbook configuration sheets, and `case_manifest.json` keep `vignette` as an
  identifier.
* Section 6.6 (Case V): the forecasters are built only from the information
  available when a forecast is issued. The AR(1) innovation is split into a
  leading signal, observable at issue time, and a residual that no forecaster
  sees, and each forecast mean is the conditional mean given the previous value
  plus a noisy reading of the signal, so the realized value never enters a
  forecast. `fig6-12_hm_calsharp.xlsx`, Figure 6.12, and Table 6.20 are
  regenerated, and the workbook gains the columns and aggregates the revised
  text prints (the signal split, the error of each mean, the leave-one-out
  robustness, the spread and error of the calibrated pool).
* Case III: `fig6-8_bdt_case.xlsx` records, for every intra-group program, how
  many starts reach the retained value and how far the best of the others falls
  short, the multi-start agreement the book states in place of a claim of
  stationarity.
