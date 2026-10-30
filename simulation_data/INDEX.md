# Simulation data index

Every dataset behind a figure or a data table of the book is an illustrative synthetic
construction written by `build_workbooks.py` in this folder. Some inputs are stated directly
(the elicited percentiles of the Chapter 2 running example and the seven-expert fairness panel
of Case I, for instance), and every random draw and every random solver start behind a reported
result comes from the seed 202603, or, where a config sheet records an offset of it, from a
generator seeded by the seed plus that offset (`fig6-8_bdt_case.xlsx` probe_seed_offset: seed 202604); a robustness sweep that regenerates a panel
records the further seeds it uses, and a workbook that involves no random draw says so in its
`seed` entry.
The workbooks are the authoritative numbers: every data figure is plotted from them, and
`gen_case_manifest.py` reads from them into `case_manifest.json` a summary of each study, the
numbers of the tables whose cells are workbook values, and the key numbers of the prose, each
with the words printed before it, which the delivery copy of the book compares with the LaTeX
sources (`tools/check_printed_numbers.py`).
This table is generated from the files by `tools/pack_index.py`.

## How to regenerate

```bash
python build_workbooks.py      # rewrites every workbook deterministically and runs its self-checks
python gen_case_manifest.py    # re-reads the workbooks and rewrites case_manifest.json
```

Each workbook carries a `config` or `provenance` sheet naming its generator, seed and the
`synthetic` flag, so the provenance travels with the data.

## Workbooks

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

## Other files

| File | Role |
|---|---|
| `build_workbooks.py` | the deterministic generator of every workbook; its rule-12 self-checks re-read the written workbooks, and most of them test invariants recomputed from the stored data (simplex, normalization, optimality, the identities and generator claims the book states), while the others compare structural sizes, a few printed bounds, and some orderings the text states with the values as the book prints them |
| `gen_case_manifest.py` | reads the workbooks and writes `case_manifest.json` |
| `case_manifest.json` | every number in it read from the workbooks: the numbers of the tables whose cells are workbook values (Table 2.1, Table 2.5, Table 5.3, Table 6.6, Table 6.7, Table 6.8, Table 6.9, Table 6.11, Table 6.12, Table 6.16, Table 6.17, Table 6.18, Table 6.19, Table 6.20) and the key numbers of the prose, each with the words printed before it, at full precision, and a readable summary of each study (those of Cases III and IV and of the Chapter 5 panel rounded as the book prints them) |
