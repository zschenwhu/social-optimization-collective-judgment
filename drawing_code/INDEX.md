# Drawing code index

Every file in this folder draws a data-driven figure of the book. This table is generated
from the sources by the release builder, so it cannot drift from them.

Rendered outputs go to `../figures_out/` (PNG and PDF); the copies exactly as printed in the
book are in `../reference_figures/`, and `../verify_figures.py` proves that every figure the
book prints is reproduced by the code here, with identical PDF content streams (only the
embedded creation date may differ).

## How to run

```bash
python run_all.py            # rebuild all workbooks in ../simulation_data, then render every figure
python fig4-4_pareto_front.py   # or any single figure
```

## Figures of the book, in order

| Book number | Figure file | Label in the book | Drawn by | Reads |
|---|---|---|---|---|
| Figure 2.1 | `fig2-1_linear_pool_vs_qa` | `fig:lp-vs-qa` | `fig2-1_linear_pool_vs_qa.py` | `fig2-1_linear_pool_vs_qa.xlsx` |
| Figure 2.2 | `fig2-2_three_pooling_operators` | `fig:lp-lop-qa` | `fig2-2_three_pooling_operators.py` | `fig2-2_three_pooling_operators.xlsx` |
| Figure 3.1 | `fig3-1_lorenz_gini_hoover` | `fig:lorenz` | `fig3-1_lorenz_gini_hoover.py` | `fig3-1_lorenz_gini_hoover.xlsx` |
| Figure 4.3 | `fig4-3_scalability` | `fig:scalability` | `fig4-3_scalability.py` | `fig4-3_scalability.xlsx` |
| Figure 4.4 | `fig4-4_pareto_front` | `fig:pareto` | `fig4-4_pareto_front.py` | `fig4-4_pareto_front.xlsx` |
| Figure 5.2 | `fig5-2_calibration_sharpness` | `fig:calsharp` | `fig5-2_calibration_sharpness.py` | `fig5-2_calibration_sharpness.xlsx` |
| Figure 5.3 | `fig5-3_loo` | `fig:loo-chart` | `fig5-3_loo.py` | `fig5-3_loo.xlsx` |
| Figure 5.4 | `fig5-4_depnet` | `fig:depnet` | `fig5-4_depnet.py` | `fig5-4_depnet.xlsx` |
| Figure 6.1 | `fig6-1_bim_groupagg` | `fig:bim-groupagg` | `fig6-1_bim_expansion.py` | `fig6-1_bim_case.xlsx` |
| Figure 6.2 | `fig6-2_bim_dimensions` | `fig:bim-dimensions` | `fig6-1_bim_expansion.py` | `fig6-1_bim_case.xlsx` |
| Figure 6.3 | `fig6-3_bim_densities` | `fig:bim-densities` | `fig6-3_bim_densities.py` | `fig6-3_bim_densities.xlsx` |
| Figure 6.4 | `fig6-4_bim_tradeoff` | `fig:bim-tradeoff` | `fig6-4_bim_tradeoff.py` | `fig6-4_bim_tradeoff.xlsx` |
| Figure 6.5 | `fig6-5_dt_ranking14` | `fig:dt-ranking14` | `fig6-5_dt_expansion.py` | `fig6-5_dt_case.xlsx` |
| Figure 6.6 | `fig6-6_dt_crossswf` | `fig:dt-crossswf` | `fig6-5_dt_expansion.py` | `fig6-5_dt_case.xlsx` |
| Figure 6.7 | `fig6-7_dt_triad` | `fig:dt-triad` | `fig6-5_dt_expansion.py` | `fig6-5_dt_case.xlsx` |
| Figure 6.8 | `fig6-8_bdt_q1agg` | `fig:bdt-q1agg` | `fig6-8_bdt_expansion.py` | `fig6-8_bdt_case.xlsx` |
| Figure 6.9 | `fig6-9_bdt_profile` | `fig:bdt-profile` | `fig6-8_bdt_expansion.py` | `fig6-8_bdt_case.xlsx` |
| Figure 6.10 | `fig6-10_bdt_sensitivity` | `fig:bdt-sensitivity` | `fig6-8_bdt_expansion.py` | `fig6-8_bdt_case.xlsx` |
| Figure 6.11 | `fig6-11_rt_equity` | `fig:rt-equity` | `fig6-11_rt_expansion.py` | `fig6-11_rt_case.xlsx` |
| Figure 6.12 | `fig6-12_hm_calsharp` | `fig:hm-calsharp` | `fig6-12_hm_calsharp.py` | `fig6-12_hm_calsharp.xlsx` |

## Shared modules

| File | Role |
|---|---|
| `_common.py` | palette, Times New Roman text and math, the shared density plotter, the save routine that writes PNG and PDF |
| `_scriptio.py` | reads each script's code (not its docstring) to list the workbooks it loads and the figures it saves, for the indexes and `../verify_figures.py` |
| `run_all.py` | runs `../simulation_data/build_workbooks.py` and `gen_case_manifest.py`, then every `fig*.py` |

15 Python figure scripts.
