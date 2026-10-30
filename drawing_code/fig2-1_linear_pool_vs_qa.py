#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Figure 2.1 (fig:lp-vs-qa): linear pool vs quantile averaging of two
normal experts. Data source: ../simulation_data/fig2-1_linear_pool_vs_qa.xlsx
(the quantile-average row of that workbook is computed by build_workbooks.py
from the expert rows; the linear pool is computed here from the same rows)"""
from _common import density_figure
density_figure("fig2-1_linear_pool_vs_qa.xlsx", "fig2-1_linear_pool_vs_qa")
