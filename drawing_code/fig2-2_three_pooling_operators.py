#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Figure 2.2 (fig:lp-lop-qa): linear, logarithmic, and quantile pools of the
four-expert running example. Data: ../simulation_data/fig2-2_three_pooling_operators.xlsx
(the quantile-average and logarithmic-pool rows are computed by
build_workbooks.py from the expert rows; the linear pool is computed here)"""
from _common import density_figure
density_figure("fig2-2_three_pooling_operators.xlsx", "fig2-2_three_pooling_operators")
