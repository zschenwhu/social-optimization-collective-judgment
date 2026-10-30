#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""run_all.py -- rebuild every data workbook, then regenerate every figure.

    python run_all.py

Rebuilds ../simulation_data/*.xlsx and case_manifest.json, then writes
../figures_out/*.png (+ .pdf) and ../svg/*.svg.

The outputs equal the shipped ones exactly on the software stack named in
../environment.py (the pins of ../requirements.txt); the generator prints a notice
when another stack is running.
"""
import glob
import os
import runpy
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SIM = os.path.normpath(os.path.join(HERE, "..", "simulation_data"))
sys.path.insert(0, os.path.normpath(os.path.join(HERE, "..")))
try:
    import environment
except ImportError:
    environment = None

print("== building data workbooks ==")
try:
    runpy.run_path(os.path.join(SIM, "build_workbooks.py"), run_name="__main__")
except SystemExit as stop:
    # On a stack other than the reference one the generator either refuses to write (a
    # message, passed on here) or, with SOCJ_ALLOW_OTHER_STACK=1, builds with reported
    # self-checks and then ends with status 0, after which the figures are rendered.
    if stop.code not in (None, 0):
        raise
runpy.run_path(os.path.join(SIM, "gen_case_manifest.py"), run_name="__main__")

print("\n== rendering figures ==")
for f in sorted(glob.glob(os.path.join(HERE, "fig*.py"))):
    print("--", os.path.basename(f))
    runpy.run_path(f, run_name="__main__")

print("\nDone. See ../figures_out/")
if environment is not None:
    environment.reminder()
