#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""reproduce.py -- one command to regenerate and verify everything.

    python reproduce.py            # rebuild every workbook and figure, then verify

On the reference software stack (environment.py, requirements.txt) the workbooks and
figures are rebuilt in place, identically, and then verified.  On another stack the
shipped files are left untouched (unless SOCJ_ALLOW_OTHER_STACK=1 is set), and only the
verification runs, which rebuilds everything in a temporary copy and reports how far
the result is from the shipped one.
"""
import os, subprocess, sys
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import environment
environment.require_packages()     # a plain message when a package is not installed
env = dict(os.environ, MPLBACKEND="Agg", PYTHONIOENCODING="utf-8")
if environment.mismatches() and not environment.allowed():
    environment.report(force=True)
    print("\nThe shipped workbooks and figures are left untouched; the rebuild and the comparison"
          " below run in a temporary copy.\n")
else:
    r = subprocess.run([sys.executable, "run_all.py"], cwd=os.path.join(HERE, "drawing_code"), env=env)
    if r.returncode != 0:
        sys.exit("run_all.py failed")
sys.exit(subprocess.run([sys.executable, "verify_figures.py"] + sys.argv[1:], cwd=HERE, env=env).returncode)
