# -*- coding: utf-8 -*-
"""What a figure script reads and writes, taken from its code rather than its text.

    reads, writes = script_io("fig5-3_loo.py")

reads   the workbooks passed to load(...) or density_figure(...), either as a string
        literal or through a name bound to one (WB = "fig5-3_loo.xlsx"; load(WB))
writes  the figure names passed to save(fig, ...) and to density_figure(..., name)

Docstrings and comments are never consulted, so a script that mentions another
workbook or figure in its documentation is not credited with reading or drawing it.
The indexes (tools/pack_index.py, tools/build_release.py) and the reproducibility
gates (verify_figures.py) share this one definition.
"""
import ast
import os


def _name(func):
    if isinstance(func, ast.Name):
        return func.id
    if isinstance(func, ast.Attribute):
        return func.attr
    return None


def _literal(node, bound):
    if isinstance(node, ast.Constant) and isinstance(node.value, str):
        return node.value
    if isinstance(node, ast.Name):
        return bound.get(node.id)
    return None


def script_io(path):
    """(sorted workbooks read, sorted figure names written) of one drawing script."""
    with open(path, encoding="utf-8") as fh:
        tree = ast.parse(fh.read(), filename=path)
    bound = {}
    for node in ast.walk(tree):
        if (isinstance(node, ast.Assign) and len(node.targets) == 1
                and isinstance(node.targets[0], ast.Name)
                and isinstance(node.value, ast.Constant) and isinstance(node.value.value, str)):
            bound[node.targets[0].id] = node.value.value
    reads, writes = set(), set()
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        fn = _name(node.func)
        if fn in ("load", "density_figure") and node.args:
            wb = _literal(node.args[0], bound)
            if wb and wb.endswith(".xlsx"):
                reads.add(os.path.basename(wb))
        if fn in ("save", "density_figure") and len(node.args) >= 2:
            out = _literal(node.args[1], bound)
            if out:
                writes.add(out)
    return sorted(reads), sorted(writes)


def pack_io(code_dir):
    """{script file name: (reads, writes)} for every fig*.py in code_dir."""
    return {f: script_io(os.path.join(code_dir, f))
            for f in sorted(os.listdir(code_dir)) if f.startswith("fig") and f.endswith(".py")}
