#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
verify_figures.py -- the reproducibility gate of this repository.

Rebuilds every workbook and every figure from source in a throwaway copy of the
repository and checks, without modifying anything, that

  (1) every data-driven figure of the book (listed in reference_figures/manifest.json)
      has a generator here;
  (2) re-running that generator reproduces the PDF printed in the book
      (reference_figures/<name>.pdf) exactly, comparing the decompressed PDF
      content streams so that embedded creation timestamps are ignored (the PDF
      bytes themselves differ by those timestamps);
  (3) every workbook rebuilt by simulation_data/build_workbooks.py, including the
      ones that back tables only, has the same content in every sheet as the shipped
      one (the gate also reports which workbooks are byte-identical; on the software
      stack that built the shipped copy all of them are), and the rebuilt
      case_manifest.json equals the shipped one;
  (4) no generator, workbook, or rendered figure in this repository is an orphan,
      what a script reads and draws being taken from its code
      (drawing_code/_scriptio.py), not from its documentation.

The figures are rebuilt with drawing_code/run_all.py, which first rebuilds every
workbook with build_workbooks.py and its self-checks.

    python verify_figures.py

Exit status is 0 only when every check passes.

The comparisons demand identity, and identity is to be expected only on the software
stack that built the shipped copy (environment.py, requirements.txt).  The gate prints
the running stack first.  On another stack the multi-start solves of the nonconvex
programs follow other round-off, so weight vectors of programs with a non-unique
optimum, counts of solver starts, and the placement of figure text can differ while
the collective results agree to rounding.  For a workbook that differs the gate
therefore also prints the largest numerical difference.
"""
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zlib

HERE = os.path.dirname(os.path.abspath(__file__))
REF = os.path.join(HERE, "reference_figures")
DATA = os.path.join(HERE, "simulation_data")
OUT = os.path.join(HERE, "figures_out")
CODE = os.path.join(HERE, "drawing_code")

sys.path.insert(0, CODE)
sys.path.insert(0, HERE)
sys.dont_write_bytecode = True   # the gate leaves no __pycache__ in the repository
from _scriptio import pack_io   # noqa: E402
import environment   # noqa: E402  (the reference software stack and the check of the running one)


def pdf_streams_digest(path):
    """MD5 over the decompressed content streams -- ignores /CreationDate."""
    raw = open(path, "rb").read()
    h = hashlib.md5()
    i = 0
    while True:
        a = raw.find(b"stream", i)
        if a < 0:
            break
        a += 6
        if raw[a:a + 2] == b"\r\n":
            a += 2
        elif raw[a:a + 1] in (b"\n", b"\r"):
            a += 1
        b = raw.find(b"endstream", a)
        if b < 0:
            break
        chunk = raw[a:b]
        try:
            chunk = zlib.decompress(chunk)
        except zlib.error:
            pass
        h.update(chunk)
        i = b + 9
    return h.hexdigest()


def sheet_digests(path):
    """{sheet: SHA-256 of its CSV rendering} -- the content of a workbook."""
    import pandas as pd
    return {name: hashlib.sha256(df.to_csv(index=False).encode("utf-8")).hexdigest()
            for name, df in pd.read_excel(path, sheet_name=None).items()}


def compare_data(rebuilt, shipped, problems):
    """Every rebuilt workbook and the manifest against the shipped copies."""
    new = sorted(f for f in os.listdir(rebuilt) if f.endswith(".xlsx"))
    old = sorted(f for f in os.listdir(shipped) if f.endswith(".xlsx"))
    for wb in sorted(set(old) - set(new)):
        problems.append("shipped workbook %s is not produced by build_workbooks.py" % wb)
        print("   %-44s SHIPPED BUT NOT REBUILT" % wb)
    n_bytes = 0
    for wb in new:
        a, b = os.path.join(rebuilt, wb), os.path.join(shipped, wb)
        if not os.path.exists(b):
            problems.append("rebuilt workbook %s is not shipped in simulation_data" % wb)
            print("   %-44s NOT SHIPPED" % wb)
            continue
        if open(a, "rb").read() == open(b, "rb").read():
            n_bytes += 1
            print("   %-44s identical bytes" % wb)
            continue
        da, db = sheet_digests(a), sheet_digests(b)
        if da == db:
            print("   %-44s same content in every sheet (container bytes differ)" % wb)
        else:
            diff = sorted(s for s in set(da) | set(db) if da.get(s) != db.get(s))
            problems.append("rebuilt workbook %s differs from the shipped one in sheet(s) %s" % (wb, ", ".join(diff)))
            print("   %-44s CONTENT DIFFERS in %s" % (wb, ", ".join(diff)))
            worst, where, n_num, n_txt, shapes = environment.workbook_gap(a, b)
            print("   %-44s   %d numeric cells differ, the largest by %.2e (sheet %s); %d text cells differ%s" % (
                "", n_num, worst, where or "-", n_txt,
                "; sheets of another shape: " + ", ".join(shapes) if shapes else ""))
    ma, mb = os.path.join(rebuilt, "case_manifest.json"), os.path.join(shipped, "case_manifest.json")
    if not (os.path.exists(ma) and os.path.exists(mb)):
        problems.append("case_manifest.json missing from the %s tree" % ("rebuilt" if not os.path.exists(ma) else "shipped"))
    elif json.load(open(ma, encoding="utf-8")) != json.load(open(mb, encoding="utf-8")):
        ja, jb = json.load(open(ma, encoding="utf-8")), json.load(open(mb, encoding="utf-8"))
        keys = sorted(k for k in set(ja) | set(jb) if ja.get(k) != jb.get(k))
        problems.append("rebuilt case_manifest.json differs from the shipped one in %s" % ", ".join(keys))
        print("   %-44s DIFFERS in %s" % ("case_manifest.json", ", ".join(keys)))
    else:
        print("   %-44s identical" % "case_manifest.json")
    print("   (%d of %d workbooks byte-identical)" % (n_bytes, len(new)))


def run(cmd, cwd, env=None):
    # the rebuild runs in a throwaway copy, so the generator may write there on any stack
    return subprocess.run(cmd, cwd=cwd, capture_output=True, text=True,
                          encoding="utf-8", errors="replace",
                          env=env or dict(os.environ, MPLBACKEND="Agg",
                                          PYTHONIOENCODING="utf-8", **{environment.ALLOW: "1"}))


def soft_note(stdout):
    """The generator's NOTE on self-checks that did not hold on a stack other than the
    reference one (empty on the reference stack, where they are enforced)."""
    lines, keep = [], False
    for l in stdout.splitlines():
        if l.startswith("NOTE  ") and "self-check" in l:
            keep = True
        elif keep and not l.startswith("      "):
            keep = False
        if keep:
            lines.append(l)
    return lines


def main():
    problems = []
    environment.require_packages()                 # a plain message when a package is not installed
    other_stack = environment.report(force=True)   # prints the running stack; [] on the reference stack
    manifest = json.load(open(os.path.join(REF, "manifest.json"), encoding="utf-8"))
    figures = {m["file"]: m for m in manifest["figures"]}
    print("the book prints %d data-driven figures (reference_figures/manifest.json)" % len(figures))

    # --- 1. rebuild the workbooks and the data figures into a throwaway tree ---
    tmp = tempfile.mkdtemp(prefix="figverify_")
    try:
        sandbox = os.path.join(tmp, "repo")
        shutil.copytree(HERE, sandbox, ignore=shutil.ignore_patterns(
            "figures_out", "svg", "reference_figures", ".git", "__pycache__", "*.pyc"))
        # drop the shipped workbooks and manifest so the generator must rebuild every one
        for wb in os.listdir(os.path.join(sandbox, "simulation_data")):
            if wb.endswith(".xlsx") or wb == "case_manifest.json":
                os.remove(os.path.join(sandbox, "simulation_data", wb))
        os.makedirs(os.path.join(sandbox, "figures_out"), exist_ok=True)
        print("\nrebuilding workbooks and figures from scratch ...")
        r = run([sys.executable, "run_all.py"], os.path.join(sandbox, "drawing_code"))
        rebuilt_ok = r.returncode == 0
        if not rebuilt_ok:
            print(r.stdout[-3000:])
            print(r.stderr[-3000:])
            problems.append("run_all.py failed with exit %d" % r.returncode)
            fresh = OUT
        else:
            fresh = os.path.join(sandbox, "figures_out")
            print("   ok")
            for l in soft_note(r.stdout):
                print("   " + l)

        generated = sorted(f[:-4] for f in os.listdir(fresh) if f.endswith(".pdf"))

        # --- 2. every printed figure must have a generator -----------------------
        print("\n-- printed figure -> generator --")
        for fig in sorted(figures, key=lambda f: figures[f]["order"]):
            if fig in generated:
                kind = "python"
            else:
                kind = "MISSING GENERATOR"
                problems.append("no generator for printed figure %s" % fig)
            if not os.path.exists(os.path.join(REF, fig + ".pdf")):
                problems.append("reference PDF missing for %s" % fig)
            print("   %-10s %-32s %s" % (figures[fig]["number"], fig, kind))

        # --- 3. reproduction: regenerated == printed ----------------------------
        print("\n-- reproduction check (%d python figures, PDF content streams) --" % len(generated))
        for fig in generated:
            ref = os.path.join(REF, fig + ".pdf")
            if fig not in figures:
                print("   %-32s not printed in the book (orphan output)" % fig)
                problems.append("generated figure %s is not printed in the book" % fig)
                continue
            same = pdf_streams_digest(ref) == pdf_streams_digest(os.path.join(fresh, fig + ".pdf"))
            print("   %-32s %s" % (fig, "reproduces exactly" if same else "DIFFERS"))
            if not same:
                problems.append("regenerated %s does not match the printed figure" % fig)

        # --- 4. workbooks: rebuilt == shipped ------------------------------------
        print("\n-- rebuilt workbooks and manifest == shipped --")
        if rebuilt_ok:
            compare_data(os.path.join(sandbox, "simulation_data"), DATA, problems)
        else:
            print("   (skipped: the rebuild failed)")

        # --- 5. no orphans, from what the scripts' code loads --------------------
        print("\n-- workbook usage --")
        io_map = pack_io(CODE)
        read = set(sum([r for r, _ in io_map.values()], []))
        for wb in sorted(f for f in os.listdir(DATA) if f.endswith(".xlsx")):
            used = wb in read
            role = "read by a figure script" if used else "table-only (no plot)"
            if not used and not wb.startswith("tab"):
                problems.append("workbook %s is neither read by a script nor a table" % wb)
                role = "ORPHAN"
            print("   %-44s %s" % (wb, role))
        for s, (r, _) in io_map.items():
            for wb in r:
                if not os.path.exists(os.path.join(DATA, wb)):
                    problems.append("%s reads %s, which the repository does not ship" % (s, wb))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    print("\n" + "=" * 66)
    if problems:
        print("FAILED -- %d problem(s):" % len(problems))
        for p in problems:
            print("  *", p)
        if other_stack:
            environment.explain(other_stack)
        else:
            environment.explain_reference()
        return 1
    print("PASSED -- every data-driven figure of the book has a generator here and reproduces exactly,"
          " and every rebuilt workbook matches the shipped one.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
