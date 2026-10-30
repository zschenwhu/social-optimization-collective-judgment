#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""environment.py -- the software stack of the shipped data and figures.

    python environment.py        # compare the stack running now with the reference stack

The shipped workbooks (simulation_data/*.xlsx), the manifest, and the figures were built with
exactly the versions in REFERENCE, which are also the pins of requirements.txt, on the
platform named in REFERENCE_PLATFORM.

On the reference stack a rebuild returns every workbook byte for byte and every figure with
the same PDF content streams, so `verify_figures.py` passes.

The versions alone do not fix the arithmetic. NumPy and SciPy each bundle an OpenBLAS library
that chooses its kernels from the processor it detects and splits larger matrix products over
as many threads as the machine has logical processors (24 at most). Two kernels round
differently, and so do two thread counts, because the split changes the order of the sums.
So the pack also fixes both: before numpy is imported, pin_numerics() requests the kernels
that the reference machine runs (KERNELS) and switches off the AVX-512 code paths of NumPy,
which the reference processor does not have, and after the import it sets the number of
OpenBLAS threads to that of the reference machine (THREADS). The thread count is set through
the OpenBLAS interface with the package threadpoolctl, which also works on a machine with
fewer processors, where the environment variable alone would be capped. With the reference
versions on an x86-64 processor with AVX2 the pack therefore does the same arithmetic as
the reference machine.

On any other stack the rebuild is still deterministic, and every table of the book and the
collective distributions, rankings, and scores behind it agree with the shipped ones to
rounding. The following can differ, because the multi-start solves of the nonconvex programs
are sensitive to round-off in the numerical libraries:

  * the weight vector recorded for a program whose optimum is not unique (other weights, the
    same aggregate),
  * the number of solver starts that reach a retained value, which start is retained, and the
    small gains of the further starts and of the local polish (solver diagnostics),
  * the placement of text in the figures, by about one point (font metrics of matplotlib).

Because of this the generator protects the shipped workbooks:

  * on the reference stack it rebuilds them in place, and the result is identical;
  * on another stack it stops before writing anything and says why. Set the environment
    variable SOCJ_ALLOW_OTHER_STACK=1 to build anyway. The self-checks that compare a solver
    diagnostic with the value printed in the book are then reported, not enforced, and the
    workbooks written are those of that stack, not the shipped ones. `verify_figures.py`
    does this in a temporary copy, so it never touches the shipped files.

For an exact rebuild install the reference stack:

    pip install -r requirements.txt

The pins install on Python 3.9 to 3.12. The rebuild has been verified on the reference
platform only. Other operating systems, processors without AVX2, and processors that are not
x86-64 are untested, and a difference there is to be read as round-off, not as an error in
the shipped data.

Environment variables read here
    SOCJ_ALLOW_OTHER_STACK=1   let the generator write on a stack other than the reference one
    SOCJ_NO_PIN=1              leave the kernels and the thread count to the machine (for experiments)
"""
import io
import os
import sys

# name -> version that built the shipped workbooks and figures
REFERENCE = [
    ("python", "3.9"),
    ("numpy", "1.26.4"),
    ("scipy", "1.13.1"),
    ("pandas", "2.3.3"),
    ("openpyxl", "3.1.5"),
    ("matplotlib", "3.8.4"),
]
# The OpenBLAS kernels the reference machine runs, as threadpoolctl names them.  NumPy's
# OpenBLAS (0.3.23) does not recognize the reference processor and falls back to its generic
# Prescott kernels; SciPy's (0.3.27) recognizes it and runs its Haswell kernels.
KERNELS = [("numpy", "Prescott"), ("scipy", "Haswell")]
REFERENCE_PLATFORM = ("CPython 3.9.18 on Windows 11 (10.0.26200), Intel Core i9-13950HX (x86-64 with AVX2, without AVX-512), "
                      "NumPy on its bundled OpenBLAS 0.3.23 (Prescott kernels, 24 threads), SciPy on its bundled OpenBLAS "
                      "0.3.27 (Haswell kernels, 24 threads), Times New Roman installed")

ALLOW = "SOCJ_ALLOW_OTHER_STACK"      # set to 1 to let the generator write on a stack other than the reference one
NO_PIN = "SOCJ_NO_PIN"                # set to 1 to leave the numerical kernels to the machine
KERNEL_VAR = "OPENBLAS_CORETYPE"      # read by OpenBLAS when it is loaded
THREAD_VAR = "OPENBLAS_NUM_THREADS"   # read by OpenBLAS when it is loaded (capped at the number of processors)
THREADS = 24                          # OpenBLAS threads on the reference machine; one thread rounds differently
SIMD_VAR = "NPY_DISABLE_CPU_FEATURES" # read by NumPy when it is imported
SIMD_OFF = "AVX512F AVX512CD AVX512_SKX AVX512_CLX AVX512_CNL AVX512_ICL"
_FLAG = "SOCJ_STACK_REPORTED"         # set once the stack has been printed (inherited by child processes)
_STATE = {"written": False, "failed": [], "pin": None, "kernels": None, "threads": None}


def _installed(name):
    """The installed version of a package, read from its metadata without importing it."""
    try:
        from importlib import metadata
        return metadata.version(name)
    except Exception:
        return None


def _pin():
    import platform
    if os.environ.get(NO_PIN) == "1":
        return "left to this machine (%s=1)" % NO_PIN
    if platform.machine().lower() not in ("amd64", "x86_64"):
        return "left to this machine (not an x86-64 processor)"
    ref = dict(REFERENCE)
    if _installed("numpy") != ref["numpy"] or _installed("scipy") != ref["scipy"]:
        return "left to this machine (numpy and scipy are not the reference versions)"
    if "numpy" in sys.modules or "scipy" in sys.modules:
        return "left to this machine (numpy was imported before the kernels could be requested)"
    saved = dict((v, os.environ.get(v)) for v in (KERNEL_VAR, THREAD_VAR, SIMD_VAR))
    asked = []
    try:
        os.environ[SIMD_VAR] = SIMD_OFF
        os.environ[THREAD_VAR] = str(THREADS)
        for package, core in KERNELS:
            if core == "Haswell" and not _cpu_has(("AVX2", "FMA3")):
                # the Haswell kernels need AVX2 and FMA; without them OpenBLAS keeps its own choice
                os.environ.pop(KERNEL_VAR, None)
            else:
                os.environ[KERNEL_VAR] = core
                asked.append("%s %s" % (package, core))
            try:
                __import__(package + ".linalg")    # loads the OpenBLAS of the package, which reads the variable
            except ImportError:
                pass
    finally:
        for v, old in saved.items():
            if old is None:
                os.environ.pop(v, None)
            else:
                os.environ[v] = old
    # the thread count, through the OpenBLAS interface: unlike the environment variable it is
    # not capped at the number of processors of this machine
    try:
        from threadpoolctl import threadpool_limits
        threadpool_limits(limits=THREADS, user_api="blas")
        how = "%d OpenBLAS threads" % THREADS
    except Exception:
        how = "%d OpenBLAS threads asked through the environment only, since threadpoolctl is not installed" % THREADS
    if len(asked) < len(KERNELS):
        return "requested in part (%s, %s; this processor lacks AVX2 or FMA)" % (", ".join(asked), how)
    return "requested as on the reference platform (%s, %s)" % (", ".join(asked), how)


def _cpu_has(features):
    """True when the processor has every named feature, as NumPy (already imported) detects them."""
    try:
        from numpy.core._multiarray_umath import __cpu_features__ as f
        return all(bool(f.get(x)) for x in features)
    except Exception:
        return False


def pin_numerics():
    """Request the numerical kernels of the reference platform, once per process.  It has to
    run before numpy or scipy is imported and imports both itself; every function below that
    imports them calls it first.  Returns a phrase saying what was done."""
    if _STATE["pin"] is None:
        _STATE["pin"] = _pin()
    return _STATE["pin"]


def kernels():
    """{"numpy": kernel, "scipy": kernel} of the OpenBLAS libraries in use, read with
    threadpoolctl; {} when that package is not installed (it is not required)."""
    if _STATE["kernels"] is None:
        out, threads = {}, {}
        try:
            from threadpoolctl import threadpool_info
            for d in threadpool_info():
                if d.get("internal_api") != "openblas":
                    continue
                path = (d.get("filepath") or "").replace("\\", "/").lower()
                for package, _ in KERNELS:
                    if "/%s.libs/" % package in path or "/%s/" % package in path:
                        out[package] = d.get("architecture")
                        threads[package] = d.get("num_threads")
        except Exception:
            out, threads = {}, {}
        _STATE["kernels"], _STATE["threads"] = out, threads
    return _STATE["kernels"]


def in_use():
    """'; in use: numpy Prescott (1 thread), scipy Haswell (1 thread)', or '' when threadpoolctl
    is not installed."""
    k = kernels()
    if not k:
        return ""
    t = _STATE.get("threads") or {}
    return "; in use: " + ", ".join("%s %s (%s thread%s)" % (n, k[n], t.get(n), "" if t.get(n) == 1 else "s")
                                    for n, _ in KERNELS if n in k)


def current():
    """{name: running version, or None when the package cannot be imported}."""
    pin_numerics()
    out = {"python": "%d.%d" % sys.version_info[:2]}
    for name, _ in REFERENCE[1:]:
        try:
            out[name] = __import__(name).__version__
        except Exception:
            out[name] = None
    return out


def mismatches():
    """[(name, running, reference)] for every component that is not the reference one: the
    versions and, when the versions match and threadpoolctl can read them, the kernels and
    the thread counts."""
    cur = current()
    bad = [(n, cur[n], ref) for n, ref in REFERENCE if cur[n] != ref]
    if not bad:
        try:
            import threadpoolctl  # noqa: F401  (sets and reads back the OpenBLAS thread count)
        except Exception:
            return [("threadpoolctl", None, "any version")]
        k = kernels()
        t = _STATE.get("threads") or {}
        bad = [("%s kernel" % n, k[n], core) for n, core in KERNELS if k.get(n) and k[n] != core]
        bad += [("%s threads" % n, t[n], THREADS) for n, _ in KERNELS if t.get(n) and t[n] != THREADS]
    return bad


def is_reference():
    return not mismatches()


def allowed():
    """True when the user asked the generator to write on a stack other than the reference one."""
    return os.environ.get(ALLOW) == "1"


NOT_INSTALLED = """
The Python that is running lacks packages this pack needs.

    interpreter   %s
    not installed %s

Install the versions that built the shipped data and figures (requirements.txt is in the
folder of environment.py):
    pip install -r requirements.txt
or run with an interpreter that already has them. In an IDE, select that interpreter for
the project (a newly created virtual environment starts empty).
"""


def missing():
    """Names of the packages of the pack that the running Python cannot import."""
    cur = current()
    return [n for n, _ in REFERENCE[1:] if cur[n] is None]


def require_packages():
    """Stop with a plain message, instead of a traceback, when a package is not installed."""
    miss = missing()
    if miss:
        raise SystemExit(NOT_INSTALLED % (sys.executable, ", ".join(miss)))


def report(force=False):
    """Print the stack once per process tree and return the list of mismatches."""
    bad = mismatches()
    if os.environ.get(_FLAG) and not force:
        return bad
    os.environ[_FLAG] = "1"
    cur = current()
    print("software stack: " + ", ".join("%s %s" % (n, cur[n] or "not installed") for n, _ in REFERENCE))
    print("   numerical kernels: " + pin_numerics() + in_use())
    if not bad:
        print("   this is the reference stack of the shipped workbooks and figures; a rebuild reproduces them exactly")
        return bad
    print("   NOTE  this is NOT the reference stack of the shipped workbooks and figures:")
    for n, have, ref in bad:
        print("         %-13s running %-14s reference %s" % (n, have if have is not None else "not installed", ref))
    print("   Results agree with the shipped ones to rounding on such a stack, but weight vectors of programs")
    print("   with a non-unique optimum, solver diagnostics (counts of starts, small gains of further starts),")
    print("   and the placement of text in the figures can differ, so an exact comparison reports differences.")
    if all(n.endswith((" kernel", " threads")) for n, _, _ in bad):
        print("   The versions are the reference ones; the numerical kernels or thread counts are not (see environment.py).")
    else:
        print("   For an exact rebuild:  pip install -r requirements.txt   (see environment.py)")
    return bad


REFUSAL = """
The generator has stopped before writing anything, to protect the shipped workbooks.

The software stack running now is not the one that built them (see the NOTE above). A rebuild
here would replace the shipped workbooks with ones that differ in non-unique weight vectors
and in solver diagnostics, and they would no longer match the numbers printed in the book.

Do one of the following.
  1. Install the reference stack and run again (the rebuild is then identical):
         pip install -r requirements.txt
  2. Compare without touching the shipped files (rebuilds in a temporary copy):
         python verify_figures.py
  3. Build on this stack anyway, in a copy you are free to overwrite:
         set SOCJ_ALLOW_OTHER_STACK=1        (Windows cmd;  $env:SOCJ_ALLOW_OTHER_STACK=1 in PowerShell;
                                              export SOCJ_ALLOW_OTHER_STACK=1 on Linux or macOS)
"""


def workbooks_written():
    """Called by the generator once every workbook is written; from here on its statements
    are self-checks."""
    _STATE["written"] = True


def soft_failures():
    """[(line, message)] of the self-checks that did not hold in a build on another stack."""
    return list(_STATE["failed"])


def run_softly(path):
    """Execute the generator statement by statement in one namespace, exactly as Python would
    run the file, except that an assertion failing after the workbooks are written is
    recorded and the next statement runs.  Used only on a stack other than the reference
    one: the self-checks include comparisons of solver diagnostics with the values printed
    in the book, which hold on the reference stack and need not hold elsewhere."""
    import ast
    src = io.open(path, encoding="utf-8").read()
    tree = ast.parse(src, path)
    g = {"__name__": "__main__", "__file__": path, "_SOCJ_SOFT_RUN": True}
    for node in tree.body:
        code = compile(ast.Module(body=[node], type_ignores=[]), path, "exec")
        try:
            exec(code, g)
        except AssertionError as e:
            if not _STATE["written"]:
                raise                      # a failure while building is a real failure
            # the message of the assertion, or the statement itself when it carries none
            text = str(e) or " ".join((ast.get_source_segment(src, node) or "").split())
            _STATE["failed"].append((node.lineno, text or "(assertion without a message)"))
    failed = soft_failures()
    print("")
    if failed:
        print("NOTE  %d self-check(s) did not hold on this software stack (generator line, message):" % len(failed))
        for line, msg in failed:
            print("        line %5d  %s" % (line, msg[:200]))
        print("      They hold on the reference stack.  Each compares a quantity that depends on round-off in the")
        print("      solver (a count of starts, a small gain, a non-unique weight) with the value printed in the book.")
    else:
        print("NOTE  every self-check held on this software stack as well.")
    print("      The workbooks just written are those of this stack, not the shipped ones.")
    return failed


def other_stack(path):
    """Called by the generator when the running stack is not the reference one.  Either stops
    the build (the default) or runs the generator with reported, not enforced, self-checks.
    Never returns: the caller must not continue, since the file has been executed here."""
    if not allowed():
        raise SystemExit(REFUSAL)
    print("   %s=1 is set: building on this stack; self-checks tied to printed solver diagnostics are reported, not enforced" % ALLOW)
    run_softly(path)
    raise SystemExit(0)


def reminder():
    """One closing line after a build on a stack other than the reference one."""
    if mismatches():
        print("   (built on a stack other than the reference one; see the NOTE at the top of this output)")


def explain(bad):
    """The paragraph a gate prints under its list of problems when the stack is not the reference one."""
    print("\nNOTE  This run used a software stack other than the reference one (%s)." % ", ".join(
        "%s %s instead of %s" % (n, have or "not installed", ref) for n, have, ref in bad))
    print("      Differences in non-unique weight vectors, in solver diagnostics, and in the placement of figure")
    print("      text are to be expected on such a stack and do not indicate that the shipped data are wrong.")
    print("      For an exact rebuild install the reference stack (pip install -r requirements.txt) and run")
    print("      the gate again.")


def explain_reference():
    """The paragraph a gate prints under its list of problems when the versions ARE the
    reference ones and the result still differs."""
    print("\nNOTE  The versions are the reference ones, yet the rebuild differs from the shipped files.")
    print("      Numerical kernels: %s%s." % (pin_numerics(), in_use()))
    if not kernels():
        print("      The kernels and thread counts in use could not be read back, because threadpoolctl is not")
        print("      installed. Without it the thread count is capped at the number of processors of this machine,")
        print("      and a count other than %d rounds differently (pip install -r requirements.txt)." % THREADS)
    print("      The rebuild is verified on the reference platform only (%s)." % REFERENCE_PLATFORM)
    print("      On another operating system or processor the remaining differences are round-off in the")
    print("      system libraries: the tables and aggregates agree to rounding, and non-unique weight vectors")
    print("      and solver diagnostics can differ. Each differing workbook is listed above with its largest")
    print("      numerical difference.")


def workbook_gap(a, b):
    """How far two workbooks with different content are apart: (largest absolute difference
    over the numeric cells, the sheet it is in, number of numeric cells that differ, number of
    text cells that differ, sheets missing from one file or of another shape)."""
    import math
    import numbers
    pin_numerics()
    import pandas as pd
    A = pd.read_excel(a, sheet_name=None, header=None)
    B = pd.read_excel(b, sheet_name=None, header=None)
    worst, where, n_num, n_txt, shapes = 0.0, "", 0, 0, []
    for s in sorted(set(A) | set(B)):
        if s not in A or s not in B or A[s].shape != B[s].shape:
            shapes.append(s)
            continue
        x, y = A[s].to_numpy(dtype=object), B[s].to_numpy(dtype=object)
        for i in range(x.shape[0]):
            for j in range(x.shape[1]):
                u, v = x[i, j], y[i, j]
                nu = isinstance(u, numbers.Real) and not isinstance(u, bool)
                nv = isinstance(v, numbers.Real) and not isinstance(v, bool)
                if nu and nv:
                    u, v = float(u), float(v)
                    if math.isnan(u) and math.isnan(v):
                        continue
                    d = abs(u - v)
                    if d > 0 or math.isnan(d):
                        n_num += 1
                        if d > worst:
                            worst, where = d, s
                elif u != v and not (u != u and v != v):
                    n_txt += 1
    return worst, where, n_num, n_txt, shapes


if __name__ == "__main__":
    print("reference platform: " + REFERENCE_PLATFORM)
    sys.exit(1 if report(force=True) else 0)
