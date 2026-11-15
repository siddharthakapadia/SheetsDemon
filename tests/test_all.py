"""End-to-end checks: every function runs, and the audits find exactly the
planted bugs without inventing extra ones."""

import json
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
# Use the interpreter actually running this suite. Hardcoding
# .venv/bin/python assumed one specific layout and broke anywhere the
# dependencies were installed differently — CI, conda, uv, pipx.
SD = [sys.executable, "-m", "sheetdoc.cli"]
sys.path.insert(0, str(ROOT))

PASS, FAIL = [], []


def run(args, expect_ok=True):
    r = subprocess.run(SD + args, capture_output=True, text=True, cwd=ROOT)
    if r.returncode == 2 or "Traceback" in r.stderr:
        FAIL.append(f"{' '.join(args[:3])}: crashed\n{r.stderr[-600:]}")
        return None
    return r.stdout


def codes(args):
    out = run(args + ["--json"])
    if out is None:
        return set()
    try:
        return {f["code"] for f in json.loads(out)["findings"]}
    except Exception as e:
        FAIL.append(f"{' '.join(args[:2])}: bad JSON ({e})")
        return set()


def check(name, cond, detail=""):
    (PASS if cond else FAIL).append(name + (f" — {detail}" if detail and not cond else ""))


def main():
    tmp = Path(tempfile.mkdtemp())
    sys.path.insert(0, str(ROOT / "tests"))
    import make_broken

    broken = str(make_broken.build())

    # ---- 1. auditing: exactly the planted bugs ---------------------------
    found = codes(["audit", broken])
    expected = {
        "CONSTANT_NOT_LOCKED",      # D: tax rate not anchored
        "HARDCODED_LITERAL",        # E: 0.08 in the formula
        "VLOOKUP_APPROX",           # F: no 4th argument
        "VLOOKUP_TABLE_RELATIVE",   # F: table not locked
        "XLOOKUP_LENGTH_MISMATCH",  # G: 3 vs 4 cells
        "INDEX_MATCH_OFFSET",       # H: ranges off by one row
        "RANGE_TOO_SHORT",          # C12
        "RANGE_TRAILING_BLANKS",    # E13
        "RANGE_INCLUDES_HEADER",    # E14
        "RANGE_INCLUDES_SELF",      # C17 circular
    }
    for code in expected:
        check(f"audit finds {code}", code in found, f"got {sorted(found)}")
    noise = found - expected - {"XLOOKUP_NO_FALLBACK"}
    check("audit produces no false positives", not noise, f"unexpected: {sorted(noise)}")

    # A correct sheet must stay silent. This guards the false-positive class that
    # made the tool unusable on real coursework: aggregates sitting beside their
    # source data ("Total", "Average") read as though the data continued, and a
    # row-wise sum over one deliberate sub-group read as an omission.
    clean = codes(["audit", broken, "-s", "Clean"])
    check("correct layout produces no findings at all", not clean,
          f"false positives on a clean sheet: {sorted(clean)}")
    out = run(["audit", broken, "-s", "Clean"]) or ""
    check("clean sheet reports zero errors", "0 error(s)" in out)

    # the benign literal 1 must never be flagged
    out = run(["hardcoded", broken]) or ""
    check("benign literal 1 not flagged", "1 is typed directly" not in out)

    # individual audit commands all work
    for cmd in ("hardcoded", "refs", "ranges", "lookups", "formats", "inspect"):
        check(f"`{cmd}` runs", run([cmd, broken]) is not None)

    # ---- 2. formatting ----------------------------------------------------
    fmt = codes(["formats", broken])
    check("formats flags unformatted money", "FORMAT_WRONG_FAMILY" in fmt, f"got {sorted(fmt)}")

    st = codes(["structure", broken])
    check("structure flags unbolded header", "HEADER_NOT_BOLD" in st, f"got {sorted(st)}")
    check("structure flags forced alignment", "ALIGNMENT" in st, f"got {sorted(st)}")

    cf_dry = run(["cf", broken, "C2:C11", "above", "--value", "40", "--dry-run"])
    check("cf dry-run reports hits", cf_dry is not None and "would highlight" in cf_dry)

    fixed = str(tmp / "fixed.xlsx")
    run(["structure", broken, "--fix", "--out", fixed, "--no-backup"])
    st2 = codes(["structure", fixed])
    check("structure --fix bolds the header", "HEADER_NOT_BOLD" not in st2, f"got {sorted(st2)}")

    cf_out = str(tmp / "cf.xlsx")
    run(["cf", broken, "D2:D11", "negative_red", "--out", cf_out, "--no-backup"])
    check("cf writes a rule", Path(cf_out).exists())
    import openpyxl

    check("cf rule persists in the file",
          bool(openpyxl.load_workbook(cf_out)["Homework"].conditional_formatting._cf_rules))

    # ---- 3. tutoring -------------------------------------------------------
    h = run(["hint", "--file", broken, "--cell", "F2"]) or ""
    check("hint diagnoses F2 from the sheet", "Tier 1" in h and "match" in h.lower())
    check("hint gives all three tiers", h.count("Tier") >= 3)

    h1 = run(["hint", "--file", broken, "--cell", "F2", "--tier", "1"]) or ""
    check("hint --tier 1 withholds tiers 2-3", "Tier 2" not in h1 and "Tier 1" in h1)

    hf = run(["hint", "--formula", "=VLOOKUP(A2,E2:G11,3)"]) or ""
    check("hint works on a pasted formula", "Tier 1" in hf)

    ht = run(["hint", "--topic", "absolute references"]) or ""
    check("hint works from a topic", "Tier 1" in ht)

    for code in ("#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NAME?", "#NUM!", "#SPILL!", "####", "circular"):
        e = run(["error", code]) or ""
        check(f"error {code} explained", "Likely cause 1" in e or "In plain English" in e)

    for topic in ("vlookup", "xlookup", "index_match", "absolute", "sumif",
                  "if_nested", "percentage", "payroll", "gradebook", "budget"):
        p = str(tmp / f"practice_{topic}.xlsx")
        run(["practice", topic, p, "--seed", "7"])
        ok = Path(p).exists()
        if ok:
            wb = openpyxl.load_workbook(p)
            ok = {"Instructions", "Data", "Answer Key"} <= set(wb.sheetnames)
            ok = ok and wb["Answer Key"].sheet_state == "hidden"
        check(f"practice {topic} builds a valid workbook", ok)

    # ---- 4. data and analysis ---------------------------------------------
    for kind in ("sales", "payroll", "grades", "inventory", "expenses", "budget"):
        m = str(tmp / f"mock_{kind}.xlsx")
        run(["mock", kind, m, "--rows", "12", "--seed", "3"])
        ok = Path(m).exists()
        if ok:
            wb = openpyxl.load_workbook(m)
            ws = wb.active
            ok = ws.max_row == 13 and ws.max_column >= 5
        check(f"mock {kind} writes 12 rows", ok)

    sales = str(tmp / "mock_sales.xlsx")
    piv = run(["pivot", sales]) or ""
    check("pivot maps Rows", "Rows ▸" in piv)
    check("pivot maps Values", "Values ▸" in piv)
    check("pivot gives click steps", "Step 1" in piv)

    ch = run(["chart", sales]) or ""
    check("chart picks a best choice", "BEST CHOICE" in ch)
    check("chart explains why", "because" in ch.lower() or "definition of" in ch.lower())

    grades = str(tmp / "mock_grades.xlsx")
    chg = run(["chart", grades]) or ""
    check("chart on grades recommends something", "BEST CHOICE" in chg)

    # a real time series should prefer a line chart
    check("chart picks Line for date+measure", "BEST CHOICE: Line chart" in ch,
          ch.split("BEST CHOICE")[1][:60] if "BEST CHOICE" in ch else "")

    # ---- 5. build and edit -------------------------------------------------
    spec = {
        "file": str(tmp / "built.xlsx"),
        "sheets": [{
            "name": "Budget",
            "table": {
                "at": "A3",
                "headers": ["Category", "Budget", "Actual", "Variance"],
                "rows": [["Rent", 1200, 1150, "=B4-C4"],
                         ["Utilities", 300, 342, "=B5-C5"],
                         ["Travel", 800, 615, "=B6-C6"]],
                "formats": [None, "currency", "currency", "currency"],
                "totals": [None, "SUM", "SUM", "SUM"],
            },
            "cells": {"A1": {"v": "Q1 Budget", "bold": True, "size": 14}},
            "conditional": [{"range": "D4:D6", "preset": "negative_red"}],
            "charts": [{"type": "column", "data": "B3:C6", "cats": "A4:A6",
                        "at": "F3", "title": "Budget vs Actual"}],
        }],
    }
    spec_file = tmp / "spec.json"
    spec_file.write_text(json.dumps(spec))
    run(["build", str(spec_file)])
    built = Path(spec["file"])
    check("build creates the workbook", built.exists())
    if built.exists():
        wb = openpyxl.load_workbook(built)
        ws = wb["Budget"]
        check("build writes the title", ws["A1"].value == "Q1 Budget")
        check("build writes formulas", str(ws["D4"].value).startswith("="))
        check("build adds a totals row", str(ws["B7"].value).upper().startswith("=SUM"))
        check("build applies currency format", "$" in ws["B4"].number_format)
        check("build adds the chart", len(ws._charts) == 1)
        check("build adds conditional formatting",
              bool(ws.conditional_formatting._cf_rules))

    edited = str(tmp / "edited.xlsx")
    run(["set", str(built), "E3=Note", 'E4="=B4*2"', "--out", edited, "--no-backup"])
    wb = openpyxl.load_workbook(edited)
    check("set writes a text value", wb["Budget"]["E3"].value == "Note")
    check("set writes a formula", wb["Budget"]["E4"].value == "=B4*2")

    filled = str(tmp / "filled.xlsx")
    run(["fill", edited, "E4", "E6", "--out", filled, "--no-backup"])
    wb = openpyxl.load_workbook(filled)
    check("fill translates rows", wb["Budget"]["E6"].value == "=B6*2",
          str(wb["Budget"]["E6"].value))

    # The table has a totals row at row 7, so the new row must land at row 7
    # (above the totals) and the totals must stretch to cover it.
    rowed = str(tmp / "rowed.xlsx")
    run(["addrow", filled, "Software", "450", "512", "_", "--out", rowed, "--no-backup"])
    wb = openpyxl.load_workbook(rowed)
    ws = wb["Budget"]
    check("addrow inserts above the totals row", ws["A7"].value == "Software", str(ws["A7"].value))
    check("addrow translates the formula", ws["D7"].value == "=B7-C7", str(ws["D7"].value))
    check("addrow copies number formats", "$" in ws["B7"].number_format)
    check("addrow extends the totals formula", ws["B8"].value == "=SUM(B4:B7)", str(ws["B8"].value))

    upd = str(tmp / "upd.xlsx")
    run(["updaterow", rowed, "--where", "Category=Travel", "--set", "Actual=999",
         "--out", upd, "--no-backup"])
    wb = openpyxl.load_workbook(upd)
    check("updaterow matches by column value", wb["Budget"]["C6"].value == 999,
          str(wb["Budget"]["C6"].value))

    upd2 = str(tmp / "upd2.xlsx")
    run(["updaterow", rowed, "--row", "1", "--set", "Budget=1500", "--out", upd2, "--no-backup"])
    wb = openpyxl.load_workbook(upd2)
    check("updaterow matches by index", wb["Budget"]["B4"].value == 1500,
          str(wb["Budget"]["B4"].value))

    # ---- audits must pass on a well-built workbook -------------------------
    clean = codes(["audit", str(built)])
    check("clean workbook audits clean", not (clean - {"XLOOKUP_NO_FALLBACK"}),
          f"got {sorted(clean)}")

    # ---- backup behaviour ---------------------------------------------------
    bak_src = tmp / "baktest.xlsx"
    openpyxl.load_workbook(built).save(bak_src)
    run(["set", str(bak_src), "A20=x"])
    check("in-place edit leaves a .bak", (tmp / "baktest.xlsx.bak").exists())

    # ---- report -------------------------------------------------------------
    print(f"\n{'=' * 70}")
    for f in FAIL:
        print(f"  FAIL  {f}")
    print(f"{'=' * 70}")
    print(f"  {len(PASS)} passed, {len(FAIL)} failed")
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
