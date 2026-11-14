"""Build a workbook with known, deliberately planted mistakes.

Each bug is labelled with the finding code the audit engine should produce, so
the test harness can assert that we catch exactly these and not much else.
"""

import openpyxl
from openpyxl.styles import Alignment, Font
from pathlib import Path

OUT = Path(__file__).parent / "broken.xlsx"

EXPECTED = {
    # cell            code                      why it is wrong
    "D2":  "MISSING_ABSOLUTE",       # tax rate B15 referenced relatively, filled down
    "E2":  "HARDCODED_LITERAL",      # 0.08 typed into the formula
    "F2":  "VLOOKUP_APPROX",         # 4th arg missing
    "F2#": "VLOOKUP_TABLE_RELATIVE", # table not locked
    "G2":  "XLOOKUP_LENGTH_MISMATCH",# vectors of different length
    "H2":  "INDEX_MATCH_OFFSET",     # INDEX starts one row below MATCH
    "C12": "RANGE_TOO_SHORT",        # SUM(C2:C10) but data runs to C11
    "C13": "RANGE_TRAILING_BLANKS",  # AVERAGE(C2:C20) with blanks after C11
    "C14": "RANGE_INCLUDES_HEADER",  # AVERAGE(C1:C11) includes the header
    "D9":  "INCONSISTENT_FILL",      # one cell in the run retyped by hand
}


def build():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Homework"

    # -- header row (deliberately NOT bold, one column forced right) --------
    headers = ["Employee", "Dept", "Hours", "Pay", "Bonus", "Rate (VLOOKUP)",
               "Rate (XLOOKUP)", "Dept Head"]
    for c, h in enumerate(headers, 1):
        ws.cell(row=1, column=c, value=h)
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="right")  # text forced right

    names = ["Adler", "Bauer", "Castro", "Dunn", "Ergin",
             "Fischer", "Gomez", "Haas", "Iqbal", "Jensen"]
    depts = ["Sales", "Eng", "Sales", "Support", "Eng",
             "Sales", "Support", "Eng", "Sales", "Support"]
    hours = [38, 41, 40, 36, 45, 40, 39, 43, 40, 37]

    for i, (n, d, h) in enumerate(zip(names, depts, hours)):
        r = 2 + i
        ws.cell(row=r, column=1, value=n)
        ws.cell(row=r, column=2, value=d)
        cell_c = ws.cell(row=r, column=3, value=h)  # C2:C11 -> data runs to row 11
        if i == 0:
            cell_c.alignment = Alignment(horizontal="left")  # BUG: number forced left

        # D: BUG — $ missing on the rate cell, so it drifts as it fills down
        ws.cell(row=r, column=4, value=f"=C{r}*B15")

        # E: BUG — 0.08 hardcoded (and 1 is benign, must NOT be flagged)
        ws.cell(row=r, column=5, value=f"=D{r}*0.08")

        # F: BUG — VLOOKUP with no 4th arg and an unlocked table
        ws.cell(row=r, column=6, value=f"=VLOOKUP(B{r},J2:K4,2)")

        # G: BUG — XLOOKUP vectors of different lengths
        ws.cell(row=r, column=7, value=f"=XLOOKUP(B{r},$J$2:$J$4,$K$2:$K$5)")

        # H: BUG — INDEX range starts a row below the MATCH range
        ws.cell(row=r, column=8, value=f"=INDEX($L$3:$L$5,MATCH(B{r},$J$2:$J$4,0))")

    # D9: BUG — one cell in the filled run retyped by hand (wrong row on C)
    ws["D9"] = "=C8*B15"

    # -- money left in General format (BUG: FORMAT_WRONG_FAMILY) ------------
    # (no number_format applied to columns D and E at all)

    # -- aggregate bugs ------------------------------------------------------
    # Kept out of column C so they do not sit inside their own ranges.
    ws["B12"] = "Total hours"
    ws["C12"] = "=SUM(C2:C10)"        # BUG: data runs to C11
    ws["B13"] = "Average hours"
    ws["E13"] = "=AVERAGE(C2:C20)"    # BUG: trailing blanks
    ws["B14"] = "Avg (bad)"
    ws["E14"] = "=AVERAGE(C1:C11)"    # BUG: includes the header
    ws["B17"] = "Circular"
    ws["C17"] = "=SUM(C2:C17)"        # BUG: range swallows the formula's own cell

    # -- assumptions block (referenced relatively above) ---------------------
    ws["A15"] = "Tax rate"
    ws["B15"] = 0.15

    # -- lookup table --------------------------------------------------------
    ws["J1"] = "Dept"
    ws["K1"] = "Rate"
    ws["L1"] = "Head"
    for i, (d, rate, head) in enumerate([("Sales", 22.5, "Novak"),
                                         ("Eng", 34.0, "Okafor"),
                                         ("Support", 19.75, "Petrov")]):
        ws.cell(row=2 + i, column=10, value=d)
        ws.cell(row=2 + i, column=11, value=rate)
        ws.cell(row=2 + i, column=12, value=head)

    # ---- a CORRECT layout that must NOT be flagged -------------------------
    # Aggregates sitting beside their source data, and a row-wise sum covering
    # one deliberate sub-group. Both were false-positive generators: the engine
    # read the aggregates as "the data continues".
    ok = wb.create_sheet("Clean")
    ok["A1"] = "Name"
    for i, m in enumerate(["Jan", "Feb", "Mar"]):
        ok.cell(row=1, column=2 + i, value=m)
    ok["E1"] = "Total"; ok["F1"] = "Average"
    ok["G1"] = "Exam 1"; ok["H1"] = "Exam 2"; ok["I1"] = "Exam Total"
    for r in range(2, 6):
        ok.cell(row=r, column=1, value=f"P{r}")
        for c in range(2, 5):
            ok.cell(row=r, column=c, value=100 * r + c)
        ok.cell(row=r, column=5, value=f"=SUM(B{r}:D{r})")     # aggregate beside data
        ok.cell(row=r, column=6, value=f"=AVERAGE(B{r}:D{r})")  # another aggregate
        ok.cell(row=r, column=7, value=70 + r)
        ok.cell(row=r, column=8, value=80 + r)
        ok.cell(row=r, column=9, value=f"=SUM(G{r}:H{r})")      # different group

    wb.save(OUT)
    return OUT


if __name__ == "__main__":
    print(build())
