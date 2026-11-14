"""Category 3 — pedagogy and tutoring.

hint_generator_socratic, explain_error_code, generate_practice_variant.

The rule these obey: Tier 3 shows the *shape* of the answer with placeholder
names, never the finished formula against the student's own cells.
"""

from __future__ import annotations

import random
import re
from dataclasses import dataclass

from .formula import find_calls, tokenize
from .report import Finding, Report
from .workbook import Book

# ==========================================================================
# 3.1 hint_generator_socratic
# ==========================================================================


@dataclass
class HintSet:
    concept: str
    tier1: str
    tier2: str
    tier3: str
    check_yourself: str = ""


# Diagnosis -> hint ladder. Keyed by the finding codes the audit engine emits,
# so a hint is always anchored to something actually wrong in the sheet.
HINTS: dict[str, HintSet] = {
    "MISSING_ABSOLUTE": HintSet(
        concept="Absolute vs relative references",
        tier1=(
            "When you drag a formula down, every reference in it moves down too — unless you "
            "stop it. Ask yourself: in this formula, which cell should be the *same* for every "
            "row, and which should change with the row?"
        ),
        tier2=(
            "A dollar sign freezes the part it sits in front of. $B$2 never moves. B$2 keeps "
            "row 2 but lets the column slide. $B2 keeps column B but lets the row slide. "
            "Click inside the reference in the formula bar and press ⌘T (F4 on Windows) to "
            "cycle through the four combinations."
        ),
        tier3=(
            "Shape:  =[this row's value] * $[fixed cell]$[fixed row]\n"
            "So a column of results that all share one rate looks like:\n"
            "    =B2*$F$1\n"
            "Fill that down and B2 becomes B3, B4… while $F$1 stays put."
        ),
        check_yourself=(
            "Click the last cell of your filled column and read its formula. Does the fixed "
            "part still point at the same cell as in the first row? If not, it needed a $."
        ),
    ),
    "OVER_ANCHORED": HintSet(
        concept="Absolute vs relative references",
        tier1=(
            "You locked everything with $. Now ask: if every reference is frozen, what changes "
            "when you drag the formula down?"
        ),
        tier2=(
            "Exactly one part of most fill-down formulas must move: the reference to the current "
            "row's own data. Only the shared constant should be locked."
        ),
        tier3=(
            "Shape:  =B2*$F$1     ← B2 moves, $F$1 does not\n"
            "Not:    =$B$2*$F$1   ← nothing moves, every row shows the same answer"
        ),
    ),
    "INCONSISTENT_FILL": HintSet(
        concept="Filling formulas consistently",
        tier1=(
            "One cell in this column does not match its neighbours. A column of results should "
            "come from one formula filled down, not from formulas typed one at a time."
        ),
        tier2=(
            "Click the first cell of the column, then click the odd one out, and compare the two "
            "formulas in the formula bar. Read them as a pattern — which reference is off?"
        ),
        tier3=(
            "Repair method: click the cell that is definitely correct, copy it (⌘C), select the "
            "whole column of results, paste (⌘V). Or grab the small square at the bottom-right "
            "of the correct cell and drag it down the column."
        ),
        check_yourself="Select the whole column and check that every formula reads the same shape.",
    ),
    "HARDCODED_LITERAL": HintSet(
        concept="Parameterising values instead of typing them",
        tier1=(
            "There is a number living inside your formula. Ask: if the professor changes that "
            "number next week, how many formulas would you have to open and edit?"
        ),
        tier2=(
            "The convention is one cell per input, with a text label beside it, usually in an "
            "'Assumptions' block at the top of the sheet. Formulas then point at that cell — "
            "and because every row shares it, that reference is absolute."
        ),
        tier3=(
            "Instead of:  =B2*0.07\n"
            "Build:       A1: Tax rate      B1: 0.07\n"
            "             =B2*$B$1\n"
            "Now changing B1 updates every row at once, and the grader can see your assumption."
        ),
        check_yourself="Read your formula aloud. If you say a number, it probably belongs in a cell.",
    ),
    "RANGE_TOO_SHORT": HintSet(
        concept="Range boundaries in aggregate functions",
        tier1=(
            "Your total does not cover all the data. Look at where the numbers actually start "
            "and stop, then look at the range in your formula."
        ),
        tier2=(
            "Click the cell with the SUM in it. Excel draws a coloured box around the range it "
            "is really using. Does that box reach the last row of data?"
        ),
        tier3=(
            "Shape:  =SUM(first_data_cell:last_data_cell)\n"
            "For data in B2 down to B21:  =SUM(B2:B21)\n"
            "Drag the edge of the coloured box to extend it rather than retyping."
        ),
        check_yourself=(
            "Add up two or three values by hand and compare against your formula's answer."
        ),
    ),
    "RANGE_TRAILING_BLANKS": HintSet(
        concept="Range boundaries in aggregate functions",
        tier1=(
            "Your range runs past where the data ends. SUM survives that, but AVERAGE and COUNT "
            "do not always — and anything typed into those blank rows later joins the total "
            "without warning."
        ),
        tier2="Shrink the range so its last cell is the last row that actually holds a value.",
        tier3="Shape:  =AVERAGE(B2:B21)   not   =AVERAGE(B2:B100)",
    ),
    "RANGE_INCLUDES_HEADER": HintSet(
        concept="Range boundaries in aggregate functions",
        tier1="Your range starts on the label row rather than on the first number.",
        tier2=(
            "Aggregates should cover data only. Headers are text, and text inside a numeric "
            "range is the kind of thing that behaves differently in the grader's copy."
        ),
        tier3="Shape:  =SUM(B2:B21)   where row 1 holds the header",
    ),
    "VLOOKUP_APPROX": HintSet(
        concept="Exact vs approximate match",
        tier1=(
            "VLOOKUP has a 4th argument you left out, and its default is not the one you want. "
            "Ask: are you looking for an exact ID, or for a band like a grade scale?"
        ),
        tier2=(
            "The 4th argument is range_lookup. FALSE (or 0) means 'exact match — return #N/A if "
            "it isn't there'. TRUE (or omitted) means 'approximate — return the nearest smaller "
            "value', which requires the first column to be sorted ascending."
        ),
        tier3=(
            "Shape:  =VLOOKUP(lookup_value, $table$range, column_number, FALSE)\n"
            "The FALSE at the end is the part almost everyone forgets."
        ),
        check_yourself=(
            "Look up a value you know is NOT in the table. Exact match gives #N/A; approximate "
            "match quietly gives you a wrong answer."
        ),
    ),
    "VLOOKUP_TABLE_RELATIVE": HintSet(
        concept="Locking the lookup table",
        tier1=(
            "Your lookup table has no $ on it. Picture dragging this formula down twenty rows — "
            "where has the table moved to by row 20?"
        ),
        tier2=(
            "The table_array is the one thing that must never move. Select it in the formula bar "
            "and press ⌘T (F4) to lock all four corners."
        ),
        tier3="Shape:  =VLOOKUP(A2, $E$2:$G$11, 3, FALSE)   ← note the four $ signs",
        check_yourself=(
            "Fill the formula to the bottom of the column. If the last few rows show #N/A while "
            "the first rows work, the table slid off the end."
        ),
    ),
    "VLOOKUP_COL_OUT_OF_RANGE": HintSet(
        concept="Counting columns in a lookup table",
        tier1=(
            "The column number you asked for is bigger than the table you selected. Count the "
            "columns in your table_array."
        ),
        tier2=(
            "col_index_num counts from the LEFT EDGE of table_array, not from column A of the "
            "sheet. If table_array is E2:G11, then E is 1, F is 2, G is 3."
        ),
        tier3=(
            "If your table is $E$2:$G$11 and you want the value from column G:\n"
            "    =VLOOKUP(A2, $E$2:$G$11, 3, FALSE)"
        ),
    ),
    "XLOOKUP_LENGTH_MISMATCH": HintSet(
        concept="Matching lookup and return vectors",
        tier1=(
            "XLOOKUP compares two lists side by side. Yours are different lengths — so which "
            "row is it supposed to line up with which?"
        ),
        tier2=(
            "lookup_array and return_array must cover exactly the same records: same first row, "
            "same last row. Mismatched sizes give #VALUE!."
        ),
        tier3=(
            "Shape:  =XLOOKUP(A2, $E$2:$E$11, $G$2:$G$11, \"Not found\")\n"
            "Both vectors span rows 2 to 11. Nothing else is allowed."
        ),
    ),
    "XLOOKUP_ORIENTATION": HintSet(
        concept="Matching lookup and return vectors",
        tier1="One of your two arrays is a column and the other is a row. They cannot line up.",
        tier2="Both arrays must run the same direction and cover the same records.",
        tier3="Shape:  =XLOOKUP(A2, $E$2:$E$11, $G$2:$G$11, \"Not found\")",
    ),
    "INDEX_MATCH_SIZE": HintSet(
        concept="Aligning INDEX and MATCH ranges",
        tier1=(
            "MATCH does not return a cell — it returns a position, like 'the 4th one'. INDEX "
            "then counts that many rows into its own range. What happens if the two ranges do "
            "not start on the same row?"
        ),
        tier2=(
            "Both ranges must span the same rows. If MATCH searches E2:E11 then INDEX must "
            "return from a range that also runs from row 2 to row 11."
        ),
        tier3=(
            "Shape:  =INDEX($G$2:$G$11, MATCH(A2, $E$2:$E$11, 0))\n"
            "Same rows on both sides, and a 0 on the end of MATCH for exact match."
        ),
        check_yourself=(
            "Look up the very first and very last item in your table. If both are right but "
            "middle ones are wrong, the ranges are misaligned."
        ),
    ),
    "INDEX_MATCH_OFFSET": HintSet(
        concept="Aligning INDEX and MATCH ranges",
        tier1="Your two ranges start on different rows, so every answer is shifted.",
        tier2="Line up the first row of the INDEX range with the first row of the MATCH range.",
        tier3="Shape:  =INDEX($G$2:$G$11, MATCH(A2, $E$2:$E$11, 0))",
    ),
    "MATCH_NO_TYPE": HintSet(
        concept="Exact vs approximate match",
        tier1="MATCH has a 3rd argument you left off, and its default is approximate.",
        tier2="Add 0 as the third argument to force an exact match.",
        tier3="Shape:  MATCH(lookup_value, $lookup$range, 0)",
    ),
    "FORMAT_WRONG_FAMILY": HintSet(
        concept="Number formatting",
        tier1=(
            "The values are right but they do not read as what they are. What should a reader "
            "see at a glance — dollars? a percent? a date?"
        ),
        tier2=(
            "Formatting changes only the display, never the stored number. Select the column, "
            "then Home ▸ Number and choose the family."
        ),
        tier3=(
            "Money: Currency or Accounting, 2 decimals.\n"
            "Rates: Percentage — but note that 0.075 displays as 7.5%, while 7.5 displays as "
            "750%. Check which one you stored before you apply it."
        ),
    ),
    "HEADER_NOT_BOLD": HintSet(
        concept="Table presentation",
        tier1="A reader should be able to tell labels from data without thinking about it.",
        tier2="Select the header row and bold it (⌘B); add a bottom border underneath.",
        tier3="Header row: bold, bottom border, and freeze panes below it so it stays visible.",
    ),
    "GENERIC": HintSet(
        concept="General formula review",
        tier1=(
            "Read your formula out loud in plain English, one piece at a time. The step where "
            "the sentence stops making sense is where the bug is."
        ),
        tier2=(
            "Break it up: put each inner piece in its own spare cell and check that each one "
            "gives what you expect before nesting them back together."
        ),
        tier3=(
            "Use Formulas ▸ Evaluate Formula to step through the calculation one operation at "
            "a time and watch where the value turns wrong."
        ),
    ),
}


def hint_generator_socratic(
    book: Book | None,
    sheet_name: str | None = None,
    cell: str | None = None,
    formula: str | None = None,
    topic: str | None = None,
    tier: int = 0,
) -> Report:
    """Produce a 3-tier hint ladder for a cell, a pasted formula, or a topic.

    tier=0 shows all three tiers; tier=1..3 reveals only up to that level, which
    is how you use it when you want to stay stuck on purpose.
    """
    rep = Report(title="Socratic hints")
    diagnosis: list[str] = []
    target_formula = formula or ""

    if book is not None and cell:
        from .formula import parse_ref

        sh = book.sheet(sheet_name)
        ref = parse_ref(cell)
        if ref:
            target_formula = sh.formula(ref.row, ref.col) or ""
            if not target_formula:
                v = sh.value(ref.row, ref.col)
                rep.notes.append(
                    f"{cell} holds no formula (value: {v!r}). Hints below are topic-based."
                )
        diagnosis = _diagnose(book, sheet_name, cell)

    if topic:
        code = _topic_to_code(topic)
        if code:
            diagnosis.insert(0, code)

    if not diagnosis and target_formula:
        diagnosis = _diagnose_formula_text(target_formula)

    codes = list(dict.fromkeys(diagnosis)) or ["GENERIC"]
    primary = codes[0]
    hint = HINTS.get(primary, HINTS["GENERIC"])

    rep.summary = {
        "cell": cell or "(pasted formula)",
        "formula": target_formula or "(none)",
        "concept": hint.concept,
        "diagnosis": primary,
    }

    rep.add(Finding("info", "HINT_T1", cell, sheet_name,
                    "Tier 1 — the concept", hint.tier1))
    if tier == 0 or tier >= 2:
        rep.add(Finding("info", "HINT_T2", cell, sheet_name,
                        "Tier 2 — the syntax", hint.tier2))
    if tier == 0 or tier >= 3:
        rep.add(Finding("info", "HINT_T3", cell, sheet_name,
                        "Tier 3 — the shape (placeholders, not your answer)", hint.tier3))
    if hint.check_yourself and (tier == 0 or tier >= 2):
        rep.add(Finding("info", "HINT_CHECK", cell, sheet_name,
                        "How to check yourself", hint.check_yourself))

    if len(codes) > 1:
        rep.notes.append(
            "Other issues detected in this cell: "
            + ", ".join(HINTS.get(c, HINTS["GENERIC"]).concept for c in codes[1:4])
        )
    if tier and tier < 3:
        rep.notes.append(f"Showing tiers 1–{tier}. Re-run with --tier {tier + 1} for more.")
    return rep


def _diagnose(book: Book, sheet_name: str | None, cell: str) -> list[str]:
    """Run the real audits and keep the findings that touch this cell."""
    from . import audit

    codes: list[str] = []
    target = cell.upper().replace("$", "")
    reports = [
        audit.check_reference_types(book, sheet_name),
        audit.verify_lookup_syntax(book, sheet_name),
        audit.validate_range_bounds(book, sheet_name),
        audit.detect_hardcoded_values(book, sheet_name),
    ]
    for r in reports:
        for f in r.findings:
            if f.severity == "pass":
                continue
            if f.cell and f.cell.upper().replace("$", "").split(":")[0] == target:
                codes.append(f.code)
    return codes


def _diagnose_formula_text(formula: str) -> list[str]:
    """Best-effort diagnosis when we only have the formula string, no workbook."""
    codes: list[str] = []
    calls = find_calls(formula)
    names = {c.name for c in calls}

    for c in calls:
        if c.name == "VLOOKUP":
            if len(c.args) < 4:
                codes.append("VLOOKUP_APPROX")
            if len(c.args) >= 2 and "$" not in (c.arg(1) or ""):
                codes.append("VLOOKUP_TABLE_RELATIVE")
        if c.name == "MATCH" and len(c.args) < 3:
            codes.append("MATCH_NO_TYPE")
        if c.name == "XLOOKUP" and len(c.args) >= 3:
            from .formula import parse_range

            la, ra = parse_range(c.arg(1) or ""), parse_range(c.arg(2) or "")
            if la and ra and (la.height, la.width) != (ra.height, ra.width):
                codes.append("XLOOKUP_LENGTH_MISMATCH")

    for i, t in enumerate(tokenize(formula)):
        if t.kind == "number":
            from .formula import classify_number_context

            if classify_number_context(tokenize(formula), i) == "arithmetic":
                try:
                    if float(t.value) not in (0.0, 1.0):
                        codes.append("HARDCODED_LITERAL")
                        break
                except ValueError:
                    pass

    if "$" not in formula and any(n in names for n in ("VLOOKUP", "XLOOKUP", "INDEX", "MATCH")):
        codes.append("MISSING_ABSOLUTE")
    return codes


TOPIC_MAP = {
    "absolute": "MISSING_ABSOLUTE",
    "relative": "MISSING_ABSOLUTE",
    "reference": "MISSING_ABSOLUTE",
    "anchor": "MISSING_ABSOLUTE",
    "dollar": "MISSING_ABSOLUTE",
    "vlookup": "VLOOKUP_APPROX",
    "xlookup": "XLOOKUP_LENGTH_MISMATCH",
    "index": "INDEX_MATCH_SIZE",
    "match": "INDEX_MATCH_SIZE",
    "lookup": "VLOOKUP_APPROX",
    "sum": "RANGE_TOO_SHORT",
    "average": "RANGE_TOO_SHORT",
    "range": "RANGE_TOO_SHORT",
    "hardcode": "HARDCODED_LITERAL",
    "hardcoded": "HARDCODED_LITERAL",
    "format": "FORMAT_WRONG_FAMILY",
    "percent": "FORMAT_WRONG_FAMILY",
    "currency": "FORMAT_WRONG_FAMILY",
    "header": "HEADER_NOT_BOLD",
    "fill": "INCONSISTENT_FILL",
    "drag": "MISSING_ABSOLUTE",
}


def _topic_to_code(topic: str) -> str | None:
    t = topic.lower()
    for key, code in TOPIC_MAP.items():
        if key in t:
            return code
    return None


# ==========================================================================
# 3.2 explain_error_code
# ==========================================================================

ERRORS: dict[str, dict[str, object]] = {
    "#N/A": {
        "means": "Not Available — a lookup ran and found nothing.",
        "plain": (
            "Excel searched for your value and came back empty-handed. This is the one error "
            "that is often *correct behaviour*: it is Excel refusing to guess."
        ),
        "causes": [
            "The value genuinely is not in the lookup table.",
            "Trailing spaces: \"Smith \" does not equal \"Smith\".",
            "A number stored as text — 1001 typed into a text-formatted cell will not match the number 1001.",
            "The lookup table was not locked with $, so it slid away as you filled down.",
            "VLOOKUP searching a column that is not the leftmost of table_array.",
        ],
        "find_it": [
            "Does it fail on every row, or only some? Only some → the data does not match. Every row → the formula or the range is wrong.",
            "Does it start working at the top and fail lower down? → the table needs $ locks.",
            "Test the match directly in a spare cell: =A2=E5 should give TRUE.",
            "Clean invisible spaces with TRIM: =TRIM(A2).",
            "Check the type mismatch: =ISNUMBER(A2) and =ISNUMBER(E5) should agree.",
        ],
        "hide_it": (
            "Once you know the misses are legitimate: =IFERROR(your_formula, \"Not found\") "
            "or XLOOKUP's 4th argument. Never wrap in IFERROR before you understand why it "
            "was failing — that hides real bugs."
        ),
    },
    "#VALUE!": {
        "means": "Wrong type of value — the formula got text where it needed a number.",
        "plain": (
            "You asked Excel to do arithmetic on something that is not a number, or handed a "
            "function an argument of the wrong shape."
        ),
        "causes": [
            "Text in a cell you are multiplying or adding, including a space typed into a 'blank' cell.",
            "A number imported as text (usually left-aligned, often with a small green triangle).",
            "XLOOKUP with lookup_array and return_array of different sizes.",
            "Date arithmetic where one side is a date written as text.",
            "A function given a range where it expected a single cell.",
        ],
        "find_it": [
            "Test the suspect cell: =ISNUMBER(B2). FALSE means it is text.",
            "Left-aligned numbers are a giveaway — Excel right-aligns real numbers by default.",
            "Convert text to number: =VALUE(B2), or Data ▸ Text to Columns ▸ Finish on the column.",
            "For SUM specifically: SUM ignores text, so if SUM works but B2+B3 fails, the cell is text.",
        ],
    },
    "#REF!": {
        "means": "Reference destroyed — the cell the formula pointed at no longer exists.",
        "plain": (
            "You deleted a row, column, or sheet that a formula was using. The reference is "
            "gone for good; Excel cannot infer what you meant."
        ),
        "causes": [
            "Deleting a row or column that other formulas referenced.",
            "VLOOKUP col_index_num larger than the width of table_array.",
            "Cutting and pasting cells over a range that formulas depend on.",
            "A formula referring to a deleted or renamed sheet.",
        ],
        "find_it": [
            "⌘Z immediately — if you just deleted something, undo is the whole fix.",
            "Look at the formula: it will literally contain the text #REF! where the reference was.",
            "For VLOOKUP: count the columns in table_array. col_index_num cannot exceed that count.",
            "Retype the reference by clicking the cell you actually meant.",
        ],
    },
    "#DIV/0!": {
        "means": "Division by zero — the denominator is 0 or empty.",
        "plain": (
            "You divided by a cell that holds 0, or by a cell that is blank (which Excel treats "
            "as 0). This one usually means data is missing, not that the formula is wrong."
        ),
        "causes": [
            "The denominator cell is genuinely empty because that row's data has not been entered.",
            "The denominator is a formula that itself returned 0.",
            "AVERAGE over a range with no numbers in it.",
            "A percent-change formula where the prior period was 0.",
        ],
        "find_it": [
            "Click the formula and look at what the bottom of the division points to. Is it blank?",
            "Blank is the common case in a half-filled assignment — fill the data and it resolves itself.",
            "Guard it once you know it is expected: =IFERROR(A2/B2, 0) or =IF(B2=0, \"\", A2/B2).",
        ],
    },
    "#NAME?": {
        "means": "Unrecognised name — Excel does not know a word in your formula.",
        "plain": (
            "Almost always a typo in a function name, or text that is missing its quotation marks."
        ),
        "causes": [
            "Misspelled function: =VLOOKUPP, =AVRAGE, =SUMM.",
            "Text without quotes: =IF(A2>90, Pass, Fail) instead of =IF(A2>90, \"Pass\", \"Fail\").",
            "A named range that was never defined, or was deleted.",
            "A function from a newer Excel version (XLOOKUP, TEXTJOIN) in an older file.",
            "A missing colon in a range: SUM(B2 B21).",
        ],
        "find_it": [
            "Delete the = and retype it. Excel's autocomplete list confirms the correct spelling as you type.",
            "Scan for bare words in the formula — any text output needs \"quotes\".",
            "Formulas ▸ Name Manager shows which named ranges actually exist.",
        ],
    },
    "#NUM!": {
        "means": "Invalid number — the maths is impossible or the result is too large.",
        "plain": "Excel tried the calculation and the answer is not a real number it can store.",
        "causes": [
            "Square root of a negative: =SQRT(-4).",
            "A financial function (IRR, RATE) that cannot converge on an answer.",
            "A result bigger than about 1E+308.",
            "A date before 1900.",
        ],
        "find_it": [
            "Check the sign of the input to SQRT, LOG, or a rate function.",
            "For IRR/RATE, supply a starting guess: =IRR(range, 0.1).",
        ],
    },
    "#NULL!": {
        "means": "Empty intersection — usually a missing comma.",
        "plain": (
            "A space between two references means 'the cells they share'. If they share nothing, "
            "you get #NULL!. Nine times in ten you meant to type a comma or a colon."
        ),
        "causes": [
            "=SUM(A1:A5 C1:C5) — space where a comma belonged.",
            "=SUM(A1 A5) — space where a colon belonged.",
        ],
        "find_it": ["Look for a space between references and replace it with , or :."],
    },
    "#SPILL!": {
        "means": "A dynamic array cannot spill — something is in its way.",
        "plain": (
            "Modern functions like XLOOKUP, FILTER, SORT and UNIQUE return several cells at "
            "once. The area they need is not empty."
        ),
        "causes": [
            "Data sitting in the cells below or to the right of the formula.",
            "The spill area running into a merged cell.",
            "The formula sitting inside a formatted Excel Table.",
        ],
        "find_it": [
            "Excel outlines the blocked area with a dashed border — clear whatever is inside it.",
            "Merged cells inside the spill area must be unmerged.",
        ],
    },
    "#####": {
        "means": "Not an error — the column is too narrow to display the number.",
        "plain": (
            "The value is fine. Excel just refuses to show a number partially, because a "
            "truncated number would be misleading."
        ),
        "causes": ["Column too narrow.", "A negative date or time value."],
        "find_it": [
            "Double-click the right edge of the column header to auto-fit.",
            "Or select all ▸ Format ▸ AutoFit Column Width.",
        ],
    },
    "CIRCULAR": {
        "means": "Circular reference — a formula includes its own cell.",
        "plain": (
            "The formula needs its own answer in order to calculate its own answer. Excel gives "
            "up and shows 0."
        ),
        "causes": [
            "=SUM(B2:B10) typed into B10 — the range swallowed the total cell.",
            "Two cells that reference each other.",
        ],
        "find_it": [
            "Formulas ▸ Error Checking ▸ Circular References lists the offending cell.",
            "Shrink the range so it stops one row above the total.",
        ],
    },
}

_ALIASES = {
    "NA": "#N/A", "N/A": "#N/A", "VALUE": "#VALUE!", "REF": "#REF!",
    "DIV0": "#DIV/0!", "DIV/0": "#DIV/0!", "NAME": "#NAME?", "NUM": "#NUM!",
    "NULL": "#NULL!", "SPILL": "#SPILL!", "HASH": "#####", "CIRCULAR": "CIRCULAR",
    "CIRC": "CIRCULAR",
}


def explain_error_code(code: str, context_formula: str | None = None) -> Report:
    """Translate an Excel error into a beginner explanation plus a debugging route."""
    key = code.strip().upper()
    if key not in ERRORS:
        if set(key) == {"#"}:  # any run of hashes means the column is too narrow
            key = "#####"
        else:
            key = _ALIASES.get(key.lstrip("#").rstrip("!?"), key)
    if key not in ERRORS:
        rep = Report(title=f"Unknown error code: {code}")
        rep.add(Finding("warning", "UNKNOWN_ERROR", None, None,
                        f"{code!r} is not a standard Excel error",
                        "Known codes: " + ", ".join(ERRORS),
                        "Check the spelling, or paste the exact text from the cell."))
        return rep

    info = ERRORS[key]
    rep = Report(title=f"{key} — {info['means']}")
    rep.summary = {"error": key, "what it means": info["means"]}
    rep.add(Finding("info", "ERR_PLAIN", None, None, "In plain English", str(info["plain"])))

    for i, cause in enumerate(info["causes"], 1):  # type: ignore[union-attr]
        rep.add(Finding("info", "ERR_CAUSE", None, None, f"Likely cause {i}", str(cause)))
    for i, step in enumerate(info["find_it"], 1):  # type: ignore[union-attr]
        rep.add(Finding("info", "ERR_DEBUG", None, None, f"Debug step {i}", str(step)))
    if "hide_it" in info:
        rep.add(Finding("warning", "ERR_SUPPRESS", None, None,
                        "Only after you understand it", str(info["hide_it"])))

    if context_formula:
        hints = _diagnose_formula_text(context_formula)
        if hints:
            rep.notes.append(
                "Your formula also shows: "
                + ", ".join(HINTS.get(h, HINTS["GENERIC"]).concept for h in dict.fromkeys(hints))
            )
    return rep


# ==========================================================================
# 3.3 generate_practice_variant
# ==========================================================================

FIRST_NAMES = [
    "Amara", "Ben", "Chloe", "Dmitri", "Elena", "Farid", "Grace", "Hugo", "Ines",
    "Jonas", "Kira", "Liam", "Mira", "Noah", "Omar", "Priya", "Quinn", "Rosa",
    "Sven", "Tara", "Ugo", "Vera", "Wes", "Xenia", "Yusuf", "Zoe",
]
LAST_NAMES = [
    "Adler", "Bauer", "Castro", "Dunn", "Ergin", "Fischer", "Gomez", "Haas",
    "Iqbal", "Jensen", "Klein", "Lopez", "Meier", "Novak", "Okafor", "Petrov",
    "Rahman", "Schmidt", "Tanaka", "Ueda", "Vogel", "Weber", "Yilmaz", "Zhang",
]

VARIANTS: dict[str, dict[str, object]] = {
    "vlookup": {
        "skill": "VLOOKUP with exact match and a locked table",
        "why": "Tests whether you lock table_array and remember the FALSE.",
    },
    "xlookup": {
        "skill": "XLOOKUP with matched vectors and a not-found fallback",
        "why": "Tests vector alignment and the 4th argument.",
    },
    "index_match": {
        "skill": "INDEX + MATCH looking leftwards",
        "why": "Tests whether your two ranges span the same rows.",
    },
    "absolute": {
        "skill": "Absolute vs relative references in a fill-down column",
        "why": "Tests the single most-marked concept in an intro course.",
    },
    "sumif": {
        "skill": "SUMIF / COUNTIF / AVERAGEIF with criteria in cells",
        "why": "Tests criteria ranges and not burying thresholds in strings.",
    },
    "if_nested": {
        "skill": "Nested IF for a grade scale",
        "why": "Tests boundary conditions and order of comparisons.",
    },
    "percentage": {
        "skill": "Percent change and percent of total",
        "why": "Tests the $ on the total cell and percentage formatting.",
    },
    "payroll": {
        "skill": "Gross-to-net payroll with overtime and a shared tax rate",
        "why": "Tests fill-down with one locked assumption cell.",
    },
    "gradebook": {
        "skill": "Weighted average grades",
        "why": "Tests SUMPRODUCT or locked weight cells.",
    },
    "budget": {
        "skill": "Budget vs actual variance",
        "why": "Tests negative-number formatting and conditional formatting.",
    },
}


def generate_practice_variant(
    topic: str,
    out_path: str,
    rows: int = 15,
    seed: int | None = None,
    difficulty: str = "same",
) -> Report:
    """Build a parallel practice workbook: same technique, different numbers.

    Writes three sheets — Practice (the task), Data (fresh numbers), and
    Answer Key (hidden by default, so you can check yourself after trying).
    """
    from . import datagen

    topic_key = _match_topic(topic)
    if topic_key is None:
        rep = Report(title=f"No practice template for {topic!r}")
        rep.add(Finding("warning", "NO_TEMPLATE", None, None,
                        f"No template matches {topic!r}",
                        "Available topics: " + ", ".join(sorted(VARIANTS)),
                        "Pick one of the topics above."))
        return rep

    rng = random.Random(seed if seed is not None else random.randrange(1 << 30))
    return datagen.write_practice_workbook(topic_key, out_path, rows, rng, difficulty)


def _match_topic(topic: str) -> str | None:
    t = re.sub(r"[^a-z]+", "_", topic.lower()).strip("_")
    if t in VARIANTS:
        return t
    for key in VARIANTS:
        if key.replace("_", "") in t.replace("_", ""):
            return key
    keywords = {
        "vlookup": "vlookup", "lookup": "vlookup", "xlookup": "xlookup",
        "index": "index_match", "match": "index_match",
        "absolute": "absolute", "dollar": "absolute", "anchor": "absolute",
        "relative": "absolute", "drag": "absolute", "fill": "absolute",
        "sumif": "sumif", "countif": "sumif", "criteria": "sumif",
        "if": "if_nested", "nested": "if_nested", "grade scale": "if_nested",
        "percent": "percentage", "%": "percentage",
        "payroll": "payroll", "salary": "payroll", "wage": "payroll",
        "grade": "gradebook", "weighted": "gradebook",
        "budget": "budget", "variance": "budget",
    }
    for k, v in keywords.items():
        if k in topic.lower():
            return v
    return None
