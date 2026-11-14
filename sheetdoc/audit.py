"""Category 1 — homework and assignment auditing.

detect_hardcoded_values, check_reference_types, validate_range_bounds,
verify_lookup_syntax. These are the four that decide whether a submission
follows intro-course formula standards.
"""

from __future__ import annotations

from openpyxl.utils import get_column_letter

from .formula import (
    AGGREGATE_FUNCS,
    BENIGN_LITERALS,
    CONDITIONAL_AGGREGATES,
    STRUCTURAL_ARGS,
    classify_number_context,
    col_letter,
    enclosing_arg,
    extract_ranges,
    extract_refs,
    find_calls,
    parse_range,
    parse_ref,
    tokenize,
)
from .report import Finding, Report
from .workbook import Book, Sheet, is_formula

# ==========================================================================
# 1.1 detect_hardcoded_values
# ==========================================================================


def detect_hardcoded_values(
    book: Book,
    sheet_name: str | None = None,
    cell_range: str | None = None,
    strict: bool = False,
) -> Report:
    """Flag numbers typed into formulas that belong in their own labelled cell.

    The heuristic that keeps this from crying wolf: a literal is only "magic" if
    it participates in arithmetic or a comparison. A 2 in VLOOKUP(...,2,FALSE)
    is column-index syntax, and a 2 in ROUND(x,2) is a decimal count — neither
    is a value the student should have parameterised.
    """
    rep = Report(title="Hardcoded values in formulas")
    scanned = flagged = 0

    for sh in book.targets(sheet_name):
        bounds = parse_range(cell_range) if cell_range else None
        for row, col, formula in sh.iter_formulas():
            if bounds and not (
                bounds.min_row <= row <= bounds.max_row
                and bounds.min_col <= col <= bounds.max_col
            ):
                continue
            scanned += 1
            addr = sh.addr(row, col)
            tokens = tokenize(formula)

            for i, t in enumerate(tokens):
                if t.kind != "number":
                    continue
                try:
                    val = float(t.value)
                except ValueError:
                    continue

                context = classify_number_context(tokens, i)
                owner = enclosing_arg(formula, t)

                # Structural syntax slot -> not a magic number.
                if owner and context == "argument":
                    fname, argi = owner
                    if argi in STRUCTURAL_ARGS.get(fname, set()):
                        continue

                # Criteria strings in SUMIF/COUNTIF families are thresholds, but
                # they are usually written as ">100" strings and caught below.
                if context == "argument" and not strict:
                    continue

                # Exponents define the maths, not a parameter to lift out.
                if context == "exponent" and not strict:
                    continue

                if val in BENIGN_LITERALS and not strict:
                    continue

                severity = "warning" if context == "arithmetic" else "info"
                label = (
                    "arithmetic constant"
                    if context == "arithmetic"
                    else "comparison threshold"
                )
                suggestion = _suggest_named_cell(val, context, formula)
                flagged += 1
                rep.add(
                    Finding(
                        severity=severity,
                        code="HARDCODED_LITERAL",
                        cell=addr,
                        sheet=sh.name,
                        message=f"{t.value} is typed directly into the formula ({label})",
                        detail=(
                            "If this number ever changes you must edit every formula that "
                            "contains it, and a grader cannot see where it came from."
                        ),
                        fix=suggestion,
                        formula=formula,
                        data={"literal": val, "context": context},
                    )
                )

            # String criteria like ">=90" hide thresholds too.
            for call in find_calls(formula):
                if call.name not in CONDITIONAL_AGGREGATES:
                    continue
                for arg in call.args:
                    a = arg.strip()
                    if (
                        a.startswith('"')
                        and a.endswith('"')
                        and any(op in a for op in (">", "<", "="))
                        and any(ch.isdigit() for ch in a)
                    ):
                        flagged += 1
                        rep.add(
                            Finding(
                                severity="info",
                                code="HARDCODED_CRITERIA",
                                cell=addr,
                                sheet=sh.name,
                                message=f"{call.name} criteria {a} hides a threshold inside text",
                                detail="The cut-off is buried in a string, so it cannot be seen or changed from the sheet.",
                                fix=(
                                    'Put the number in a cell and build the criteria: '
                                    f'{call.name}(..., ">="&$B$1)'
                                ),
                                formula=formula,
                            )
                        )

    rep.summary = {"formulas scanned": scanned, "literals flagged": flagged}
    if scanned == 0:
        rep.notes.append("No formulas found in the scanned area — check the sheet or range.")
    return rep


def _suggest_named_cell(val: float, context: str, formula: str) -> str:
    if context == "comparison":
        return (
            f"Move {_fmt(val)} into a labelled cell (e.g. a 'Threshold' row in an "
            "assumptions block) and reference it absolutely, like $B$2."
        )
    guess = ""
    if 0 < val < 1:
        guess = " This looks like a rate — label it (e.g. 'Tax rate') in its own cell."
    elif val in (12, 52, 365, 4, 24, 60):
        guess = " Periods-per-year style constants are still better held in a labelled cell."
    elif val > 1:
        guess = " Label it (e.g. 'Unit price', 'Multiplier') in its own cell."
    return (
        f"Put {_fmt(val)} in its own cell with a text label beside it, then reference "
        f"that cell absolutely ($B$1) so the formula still works when you drag it.{guess}"
    )


def _fmt(v: float) -> str:
    return str(int(v)) if v == int(v) else str(v)


# ==========================================================================
# 1.2 check_reference_types
# ==========================================================================


def check_reference_types(
    book: Book, sheet_name: str | None = None, min_run: int = 3
) -> Report:
    """Decide whether $ anchoring is right for formulas meant to be dragged.

    Two independent signals:
      1. Consistency — every cell in a filled run must share one R1C1 shape.
         A cell that differs was retyped by hand and usually differs by a bug.
      2. Drift — for each relative reference, simulate the fill. If the anchor
         row points at real data but a later row points at a blank cell, the
         reference walked off its target and needed a $.
    """
    rep = Report(title="Cell reference types ($ anchoring)")
    runs_checked = 0

    for sh in book.targets(sheet_name):
        for run in sh.formula_runs(min_len=min_run):
            run = _trim_totals(run)
            if len(run.cells) < min_run:
                continue
            groups = run.r1c1_groups()
            total = sum(len(v) for v in groups.values())
            dominant = max(len(v) for v in groups.values())

            runs_checked += 1

            # A reference held identical across every cell is a shared constant
            # however the cells were produced, so this check does not depend on
            # the run being a coherent fill.
            _check_constant_by_hand(sh, run, rep)

            # Adjacent formulas are not automatically a *fill*. Five different
            # formulas side by side are five columns of a table, and comparing
            # them to each other produces nothing but noise. Only treat a run as
            # a fill when most of it really does share one shape.
            if dominant < max(2, total * 0.6):
                continue

            _check_run_consistency(sh, run, rep, groups)
            if len(groups) == 1:
                _check_run_drift(sh, run, rep)

        _check_constants_referenced_relatively(sh, rep)

    rep.summary = {"filled runs analysed": runs_checked}
    if runs_checked == 0:
        rep.notes.append(
            f"No runs of {min_run}+ adjacent formulas found — nothing here is meant to be dragged."
        )
    return rep


def _trim_totals(run):
    """Drop a trailing totals cell from a run.

    A SUM sitting directly under a filled column is supposed to look different
    from the rows above it. Left in the run it reads as an inconsistent fill,
    which is the single most annoying false positive this tool can produce.
    """
    if len(run.cells) < 2:
        return run
    last_row, last_col, last_formula = run.cells[-1]
    for _, rng in extract_ranges(last_formula):
        if rng.max_row == last_row - 1 and rng.min_row >= run.cells[0][0]:
            from dataclasses import replace

            return replace(run, cells=run.cells[:-1])
    return run


def _check_run_consistency(sh: Sheet, run, rep: Report, groups=None) -> None:
    groups = groups if groups is not None else run.r1c1_groups()
    if len(groups) <= 1:
        return
    ranked = sorted(groups.items(), key=lambda kv: -len(kv[1]))
    majority_shape, majority_cells = ranked[0]
    for shape, cells in ranked[1:]:
        for addr in cells:
            rep.add(
                Finding(
                    severity="error",
                    code="INCONSISTENT_FILL",
                    cell=addr,
                    sheet=sh.name,
                    message=f"Formula does not match the rest of {run.a1}",
                    detail=(
                        f"{len(majority_cells)} of the {sum(len(v) for v in groups.values())} "
                        f"cells in this run share one pattern; this one is different. "
                        "That normally means it was typed by hand instead of filled."
                    ),
                    fix=(
                        f"Select {majority_cells[0]}, copy it, and paste over {addr} — "
                        "or drag the fill handle down the whole run."
                    ),
                    formula=dict((f"{get_column_letter(c)}{r}", f) for r, c, f in run.cells).get(addr, ""),
                    data={"expected_r1c1": majority_shape, "found_r1c1": shape},
                )
            )


def _check_run_drift(sh: Sheet, run, rep: Report) -> None:
    """Simulate the fill and see which relative refs stop pointing at data."""
    anchor_row, anchor_col, anchor_formula = run.anchor
    refs = extract_refs(anchor_formula)
    if not refs:
        return

    for token, ref in refs:
        if ref.sheet:
            continue  # cross-sheet refs need the other sheet's data; skip
        if ref.col_abs and ref.row_abs:
            continue  # already locked

        anchor_has_data = not sh.is_blank(ref.row, ref.col)
        if not anchor_has_data:
            continue

        drifted_blank: list[str] = []
        first_example: tuple[str, str] | None = None
        for i in range(1, len(run.cells)):
            drow, dcol = run.step(i)
            moved = ref.offset(drow, dcol)
            off_sheet = moved.row < 1 or moved.col < 1 or moved.row > sh.max_row + 50
            if off_sheet or sh.is_blank(moved.row, moved.col):
                drifted_blank.append(run.addr(i))
                if first_example is None:
                    first_example = (run.addr(i), moved.a1)

        if not drifted_blank:
            continue

        share = len(drifted_blank) / max(1, len(run.cells) - 1)
        if share < 0.5:
            continue  # a couple of blanks is ragged data, not a missing $

        needs_row_lock = run.axis == "vertical" and not ref.row_abs
        needs_col_lock = run.axis == "horizontal" and not ref.col_abs
        suggested = _lock(ref, row=needs_row_lock or ref.row_abs, col=needs_col_lock or ref.col_abs)

        rep.add(
            Finding(
                severity="error",
                code="MISSING_ABSOLUTE",
                cell=f"{get_column_letter(anchor_col)}{anchor_row}",
                sheet=sh.name,
                message=(
                    f"{ref.a1} is {ref.anchor_style} but must stay put when this "
                    f"{'column' if run.axis == 'vertical' else 'row'} is filled"
                ),
                detail=(
                    f"Filling {run.a1} walks this reference off its target: "
                    f"{len(drifted_blank)} of {len(run.cells) - 1} filled cells end up pointing "
                    f"at an empty cell"
                    + (f" (in {first_example[0]} it becomes {first_example[1]})." if first_example else ".")
                ),
                fix=f"Change {ref.a1} to {suggested}. In Excel, click the reference and press ⌘T (F4 on Windows).",
                formula=anchor_formula,
                data={"reference": ref.a1, "suggested": suggested, "drifted": drifted_blank[:8]},
            )
        )

    # The inverse mistake: everything locked, so the fill does nothing.
    if all(r.col_abs and r.row_abs for _, r in refs) and len(refs) >= 2:
        rep.add(
            Finding(
                severity="warning",
                code="OVER_ANCHORED",
                cell=f"{get_column_letter(anchor_col)}{anchor_row}",
                sheet=sh.name,
                message="Every reference is fully absolute, so filling this run copies one identical answer",
                detail=(
                    "At least one reference normally has to move with the fill — usually the "
                    "one pointing at the current row's data."
                ),
                fix="Remove the $ from the references that should follow each row (e.g. $B$2 → B2).",
                formula=anchor_formula,
            )
        )


def _check_constant_by_hand(sh: Sheet, run, rep: Report) -> None:
    """Find references kept identical across a run *without* a $ holding them there.

    The formulas happen to be right today, because the student retyped the same
    address on every row. Nothing enforces it: one drag of the fill handle and
    the whole column silently changes. This is the mistake the drift check
    cannot see, because the cells never actually drifted.
    """
    if len(run.cells) < 3:
        return

    per_cell: list[dict[str, object]] = []
    for row, col, formula in run.cells:
        addrs: dict[str, object] = {}
        for _, ref in extract_refs(formula):
            if ref.sheet:
                continue
            addrs[f"{col_letter(ref.col)}{ref.row}"] = ref
        per_cell.append(addrs)

    shared = set(per_cell[0])
    for d in per_cell[1:]:
        shared &= set(d)
    if not shared:
        return

    for addr in sorted(shared):
        ref = per_cell[0][addr]
        if ref.col_abs and ref.row_abs:
            continue  # correctly locked already
        if not sh.is_numeric(ref.row, ref.col):
            continue
        # A reference that never varies across the run is a shared constant.
        locked = f"${col_letter(ref.col)}${ref.row}"
        label = _label_left_of(sh, ref.row, ref.col)
        rep.add(
            Finding(
                severity="error",
                code="CONSTANT_NOT_LOCKED",
                cell=run.a1,
                sheet=sh.name,
                message=(
                    f"{ref.a1} is the same in all {len(run.cells)} cells of {run.a1}, "
                    "but nothing locks it there"
                ),
                detail=(
                    f"Every formula in this run points at {ref.a1}"
                    + (f" ({label})" if label else "")
                    + ", so it is a shared constant. Written without $ it only stays correct "
                    "while nobody touches it — dragging the fill handle down this column, or "
                    "inserting a row above, moves it on every row at once."
                ),
                fix=f"Change {ref.a1} to {locked} in the first cell, then fill the run down from there.",
                formula=run.cells[0][2],
                data={"reference": ref.a1, "suggested": locked},
            )
        )


def _lock(ref, row: bool, col: bool) -> str:
    c = f"{'$' if col else ''}{col_letter(ref.col)}"
    r = f"{'$' if row else ''}{ref.row}"
    return c + r


def _check_constants_referenced_relatively(sh: Sheet, rep: Report) -> None:
    """A relative pointer at an assumptions cell above the table is a classic bug."""
    region = sh.main_region()
    if not region:
        return
    header = sh.header_row(region)
    if not header:
        return
    for row, col, formula in sh.iter_formulas():
        if row <= header:
            continue
        for _, ref in extract_refs(formula):
            if ref.sheet or (ref.row_abs and ref.col_abs):
                continue
            if ref.row < header and sh.is_numeric(ref.row, ref.col):
                label = _label_left_of(sh, ref.row, ref.col)
                rep.add(
                    Finding(
                        severity="warning",
                        code="RELATIVE_TO_CONSTANT",
                        cell=sh.addr(row, col),
                        sheet=sh.name,
                        message=f"{ref.a1} points at an assumption above the table but is not locked",
                        detail=(
                            f"{ref.a1}{f' ({label})' if label else ''} sits above the header row, "
                            "so it is a single constant every row should share."
                        ),
                        fix=f"Lock it: {ref.a1} → ${col_letter(ref.col)}${ref.row}",
                        formula=formula,
                    )
                )
                break


def _label_left_of(sh: Sheet, row: int, col: int) -> str:
    for c in range(col - 1, max(0, col - 3), -1):
        if sh.is_text(row, c):
            return str(sh.raw(row, c)).strip()
    return ""


# ==========================================================================
# 1.3 validate_range_bounds
# ==========================================================================

_ROW_TOLERANCE = 0


def validate_range_bounds(book: Book, sheet_name: str | None = None) -> Report:
    """Check that SUM/AVERAGE/lookup ranges cover exactly the intended data.

    Four failure modes, all common in intro assignments:
      short   — real data continues past the end of the range
      late    — the range starts below the first data row
      padded  — the range trails into blank rows
      header  — the range swallows the header row (breaks AVERAGE/COUNT)
    """
    rep = Report(title="Range boundaries")
    checked = 0

    for sh in book.targets(sheet_name):
        header = sh.header_row()
        for row, col, formula in sh.iter_formulas():
            addr = sh.addr(row, col)
            for call in find_calls(formula):
                if call.name not in AGGREGATE_FUNCS and call.name not in CONDITIONAL_AGGREGATES:
                    continue
                for token, rng in extract_ranges(formula):
                    if rng.sheet:
                        continue
                    checked += 1
                    _check_one_range(sh, rng, addr, formula, call.name, header, row, col, rep)
                break  # one report per formula is enough

    rep.summary = {"ranges checked": checked}
    return rep


def _check_one_range(sh, rng, addr, formula, fname, header, frow, fcol, rep) -> None:
    single_col = rng.width == 1
    single_row = rng.height == 1
    if not (single_col or single_row):
        return  # 2-D blocks are lookup tables; bounds logic below doesn't apply

    col = rng.min_col
    row = rng.min_row

    # -- self-inclusion (circular) --------------------------------------
    if rng.min_row <= frow <= rng.max_row and rng.min_col <= fcol <= rng.max_col:
        rep.add(
            Finding(
                severity="error",
                code="RANGE_INCLUDES_SELF",
                cell=addr,
                sheet=sh.name,
                message=f"{fname} range {rng.a1} includes the formula's own cell",
                detail="Excel will refuse this with a circular reference warning and show 0.",
                fix=f"Shrink the range so it stops before {addr}.",
                formula=formula,
            )
        )
        return

    if single_col:
        _bounds_1d(sh, rng, addr, formula, fname, header, rep, vertical=True)
    else:
        _bounds_1d(sh, rng, addr, formula, fname, header, rep, vertical=False)


def _neighbour_headers(sh, cells: list[str], header: int | None, vertical: bool) -> str:
    """Column headers for the cells sitting beside a range, to give the reader
    the context that decides whether the omission matters."""
    if vertical or not header or not cells:
        return ""
    names = []
    for name in cells[:4]:
        ref = parse_ref(name)
        if ref and sh.is_text(header, ref.col):
            names.append(str(sh.raw(header, ref.col)).strip())
    return ", ".join(names)


def _derives_from(sh, row: int, col: int, rng) -> bool:
    """Does the cell at (row, col) reference `rng`?

    A cell that reads the range cannot be part of it — `=SUM(B4:G4)` sitting in
    H4 is a Total column, not a seventh month. Without this, every well-built
    sheet with aggregates beside its data gets flagged for a "short" range.
    """
    f = sh.formula(row, col)
    if not f:
        return False
    for _, other in extract_ranges(f):
        if other.sheet:
            continue
        if (other.min_row <= rng.max_row and other.max_row >= rng.min_row
                and other.min_col <= rng.max_col and other.max_col >= rng.min_col):
            return True
    for _, ref in extract_refs(f):
        if ref.sheet:
            continue
        if (rng.min_row <= ref.row <= rng.max_row
                and rng.min_col <= ref.col <= rng.max_col):
            return True
    return False


def _bounds_1d(sh, rng, addr, formula, fname, header, rep, vertical: bool) -> None:
    if vertical:
        col = rng.min_col
        first, last = rng.min_row, rng.max_row
        occupied = lambda i: not sh.is_blank(i, col)  # noqa: E731
        numeric = lambda i: sh.is_numeric(i, col)  # noqa: E731
        texty = lambda i: sh.is_text(i, col)  # noqa: E731
        cellname = lambda i: f"{get_column_letter(col)}{i}"  # noqa: E731
        limit = sh.max_row
        axis_word = "row"
    else:
        row = rng.min_row
        first, last = rng.min_col, rng.max_col
        occupied = lambda i: not sh.is_blank(row, i)  # noqa: E731
        numeric = lambda i: sh.is_numeric(row, i)  # noqa: E731
        texty = lambda i: sh.is_text(row, i)  # noqa: E731
        cellname = lambda i: f"{get_column_letter(i)}{row}"  # noqa: E731
        limit = sh.max_col
        axis_word = "column"

    # -- data continues past the end -------------------------------------
    # Stop at the first cell that derives from this range: that is an aggregate
    # column/row (Total, Average, …), and everything beyond it is too.
    def _pos(i: int) -> tuple[int, int]:
        return (i, col) if vertical else (row, i)

    missing_after: list[str] = []
    i = last + 1
    while i <= limit and numeric(i) and not _derives_from(sh, *_pos(i), rng):
        missing_after.append(cellname(i))
        i += 1
    if missing_after:
        # Rows and columns are not symmetric here. Down a column, rows are
        # repeated records: a SUM that stops one row short is a real bug.
        # Across a row, columns are distinct fields — "Homework 1..5" then
        # "Exam 1..3" — so a partial span is usually deliberate. Asserting a bug
        # there produces dozens of false alarms on ordinary coursework, so the
        # horizontal case is reported as something to glance at, not an error.
        nbrs = _neighbour_headers(sh, missing_after, header, vertical)
        rep.add(
            Finding(
                severity="error" if vertical else "info",
                code="RANGE_TOO_SHORT" if vertical else "RANGE_STOPS_EARLY",
                cell=addr,
                sheet=sh.name,
                message=(
                    f"{fname} range {rng.a1} stops short — "
                    f"{len(missing_after)} more value(s) sit just past it"
                    if vertical else
                    f"{fname} range {rng.a1} covers only part of the row"
                ),
                detail=(
                    f"Left out: {', '.join(missing_after[:6])}"
                    + (" …" if len(missing_after) > 6 else "")
                    if vertical else
                    f"{', '.join(missing_after[:6])} "
                    + ("sits" if len(missing_after) == 1 else "sit") + " beside it"
                    + (f" ({nbrs})" if nbrs else "")
                    + ". That is fine if they are a different group; worth a "
                      "glance if they belong in the same total."
                ),
                fix=(
                    f"Extend the range to {rng.a1.split(':')[0]}:{missing_after[-1]}."
                    if vertical else ""
                ),
                formula=formula,
                data={"missing": missing_after},
            )
        )

    # -- range starts late ------------------------------------------------
    missing_before: list[str] = []
    i = first - 1
    while (i >= 1 and numeric(i) and not (header and vertical and i <= header)
           and not _derives_from(sh, *_pos(i), rng)):
        missing_before.append(cellname(i))
        i -= 1
    if missing_before and not vertical:
        # Same asymmetry as above: a row-wise sum starting partway along is
        # normally a deliberate sub-group, not an omission.
        missing_before = []
    if missing_before:
        rep.add(
            Finding(
                severity="error",
                code="RANGE_STARTS_LATE",
                cell=addr,
                sheet=sh.name,
                message=(
                    f"{fname} range {rng.a1} starts one {axis_word} too low — "
                    f"{len(missing_before)} value(s) above it are excluded"
                ),
                detail=f"Left out: {', '.join(reversed(missing_before[:6]))}",
                fix=f"Extend the range up to {missing_before[-1]}.",
                formula=formula,
                data={"missing": missing_before},
            )
        )

    # -- trailing blanks ---------------------------------------------------
    trailing = 0
    i = last
    while i >= first and not occupied(i):
        trailing += 1
        i -= 1
    if trailing:
        sev = "warning" if fname in ("SUM", "COUNTA", "MAX", "MIN", "PRODUCT") else "error"
        why = (
            "SUM ignores blanks, so the total is still right — but the range no longer "
            "describes the data, and a value pasted there later would be counted silently."
            if sev == "warning"
            else f"{fname} is distorted by blank cells in its range."
        )
        rep.add(
            Finding(
                severity=sev,
                code="RANGE_TRAILING_BLANKS",
                cell=addr,
                sheet=sh.name,
                message=f"{fname} range {rng.a1} includes {trailing} empty {axis_word}(s) at the end",
                detail=why,
                fix=f"Trim the range to end at {cellname(i)}.",
                formula=formula,
            )
        )

    # -- header swallowed --------------------------------------------------
    if texty(first):
        sev = "error" if fname in ("AVERAGE", "COUNT", "MEDIAN", "MIN", "MAX") else "warning"
        rep.add(
            Finding(
                severity=sev,
                code="RANGE_INCLUDES_HEADER",
                cell=addr,
                sheet=sh.name,
                message=f"{fname} range {rng.a1} starts on a text cell ({cellname(first)})",
                detail=(
                    "Header text in the range is a silent bug: AVERAGE and COUNT skip it, "
                    "which is usually fine, but if the grader's file stores that cell as a "
                    "number the result changes."
                ),
                fix=f"Start the range at {cellname(first + 1)} instead.",
                formula=formula,
            )
        )


# ==========================================================================
# 1.4 verify_lookup_syntax
# ==========================================================================


def verify_lookup_syntax(book: Book, sheet_name: str | None = None) -> Report:
    """Audit VLOOKUP / XLOOKUP / INDEX+MATCH / HLOOKUP argument correctness."""
    rep = Report(title="Lookup function syntax")
    n = 0

    for sh in book.targets(sheet_name):
        for row, col, formula in sh.iter_formulas():
            addr = sh.addr(row, col)
            calls = find_calls(formula)
            names = {c.name for c in calls}
            for call in calls:
                if call.name == "VLOOKUP":
                    n += 1
                    _vlookup(sh, call, addr, formula, rep)
                elif call.name == "HLOOKUP":
                    n += 1
                    _hlookup(sh, call, addr, formula, rep)
                elif call.name == "XLOOKUP":
                    n += 1
                    _xlookup(sh, call, addr, formula, rep)
                elif call.name == "INDEX" and "MATCH" in names:
                    n += 1
                    _index_match(sh, call, calls, addr, formula, rep)
                elif call.name == "MATCH" and "INDEX" not in names:
                    n += 1
                    _match_alone(sh, call, addr, formula, rep)

    rep.summary = {"lookup calls checked": n}
    if n == 0:
        rep.notes.append("No lookup functions found in this workbook.")
    return rep


def _vlookup(sh, call, addr, formula, rep) -> None:
    if len(call.args) < 3:
        rep.add(Finding("error", "VLOOKUP_ARGS", addr, sh.name,
                        f"VLOOKUP needs at least 3 arguments, got {len(call.args)}",
                        "Signature: VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])",
                        "Add the missing arguments.", formula))
        return

    table = parse_range(call.arg(1) or "")
    lookup_value = call.arg(0) or ""
    col_index_raw = (call.arg(2) or "").strip()

    # -- 4th argument ---------------------------------------------------
    if len(call.args) < 4:
        rep.add(Finding("error", "VLOOKUP_APPROX", addr, sh.name,
                        "VLOOKUP is missing its 4th argument, so it defaults to approximate match",
                        "With the 4th argument omitted Excel assumes TRUE. It then requires the "
                        "table's first column to be sorted ascending, and silently returns the "
                        "nearest smaller match instead of an error when there is no exact hit.",
                        "Add FALSE (or 0) as the 4th argument: VLOOKUP(...,FALSE)", formula))
    else:
        arg4 = (call.arg(3) or "").strip().upper()
        if arg4 in ("TRUE", "1"):
            rep.add(Finding("warning", "VLOOKUP_APPROX", addr, sh.name,
                            f"VLOOKUP uses approximate match ({arg4})",
                            "Correct only for banded lookups like tax brackets or grade scales, "
                            "and only when the first column is sorted ascending.",
                            "If you want an exact match on an ID or name, use FALSE.", formula))

    # -- table_array anchoring ------------------------------------------
    if table and not table.is_fully_absolute:
        rep.add(Finding("error", "VLOOKUP_TABLE_RELATIVE", addr, sh.name,
                        f"table_array {table.a1} is not locked with $",
                        "When you fill this formula down, the lookup table slides down with it "
                        "and rows fall off the bottom, producing #N/A partway through.",
                        f"Lock it: ${get_column_letter(table.min_col)}${table.min_row}:"
                        f"${get_column_letter(table.max_col)}${table.max_row}", formula))

    # -- col_index sanity -------------------------------------------------
    if table and col_index_raw.isdigit():
        ci = int(col_index_raw)
        if ci > table.width:
            rep.add(Finding("error", "VLOOKUP_COL_OUT_OF_RANGE", addr, sh.name,
                            f"col_index_num {ci} exceeds the {table.width}-column table {table.a1}",
                            "Excel returns #REF! because there is no such column in the range.",
                            f"Use a number between 1 and {table.width}, or widen table_array.",
                            formula))
        elif ci == 1:
            rep.add(Finding("info", "VLOOKUP_COL_ONE", addr, sh.name,
                            "col_index_num is 1, which returns the lookup column itself",
                            "This gives back the value you searched for, which is rarely the goal.",
                            "Count columns from the left edge of table_array — the first column is 1.",
                            formula))

    # -- lookup value must be in the table's first column ------------------
    lv = parse_ref(lookup_value)
    if table and lv and not lv.sheet and table.min_col <= lv.col <= table.max_col:
        if lv.col != table.min_col:
            rep.add(Finding("warning", "VLOOKUP_VALUE_INSIDE_TABLE", addr, sh.name,
                            f"lookup_value {lv.a1} sits inside table_array but not in its first column",
                            "VLOOKUP only ever searches the leftmost column of table_array.",
                            f"Start table_array at column {get_column_letter(lv.col)}, or switch to "
                            "XLOOKUP / INDEX+MATCH, which can look left.", formula))


def _hlookup(sh, call, addr, formula, rep) -> None:
    if len(call.args) < 4:
        rep.add(Finding("error", "HLOOKUP_APPROX", addr, sh.name,
                        "HLOOKUP is missing its 4th argument, so it defaults to approximate match",
                        "Same trap as VLOOKUP: omitted means TRUE.",
                        "Add FALSE as the 4th argument.", formula))
    table = parse_range(call.arg(1) or "")
    if table and not table.is_fully_absolute:
        rep.add(Finding("error", "HLOOKUP_TABLE_RELATIVE", addr, sh.name,
                        f"table_array {table.a1} is not locked with $",
                        "The table shifts as the formula is copied.",
                        "Lock every corner with $.", formula))
    if table and (call.arg(2) or "").isdigit() and int(call.arg(2)) > table.height:
        rep.add(Finding("error", "HLOOKUP_ROW_OUT_OF_RANGE", addr, sh.name,
                        f"row_index_num {call.arg(2)} exceeds the {table.height}-row table",
                        "Excel returns #REF!.", f"Use 1–{table.height}.", formula))


def _xlookup(sh, call, addr, formula, rep) -> None:
    if len(call.args) < 3:
        rep.add(Finding("error", "XLOOKUP_ARGS", addr, sh.name,
                        f"XLOOKUP needs at least 3 arguments, got {len(call.args)}",
                        "Signature: XLOOKUP(lookup_value, lookup_array, return_array, "
                        "[if_not_found], [match_mode], [search_mode])",
                        "Add the missing arguments.", formula))
        return

    la = parse_range(call.arg(1) or "")
    ra = parse_range(call.arg(2) or "")

    if la and ra:
        # The classic XLOOKUP bug: vectors of different length -> #VALUE!.
        la_len, ra_len = (la.height, ra.height) if la.width == 1 else (la.width, ra.width)
        la_vert, ra_vert = la.width == 1, ra.width == 1
        if la_vert != ra_vert:
            rep.add(Finding("error", "XLOOKUP_ORIENTATION", addr, sh.name,
                            f"lookup_array {la.a1} and return_array {ra.a1} run in different directions",
                            f"One is {'a column' if la_vert else 'a row'} and the other is "
                            f"{'a column' if ra_vert else 'a row'}. XLOOKUP returns #VALUE!.",
                            "Make both vectors run the same way and cover the same records.",
                            formula))
        elif la_len != ra_len:
            rep.add(Finding("error", "XLOOKUP_LENGTH_MISMATCH", addr, sh.name,
                            f"lookup_array {la.a1} has {la_len} cells but return_array {ra.a1} has {ra_len}",
                            "XLOOKUP requires the two vectors to be the same size; mismatched "
                            "vectors return #VALUE!.",
                            f"Resize one of them so both cover the same {la_len if la_len > ra_len else ra_len} records.",
                            formula))

    for name, rng in (("lookup_array", la), ("return_array", ra)):
        if rng and not rng.is_fully_absolute:
            rep.add(Finding("warning", f"XLOOKUP_{name.upper()}_RELATIVE", addr, sh.name,
                            f"{name} {rng.a1} is not locked with $",
                            "Filling the formula down slides this vector down with it.",
                            f"Lock it: ${get_column_letter(rng.min_col)}${rng.min_row}:"
                            f"${get_column_letter(rng.max_col)}${rng.max_row}", formula))

    if len(call.args) < 4:
        rep.add(Finding("info", "XLOOKUP_NO_FALLBACK", addr, sh.name,
                        "No if_not_found argument — misses will show #N/A",
                        "XLOOKUP's 4th argument is the friendly answer for 'not found'. "
                        "This is the main advantage it has over VLOOKUP.",
                        'Add a fallback: XLOOKUP(..., "Not found")', formula))

    if len(call.args) >= 5:
        mm = (call.arg(4) or "").strip()
        if mm not in ("0", ""):
            rep.add(Finding("info", "XLOOKUP_MATCH_MODE", addr, sh.name,
                            f"match_mode is {mm} (not exact)",
                            "0 = exact, -1 = exact or next smaller, 1 = exact or next larger, "
                            "2 = wildcard. Anything but 0 can return a neighbour.",
                            "Use 0 unless the assignment asks for banded matching.", formula))


def _index_match(sh, index_call, calls, addr, formula, rep) -> None:
    matches = [c for c in calls if c.name == "MATCH"]
    index_array = parse_range(index_call.arg(0) or "")

    for m in matches:
        lookup_array = parse_range(m.arg(1) or "")
        if len(m.args) < 3:
            rep.add(Finding("error", "MATCH_NO_TYPE", addr, sh.name,
                            "MATCH is missing its 3rd argument, so it defaults to approximate (1)",
                            "Omitted match_type means 1, which needs sorted data and returns the "
                            "nearest smaller value rather than #N/A.",
                            "Add 0 for an exact match: MATCH(value, range, 0)", formula))
        elif (m.arg(2) or "").strip() not in ("0",):
            rep.add(Finding("warning", "MATCH_APPROX", addr, sh.name,
                            f"MATCH uses match_type {m.arg(2)} instead of 0",
                            "1 and -1 are approximate and require sorted data.",
                            "Use 0 for exact match.", formula))

        if index_array and lookup_array:
            # MATCH returns a position; that position must be valid inside INDEX's array.
            if lookup_array.width == 1 and index_array.height != lookup_array.height:
                rep.add(Finding("error", "INDEX_MATCH_SIZE", addr, sh.name,
                                f"INDEX array {index_array.a1} has {index_array.height} rows but "
                                f"MATCH searches {lookup_array.height} rows in {lookup_array.a1}",
                                "MATCH returns a position counted inside its own range. If the two "
                                "ranges do not start on the same row and cover the same records, "
                                "that position lands on the wrong row of the INDEX array — usually "
                                "an answer that looks plausible and is silently wrong.",
                                f"Make both ranges span the same rows "
                                f"({lookup_array.min_row}–{lookup_array.max_row}).", formula))
            elif lookup_array.height == 1 and index_array.width != lookup_array.width:
                rep.add(Finding("error", "INDEX_MATCH_SIZE", addr, sh.name,
                                f"INDEX array {index_array.a1} has {index_array.width} columns but "
                                f"MATCH searches {lookup_array.width}",
                                "The position MATCH returns will not line up with the INDEX array.",
                                "Make both ranges cover the same columns.", formula))
            elif lookup_array.width == 1 and index_array.min_row != lookup_array.min_row:
                rep.add(Finding("error", "INDEX_MATCH_OFFSET", addr, sh.name,
                                f"INDEX array starts on row {index_array.min_row} but MATCH starts "
                                f"on row {lookup_array.min_row}",
                                f"The two ranges are offset by "
                                f"{abs(index_array.min_row - lookup_array.min_row)} row(s), so every "
                                "answer is shifted by that much.",
                                "Align both ranges to the same first row.", formula))

        if lookup_array and not lookup_array.is_fully_absolute:
            rep.add(Finding("warning", "MATCH_RANGE_RELATIVE", addr, sh.name,
                            f"MATCH range {lookup_array.a1} is not locked with $",
                            "It slides when the formula is filled down.",
                            f"Lock it: ${get_column_letter(lookup_array.min_col)}"
                            f"${lookup_array.min_row}:${get_column_letter(lookup_array.max_col)}"
                            f"${lookup_array.max_row}", formula))

    if index_array and not index_array.is_fully_absolute:
        rep.add(Finding("warning", "INDEX_RANGE_RELATIVE", addr, sh.name,
                        f"INDEX array {index_array.a1} is not locked with $",
                        "It slides when the formula is filled down.",
                        f"Lock it: ${get_column_letter(index_array.min_col)}${index_array.min_row}:"
                        f"${get_column_letter(index_array.max_col)}${index_array.max_row}", formula))


def _match_alone(sh, call, addr, formula, rep) -> None:
    if len(call.args) < 3:
        rep.add(Finding("error", "MATCH_NO_TYPE", addr, sh.name,
                        "MATCH is missing its 3rd argument, so it defaults to approximate (1)",
                        "This needs sorted data and hides misses.",
                        "Add 0: MATCH(value, range, 0)", formula))


# ==========================================================================
# Combined pass
# ==========================================================================


def audit_all(book: Book, sheet_name: str | None = None, strict: bool = False) -> Report:
    rep = Report(title=f"Full formula audit — {book.path.name}")
    parts = [
        detect_hardcoded_values(book, sheet_name, strict=strict),
        check_reference_types(book, sheet_name),
        validate_range_bounds(book, sheet_name),
        verify_lookup_syntax(book, sheet_name),
    ]
    for p in parts:
        rep.extend(p)
    rep.summary = {
        "sheets": ", ".join(s.name for s in book.targets(sheet_name)),
        "checks run": "hardcoded values, reference types, range bounds, lookup syntax",
    }
    # Carry the sub-reports' numeric totals up, so callers that report "how much
    # was examined" (the CLI header, the chatbot's stat tiles) have real numbers
    # rather than nothing.
    for p in parts:
        for key, val in p.summary.items():
            if isinstance(val, int):
                rep.summary[key] = rep.summary.get(key, 0) + val
    return rep
