"""Category 2 — sheet formatting and styling.

audit_number_formats, apply_conditional_formatting_rule,
check_header_and_grid_structure.
"""

from __future__ import annotations

from openpyxl.formatting.rule import (
    CellIsRule,
    ColorScaleRule,
    DataBarRule,
    FormulaRule,
    IconSetRule,
)
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .report import Finding, Report
from .workbook import (
    Book,
    Sheet,
    decimals_in_format,
    expected_format_family,
    format_family,
)

# ==========================================================================
# 2.1 audit_number_formats
# ==========================================================================

FAMILY_EXAMPLE = {
    "currency": '$#,##0.00  →  Home ▸ Number ▸ Currency',
    "accounting": '_($* #,##0.00_)  →  Home ▸ Number ▸ Accounting',
    "percentage": '0.0%  →  Home ▸ Number ▸ Percentage',
    "date": 'mm/dd/yyyy  →  Home ▸ Number ▸ Short Date',
    "integer": '#,##0  →  Home ▸ Number ▸ Number, 0 decimals',
    "number": '#,##0.00',
}


def audit_number_formats(book: Book, sheet_name: str | None = None) -> Report:
    """Check that numeric columns carry the format their header implies.

    Detects the four things graders actually take marks for: money left in
    General, rates shown as 0.075 instead of 7.5%, mixed formats inside one
    column, and decimal counts that disagree row to row.
    """
    rep = Report(title="Number formatting")
    cols_checked = 0

    for sh in book.targets(sheet_name):
        region = sh.main_region()
        if not region:
            rep.notes.append(f"{sh.name}: no data block found.")
            continue
        header_row = sh.header_row(region)

        for col in range(region.min_col, region.max_col + 1):
            prof = sh.column_profile(col, region)
            if prof.kind in ("empty", "text"):
                continue
            cols_checked += 1
            colname = prof.header or f"column {prof.letter}"
            actual = format_family(prof.dominant_format)
            expected = expected_format_family(prof.header, prof)
            span = f"{prof.letter}{prof.first_data_row}:{prof.letter}{prof.last_data_row}"

            # -- inconsistent formats inside one column ------------------
            if not prof.format_consistent:
                variants = sorted(set(prof.formats))
                rep.add(Finding(
                    "warning", "FORMAT_INCONSISTENT", span, sh.name,
                    f"'{colname}' mixes {len(variants)} different number formats",
                    "Found: " + ", ".join(repr(v) for v in variants[:4]),
                    f"Select {span}, then apply one format to the whole column.",
                ))

            # -- decimals disagree ---------------------------------------
            decs = {decimals_in_format(f) for f in prof.formats if format_family(f) != "general"}
            if len(decs) > 1:
                rep.add(Finding(
                    "warning", "FORMAT_DECIMALS", span, sh.name,
                    f"'{colname}' shows {sorted(decs)} decimal places in different rows",
                    "Money columns should show the same number of decimals in every row.",
                    f"Select {span} and set a single decimal count (Home ▸ Number ▸ .00 buttons).",
                ))

            if expected is None:
                # If the only evidence is uncalculated formulas we are guessing
                # the column is numeric, so stay quiet rather than guess loudly.
                if actual == "general" and prof.numeric >= 3:
                    rep.add(Finding(
                        "info", "FORMAT_GENERAL", span, sh.name,
                        f"'{colname}' is numeric but left as General",
                        "General is Excel's 'no format chosen' state. It is not wrong, but a "
                        "graded sheet normally has a deliberate format on every numeric column.",
                        "Apply Number, Currency, or Percentage as appropriate.",
                    ))
                continue

            # -- expected family not met ---------------------------------
            ok = _family_matches(expected, actual)
            if ok:
                rep.add(Finding("pass", "FORMAT_OK", span, sh.name,
                                f"'{colname}' correctly formatted as {actual}"))
                continue

            sev, detail = _severity_for(expected, actual, prof)
            rep.add(Finding(
                sev, "FORMAT_WRONG_FAMILY", span, sh.name,
                f"'{colname}' looks like {expected} data but is formatted as {actual}",
                detail,
                f"Select {span} and apply: {FAMILY_EXAMPLE.get(expected, expected)}",
                data={"expected": expected, "actual": actual,
                      "number_format": prof.dominant_format},
            ))

    rep.summary = {"numeric columns checked": cols_checked}
    return rep


def _family_matches(expected: str, actual: str) -> bool:
    if expected == "currency":
        return actual in ("currency", "accounting")
    if expected == "integer":
        return actual in ("number", "integer", "general")
    if expected == "text_or_plain":
        return actual in ("text", "general", "number")
    return expected == actual


def _severity_for(expected: str, actual: str, prof) -> tuple[str, str]:
    if expected == "percentage" and actual in ("general", "number"):
        vals = prof.numeric_values
        if vals and all(0 < abs(v) <= 1 for v in vals):
            return ("error", (
                "The values are stored as decimals (e.g. 0.075) and displayed raw. Excel's "
                "Percentage format multiplies the display by 100 without changing the stored "
                "value, so this reads as 0.075 where it should read 7.5%."
            ))
        if vals and any(abs(v) > 1.5 for v in vals):
            return ("warning", (
                "Values are above 1, so they may already be stored as whole percents (7.5 "
                "meaning 7.5%). Applying Percentage format would display 750%. Divide by 100 "
                "first, or keep a plain number format and put '%' in the header."
            ))
    if expected == "currency":
        return ("error", (
            "A money column shown as General has no $ sign, no thousands separator, and a "
            "ragged number of decimals — the three things a formatting rubric checks first."
        ))
    if expected == "date":
        return ("warning",
                "Dates stored as serial numbers display as 5-digit numbers like 45678.")
    return ("warning", f"Expected {expected} formatting, found {actual}.")


# ==========================================================================
# 2.2 apply_conditional_formatting_rule
# ==========================================================================

PRESETS = {
    "negative_red": "Light-red fill on values below 0",
    "below_zero": "Light-red fill on values below 0 (alias of negative_red)",
    "above": "Green fill on values above a threshold you supply",
    "below": "Red fill on values below a threshold you supply",
    "between": "Yellow fill on values between two thresholds",
    "top_n": "Green fill on the N largest values",
    "bottom_n": "Red fill on the N smallest values",
    "above_average": "Green fill on values above the range average",
    "duplicates": "Orange fill on duplicated values",
    "unique": "Blue fill on values that appear only once",
    "blanks": "Grey fill on empty cells",
    "errors": "Red fill on cells showing an error",
    "text_contains": "Yellow fill where the cell contains your text",
    "data_bar": "In-cell blue bars proportional to value",
    "color_scale": "Green-yellow-red 3-colour scale",
    "icon_set": "Traffic-light icons across the range",
    "highlight_row": "Fill the whole row when one column meets a condition",
}

FILLS = {
    "red": PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
    "green": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
    "yellow": PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
    "orange": PatternFill(start_color="FFCC99", end_color="FFCC99", fill_type="solid"),
    "blue": PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid"),
    "grey": PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid"),
}
FONTS = {
    "red": Font(color="9C0006"),
    "green": Font(color="006100"),
    "yellow": Font(color="9C6500"),
    "orange": Font(color="9C6500"),
    "blue": Font(color="1F4E78"),
    "grey": Font(color="595959"),
}


def apply_conditional_formatting_rule(
    book: Book,
    sheet_name: str,
    cell_range: str,
    preset: str,
    value: float | str | None = None,
    value2: float | str | None = None,
    dry_run: bool = False,
) -> Report:
    """Configure a conditional-formatting rule, or simulate it and report hits.

    dry_run=True never touches the file: it evaluates the condition against the
    current values and tells you exactly which cells would light up. That is the
    "test" half of the function and the safe way to check an assignment's rule.
    """
    rep = Report(title=f"Conditional formatting — {preset} on {cell_range}")
    sh = book.sheet(sheet_name)
    preset = preset.lower().strip()

    if preset not in PRESETS:
        rep.add(Finding("error", "UNKNOWN_PRESET", None, sh.name,
                        f"Unknown preset {preset!r}",
                        "Available: " + ", ".join(sorted(PRESETS)),
                        "Pick one of the presets listed above."))
        return rep

    from .formula import parse_range

    rng = parse_range(cell_range)
    if not rng:
        rep.add(Finding("error", "BAD_RANGE", None, sh.name,
                        f"Could not parse range {cell_range!r}",
                        "", "Use A1 notation like B2:B21."))
        return rep

    hits, total = _simulate(sh, rng, preset, value, value2)

    if dry_run:
        rep.summary = {
            "rule": PRESETS[preset],
            "range": cell_range,
            "cells that would highlight": f"{len(hits)} of {total}",
        }
        for addr, v in hits[:40]:
            rep.add(Finding("info", "CF_HIT", addr, sh.name,
                            f"would highlight (value: {v})"))
        if len(hits) > 40:
            rep.notes.append(f"…and {len(hits) - 40} more.")
        if not hits:
            rep.add(Finding("warning", "CF_NO_HITS", None, sh.name,
                            "No cells in this range meet the condition",
                            "The rule is valid but nothing would change colour. If the "
                            "assignment expects visible highlighting, re-check the threshold.",
                            "Verify the comparison value and that the range holds numbers, not text."))
        return rep

    rule = _build_rule(preset, value, value2, rng)
    if rule is None:
        rep.add(Finding("error", "PRESET_NEEDS_VALUE", None, sh.name,
                        f"Preset {preset!r} needs a value argument",
                        "", "Pass --value (and --value2 for 'between')."))
        return rep

    sh.ws.conditional_formatting.add(cell_range, rule)
    rep.summary = {
        "rule applied": PRESETS[preset],
        "range": cell_range,
        "cells that now highlight": f"{len(hits)} of {total}",
    }
    rep.add(Finding("done", "CF_APPLIED", cell_range, sh.name,
                    f"Applied '{PRESETS[preset]}' to {cell_range}",
                    f"{len(hits)} of {total} cells currently meet the condition.",
                    "Open the file in Excel and check Home ▸ Conditional Formatting ▸ Manage Rules."))
    return rep


def _build_rule(preset: str, value, value2, rng):
    red, green, yellow, orange, blue, grey = (
        FILLS["red"], FILLS["green"], FILLS["yellow"], FILLS["orange"], FILLS["blue"], FILLS["grey"]
    )
    if preset in ("negative_red", "below_zero"):
        return CellIsRule(operator="lessThan", formula=["0"], fill=red, font=FONTS["red"])
    if preset == "above":
        if value is None:
            return None
        return CellIsRule(operator="greaterThan", formula=[str(value)], fill=green, font=FONTS["green"])
    if preset == "below":
        if value is None:
            return None
        return CellIsRule(operator="lessThan", formula=[str(value)], fill=red, font=FONTS["red"])
    if preset == "between":
        if value is None or value2 is None:
            return None
        return CellIsRule(operator="between", formula=[str(value), str(value2)],
                          fill=yellow, font=FONTS["yellow"])
    if preset == "top_n":
        from openpyxl.formatting.rule import Rule
        from openpyxl.styles.differential import DifferentialStyle
        n = int(value or 10)
        return Rule(type="top10", rank=n, dxf=DifferentialStyle(fill=green, font=FONTS["green"]))
    if preset == "bottom_n":
        from openpyxl.formatting.rule import Rule
        from openpyxl.styles.differential import DifferentialStyle
        n = int(value or 10)
        return Rule(type="top10", rank=n, bottom=True,
                    dxf=DifferentialStyle(fill=red, font=FONTS["red"]))
    if preset == "above_average":
        from openpyxl.formatting.rule import Rule
        from openpyxl.styles.differential import DifferentialStyle
        return Rule(type="aboveAverage", dxf=DifferentialStyle(fill=green, font=FONTS["green"]))
    if preset == "duplicates":
        from openpyxl.formatting.rule import Rule
        from openpyxl.styles.differential import DifferentialStyle
        return Rule(type="duplicateValues", dxf=DifferentialStyle(fill=orange, font=FONTS["orange"]))
    if preset == "unique":
        from openpyxl.formatting.rule import Rule
        from openpyxl.styles.differential import DifferentialStyle
        return Rule(type="uniqueValues", dxf=DifferentialStyle(fill=blue, font=FONTS["blue"]))
    if preset == "blanks":
        from openpyxl.formatting.rule import Rule
        from openpyxl.styles.differential import DifferentialStyle
        return Rule(type="containsBlanks", dxf=DifferentialStyle(fill=grey))
    if preset == "errors":
        from openpyxl.formatting.rule import Rule
        from openpyxl.styles.differential import DifferentialStyle
        return Rule(type="containsErrors", dxf=DifferentialStyle(fill=red, font=FONTS["red"]))
    if preset == "text_contains":
        if value is None:
            return None
        first = f"{get_column_letter(rng.min_col)}{rng.min_row}"
        return FormulaRule(
            formula=[f'NOT(ISERROR(SEARCH("{value}",{first})))'],
            fill=yellow, font=FONTS["yellow"],
        )
    if preset == "data_bar":
        return DataBarRule(start_type="min", end_type="max", color="638EC6", showValue=True)
    if preset == "color_scale":
        return ColorScaleRule(
            start_type="min", start_color="F8696B",
            mid_type="percentile", mid_value=50, mid_color="FFEB84",
            end_type="max", end_color="63BE7B",
        )
    if preset == "icon_set":
        return IconSetRule("3TrafficLights1", "percent", [0, 33, 67], showValue=True)
    if preset == "highlight_row":
        if value is None:
            return None
        # value is a full condition formula anchored to the first row, e.g. $D2<0
        return FormulaRule(formula=[str(value)], fill=yellow, font=FONTS["yellow"])
    return None


def _simulate(sh: Sheet, rng, preset: str, value, value2) -> tuple[list[tuple[str, object]], int]:
    """Evaluate the condition against current values without writing anything."""
    cells: list[tuple[str, object]] = []
    for r, c in rng.cells():
        cells.append((sh.addr(r, c), sh.value(r, c)))
    nums = [(a, v) for a, v in cells if isinstance(v, (int, float)) and not isinstance(v, bool)]
    hits: list[tuple[str, object]] = []

    if preset in ("negative_red", "below_zero"):
        hits = [(a, v) for a, v in nums if v < 0]
    elif preset == "above" and value is not None:
        hits = [(a, v) for a, v in nums if v > float(value)]
    elif preset == "below" and value is not None:
        hits = [(a, v) for a, v in nums if v < float(value)]
    elif preset == "between" and value is not None and value2 is not None:
        lo, hi = sorted((float(value), float(value2)))
        hits = [(a, v) for a, v in nums if lo <= v <= hi]
    elif preset == "top_n":
        n = int(value or 10)
        hits = sorted(nums, key=lambda x: -x[1])[:n]
    elif preset == "bottom_n":
        n = int(value or 10)
        hits = sorted(nums, key=lambda x: x[1])[:n]
    elif preset == "above_average":
        if nums:
            avg = sum(v for _, v in nums) / len(nums)
            hits = [(a, v) for a, v in nums if v > avg]
    elif preset in ("duplicates", "unique"):
        seen: dict[object, int] = {}
        for _, v in cells:
            if v is not None:
                seen[v] = seen.get(v, 0) + 1
        want_dupe = preset == "duplicates"
        hits = [(a, v) for a, v in cells
                if v is not None and ((seen[v] > 1) if want_dupe else (seen[v] == 1))]
    elif preset == "blanks":
        hits = [(a, v) for a, v in cells if v is None or (isinstance(v, str) and not v.strip())]
    elif preset == "errors":
        hits = [(a, v) for a, v in cells if isinstance(v, str) and v.startswith("#")]
    elif preset == "text_contains" and value is not None:
        needle = str(value).lower()
        hits = [(a, v) for a, v in cells if isinstance(v, str) and needle in v.lower()]
    elif preset in ("data_bar", "color_scale", "icon_set"):
        hits = nums
    return hits, len(cells)


# ==========================================================================
# 2.3 check_header_and_grid_structure
# ==========================================================================


def check_header_and_grid_structure(
    book: Book, sheet_name: str | None = None, fix: bool = False
) -> Report:
    """Verify the table-presentation rubric: bold headers, alignment, borders, freeze panes.

    Alignment is judged the way Excel actually behaves: General alignment already
    puts text left and numbers right, so only an *explicit* wrong setting is a
    finding. Flagging every General cell would bury the real problems.
    """
    rep = Report(title="Header and grid structure")
    changed = 0

    for sh in book.targets(sheet_name):
        region = sh.main_region()
        if not region:
            rep.notes.append(f"{sh.name}: no data block found.")
            continue
        hr = sh.header_row(region)
        if hr is None:
            rep.add(Finding("warning", "NO_HEADER", region.a1, sh.name,
                            "Could not identify a header row for this table",
                            "The first row of the block is not mostly text, so a grader "
                            "cannot tell what the columns mean.",
                            "Add a row of short text labels above the data."))
            continue

        rep.summary.setdefault("table", f"{sh.name}!{region.a1}")
        rep.summary.setdefault("header row", hr)

        # -- bold headers ------------------------------------------------
        unbolded = []
        for c in range(region.min_col, region.max_col + 1):
            cell = sh.cell(hr, c)
            if cell.value is None:
                continue
            if not (cell.font and cell.font.bold):
                unbolded.append(sh.addr(hr, c))
        if unbolded:
            if fix:
                for a in unbolded:
                    cell = sh.ws[a]
                    cell.font = Font(bold=True, color=cell.font.color if cell.font else None)
                    changed += 1
            rep.add(Finding(
                "done" if fix else "warning", "HEADER_NOT_BOLD",
                f"{sh.addr(hr, region.min_col)}:{sh.addr(hr, region.max_col)}", sh.name,
                (f"Bolded {len(unbolded)} header cell(s)" if fix
                 else f"{len(unbolded)} header cell(s) are not bold"),
                "Bold headers are the standard visual separator between labels and data.",
                "Select the header row and press ⌘B." if not fix else "",
            ))
        else:
            rep.add(Finding("pass", "HEADER_BOLD", None, sh.name, "Header row is bold"))

        # -- explicit wrong alignment -------------------------------------
        misaligned = []
        for r in range(hr + 1, region.max_row + 1):
            for c in range(region.min_col, region.max_col + 1):
                cell = sh.cell(r, c)
                h = cell.alignment.horizontal if cell.alignment else None
                if h is None or h == "general":
                    continue
                if sh.is_numeric(r, c) and h == "left":
                    misaligned.append((sh.addr(r, c), "number forced left"))
                elif sh.is_text(r, c) and h == "right":
                    misaligned.append((sh.addr(r, c), "text forced right"))
        if misaligned:
            if fix:
                for a, kind in misaligned:
                    sh.ws[a].alignment = Alignment(
                        horizontal="right" if "number" in kind else "left"
                    )
                    changed += 1
            rep.add(Finding(
                "done" if fix else "warning", "ALIGNMENT",
                misaligned[0][0], sh.name,
                (f"Realigned {len(misaligned)} cell(s)" if fix
                 else f"{len(misaligned)} cell(s) have alignment fighting their data type"),
                "Convention: text left, numbers right. Examples — "
                + ", ".join(f"{a} ({k})" for a, k in misaligned[:4]),
                "Clear the manual alignment (Home ▸ Alignment ▸ General) and Excel gets it right."
                if not fix else "",
            ))
        else:
            rep.add(Finding("pass", "ALIGNMENT_OK", None, sh.name,
                            "No cells have alignment contradicting their data type"))

        # -- header separator border ---------------------------------------
        has_border = any(
            (sh.cell(hr, c).border and sh.cell(hr, c).border.bottom
             and sh.cell(hr, c).border.bottom.style)
            for c in range(region.min_col, region.max_col + 1)
        )
        if not has_border:
            if fix:
                thin = Side(style="thin", color="000000")
                for c in range(region.min_col, region.max_col + 1):
                    sh.cell(hr, c).border = Border(bottom=thin)
                    changed += 1
            rep.add(Finding(
                "done" if fix else "info", "NO_HEADER_BORDER", None, sh.name,
                ("Added a bottom border under the header" if fix
                 else "No border separating the header row from the data"),
                "A single bottom border under the header is the cleanest separator; full "
                "gridlines on every cell usually read as cluttered.",
                "Select the header row ▸ Home ▸ Borders ▸ Bottom Border." if not fix else "",
            ))

        # -- freeze panes ---------------------------------------------------
        if region.height > 15 and not sh.ws.freeze_panes:
            if fix:
                sh.ws.freeze_panes = sh.addr(hr + 1, region.min_col)
                changed += 1
            rep.add(Finding(
                "done" if fix else "info", "NO_FREEZE_PANES", None, sh.name,
                (f"Froze panes at {sh.addr(hr + 1, region.min_col)}" if fix
                 else f"{region.height} rows of data with no frozen header row"),
                "Scrolling past row 20 hides the column labels.",
                f"Click {sh.addr(hr + 1, region.min_col)} ▸ View ▸ Freeze Panes." if not fix else "",
            ))

        # -- merged cells inside the data ------------------------------------
        merged_in_data = [
            str(m) for m in sh.ws.merged_cells.ranges
            if m.min_row > hr and m.min_row <= region.max_row
        ]
        if merged_in_data:
            rep.add(Finding(
                "warning", "MERGED_IN_DATA", merged_in_data[0], sh.name,
                f"{len(merged_in_data)} merged cell block(s) inside the data area",
                "Merged cells break sorting, filtering and pivot tables — the exact tools "
                "the next assignment will ask you to use. Merging is fine in a title row above "
                "the table, not inside it.",
                "Unmerge, then use Home ▸ Alignment ▸ Center Across Selection if you wanted "
                "the centred look.",
            ))

        # -- columns too narrow -----------------------------------------------
        narrow = []
        for c in range(region.min_col, region.max_col + 1):
            letter = get_column_letter(c)
            width = sh.ws.column_dimensions[letter].width
            longest = max(
                (len(str(sh.raw(r, c))) for r in range(hr, region.max_row + 1)
                 if sh.raw(r, c) is not None),
                default=0,
            )
            if width and width < longest * 0.7:
                narrow.append(letter)
            elif width is None and longest > 14:
                narrow.append(letter)
        if narrow:
            if fix:
                for c in range(region.min_col, region.max_col + 1):
                    letter = get_column_letter(c)
                    longest = max(
                        (len(str(sh.raw(r, c))) for r in range(hr, region.max_row + 1)
                         if sh.raw(r, c) is not None),
                        default=8,
                    )
                    sh.ws.column_dimensions[letter].width = min(max(longest + 3, 9), 45)
                    changed += 1
            rep.add(Finding(
                "done" if fix else "info", "NARROW_COLUMNS", None, sh.name,
                (f"Auto-fitted columns {region.min_col}–{region.max_col}" if fix
                 else f"Column(s) {', '.join(narrow)} may be too narrow for their contents"),
                "Numbers that do not fit display as ##### and lose marks on a printed sheet.",
                "Select all ▸ Format ▸ AutoFit Column Width." if not fix else "",
            ))

        # -- autofilter -------------------------------------------------------
        if region.height > 10 and not sh.ws.auto_filter.ref:
            rep.add(Finding(
                "info", "NO_AUTOFILTER", None, sh.name,
                "No filter buttons on the header row",
                "Not always required, but most assignments that mention sorting or filtering "
                "expect them.",
                f"Select {region.a1} ▸ Data ▸ Filter.",
            ))

    if fix:
        rep.summary["cells changed"] = changed
    return rep
