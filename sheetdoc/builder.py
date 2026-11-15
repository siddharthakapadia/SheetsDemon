"""Building and editing workbooks.

Covers the "build my spreadsheet" and "edit my spreadsheet" half of the job,
plus local equivalents of the two Excel table-row steps:
add_worksheet_tablerow and update_worksheet_tablerow.
"""

from __future__ import annotations

import datetime as dt
import json
import re
from pathlib import Path

import openpyxl
from openpyxl.chart import BarChart, LineChart, PieChart, Reference, ScatterChart, Series
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

from .formula import col_index, parse_range, parse_ref
from .report import Finding, Report
from .workbook import Book

THIN = Side(style="thin", color="B0B0B0")
MEDIUM = Side(style="medium", color="000000")

NAMED_FORMATS = {
    "currency": '$#,##0.00',
    "currency0": '$#,##0',
    "accounting": '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
    "percent": "0.0%",
    "percent0": "0%",
    "percent2": "0.00%",
    "number": "#,##0.00",
    "integer": "#,##0",
    "date": "mm/dd/yyyy",
    "longdate": "dd mmmm yyyy",
    "time": "hh:mm",
    "text": "@",
    "general": "General",
}


# ==========================================================================
# Build from spec
# ==========================================================================


def build_from_spec(spec: dict | str, out_path: str | None = None) -> Report:
    """Create a workbook from a declarative spec.

    Spec shape:
      {
        "file": "budget.xlsx",
        "sheets": [{
          "name": "Budget",
          "freeze": "A3",
          "columns": {"A": 22, "B": 14},
          "cells": {"A1": {"v": "Q1 Budget", "bold": true, "size": 14}},
          "table": {
             "at": "A3",
             "headers": ["Category", "Budget", "Actual", "Variance"],
             "rows": [["Rent", 1200, 1150, "=B4-C4"]],
             "formats": [null, "currency", "currency", "currency"],
             "style": true
          },
          "conditional": [{"range": "D4:D10", "preset": "negative_red"}],
          "charts": [{"type": "bar", "data": "B3:C10", "cats": "A4:A10",
                      "at": "F3", "title": "Budget vs Actual"}]
        }]
      }
    """
    if isinstance(spec, str):
        spec = json.loads(Path(spec).read_text(encoding="utf-8")) if Path(spec).exists() else json.loads(spec)

    target = Path(out_path or spec.get("file") or "workbook.xlsx").expanduser()
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    rep = Report(title=f"Build workbook — {target.name}")
    made = []

    for sdef in spec.get("sheets", []):
        name = sdef.get("name", "Sheet1")
        ws = wb.create_sheet(name)
        made.append(name)

        tbl = sdef.get("table")
        if tbl:
            _write_spec_table(ws, tbl, rep, name)

        for addr, cdef in (sdef.get("cells") or {}).items():
            _write_spec_cell(ws, addr, cdef)

        for letter, width in (sdef.get("columns") or {}).items():
            ws.column_dimensions[letter].width = width

        if sdef.get("freeze"):
            ws.freeze_panes = sdef["freeze"]

        for cf in sdef.get("conditional") or []:
            from .formatting import _build_rule

            rng = parse_range(cf["range"])
            rule = _build_rule(cf.get("preset", "negative_red"), cf.get("value"),
                               cf.get("value2"), rng)
            if rule is not None:
                ws.conditional_formatting.add(cf["range"], rule)
                rep.add(Finding("done", "CF_ADDED", cf["range"], name,
                                f"Conditional formatting '{cf.get('preset')}' applied"))

        for ch in sdef.get("charts") or []:
            _add_chart_to_ws(ws, ch, rep, name)

        if sdef.get("autofilter"):
            ws.auto_filter.ref = sdef["autofilter"]

    if not made:
        rep.add(Finding("error", "EMPTY_SPEC", None, None,
                        "Spec defined no sheets", "", "Add a 'sheets' array to the spec."))
        return rep

    target.parent.mkdir(parents=True, exist_ok=True)
    wb.save(target)
    rep.summary = {"file": str(target), "sheets": ", ".join(made)}
    rep.add(Finding("done", "BUILT", None, None,
                    f"Created {target.name} with {len(made)} sheet(s)", "",
                    f"Open it: open '{target}'"))
    return rep


def _write_spec_table(ws, tbl: dict, rep: Report, sheet_name: str) -> None:
    at = parse_ref(tbl.get("at", "A1"))
    r0, c0 = (at.row, at.col) if at else (1, 1)
    headers = tbl.get("headers") or []
    rows = tbl.get("rows") or []
    formats = tbl.get("formats") or []

    for i, h in enumerate(headers):
        cell = ws.cell(row=r0, column=c0 + i, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", start_color="1F4E78")
        cell.alignment = Alignment(horizontal="center")
        cell.border = Border(bottom=THIN)

    for ri, row in enumerate(rows):
        for ci, v in enumerate(row):
            cell = ws.cell(row=r0 + 1 + ri, column=c0 + ci, value=_coerce(v))
            if ci < len(formats) and formats[ci]:
                cell.number_format = NAMED_FORMATS.get(str(formats[ci]), str(formats[ci]))

    if headers:
        for i, h in enumerate(headers):
            letter = get_column_letter(c0 + i)
            longest = max([len(str(h))] + [len(str(r[i])) for r in rows[:80] if i < len(r)])
            ws.column_dimensions[letter].width = min(max(longest + 4, 10), 40)

    if tbl.get("totals"):
        tr = r0 + 1 + len(rows)
        ws.cell(row=tr, column=c0, value="Total").font = Font(bold=True)
        for i, spec in enumerate(tbl["totals"]):
            if not spec:
                continue
            letter = get_column_letter(c0 + i)
            fn = spec if isinstance(spec, str) else "SUM"
            cell = ws.cell(row=tr, column=c0 + i,
                           value=f"={fn.upper()}({letter}{r0 + 1}:{letter}{tr - 1})")
            cell.font = Font(bold=True)
            cell.border = Border(top=MEDIUM)
            if i < len(formats) and formats[i]:
                cell.number_format = NAMED_FORMATS.get(str(formats[i]), str(formats[i]))

    if tbl.get("style") and headers and rows:
        ref = (f"{get_column_letter(c0)}{r0}:"
               f"{get_column_letter(c0 + len(headers) - 1)}{r0 + len(rows)}")
        try:
            t = Table(displayName=re.sub(r"\W", "", tbl.get("name", f"Table{sheet_name}")), ref=ref)
            t.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium9", showRowStripes=True, showColumnStripes=False
            )
            ws.add_table(t)
        except Exception as e:
            rep.notes.append(f"Could not add an Excel Table object: {e}")

    ws.freeze_panes = ws.cell(row=r0 + 1, column=c0)


def _write_spec_cell(ws, addr: str, cdef) -> None:
    if not isinstance(cdef, dict):
        cdef = {"v": cdef}
    cell = ws[addr]
    if "v" in cdef:
        cell.value = _coerce(cdef["v"])
    fmt = cdef.get("format") or cdef.get("fmt")
    if fmt:
        cell.number_format = NAMED_FORMATS.get(str(fmt), str(fmt))
    font_kw = {}
    if cdef.get("bold"):
        font_kw["bold"] = True
    if cdef.get("italic"):
        font_kw["italic"] = True
    if cdef.get("size"):
        font_kw["size"] = cdef["size"]
    if cdef.get("color"):
        font_kw["color"] = str(cdef["color"]).lstrip("#").upper()
    if font_kw:
        cell.font = Font(**font_kw)
    if cdef.get("fill"):
        c = str(cdef["fill"]).lstrip("#").upper()
        cell.fill = PatternFill("solid", start_color=c, end_color=c)
    if cdef.get("align"):
        cell.alignment = Alignment(horizontal=cdef["align"])
    if cdef.get("wrap"):
        cell.alignment = Alignment(wrap_text=True, vertical="top")


def _coerce(v):
    if isinstance(v, str):
        m = re.fullmatch(r"(\d{4})-(\d{2})-(\d{2})", v.strip())
        if m:
            return dt.date(int(m[1]), int(m[2]), int(m[3]))
    return v


# ==========================================================================
# Cell editing
# ==========================================================================


def set_cells(
    book: Book, sheet_name: str | None, assignments: dict[str, object],
    number_format: str | None = None, bold: bool = False,
) -> Report:
    """Write values or formulas into specific cells of an existing workbook."""
    rep = Report(title="Edit cells")
    sh = book.sheet(sheet_name)
    for addr, value in assignments.items():
        ref = parse_ref(addr)
        if not ref:
            rep.add(Finding("error", "BAD_ADDRESS", addr, sh.name,
                            f"Not a valid cell address: {addr!r}", "", "Use A1 notation."))
            continue
        cell = sh.cell(ref.row, ref.col)
        before = cell.value
        cell.value = _coerce(value)
        if number_format:
            cell.number_format = NAMED_FORMATS.get(number_format, number_format)
        if bold:
            cell.font = Font(bold=True)
        rep.add(Finding("done", "CELL_SET", addr.upper(), sh.name,
                        f"{addr.upper()} = {value!r}",
                        f"was: {before!r}" if before is not None else "was: empty"))
    return rep


def fill_down(
    book: Book, sheet_name: str | None, source: str, through: str
) -> Report:
    """Replicate a formula down a column exactly as Excel's fill handle would."""
    from .formula import translate

    rep = Report(title=f"Fill {source} down to {through}")
    sh = book.sheet(sheet_name)
    src, dst = parse_ref(source), parse_ref(through)
    if not src or not dst:
        rep.add(Finding("error", "BAD_ADDRESS", None, sh.name,
                        "Could not parse the source or target cell", "", "Use A1 notation."))
        return rep
    formula = sh.formula(src.row, src.col)
    if not formula:
        rep.add(Finding("error", "NO_FORMULA", source, sh.name,
                        f"{source} does not contain a formula",
                        f"It holds: {sh.raw(src.row, src.col)!r}",
                        "Fill only makes sense from a formula cell."))
        return rep

    written = 0
    for r in range(src.row + 1, dst.row + 1):
        sh.cell(r, src.col).value = translate(formula, r - src.row, 0)
        written += 1
    rep.summary = {"source formula": formula, "cells written": written}
    rep.add(Finding("done", "FILLED", f"{source}:{through}", sh.name,
                    f"Filled {written} cell(s) from {source}",
                    f"Last cell now reads: {sh.formula(dst.row, src.col)}"))
    return rep


# ==========================================================================
# Charts
# ==========================================================================


def add_chart(
    book: Book, sheet_name: str | None, chart_type: str, data_range: str,
    categories: str | None = None, anchor: str = "H2", title: str | None = None,
    x_title: str | None = None, y_title: str | None = None,
) -> Report:
    rep = Report(title=f"Add {chart_type} chart")
    sh = book.sheet(sheet_name)
    spec = {
        "type": chart_type, "data": data_range, "cats": categories,
        "at": anchor, "title": title, "x": x_title, "y": y_title,
    }
    _add_chart_to_ws(sh.ws, spec, rep, sh.name)
    return rep


def _add_chart_to_ws(ws, spec: dict, rep: Report, sheet_name: str) -> None:
    kind = str(spec.get("type", "bar")).lower()
    data_rng = parse_range(spec.get("data", ""))
    if not data_rng:
        rep.add(Finding("error", "BAD_RANGE", None, sheet_name,
                        f"Could not parse chart data range {spec.get('data')!r}", "",
                        "Use A1 notation, e.g. B3:C12 (include the header row)."))
        return

    if kind in ("bar", "column"):
        chart = BarChart()
        chart.type = "col" if kind == "column" else "bar"
        chart.style = 10
    elif kind == "line":
        chart = LineChart()
    elif kind == "pie":
        chart = PieChart()
    elif kind in ("scatter", "xy"):
        chart = ScatterChart()
        chart.style = 13
    else:
        rep.add(Finding("error", "BAD_CHART_TYPE", None, sheet_name,
                        f"Unknown chart type {kind!r}", "",
                        "Use one of: bar, column, line, pie, scatter."))
        return

    data = Reference(ws, min_col=data_rng.min_col, min_row=data_rng.min_row,
                     max_col=data_rng.max_col, max_row=data_rng.max_row)

    if kind in ("scatter", "xy"):
        # Scatter needs explicit x/y series: first column is x, the rest are y.
        xref = Reference(ws, min_col=data_rng.min_col, min_row=data_rng.min_row + 1,
                         max_row=data_rng.max_row)
        for c in range(data_rng.min_col + 1, data_rng.max_col + 1):
            yref = Reference(ws, min_col=c, min_row=data_rng.min_row, max_row=data_rng.max_row)
            chart.series.append(Series(yref, xref, title_from_data=True))
    else:
        chart.add_data(data, titles_from_data=True)
        if spec.get("cats"):
            cat_rng = parse_range(spec["cats"])
            if cat_rng:
                chart.set_categories(Reference(
                    ws, min_col=cat_rng.min_col, min_row=cat_rng.min_row,
                    max_col=cat_rng.max_col, max_row=cat_rng.max_row))

    if spec.get("title"):
        chart.title = spec["title"]
    if spec.get("x"):
        chart.x_axis.title = spec["x"]
    if spec.get("y"):
        chart.y_axis.title = spec["y"]
    chart.height, chart.width = 8, 15

    ws.add_chart(chart, spec.get("at", "H2"))
    rep.add(Finding("done", "CHART_ADDED", spec.get("at", "H2"), sheet_name,
                    f"Added a {kind} chart at {spec.get('at', 'H2')}",
                    f"Data: {spec.get('data')}"
                    + (f", categories: {spec.get('cats')}" if spec.get("cats") else ""),
                    "openpyxl writes the chart definition; Excel renders it on open."))


# ==========================================================================
# Table-row operations (local equivalents of the Excel connector steps)
# ==========================================================================


def add_worksheet_tablerow(
    book: Book, sheet_name: str | None, values: list | list[list],
    table_name: str | None = None,
) -> Report:
    """Append one or more rows to a table, extending formulas and formats.

    Local counterpart of the "Microsoft Excel — Add A Worksheet Tablerow" step.
    Unlike a raw append, this copies the previous row's number formats and
    translates its formulas down, which is what makes the new row behave like
    the rest of the table.
    """
    from .formula import translate

    rep = Report(title="Add table row(s)")
    sh = book.sheet(sheet_name)
    bounds = _resolve_table(sh, table_name)
    if bounds is None:
        rep.add(Finding("error", "NO_TABLE", None, sh.name,
                        "No table found on this sheet", "",
                        "Give --table-name, or make sure the sheet has a header row and data."))
        return rep
    min_row, max_row, min_col, max_col = bounds

    rows = values if values and isinstance(values[0], list) else [values]
    written = 0
    totals_row = _find_totals_row(sh, min_row, max_row, min_col, max_col)
    if totals_row:
        # Excel inserts new rows *above* a total row rather than after it.
        # Appending below would copy the SUM formula down and orphan the total.
        max_row = totals_row - 1
        rep.notes.append(
            f"Row {totals_row} holds totals, so new rows are inserted above it "
            "and the total formulas are extended."
        )

    for row_vals in rows:
        target = max_row + 1
        if totals_row:
            sh.ws.insert_rows(target)
            totals_row += 1
        for i in range(max_col - min_col + 1):
            col = min_col + i
            src = sh.cell(max_row, col)
            cell = sh.cell(target, col)
            if i < len(row_vals) and row_vals[i] is not None:
                cell.value = _coerce(row_vals[i])
            elif isinstance(src.value, str) and src.value.startswith("="):
                cell.value = translate(src.value, 1, 0)
            cell.number_format = src.number_format
            if src.alignment and src.alignment.horizontal:
                cell.alignment = Alignment(horizontal=src.alignment.horizontal)
        max_row = target
        written += 1
        rep.add(Finding("done", "ROW_ADDED", f"{get_column_letter(min_col)}{target}", sh.name,
                        f"Row {target}: {row_vals}",
                        "Formats copied from the row above; formulas translated down."))

    if totals_row:
        _restretch_totals(sh, totals_row, min_row, max_row, min_col, max_col, rep)

    _extend_table_ref(sh, table_name, min_row, max_row, min_col, max_col, rep)
    rep.summary = {"rows added": written, "table now ends at row": max_row}
    return rep


def update_worksheet_tablerow(
    book: Book, sheet_name: str | None, row_index: int,
    values: dict[str, object] | list, table_name: str | None = None,
    match_column: str | None = None, match_value: object = None,
) -> Report:
    """Modify values across the cells of one table row.

    Local counterpart of the "Microsoft Excel — Update Worksheet Tablerow" step.
    Target the row either by its 1-based position within the table's data, or by
    matching a value in one of its columns.
    """
    rep = Report(title="Update table row")
    sh = book.sheet(sheet_name)
    bounds = _resolve_table(sh, table_name)
    if bounds is None:
        rep.add(Finding("error", "NO_TABLE", None, sh.name,
                        "No table found on this sheet", "", "Give --table-name."))
        return rep
    min_row, max_row, min_col, max_col = bounds
    headers = {
        str(sh.raw(min_row, c)).strip(): c
        for c in range(min_col, max_col + 1)
        if sh.raw(min_row, c) is not None
    }

    if match_column is not None:
        col = headers.get(match_column) or (
            col_index(match_column) if re.fullmatch(r"[A-Za-z]{1,3}", match_column) else None
        )
        if col is None:
            rep.add(Finding("error", "NO_COLUMN", None, sh.name,
                            f"No column named {match_column!r}",
                            "Available: " + ", ".join(headers), "Use one of the headers above."))
            return rep
        target = next(
            (r for r in range(min_row + 1, max_row + 1)
             if str(sh.value(r, col)).strip() == str(match_value).strip()),
            None,
        )
        if target is None:
            rep.add(Finding("error", "NO_MATCH", None, sh.name,
                            f"No row where {match_column} = {match_value!r}", "",
                            "Check the value, including capitalisation and trailing spaces."))
            return rep
    else:
        target = min_row + int(row_index)
        if not (min_row < target <= max_row):
            rep.add(Finding("error", "ROW_OUT_OF_RANGE", None, sh.name,
                            f"Row index {row_index} is outside the table's "
                            f"{max_row - min_row} data rows", "",
                            f"Use 1 to {max_row - min_row}."))
            return rep

    changes = 0
    if isinstance(values, dict):
        for key, v in values.items():
            col = headers.get(key) or (
                col_index(key) if re.fullmatch(r"[A-Za-z]{1,3}", str(key)) else None
            )
            if col is None:
                rep.add(Finding("warning", "NO_COLUMN", None, sh.name,
                                f"Skipped unknown column {key!r}",
                                "Available: " + ", ".join(headers), ""))
                continue
            before = sh.raw(target, col)
            sh.cell(target, col).value = _coerce(v)
            changes += 1
            rep.add(Finding("done", "CELL_UPDATED",
                            f"{get_column_letter(col)}{target}", sh.name,
                            f"{key}: {before!r} → {v!r}"))
    else:
        for i, v in enumerate(values):
            if v is None:
                continue
            col = min_col + i
            before = sh.raw(target, col)
            sh.cell(target, col).value = _coerce(v)
            changes += 1
            rep.add(Finding("done", "CELL_UPDATED",
                            f"{get_column_letter(col)}{target}", sh.name,
                            f"{before!r} → {v!r}"))

    rep.summary = {"row updated": target, "cells changed": changes}
    return rep


def _find_totals_row(sh, min_row: int, max_row: int, min_col: int, max_col: int) -> int | None:
    """Is the last row of this block a totals row rather than a data row?

    A totals row is one whose formulas aggregate the rows above them, so its
    SUM/AVERAGE ranges end on the row immediately before it.
    """
    if max_row <= min_row + 1:
        return None
    aggregates = 0
    for c in range(min_col, max_col + 1):
        f = sh.formula(max_row, c)
        if not f:
            continue
        for _, rng in parse_ranges_in(f):
            if rng.max_row == max_row - 1 and rng.min_row > min_row - 1:
                aggregates += 1
                break
    return max_row if aggregates else None


def parse_ranges_in(formula: str):
    from .formula import extract_ranges

    return extract_ranges(formula)


def _restretch_totals(sh, totals_row, min_row, max_row, min_col, max_col, rep) -> None:
    """Rewrite the totals formulas so they cover the rows just inserted."""
    fixed = 0
    for c in range(min_col, max_col + 1):
        f = sh.formula(totals_row, c)
        if not f:
            continue
        new = f
        for token, rng in parse_ranges_in(f):
            if rng.min_row > min_row - 1 and rng.max_row < totals_row:
                old_text = token.value
                new_text = (
                    f"{get_column_letter(rng.min_col)}{rng.min_row}:"
                    f"{get_column_letter(rng.max_col)}{max_row}"
                )
                new = new.replace(old_text, new_text)
        if new != f:
            sh.cell(totals_row, c).value = new
            fixed += 1
    if fixed:
        rep.add(Finding("done", "TOTALS_EXTENDED", f"row {totals_row}", sh.name,
                        f"Extended {fixed} total formula(s) to include the new row(s)",
                        f"They now run through row {max_row}."))


def _resolve_table(sh, table_name: str | None):
    """-> (header_row, last_row, first_col, last_col)"""
    if table_name:
        for name, ref in getattr(sh.ws, "tables", {}).items():
            if name.lower() == table_name.lower():
                rng = parse_range(ref if isinstance(ref, str) else ref.ref)
                if rng:
                    return rng.min_row, rng.max_row, rng.min_col, rng.max_col
        return None
    tables = getattr(sh.ws, "tables", {})
    if tables:
        ref = next(iter(tables.values()))
        rng = parse_range(ref if isinstance(ref, str) else ref.ref)
        if rng:
            return rng.min_row, rng.max_row, rng.min_col, rng.max_col
    region = sh.main_region()
    if not region:
        return None
    hr = sh.header_row(region) or region.min_row
    return hr, region.max_row, region.min_col, region.max_col


def _extend_table_ref(sh, table_name, min_row, max_row, min_col, max_col, rep) -> None:
    tables = getattr(sh.ws, "tables", {})
    for name, tbl in list(tables.items()):
        if table_name and name.lower() != table_name.lower():
            continue
        new_ref = (f"{get_column_letter(min_col)}{min_row}:"
                   f"{get_column_letter(max_col)}{max_row}")
        try:
            tbl.ref = new_ref
            rep.notes.append(f"Extended Excel Table '{name}' to {new_ref}.")
        except Exception:
            pass
        return


# ==========================================================================
# Inspection
# ==========================================================================


def inspect(book: Book, sheet_name: str | None = None, show_formulas: bool = True,
            max_rows: int = 30) -> Report:
    """Dump a workbook's structure — the first thing to run on an unfamiliar file."""
    rep = Report(title=f"Structure of {book.path.name}")
    rep.summary = {
        "file": str(book.path),
        "sheets": ", ".join(s.name for s in book.sheets),
    }
    for sh in book.targets(sheet_name):
        region = sh.main_region()
        hr = sh.header_row(region) if region else None
        formulas = list(sh.iter_formulas())
        rep.add(Finding("info", "SHEET", None, sh.name,
                        f"{sh.name}: {sh.max_row} rows × {sh.max_col} cols",
                        (f"data block {region.a1}, header row {hr}, "
                         f"{len(formulas)} formula cell(s)" if region else "no data block found")))
        if region and hr:
            for c in range(region.min_col, region.max_col + 1):
                p = sh.column_profile(c, region)
                if p.kind == "empty":
                    continue
                rep.add(Finding("info", "COLUMN", p.letter, sh.name,
                                f"{p.letter}: {p.header or '(no header)'}",
                                f"type {p.kind}, {p.numeric + p.text + p.dates} values, "
                                f"{p.distinct} distinct, format {p.dominant_format!r}"))
        if show_formulas and formulas:
            runs = sh.formula_runs(min_len=2)
            seen = set()
            for run in runs:
                r, c, f = run.anchor
                rep.add(Finding("info", "FORMULA_RUN", run.a1, sh.name,
                                f"{run.a1} filled {run.axis}ly",
                                f"anchor formula: {f}"))
                seen.update((rr, cc) for rr, cc, _ in run.cells)
            singles = [(r, c, f) for r, c, f in formulas if (r, c) not in seen]
            for r, c, f in singles[:max_rows]:
                rep.add(Finding("info", "FORMULA", sh.addr(r, c), sh.name,
                                f"{sh.addr(r, c)}", f))
            if len(singles) > max_rows:
                rep.notes.append(f"…and {len(singles) - max_rows} more single formulas.")
    return rep
