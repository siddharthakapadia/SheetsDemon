"""Live bridge to the running Microsoft Excel app on macOS, via AppleScript.

Why this exists: openpyxl operates on the *file on disk*. While Excel has a
workbook open it holds unsaved changes in memory and takes a lock, so a
file-based tool cannot see what the user is actually looking at, and writing to
the file underneath Excel gets silently overwritten on the next save.

Talking to Excel itself removes that conflict, and adds something openpyxl
cannot do at all: read **live calculated formula results**. openpyxl only ever
sees the values Excel cached at last save.

macOS only, and requires Excel to be running.
"""

from __future__ import annotations

import re
import shutil
import sys
import subprocess
from dataclasses import dataclass, field
from pathlib import Path

# Delimiters for getting tabular data back out of AppleScript. Chosen because
# they cannot appear in a spreadsheet cell.
CELL_SEP = "\x1f"
ROW_SEP = "\x1e"

TIMEOUT = 30


class ExcelError(Exception):
    """An AppleScript/Excel failure, with a message fit for the end user."""


# --------------------------------------------------------------------------
# AppleScript plumbing
# --------------------------------------------------------------------------


def _quote(s: str) -> str:
    """Escape a Python string for embedding in an AppleScript string literal.

    Cell values and sheet names originate from the model and the user, so this
    is the injection boundary — everything interpolated into a script must pass
    through here.
    """
    out = str(s).replace("\\", "\\\\").replace('"', '\\"')
    # AppleScript literals cannot contain raw newlines or tabs.
    return out.replace("\r\n", " ").replace("\n", " ").replace("\r", " ").replace("\t", " ")


_ADDR_RE = re.compile(r"^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Za-z]{1,3}\$?[0-9]{1,7})?$")


def _check_addr(addr: str) -> str:
    """Validate an A1 address/range before it reaches AppleScript."""
    a = str(addr).strip()
    if not _ADDR_RE.match(a):
        raise ExcelError(f"Not a valid cell or range address: {addr!r}")
    return a


def run_script(script: str) -> str:
    """Execute AppleScript, returning stdout."""
    try:
        p = subprocess.run(
            ["osascript", "-"], input=script, capture_output=True,
            text=True, timeout=TIMEOUT,
        )
    except FileNotFoundError as e:
        raise ExcelError("osascript not found — live Excel needs macOS.") from e
    except subprocess.TimeoutExpired as e:
        raise ExcelError(
            "Excel did not respond in time. It may be showing a dialog — "
            "check the Excel window and dismiss anything waiting."
        ) from e

    if p.returncode != 0:
        err = (p.stderr or "").strip()
        if "Application isn't running" in err or "-600" in err:
            raise ExcelError("Excel is not running. Open Excel and try again.")
        if "-1743" in err or "not allowed" in err.lower():
            raise ExcelError(
                "macOS blocked control of Excel. Grant permission in "
                "System Settings ▸ Privacy & Security ▸ Automation, then retry."
            )
        raise ExcelError(_clean_error(err))
    return p.stdout.strip()


def _clean_error(err: str) -> str:
    """Turn 'execution error: Microsoft Excel got an error: X. (-50)' into X."""
    m = re.search(r"got an error:\s*(.+?)\.?\s*\(-?\d+\)", err)
    if m:
        return f"Excel error: {m.group(1)}"
    return err.splitlines()[-1] if err else "Unknown Excel error"


def _tell(body: str) -> str:
    return f'tell application "Microsoft Excel"\n{body}\nend tell'


# --------------------------------------------------------------------------
# State
# --------------------------------------------------------------------------


def applescript_available() -> bool:
    """Is this machine capable of driving Excel at all?

    Distinguishes "Excel is closed" from "this is not a Mac" — telling a Linux
    or Windows user to open Excel sends them looking for a problem that is not
    theirs to fix.
    """
    return sys.platform == "darwin" and shutil.which("osascript") is not None


NO_APPLESCRIPT = (
    "Live Excel needs macOS with Microsoft Excel — this machine has no "
    "AppleScript. File-based auditing, the CLI and workbook building all "
    "still work."
)


def is_running() -> bool:
    if not applescript_available():
        return False
    try:
        out = subprocess.run(
            ["osascript", "-e",
             'tell application "System Events" to return exists application process "Microsoft Excel"'],
            capture_output=True, text=True, timeout=10,
        )
        return out.stdout.strip() == "true"
    except Exception:
        return False


@dataclass
class LiveState:
    running: bool = False
    workbooks: list[str] = field(default_factory=list)
    active_workbook: str | None = None
    active_sheet: str | None = None
    sheets: list[str] = field(default_factory=list)
    selection: str | None = None
    used_range: str | None = None
    unsaved: bool = False
    error: str | None = None


def state() -> LiveState:
    """One round-trip snapshot of what Excel is doing right now."""
    if not applescript_available():
        return LiveState(running=False, error=NO_APPLESCRIPT)
    if not is_running():
        return LiveState(running=False, error="Excel is not running.")

    # Excel for Mac cannot resolve `name of s` when iterating `worksheets of wb`
    # directly (Parameter error -50) — index the collection instead.
    script = _tell(f'''
  set S to (ASCII character 31)
  if (count of workbooks) is 0 then return "NONE"
  set wb to active workbook
  set sh to active sheet
  set shNames to ""
  repeat with i from 1 to (count of worksheets of wb)
    set shNames to shNames & (name of worksheet i of wb) & ","
  end repeat
  set wbNames to ""
  repeat with i from 1 to (count of workbooks)
    set wbNames to wbNames & (name of workbook i) & ","
  end repeat
  set sel to ""
  try
    set sel to get address of selection
  end try
  set ur to ""
  try
    set ur to get address of (used range of sh)
  end try
  return (name of wb) & S & (name of sh) & S & shNames & S & sel & S & ur & S & (saved of wb) & S & wbNames
''')
    try:
        out = run_script(script)
    except ExcelError as e:
        return LiveState(running=True, error=str(e))

    if out == "NONE":
        return LiveState(running=True, error="No workbook is open in Excel.")

    parts = out.split(CELL_SEP)
    while len(parts) < 7:
        parts.append("")
    return LiveState(
        running=True,
        active_workbook=parts[0] or None,
        active_sheet=parts[1] or None,
        sheets=[s for s in parts[2].split(",") if s],
        selection=(parts[3] or "").replace("$", "") or None,
        used_range=(parts[4] or "").replace("$", "") or None,
        unsaved=parts[5].strip().lower() in ("false", "0"),
        workbooks=[w for w in parts[6].split(",") if w],
    )


# --------------------------------------------------------------------------
# Reading
# --------------------------------------------------------------------------


@dataclass
class Grid:
    address: str
    sheet: str
    values: list[list[str]]
    formulas: list[list[str]]
    start_row: int = 1
    start_col: int = 1


def _parse_table(raw: str) -> list[list[str]]:
    if not raw:
        return []
    return [r.split(CELL_SEP) for r in raw.split(ROW_SEP) if r != ""]


def tidy_number(text: str) -> str:
    """AppleScript coerces numbers to forms no student would recognise.

    83758 comes back as "8.3758E+4" and 3 as "3.0". Render them the way Excel
    shows them, leaving anything non-numeric (text, #N/A, dates) untouched.
    """
    t = text.strip()
    if not t or t[0] not in "-+.0123456789":
        return text
    try:
        n = float(t)
    except ValueError:
        return text
    if n != n or n in (float("inf"), float("-inf")):
        return text
    if n == int(n) and abs(n) < 1e15:
        return str(int(n))
    return f"{n:.10g}"


def read_grid(cell_range: str | None = None, sheet: str | None = None,
              max_rows: int = 60, max_cols: int = 26) -> Grid:
    """Read values and formulas from the live workbook in two bulk calls."""
    st = state()
    if st.error:
        raise ExcelError(st.error)

    target_sheet = sheet or st.active_sheet or ""
    rng = _check_addr(cell_range) if cell_range else (st.used_range or "A1:H20")

    # Cap the window so a huge sheet cannot stall the UI.
    m = re.match(r"^([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)$", rng)
    if m:
        from sheetdoc.formula import col_index, col_letter

        c1, r1, c2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
        ci1, ci2 = col_index(c1), col_index(c2)
        r2 = min(r2, r1 + max_rows - 1)
        ci2 = min(ci2, ci1 + max_cols - 1)
        rng = f"{col_letter(ci1)}{r1}:{col_letter(ci2)}{r2}"
        start_row, start_col = r1, ci1
    else:
        start_row, start_col = 1, 1

    sq, rq = _quote(target_sheet), _quote(rng)
    # `value of range` returns calculated results; `formula of range` returns the
    # formulas. Both come back as 2-D lists in a single call each, which is far
    # faster than walking cells and is what makes a live grid view viable.
    script = _tell(f'''
  set S to (ASCII character 31)
  set R to (ASCII character 30)
  set sh to worksheet "{sq}" of active workbook
  set theRange to range "{rq}" of sh
  set vals to value of theRange
  set fs to formula of theRange
  -- A single-cell range yields a scalar, not a 2-D list. Without this the
  -- loops below would iterate the *characters* of that value.
  if class of vals is not list then set vals to {{{{vals}}}}
  if class of fs is not list then set fs to {{{{fs}}}}
  set vOut to ""
  repeat with rw in vals
    repeat with c in rw
      set vOut to vOut & (c as string) & S
    end repeat
    set vOut to vOut & R
  end repeat
  set fOut to ""
  repeat with rw in fs
    repeat with c in rw
      set fOut to fOut & (c as string) & S
    end repeat
    set fOut to fOut & R
  end repeat
  return vOut & (ASCII character 29) & fOut
''')
    out = run_script(script)
    vraw, _, fraw = out.partition("\x1d")
    values = _parse_table(vraw)
    formulas = _parse_table(fraw)
    # Trailing separator produces an empty last element per row; drop it.
    values = [r[:-1] if r and r[-1] == "" else r for r in values]
    formulas = [r[:-1] if r and r[-1] == "" else r for r in formulas]
    values = [[tidy_number(c) for c in r] for r in values]
    return Grid(address=rng, sheet=target_sheet, values=values, formulas=formulas,
                start_row=start_row, start_col=start_col)


def read_cell(addr: str, sheet: str | None = None) -> dict[str, str]:
    """Value, formula and number format of one cell."""
    a = _check_addr(addr)
    st = state()
    if st.error:
        raise ExcelError(st.error)
    sq = _quote(sheet or st.active_sheet or "")
    script = _tell(f'''
  set S to (ASCII character 31)
  set c to range "{a}" of worksheet "{sq}" of active workbook
  set v to ""
  try
    set v to (value of c) as string
  end try
  -- An error cell (#N/A, #DIV/0!) has no value AppleScript can coerce, and
  -- yields the literal "missing value". Its displayed text is the error code,
  -- which is what the student actually sees.
  if v is "" or v is "missing value" then
    try
      set v to (string value of c) as string
    end try
  end if
  set f to ""
  try
    set f to (formula of c) as string
  end try
  set nf to ""
  try
    set nf to (number format of c) as string
  end try
  return v & S & f & S & nf
''')
    parts = run_script(script).split(CELL_SEP)
    while len(parts) < 3:
        parts.append("")
    return {"address": a, "value": tidy_number(parts[0]),
            "formula": parts[1], "number_format": parts[2]}


# --------------------------------------------------------------------------
# Writing
# --------------------------------------------------------------------------


def write_cells(cells: dict[str, object], sheet: str | None = None) -> int:
    """Write values/formulas into the live workbook. Anything starting with '='
    is set as a formula so Excel calculates it immediately."""
    if not cells:
        return 0
    st = state()
    if st.error:
        raise ExcelError(st.error)
    sq = _quote(sheet or st.active_sheet or "")

    lines = [f'  set sh to worksheet "{sq}" of active workbook']
    for addr, val in cells.items():
        a = _check_addr(addr)
        s = "" if val is None else str(val)
        if s.startswith("="):
            lines.append(f'  set formula of range "{a}" of sh to "{_quote(s)}"')
        elif isinstance(val, (int, float)) and not isinstance(val, bool):
            lines.append(f'  set value of range "{a}" of sh to {val}')
        else:
            lines.append(f'  set value of range "{a}" of sh to "{_quote(s)}"')
    lines.append(f'  return "{len(cells)}"')
    run_script(_tell("\n".join(lines)))
    return len(cells)


def set_number_format(cell_range: str, number_format: str,
                      sheet: str | None = None) -> str:
    a = _check_addr(cell_range)
    st = state()
    if st.error:
        raise ExcelError(st.error)
    sq = _quote(sheet or st.active_sheet or "")
    run_script(_tell(
        f'  set number format of range "{a}" of worksheet "{sq}" '
        f'of active workbook to "{_quote(number_format)}"\n  return "ok"'
    ))
    return a


def close_workbook(name: str, save: bool = False) -> bool:
    """Close a workbook by name. Returns True if one was open."""
    if name not in (state().workbooks or []):
        return False
    run_script(_tell(
        f'  close (every workbook whose name is "{_quote(name)}") '
        f'saving {"yes" if save else "no"}\n  return "ok"'))
    return True


def open_file(path: str | Path, reload: bool = False) -> str:
    """Open a workbook in Excel and bring it to the front.

    If Excel already has it open, this activates the *in-memory* copy — which
    may differ from what is on disk. That is usually what the user wants
    (their unsaved work), but it silently ignores disk changes, so
    `already_open` reports it.
    """
    p = Path(path).expanduser().resolve()
    if not p.is_file():
        raise ExcelError(f"No such file: {p}")
    open_file.already_open = p.name in (state().workbooks or [])
    if reload and open_file.already_open:
        # Force Excel to re-read from disk, discarding the in-memory copy.
        close_workbook(p.name, save=False)
        open_file.already_open = False
    run_script(_tell(f'''
  open POSIX file "{_quote(str(p))}"
  activate
  return "ok"
'''))
    return p.name


open_file.already_open = False


def save_active() -> str:
    st = state()
    if st.error:
        raise ExcelError(st.error)
    run_script(_tell('  save active workbook\n  return "ok"'))
    return st.active_workbook or ""


def active_workbook_path() -> str | None:
    """Full POSIX path of the active workbook, if it has been saved to disk."""
    st = state()
    if st.error:
        return None
    try:
        # `full name` already returns a POSIX path on Excel for Mac; coercing it
        # through `as alias` fails, so use it directly.
        out = run_script(_tell('''
  try
    return (full name of active workbook) as string
  on error
    return ""
  end try
'''))
        return out if out and out.startswith("/") else None
    except ExcelError:
        return None


# --------------------------------------------------------------------------
# Charts
# --------------------------------------------------------------------------

# AppleScript identifiers differ from the friendly names, and several obvious
# guesses are not valid terms ("pie" and "line" both fail; "xy scatter" is a
# syntax error). These are the ones verified against Excel 16.
CHART_TYPES = {
    "column": "column clustered",
    "bar": "bar clustered",
    "line": "line markers",
    "pie": "pie chart",
    "scatter": "xyscatter",
    "area": "area stacked",
    "doughnut": "doughnut",
}

CHART_ALIASES = {
    "col": "column", "vertical bar": "column", "bar chart": "bar",
    "horizontal": "bar", "line graph": "line", "trend": "line",
    "pie chart": "pie", "donut": "doughnut", "xy": "scatter",
    "scatter plot": "scatter", "scatterplot": "scatter",
}


def add_chart(cell_range: str, chart_type: str = "column",
              title: str | None = None, sheet: str | None = None,
              anchor: str | None = None) -> dict[str, str]:
    """Create a native Excel chart from a data range in the live workbook."""
    rng = _check_addr(cell_range)
    kind = CHART_ALIASES.get(chart_type.lower().strip(), chart_type.lower().strip())
    applescript_type = CHART_TYPES.get(kind)
    if not applescript_type:
        raise ExcelError(
            f"I don't know the chart type {chart_type!r}. "
            f"Try: {', '.join(sorted(CHART_TYPES))}."
        )

    st = state()
    if st.error:
        raise ExcelError(st.error)
    sq = _quote(sheet or st.active_sheet or "")

    # Place it clear of the data unless told otherwise.
    top, left = 20, 420
    if anchor:
        a = parse_anchor(anchor)
        if a:
            top, left = a

    title_lines = ""
    if title:
        title_lines = (f'  set has title of ch to true\n'
                       f'  set caption of chart title of ch to "{_quote(title)}"')

    out = run_script(_tell(f'''
  set sh to worksheet "{sq}" of active workbook
  set co to make new chart object at sh with properties ¬
    {{top:{top}, left position:{left}, width:400, height:250}}
  set ch to chart of co
  set source data ch source (range "{rng}" of sh)
  set chart type of ch to {applescript_type}
{title_lines}
  return (count of chart objects of sh) as string
'''))
    return {"chart_type": kind, "range": rng, "sheet": sheet or st.active_sheet or "",
            "charts_on_sheet": out, "title": title or ""}


def parse_anchor(addr: str) -> tuple[int, int] | None:
    """Convert a cell address to rough pixel top/left for chart placement."""
    ref = None
    try:
        from sheetdoc.formula import parse_ref

        ref = parse_ref(_check_addr(addr))
    except ExcelError:
        return None
    if not ref:
        return None
    return ((ref.row - 1) * 20, (ref.col - 1) * 64)


def list_charts(sheet: str | None = None) -> list[str]:
    st = state()
    if st.error:
        raise ExcelError(st.error)
    sq = _quote(sheet or st.active_sheet or "")
    out = run_script(_tell(f'''
  set sh to worksheet "{sq}" of active workbook
  set n to count of chart objects of sh
  set names to ""
  repeat with i from 1 to n
    set names to names & (name of chart object i of sh) & (ASCII character 31)
  end repeat
  return names
'''))
    return [c for c in out.split(CELL_SEP) if c]


def delete_charts(sheet: str | None = None) -> int:
    st = state()
    if st.error:
        raise ExcelError(st.error)
    sq = _quote(sheet or st.active_sheet or "")
    out = run_script(_tell(f'''
  set sh to worksheet "{sq}" of active workbook
  set n to count of chart objects of sh
  repeat with i from n to 1 by -1
    delete chart object i of sh
  end repeat
  return n as string
'''))
    return int(out or 0)


# --------------------------------------------------------------------------
# Tables and new workbooks
# --------------------------------------------------------------------------


def _as_applescript_literal(v) -> str:
    """One cell value as an AppleScript literal."""
    if v is None:
        return '""'
    if isinstance(v, bool):
        return "true" if v else "false"
    if isinstance(v, (int, float)):
        return str(v)
    return f'"{_quote(str(v))}"'


def write_block(top_left: str, rows: list[list], sheet: str | None = None) -> str:
    """Write a rectangular block in ONE call.

    Cell-by-cell writes cost an AppleScript round-trip each; a 20x5 table would
    be 100 of them. Setting `formula` (not `value`) over the whole range means
    entries beginning with '=' become real formulas.
    """
    if not rows:
        raise ExcelError("No rows to write.")
    anchor = parse_ref_or_fail(top_left)
    width = max(len(r) for r in rows)
    padded = [list(r) + [None] * (width - len(r)) for r in rows]

    from sheetdoc.formula import col_letter

    end = f"{col_letter(anchor.col + width - 1)}{anchor.row + len(padded) - 1}"
    target = f"{top_left.upper()}:{end}"

    literal = "{" + ", ".join(
        "{" + ", ".join(_as_applescript_literal(c) for c in row) + "}"
        for row in padded) + "}"

    st = state()
    if st.error:
        raise ExcelError(st.error)
    sq = _quote(sheet or st.active_sheet or "")
    run_script(_tell(f'''
  set sh to worksheet "{sq}" of active workbook
  set formula of range "{target}" of sh to {literal}
  return "ok"
'''))
    return target


def parse_ref_or_fail(addr: str):
    from sheetdoc.formula import parse_ref

    ref = parse_ref(_check_addr(addr))
    if not ref:
        raise ExcelError(f"Not a valid cell address: {addr!r}")
    return ref


def make_table(cell_range: str, sheet: str | None = None,
               style: str = "TableStyleMedium2") -> str:
    """Turn a range into a native Excel Table (banded rows, filter buttons)."""
    rng = _check_addr(cell_range)
    st = state()
    if st.error:
        raise ExcelError(st.error)
    sq = _quote(sheet or st.active_sheet or "")
    out = run_script(_tell(f'''
  set sh to worksheet "{sq}" of active workbook
  set lo to make new list object at sh with properties {{range object:(range "{rng}" of sh)}}
  try
    set table style of lo to "{_quote(style)}"
  end try
  return name of lo
'''))
    return out


def create_table(headers: list[str], rows: list[list] | None = None,
                 at: str = "A1", sheet: str | None = None,
                 as_excel_table: bool = True,
                 number_formats: list[str | None] | None = None) -> dict:
    """Write a headed data table and (optionally) register it as an Excel Table."""
    if not headers:
        raise ExcelError("A table needs at least one column heading.")
    body = rows or []
    block = [list(headers)] + [list(r) for r in body]
    target = write_block(at, block, sheet)

    anchor = parse_ref_or_fail(at)
    from sheetdoc.formula import col_letter

    table_name = ""
    if as_excel_table:
        # With no data rows yet, span one blank row below the headings: Excel
        # needs at least one data row, and it means the student gets filter
        # buttons and auto-expansion straight away.
        span = target
        if not body:
            last = col_letter(anchor.col + len(headers) - 1)
            span = f"{at.upper()}:{last}{anchor.row + 1}"
        try:
            table_name = make_table(span, sheet)
        except ExcelError:
            table_name = ""  # styling is a nicety, not the point

    # Bold the header row even when the Table style did not apply.
    st = state()
    sq = _quote(sheet or st.active_sheet or "")
    hdr_end = f"{col_letter(anchor.col + len(headers) - 1)}{anchor.row}"
    try:
        run_script(_tell(f'''
  set sh to worksheet "{sq}" of active workbook
  set bold of font object of range "{at.upper()}:{hdr_end}" of sh to true
  autofit (entire column of range "{target}" of sh)
  return "ok"
'''))
    except ExcelError:
        pass

    if number_formats:
        for i, fmt in enumerate(number_formats):
            if not fmt or i >= len(headers):
                continue
            col = col_letter(anchor.col + i)
            try:
                set_number_format(
                    f"{col}{anchor.row + 1}:{col}{anchor.row + len(body)}", fmt, sheet)
            except ExcelError:
                pass

    return {"range": target, "rows": len(body), "columns": len(headers),
            "excel_table": table_name, "sheet": sheet or st.active_sheet or ""}


def new_workbook(save_as: str | Path | None = None,
                 sheet_name: str | None = None,
                 headers: list[str] | None = None,
                 rows: list[list] | None = None) -> dict:
    """Create a fresh workbook and open it in Excel.

    The file is written with openpyxl rather than driven through Excel's
    "save as": scripting that path needs UI-automation permission this machine
    does not grant, and it can block on a modal dialog. Building the file
    directly then opening it needs neither.
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment

    target = Path(save_as).expanduser() if save_as else None
    if target is None:
        from datetime import datetime

        from .config import WORKSPACE_DIR

        base = WORKSPACE_DIR
        base.mkdir(parents=True, exist_ok=True)
        target = base / f"new_{datetime.now():%H%M%S}.xlsx"
    target.parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_name or "Sheet1"

    if headers:
        for i, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=i, value=h)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", start_color="1F4E78")
            c.alignment = Alignment(horizontal="center")
        for r, row in enumerate(rows or [], start=2):
            for i, v in enumerate(row, 1):
                ws.cell(row=r, column=i, value=v)
        for i, h in enumerate(headers, 1):
            longest = max([len(str(h))] + [
                len(str(r[i - 1])) for r in (rows or [])[:80] if i <= len(r)])
            ws.column_dimensions[
                openpyxl.utils.get_column_letter(i)].width = min(max(longest + 4, 11), 40)
        ws.freeze_panes = "A2"

    wb.save(target)
    opened = open_file(target, reload=True)
    return {"workbook": opened, "saved_to": str(target),
            "sheet": ws.title, "columns": len(headers or []),
            "rows": len(rows or [])}
