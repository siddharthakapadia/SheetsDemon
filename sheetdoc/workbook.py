"""Workbook loading and structural inference.

openpyxl gives you cells; the audits need *structure* — where the table starts,
which row is the header, which columns are measures vs dimensions, where a
column of filled-down formulas begins and ends. That inference lives here.
"""

from __future__ import annotations

import datetime as _dt
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

import openpyxl
from openpyxl.utils import get_column_letter

from .formula import col_index, to_r1c1

NUMERIC = (int, float)


@dataclass
class Sheet:
    """A worksheet plus everything we inferred about it."""

    name: str
    ws: Any
    vs: Any | None = None  # cached-value twin (data_only=True), may be None
    max_row: int = 0
    max_col: int = 0

    # ---------------- raw access ----------------

    def cell(self, row: int, col: int):
        return self.ws.cell(row=row, column=col)

    def addr(self, row: int, col: int) -> str:
        return f"{get_column_letter(col)}{row}"

    def raw(self, row: int, col: int):
        return self.ws.cell(row=row, column=col).value

    def value(self, row: int, col: int):
        """Cached calculated value if Excel saved one, else the raw value."""
        if self.vs is not None:
            v = self.vs.cell(row=row, column=col).value
            if v is not None:
                return v
        v = self.raw(row, col)
        return None if is_formula(v) else v

    def formula(self, row: int, col: int) -> str | None:
        v = self.raw(row, col)
        return v if is_formula(v) else None

    def is_blank(self, row: int, col: int) -> bool:
        v = self.raw(row, col)
        return v is None or (isinstance(v, str) and not v.strip())

    def is_numeric(self, row: int, col: int) -> bool:
        v = self.value(row, col)
        return isinstance(v, NUMERIC) and not isinstance(v, bool)

    def is_text(self, row: int, col: int) -> bool:
        v = self.raw(row, col)
        return isinstance(v, str) and bool(v.strip()) and not is_formula(v)

    def iter_formulas(self) -> Iterable[tuple[int, int, str]]:
        for row in self.ws.iter_rows(min_row=1, max_row=self.max_row, max_col=self.max_col):
            for c in row:
                if is_formula(c.value):
                    yield c.row, c.column, c.value

    # ---------------- structure ----------------

    def header_row(self, region: "Region | None" = None) -> int | None:
        """Best guess at the header row: the first mostly-text row above data."""
        region = region or self.main_region()
        if not region:
            return None
        for r in range(region.min_row, min(region.min_row + 5, region.max_row + 1)):
            texts = sum(1 for c in range(region.min_col, region.max_col + 1) if self.is_text(r, c))
            nums = sum(1 for c in range(region.min_col, region.max_col + 1) if self.is_numeric(r, c))
            filled = sum(1 for c in range(region.min_col, region.max_col + 1) if not self.is_blank(r, c))
            span = region.max_col - region.min_col + 1
            if filled >= max(2, span * 0.6) and texts > nums:
                return r
        return None

    def regions(self, min_cells: int = 4) -> list["Region"]:
        """Contiguous blocks of non-empty cells, found by flood fill."""
        seen: set[tuple[int, int]] = set()
        out: list[Region] = []
        for r in range(1, self.max_row + 1):
            for c in range(1, self.max_col + 1):
                if (r, c) in seen or self.is_blank(r, c):
                    continue
                stack = [(r, c)]
                cells: list[tuple[int, int]] = []
                seen.add((r, c))
                while stack:
                    cr, cc = stack.pop()
                    cells.append((cr, cc))
                    for dr in (-1, 0, 1):
                        for dc in (-1, 0, 1):
                            nr, nc = cr + dr, cc + dc
                            if (
                                1 <= nr <= self.max_row
                                and 1 <= nc <= self.max_col
                                and (nr, nc) not in seen
                                and not self.is_blank(nr, nc)
                            ):
                                seen.add((nr, nc))
                                stack.append((nr, nc))
                if len(cells) >= min_cells:
                    rows = [x[0] for x in cells]
                    cols = [x[1] for x in cells]
                    out.append(Region(min(rows), max(rows), min(cols), max(cols), len(cells)))
        out.sort(key=lambda x: -x.cell_count)
        return out

    def main_region(self) -> "Region | None":
        rs = self.regions()
        return rs[0] if rs else None

    def column_profile(self, col: int, region: "Region | None" = None) -> "ColumnProfile":
        region = region or self.main_region()
        if not region:
            return ColumnProfile(col, get_column_letter(col), "", 0, 0, 0, 0, [], [])
        hr = self.header_row(region)
        start = (hr + 1) if hr else region.min_row
        header = str(self.raw(hr, col)).strip() if hr and self.raw(hr, col) is not None else ""
        nums = texts = dates = blanks = formulas = 0
        values: list[Any] = []
        formats: list[str] = []
        for r in range(start, region.max_row + 1):
            if self.is_blank(r, col):
                blanks += 1
                continue
            v = self.value(r, col)
            formats.append(self.cell(r, col).number_format)
            if isinstance(v, _dt.date | _dt.datetime):
                dates += 1
            elif isinstance(v, NUMERIC) and not isinstance(v, bool):
                nums += 1
                values.append(v)
            elif v is not None:
                texts += 1
                values.append(v)
            elif self.formula(r, col):
                # A workbook Excel has never saved carries no cached results, so
                # a computed column looks empty. Count it as a formula cell —
                # otherwise every calculated column drops out of every audit.
                formulas += 1
        return ColumnProfile(
            index=col,
            letter=get_column_letter(col),
            header=header,
            numeric=nums,
            text=texts,
            dates=dates,
            blanks=blanks,
            values=values,
            formats=formats,
            first_data_row=start,
            last_data_row=region.max_row,
            formulas=formulas,
        )

    def formula_runs(self, min_len: int = 3, axis: str = "both") -> list["FormulaRun"]:
        """Stretches of adjacent formulas — a column filled down, or a row dragged across."""
        runs: list[FormulaRun] = []
        if axis in ("both", "vertical"):
            for col in range(1, self.max_col + 1):
                r = 1
                while r <= self.max_row:
                    if not self.formula(r, col):
                        r += 1
                        continue
                    cells: list[tuple[int, int, str]] = []
                    while r <= self.max_row and self.formula(r, col):
                        cells.append((r, col, self.formula(r, col)))
                        r += 1
                    if len(cells) >= min_len:
                        runs.append(FormulaRun(cells=cells, axis="vertical", sheet=self.name))
        if axis in ("both", "horizontal"):
            for row in range(1, self.max_row + 1):
                c = 1
                while c <= self.max_col:
                    if not self.formula(row, c):
                        c += 1
                        continue
                    cells = []
                    while c <= self.max_col and self.formula(row, c):
                        cells.append((row, c, self.formula(row, c)))
                        c += 1
                    if len(cells) >= min_len:
                        runs.append(FormulaRun(cells=cells, axis="horizontal", sheet=self.name))
        return runs


@dataclass
class Region:
    min_row: int
    max_row: int
    min_col: int
    max_col: int
    cell_count: int = 0

    @property
    def a1(self) -> str:
        return (
            f"{get_column_letter(self.min_col)}{self.min_row}:"
            f"{get_column_letter(self.max_col)}{self.max_row}"
        )

    @property
    def height(self) -> int:
        return self.max_row - self.min_row + 1

    @property
    def width(self) -> int:
        return self.max_col - self.min_col + 1


@dataclass
class ColumnProfile:
    index: int
    letter: str
    header: str
    numeric: int
    text: int
    dates: int
    blanks: int
    values: list[Any]
    formats: list[str]
    first_data_row: int = 1
    last_data_row: int = 1
    formulas: int = 0  # formula cells with no cached result available

    @property
    def kind(self) -> str:
        total = self.numeric + self.text + self.dates
        if total == 0:
            # Uncalculated formula cells: treat as numeric, since a computed
            # column in these courses is virtually always a number.
            return "number" if self.formulas else "empty"
        if self.dates >= total * 0.7:
            return "date"
        if self.numeric >= total * 0.8:
            return "number"
        if self.text >= total * 0.8:
            return "text"
        return "mixed"

    @property
    def distinct(self) -> int:
        try:
            return len({v for v in self.values})
        except TypeError:
            return len(self.values)

    @property
    def cardinality_ratio(self) -> float:
        return self.distinct / len(self.values) if self.values else 0.0

    @property
    def dominant_format(self) -> str:
        if not self.formats:
            return "General"
        return max(set(self.formats), key=self.formats.count)

    @property
    def format_consistent(self) -> bool:
        return len(set(self.formats)) <= 1

    @property
    def numeric_values(self) -> list[float]:
        return [float(v) for v in self.values if isinstance(v, NUMERIC) and not isinstance(v, bool)]


@dataclass
class FormulaRun:
    """A block of adjacent formulas that were (or should have been) filled together."""

    cells: list[tuple[int, int, str]]  # (row, col, formula)
    axis: str = "vertical"
    sheet: str = ""

    @property
    def start_row(self) -> int:
        return self.cells[0][0]

    @property
    def end_row(self) -> int:
        return self.cells[-1][0]

    @property
    def col(self) -> int:
        return self.cells[0][1]

    @property
    def letter(self) -> str:
        return get_column_letter(self.col)

    @property
    def a1(self) -> str:
        first, last = self.cells[0], self.cells[-1]
        return (
            f"{get_column_letter(first[1])}{first[0]}:"
            f"{get_column_letter(last[1])}{last[0]}"
        )

    @property
    def anchor(self) -> tuple[int, int, str]:
        return self.cells[0]

    def addr(self, i: int) -> str:
        r, c, _ = self.cells[i]
        return f"{get_column_letter(c)}{r}"

    def r1c1_groups(self) -> dict[str, list[str]]:
        groups: dict[str, list[str]] = {}
        for row, col, f in self.cells:
            groups.setdefault(to_r1c1(f, row, col), []).append(f"{get_column_letter(col)}{row}")
        return groups

    def step(self, i: int) -> tuple[int, int]:
        """(drow, dcol) offset of cell i from the anchor."""
        r, c, _ = self.cells[i]
        return r - self.cells[0][0], c - self.cells[0][1]


@dataclass
class Book:
    path: Path
    wb: Any
    wb_values: Any | None
    sheets: list[Sheet] = field(default_factory=list)

    def sheet(self, name: str | None = None) -> Sheet:
        if name is None:
            return self.sheets[0]
        for s in self.sheets:
            if s.name.lower() == name.lower():
                return s
        raise KeyError(f"No sheet named {name!r}. Available: {[s.name for s in self.sheets]}")

    def targets(self, name: str | None = None) -> list[Sheet]:
        return [self.sheet(name)] if name else self.sheets

    def save(self, path: Path | str | None = None) -> Path:
        out = Path(path) if path else self.path
        self.wb.save(out)
        return out


def is_formula(v: Any) -> bool:
    return isinstance(v, str) and v.startswith("=") and len(v) > 1


def load(path: str | Path, keep_values: bool = True) -> Book:
    """Load a workbook twice: once for formulas, once for Excel's cached values."""
    p = Path(path).expanduser()
    if not p.exists():
        raise FileNotFoundError(f"No such workbook: {p}")
    wb = openpyxl.load_workbook(p, data_only=False)
    wbv = None
    if keep_values:
        try:
            wbv = openpyxl.load_workbook(p, data_only=True)
        except Exception:
            wbv = None
    sheets = []
    for ws in wb.worksheets:
        vs = wbv[ws.title] if wbv is not None and ws.title in wbv.sheetnames else None
        sheets.append(
            Sheet(
                name=ws.title,
                ws=ws,
                vs=vs,
                max_row=min(ws.max_row or 1, 20000),
                max_col=min(ws.max_column or 1, 200),
            )
        )
    return Book(path=p, wb=wb, wb_values=wbv, sheets=sheets)


# --------------------------------------------------------------------------
# Header semantics — used by the number-format audit and the pivot advisor
# --------------------------------------------------------------------------

MONEY_WORDS = re.compile(
    r"\b(price|cost|salary|wage|pay|payroll|revenue|sales|amount|total|subtotal|"
    r"income|expense|budget|profit|loss|fee|charge|balance|value|net|gross|"
    r"commission|bonus|deduction|tax|paid|due|cash|spend|earnings|usd|dollars?)\b",
    re.I,
)
PERCENT_WORDS = re.compile(
    r"(%|\bpercent(age)?\b|\brate\b|\bmargin\b|\bgrowth\b|\bchange\b|\bshare\b|"
    r"\bdiscount\b|\bcommission rate\b|\byield\b|\bapr\b|\bincrease\b|\bdecrease\b)",
    re.I,
)
DATE_WORDS = re.compile(r"\b(date|day|month|year|period|when|hired|start|end|due|dob)\b", re.I)
ID_WORDS = re.compile(r"\b(id|code|number|no\.?|sku|ref|zip|postal|phone|account)\b", re.I)
COUNT_WORDS = re.compile(r"\b(qty|quantity|units|count|hours|headcount|stock|number of)\b", re.I)


def expected_format_family(header: str, profile: "ColumnProfile | None" = None) -> str | None:
    """What number format *should* this column carry? None = no strong expectation."""
    h = header or ""
    if PERCENT_WORDS.search(h):
        return "percentage"
    if MONEY_WORDS.search(h):
        return "currency"
    if DATE_WORDS.search(h) and (profile is None or profile.dates):
        return "date"
    if ID_WORDS.search(h):
        return "text_or_plain"
    if COUNT_WORDS.search(h):
        return "integer"
    if profile and profile.numeric_values:
        vals = profile.numeric_values
        if vals and all(0 < abs(v) < 1 for v in vals) and len(vals) >= 3:
            return "percentage"
    return None


def format_family(number_format: str) -> str:
    """Bucket an Excel number-format string into a family we can reason about."""
    nf = (number_format or "General").strip()
    low = nf.lower()
    if nf == "General":
        return "general"
    if "%" in nf:
        return "percentage"
    if "$" in nf or "£" in nf or "€" in nf or "[$" in nf:
        return "accounting" if "_(" in nf or "* " in nf else "currency"
    if any(t in low for t in ("yy", "mmm", "dd/", "d/m", "m/d", "hh:", "h:mm")):
        return "date"
    if "@" in nf:
        return "text"
    if re.search(r"[0#]", nf):
        return "number"
    return "other"


def decimals_in_format(number_format: str) -> int:
    m = re.search(r"\.([0#]+)", number_format or "")
    return len(m.group(1)) if m else 0
