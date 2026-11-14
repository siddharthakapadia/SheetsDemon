"""Excel formula tokenizer and reference algebra.

Everything else in sheetdoc depends on this module being correct, so it is
deliberately a real tokenizer rather than a pile of regexes applied to whole
formulas. The distinction that matters most: a "2" inside A2 is not a number,
and a "2" inside VLOOKUP(...,2,FALSE) is not a magic number either.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Iterator

# --------------------------------------------------------------------------
# Tokenizer
# --------------------------------------------------------------------------

_TOKEN_RE = re.compile(
    r"""
      (?P<string>"(?:[^"]|"")*")
    | (?P<error>\#(?:NULL!|DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|SPILL!|CALC!|GETTING_DATA))
    | (?P<sheet>(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)
    | (?P<ref>\$?[A-Za-z]{1,3}\$?[1-9][0-9]{0,6}(?![A-Za-z0-9_.]))
    | (?P<colspan>\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}(?![A-Za-z0-9_.]))
    | (?P<rowspan>\$?[1-9][0-9]{0,6}:\$?[1-9][0-9]{0,6}(?![A-Za-z0-9_.]))
    | (?P<func>[A-Za-z_][A-Za-z0-9_.]*(?=\s*\())
    | (?P<name>[A-Za-z_\\][A-Za-z0-9_.]*)
    | (?P<number>(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?)
    | (?P<op><=|>=|<>|[-+*/^&<>=%])
    | (?P<lparen>\()
    | (?P<rparen>\))
    | (?P<comma>,)
    | (?P<colon>:)
    | (?P<semi>;)
    | (?P<lbrace>\{)
    | (?P<rbrace>\})
    | (?P<lbracket>\[)
    | (?P<rbracket>\])
    | (?P<ws>\s+)
    | (?P<unknown>.)
    """,
    re.VERBOSE,
)

ARITHMETIC_OPS = {"+", "-", "*", "/", "^", "%"}
COMPARISON_OPS = {"=", "<", ">", "<=", ">=", "<>"}


@dataclass
class Token:
    kind: str
    value: str
    pos: int

    def __repr__(self) -> str:  # pragma: no cover - debugging aid
        return f"Token({self.kind}, {self.value!r})"


def tokenize(formula: str) -> list[Token]:
    """Split a formula into typed tokens. Leading '=' is stripped."""
    if formula is None:
        return []
    text = str(formula)
    if text.startswith("="):
        text = text[1:]
    tokens: list[Token] = []
    for m in _TOKEN_RE.finditer(text):
        kind = m.lastgroup or "unknown"
        if kind == "ws":
            continue
        tokens.append(Token(kind, m.group(), m.start()))
    return _merge_ranges(tokens)


def _merge_ranges(tokens: list[Token]) -> list[Token]:
    """Fuse `ref : ref` triples into a single range token, keeping sheet prefixes."""
    out: list[Token] = []
    i = 0
    while i < len(tokens):
        t = tokens[i]
        if (
            t.kind == "ref"
            and i + 2 < len(tokens)
            and tokens[i + 1].kind == "colon"
            and tokens[i + 2].kind == "ref"
        ):
            out.append(Token("range", f"{t.value}:{tokens[i + 2].value}", t.pos))
            i += 3
            continue
        out.append(t)
        i += 1
    return out


# --------------------------------------------------------------------------
# Reference algebra
# --------------------------------------------------------------------------

_REF_RE = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)([0-9]+)$")


@dataclass
class CellRef:
    col: int  # 1-based
    row: int  # 1-based
    col_abs: bool
    row_abs: bool
    sheet: str | None = None

    @property
    def a1(self) -> str:
        c = f"{'$' if self.col_abs else ''}{col_letter(self.col)}"
        r = f"{'$' if self.row_abs else ''}{self.row}"
        base = f"{c}{r}"
        return f"{self.sheet}!{base}" if self.sheet else base

    def offset(self, drow: int, dcol: int) -> "CellRef":
        """Apply a fill translation: absolute components do not move."""
        return CellRef(
            col=self.col if self.col_abs else self.col + dcol,
            row=self.row if self.row_abs else self.row + drow,
            col_abs=self.col_abs,
            row_abs=self.row_abs,
            sheet=self.sheet,
        )

    @property
    def anchor_style(self) -> str:
        if self.col_abs and self.row_abs:
            return "absolute"
        if not self.col_abs and not self.row_abs:
            return "relative"
        return "mixed"


def col_letter(idx: int) -> str:
    """1 -> A, 27 -> AA."""
    s = ""
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def col_index(letters: str) -> int:
    """A -> 1, AA -> 27."""
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def parse_ref(text: str, sheet: str | None = None) -> CellRef | None:
    """Parse a single A1-style reference, tolerating a Sheet! prefix."""
    text = text.strip()
    if "!" in text:
        sheet_part, text = text.rsplit("!", 1)
        sheet = sheet_part.strip("'")
    m = _REF_RE.match(text)
    if not m:
        return None
    dollar_c, letters, dollar_r, digits = m.groups()
    return CellRef(
        col=col_index(letters),
        row=int(digits),
        col_abs=dollar_c == "$",
        row_abs=dollar_r == "$",
        sheet=sheet,
    )


@dataclass
class RangeRef:
    start: CellRef
    end: CellRef
    sheet: str | None = None

    @property
    def a1(self) -> str:
        base = f"{self.start.a1.split('!')[-1]}:{self.end.a1.split('!')[-1]}"
        return f"{self.sheet}!{base}" if self.sheet else base

    @property
    def height(self) -> int:
        return abs(self.end.row - self.start.row) + 1

    @property
    def width(self) -> int:
        return abs(self.end.col - self.start.col) + 1

    @property
    def min_row(self) -> int:
        return min(self.start.row, self.end.row)

    @property
    def max_row(self) -> int:
        return max(self.start.row, self.end.row)

    @property
    def min_col(self) -> int:
        return min(self.start.col, self.end.col)

    @property
    def max_col(self) -> int:
        return max(self.start.col, self.end.col)

    @property
    def is_fully_absolute(self) -> bool:
        return all(
            (c.col_abs and c.row_abs) for c in (self.start, self.end)
        )

    def cells(self) -> Iterator[tuple[int, int]]:
        for r in range(self.min_row, self.max_row + 1):
            for c in range(self.min_col, self.max_col + 1):
                yield r, c


def parse_range(text: str, sheet: str | None = None) -> RangeRef | None:
    text = text.strip()
    if "!" in text:
        sheet_part, text = text.rsplit("!", 1)
        sheet = sheet_part.strip("'")
    if ":" not in text:
        single = parse_ref(text, sheet)
        return RangeRef(single, single, sheet) if single else None
    a, b = text.split(":", 1)
    start, end = parse_ref(a, sheet), parse_ref(b, sheet)
    if not start or not end:
        return None
    return RangeRef(start, end, sheet)


def extract_refs(formula: str) -> list[tuple[Token, CellRef]]:
    """All single-cell references in a formula, with their tokens."""
    out = []
    tokens = tokenize(formula)
    for i, t in enumerate(tokens):
        if t.kind == "ref":
            sheet = tokens[i - 1].value.rstrip("!").strip("'") if i and tokens[i - 1].kind == "sheet" else None
            ref = parse_ref(t.value, sheet)
            if ref:
                out.append((t, ref))
    return out


def extract_ranges(formula: str) -> list[tuple[Token, RangeRef]]:
    out = []
    tokens = tokenize(formula)
    for i, t in enumerate(tokens):
        if t.kind == "range":
            sheet = tokens[i - 1].value.rstrip("!").strip("'") if i and tokens[i - 1].kind == "sheet" else None
            rng = parse_range(t.value, sheet)
            if rng:
                out.append((t, rng))
    return out


def to_r1c1(formula: str, base_row: int, base_col: int) -> str:
    """Normalise a formula to R1C1 relative to its own cell.

    Two cells that were filled from the same source produce identical R1C1
    strings; this is how we detect a formula somebody retyped by hand.
    """
    tokens = tokenize(formula)
    parts = []
    for t in tokens:
        if t.kind in ("ref", "range"):
            pieces = []
            for piece in t.value.split(":"):
                ref = parse_ref(piece)
                if not ref:
                    pieces.append(piece)
                    continue
                r = f"R{ref.row}" if ref.row_abs else (
                    "R" if ref.row == base_row else f"R[{ref.row - base_row:+d}]"
                )
                c = f"C{ref.col}" if ref.col_abs else (
                    "C" if ref.col == base_col else f"C[{ref.col - base_col:+d}]"
                )
                pieces.append(r + c)
            parts.append(":".join(pieces))
        elif t.kind == "func":
            parts.append(t.value.upper())
        else:
            parts.append(t.value)
    return "".join(parts)


def translate(formula: str, drow: int, dcol: int) -> str:
    """Simulate dragging a formula by (drow, dcol). Absolute parts stay put."""
    tokens = tokenize(formula)
    parts = []
    for t in tokens:
        if t.kind in ("ref", "range"):
            pieces = []
            for piece in t.value.split(":"):
                ref = parse_ref(piece)
                pieces.append(ref.offset(drow, dcol).a1 if ref else piece)
            parts.append(":".join(pieces))
        else:
            parts.append(t.value)
    prefix = "=" if str(formula).startswith("=") else ""
    return prefix + "".join(parts)


# --------------------------------------------------------------------------
# Call-tree extraction
# --------------------------------------------------------------------------


@dataclass
class FunctionCall:
    name: str
    args: list[str]
    arg_tokens: list[list[Token]] = field(default_factory=list)
    start: int = 0
    raw: str = ""

    def arg(self, i: int) -> str | None:
        return self.args[i].strip() if i < len(self.args) else None


def find_calls(formula: str) -> list[FunctionCall]:
    """Every function call in a formula, outermost first, nested included."""
    tokens = tokenize(formula)
    calls: list[FunctionCall] = []
    for i, t in enumerate(tokens):
        if t.kind != "func":
            continue
        if i + 1 >= len(tokens) or tokens[i + 1].kind != "lparen":
            continue
        depth = 0
        arg_tokens: list[list[Token]] = [[]]
        j = i + 1
        closed = False
        while j < len(tokens):
            tok = tokens[j]
            if tok.kind in ("lparen", "lbrace", "lbracket"):
                depth += 1
                if depth > 1:
                    arg_tokens[-1].append(tok)
            elif tok.kind in ("rparen", "rbrace", "rbracket"):
                depth -= 1
                if depth == 0:
                    closed = True
                    break
                arg_tokens[-1].append(tok)
            elif tok.kind in ("comma", "semi") and depth == 1:
                arg_tokens.append([])
            else:
                arg_tokens[-1].append(tok)
            j += 1
        if not closed:
            continue
        args = ["".join(tk.value for tk in group) for group in arg_tokens]
        if len(args) == 1 and args[0] == "":
            args, arg_tokens = [], []
        calls.append(
            FunctionCall(
                name=t.value.upper(),
                args=args,
                arg_tokens=arg_tokens,
                start=t.pos,
                raw=t.value,
            )
        )
    return calls


# --------------------------------------------------------------------------
# Structural-argument whitelist
# --------------------------------------------------------------------------

# (function, 0-based arg index) pairs where a bare integer is structural syntax,
# not a magic number a student should have put in a cell.
STRUCTURAL_ARGS: dict[str, set[int]] = {
    "VLOOKUP": {2, 3},
    "HLOOKUP": {2, 3},
    "XLOOKUP": {4, 5},
    "XMATCH": {2, 3},
    "INDEX": {1, 2, 3},
    "MATCH": {2},
    "ROUND": {1},
    "ROUNDUP": {1},
    "ROUNDDOWN": {1},
    "MROUND": {1},
    "TRUNC": {1},
    "LEFT": {1},
    "RIGHT": {1},
    "MID": {1, 2},
    "OFFSET": {1, 2, 3, 4},
    "LARGE": {1},
    "SMALL": {1},
    "RANK": {2},
    "RANK.EQ": {2},
    "RANK.AVG": {2},
    "WEEKDAY": {1},
    "TEXTSPLIT": {2, 3, 4},
    "SORT": {1, 2, 3},
    "SORTBY": {2, 4, 6},
    "FILTER": set(),
    "SUBTOTAL": {0},
    "CHOOSE": {0},
    "DATE": {0, 1, 2},
    "TIME": {0, 1, 2},
    "EOMONTH": {1},
    "EDATE": {1},
    "WORKDAY": {1},
    "NETWORKDAYS": set(),
    "PMT": set(),
    "FV": set(),
    "PV": set(),
    "NPER": set(),
    "RATE": set(),
    "IPMT": set(),
    "PPMT": set(),
    "REPT": {1},
    "TEXTJOIN": set(),
    "SEQUENCE": {0, 1, 2, 3},
    "RANDBETWEEN": {0, 1},
    "COLUMNS": set(),
    "ROWS": set(),
    "INDIRECT": {1},
    "CELL": set(),
    "ERROR.TYPE": set(),
}

# Numbers that are idiomatic rather than magic, e.g. =1-B2 or =A1*-1.
BENIGN_LITERALS = {0.0, 1.0, -1.0}

LOOKUP_FUNCS = {"VLOOKUP", "HLOOKUP", "XLOOKUP", "LOOKUP", "MATCH", "XMATCH", "INDEX"}
AGGREGATE_FUNCS = {
    "SUM", "AVERAGE", "COUNT", "COUNTA", "MAX", "MIN", "MEDIAN", "STDEV",
    "STDEV.S", "STDEV.P", "STDEVP", "VAR", "VAR.S", "VAR.P", "PRODUCT",
    "AVERAGEA", "COUNTBLANK", "SUBTOTAL", "AGGREGATE",
}
CONDITIONAL_AGGREGATES = {
    "SUMIF", "SUMIFS", "COUNTIF", "COUNTIFS", "AVERAGEIF", "AVERAGEIFS",
    "MAXIFS", "MINIFS",
}


def classify_number_context(tokens: list[Token], idx: int) -> str:
    """Why does this numeric literal exist? -> arithmetic | comparison | argument | exponent."""
    prev = next((tokens[i] for i in range(idx - 1, -1, -1)), None)
    nxt = tokens[idx + 1] if idx + 1 < len(tokens) else None

    # An exponent is structural maths, not a business parameter: the 2 in
    # =weight/(height^2) defines BMI. Telling a student to move it into a
    # labelled cell is actively wrong advice.
    if prev and prev.kind == "op" and prev.value == "^":
        return "exponent"
    for neighbour in (prev, nxt):
        if neighbour and neighbour.kind == "op" and neighbour.value in ARITHMETIC_OPS:
            return "arithmetic"
    for neighbour in (prev, nxt):
        if neighbour and neighbour.kind == "op" and neighbour.value in COMPARISON_OPS:
            return "comparison"
    return "argument"


def enclosing_arg(formula: str, token: Token) -> tuple[str, int] | None:
    """Which (function, arg index) does this token sit directly inside?"""
    best: tuple[str, int] | None = None
    best_start = -1
    for call in find_calls(formula):
        for i, group in enumerate(call.arg_tokens):
            if any(t.pos == token.pos and t.value == token.value for t in group):
                if call.start > best_start:
                    best_start = call.start
                    best = (call.name, i)
    return best


def replace_reference(formula: str, old_addr: str, new_addr: str) -> str:
    """Swap one cell reference for another inside a formula.

    Token-aware rather than string replacement: a naive replace of "B15" would
    also corrupt "B150", and would rewrite text inside string literals.
    """
    target = parse_ref(old_addr)
    if not target:
        return formula
    out: list[str] = []
    for t in tokenize(formula):
        if t.kind == "ref":
            ref = parse_ref(t.value)
            if ref and ref.row == target.row and ref.col == target.col and not ref.sheet:
                out.append(new_addr)
                continue
        elif t.kind == "range":
            pieces = []
            changed = False
            for piece in t.value.split(":"):
                ref = parse_ref(piece)
                if ref and ref.row == target.row and ref.col == target.col:
                    pieces.append(new_addr)
                    changed = True
                else:
                    pieces.append(piece)
            out.append(":".join(pieces) if changed else t.value)
            continue
        out.append(t.value)
    prefix = "=" if str(formula).startswith("=") else ""
    return prefix + "".join(out)


def expand_span(span: str) -> list[str]:
    """'D2:D5' -> ['D2','D3','D4','D5'];  'C12' -> ['C12']."""
    rng = parse_range(span)
    if not rng:
        return []
    return [f"{col_letter(c)}{r}" for r, c in rng.cells()]
