"""Offline mode — the chatbot working with no model provider at all.

The engine answers spreadsheet questions deterministically; a language model was
only ever choosing *which* engine call to make and wording the reply. That
routing can be done with patterns, so the chat box stays useful with no API key,
no signup and no network.

It is genuinely less capable than a model: it matches intent rather than
understanding, and it will not chain several steps from a vague request. It says
so rather than pretending otherwise. When a provider key is configured, the real
agent takes over automatically.
"""

from __future__ import annotations

import json
import re
from collections.abc import Callable

from . import tools as toolmod
from sheetdoc.formula import expand_span, parse_range, replace_reference

CELL_RE = re.compile(r"\b([A-Za-z]{1,3}\$?[0-9]{1,5})\b(?!\s*[:!])")
RANGE_RE = re.compile(r"\b([A-Za-z]{1,3}[0-9]{1,5}:[A-Za-z]{1,3}[0-9]{1,5})\b")
ERROR_RE = re.compile(r"(#N/A|#VALUE!?|#REF!?|#DIV/0!?|#NAME\??|#NUM!?|#NULL!?|#SPILL!?|#{3,})",
                      re.I)

PRACTICE_TOPICS = {
    "vlookup": "vlookup", "xlookup": "xlookup", "index": "index_match",
    "match": "index_match", "absolute": "absolute", "anchor": "absolute",
    "dollar": "absolute", "$": "absolute", "sumif": "sumif", "countif": "sumif",
    "nested if": "if_nested", "if": "if_nested", "percent": "percentage",
    "payroll": "payroll", "salary": "payroll", "grade": "gradebook",
    "weighted": "gradebook", "budget": "budget", "variance": "budget",
}

DATASET_KINDS = ("sales", "payroll", "grades", "inventory", "expenses", "budget")


# Set when Excel holds a workbook that has never been written to disk.
UNSAVED_NOTE = (
    "The workbook open in Excel (**{name}**) has never been saved, so there is no "
    "file for me to read. Press ⌘S in Excel once — after that I can audit it, chart "
    "it and fix it."
)


def _live_file() -> str | None:
    """The workbook open in Excel, saved so the engine can read it.

    A workbook that has never been saved reports a plausible-looking path that
    does not exist on disk (Excel's default save location). Reading it would fail
    with a confusing "no such workbook", so treat it as no file at all.
    """
    from pathlib import Path

    from . import excel_live as X

    try:
        path = X.active_workbook_path()
        if not path:
            return None
        X.save_active()
        return path if Path(path).exists() else None
    except X.ExcelError:
        return None


def _unsaved_workbook_name() -> str | None:
    """Name of the open workbook if it exists only in memory, else None."""
    from pathlib import Path

    from . import excel_live as X

    try:
        st = X.state()
        if not (st.running and st.active_workbook):
            return None
        path = X.active_workbook_path()
        return st.active_workbook if not path or not Path(path).exists() else None
    except Exception:
        return None


def _target(emit: Callable[[str, dict], None]) -> tuple[str | None, str]:
    """Pick the workbook to act on: what's open in Excel, else the newest upload.

    Returns (file_argument_for_tools, human_label).
    """
    live = _live_file()
    if live:
        from pathlib import Path

        # Point the tool layer at wherever that file actually lives.
        toolmod.WORKSPACE = Path(live).parent
        return Path(live).name, f"{Path(live).name} (open in Excel)"

    # Excel is showing a workbook that was never saved. Falling back to some
    # other file on disk would answer confidently about a workbook the student
    # is not looking at, which is worse than saying nothing.
    if _unsaved_workbook_name():
        return None, ""

    out, err = toolmod.run_tool("list_workbooks", {})
    if err:
        return None, ""
    books = json.loads(out).get("workbooks") or []
    if not books:
        return None, ""
    return books[-1], books[-1]


def _no_target_message(default: str) -> str:
    """Why there is nothing to work on — the unsaved case needs its own advice."""
    name = _unsaved_workbook_name()
    return UNSAVED_NOTE.format(name=name) if name else default


def _findings_text(raw: str, limit: int = 6) -> str:
    """Turn a tool's JSON report into something a person reads."""
    try:
        d = json.loads(raw)
    except json.JSONDecodeError:
        return raw[:600]

    if d.get("error"):
        return str(d["error"])

    out: list[str] = []
    counts = d.get("counts") or {}
    errs, warns = counts.get("error", 0), counts.get("warning", 0)

    findings = [f for f in (d.get("findings") or []) if f.get("severity") != "pass"]
    if not findings:
        return "Nothing flagged — that all looks correct."

    if errs or warns:
        bits = []
        if errs:
            bits.append(f"{errs} problem{'s' if errs != 1 else ''}")
        if warns:
            bits.append(f"{warns} thing{'s' if warns != 1 else ''} worth tightening")
        out.append("Found " + " and ".join(bits) + ".")

    for f in findings[:limit]:
        mark = {"error": "✗", "warning": "▲"}.get(f.get("severity"), "·")
        cell = f.get("cell")
        line = f"\n{mark} " + (f"**{cell}** — " if cell else "") + f.get("message", "")
        if f.get("affected_cells"):
            line += f" ({f['affected_cells']} cells)"
        if f.get("formula"):
            line += f"\n   `{f['formula']}`"
        if f.get("fix"):
            line += f"\n   → {f['fix']}"
        out.append(line)

    if len(findings) > limit:
        out.append(f"\n…and {len(findings) - limit} more.")
    return "\n".join(out)


def _hint_text(raw: str) -> str:
    try:
        d = json.loads(raw)
    except json.JSONDecodeError:
        return raw[:600]
    summary = d.get("summary") or {}
    out = []
    if summary.get("concept"):
        out.append(f"**{summary['concept']}**")
    if summary.get("formula") and summary["formula"] != "(none)":
        out.append(f"`{summary['formula']}`")
    for f in d.get("findings") or []:
        msg, detail = f.get("message", ""), f.get("detail", "")
        if msg.startswith("Tier") or msg.startswith("How to"):
            out.append(f"\n**{msg}**\n{detail}")
    for n in d.get("notes") or []:
        out.append(f"\n_{n}_")
    return "\n".join(out) if out else "No hint available for that."



# ==========================================================================
# Editing the live sheet
# ==========================================================================

SET_RE = re.compile(
    r"\b(?:set|change|make|put|write|enter|update)\s+(?:cell\s+)?"
    r"([A-Za-z]{1,3}[0-9]{1,5})\s*(?:to|=|as)\s*(.+)$", re.I)
ASSIGN_RE = re.compile(r"^\s*([A-Za-z]{1,3}[0-9]{1,5})\s*=\s*(.+)$")

FORMAT_WORDS = {
    "currency": "$#,##0.00", "dollar": "$#,##0.00", "money": "$#,##0.00",
    "accounting": '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
    "percent": "0.0%", "percentage": "0.0%",
    "date": "mm/dd/yyyy", "integer": "#,##0", "whole number": "#,##0",
    "number": "#,##0.00", "two decimal": "#,##0.00", "text": "@",
}
FORMAT_RE = re.compile(
    r"format\s+([A-Za-z]{1,3}[0-9]{1,5}(?::[A-Za-z]{1,3}[0-9]{1,5})?)\s+as\s+(.+)$", re.I)


def _fix_plan(file: str) -> tuple[list[tuple[str, str, str]], list[str]]:
    """Work out which cells the audit's own suggestions would change.

    Returns (changes, skipped) where a change is (cell, old_formula, new_formula).
    """
    from . import excel_live as X

    out, err = toolmod.run_tool("audit_formulas", {"file": file})
    if err:
        return [], [out]

    changes: list[tuple[str, str, str]] = []
    skipped: list[str] = []

    for f in json.loads(out).get("findings") or []:
        data = f.get("data") or {}
        span = f.get("cell")
        if not span:
            continue

        ref, suggested = data.get("reference"), data.get("suggested")
        missing = data.get("missing")

        if ref and suggested:
            for cell in expand_span(span):
                try:
                    live = X.read_cell(cell)
                except X.ExcelError:
                    continue
                formula = live.get("formula") or ""
                if not formula.startswith("="):
                    continue
                new = replace_reference(formula, ref, suggested)
                if new != formula:
                    changes.append((cell, formula, new))

        elif missing:
            # Extend an aggregate range to cover the values it missed.
            cell = expand_span(span)[0] if expand_span(span) else None
            if not cell:
                continue
            try:
                formula = (X.read_cell(cell).get("formula") or "")
            except X.ExcelError:
                continue
            new = formula
            for _, rng in __import__("sheetdoc.formula", fromlist=["x"]).extract_ranges(formula):
                last = missing[-1]
                target = parse_range(f"{rng.a1.split(':')[0]}:{last}")
                if target:
                    new = new.replace(rng.a1, target.a1, 1)
                    break
            if new != formula:
                changes.append((cell, formula, new))
        else:
            note = f.get("message", "")
            # One conceptual problem can appear on several cells; the user wants
            # the distinct list, not the same sentence four times.
            if note and f.get("severity") == "error" and note not in skipped:
                skipped.append(note)

    return changes, skipped


def _apply_changes(changes: list[tuple[str, str, str]]) -> int:
    from . import excel_live as X

    if not changes:
        return 0
    X.write_cells({cell: new for cell, _, new in changes})
    return len(changes)


# ==========================================================================
# Intent handling
# ==========================================================================


# --- table / new-workbook intents ---------------------------------------
NEW_BOOK_RE = re.compile(
    r"\b(?:create|make|start|open|give me|build|need)\b[^.?!]*?\b"
    r"(?:new|fresh|blank|empty)\b[^.?!]*?\b(?:spreadsheet|workbook|excel file|xlsx|file)\b",
    re.I)
TABLE_RE = re.compile(
    r"\b(?:create|make|build|add|set ?up|insert|give me|need|want)\b[^.?!]*?\b"
    r"(?:data\s+)?(?:table|tracker|log sheet|column headers?)\b", re.I)
# "pivot table" and "table of contents" are other intents that happen to say "table".
NOT_A_TABLE_RE = re.compile(r"\bpivot\b|\btable of contents\b", re.I)
# "with columns A, B, C" / "headers: A, B, C" / "for A, B and C"
COLS_RE = re.compile(
    r"\b(?:columns?|headers?|headings?|fields?)"
    r"(?:\s+(?:headers?|headings?|names?|titles?))?\b"
    r"\s*(?:called|named|for|of|are|=)?\s*:?\s*(.+)",
    re.I)
NAMED_RE = re.compile(
    r"\b(?:called|named|name it|titled)\s+['\"]?([\w\-. ]{1,40}?)['\"]?"
    r"\s*(?=\b(?:with|that|which|for|containing|and)\b|[.,;!?]|$)", re.I)
ROWS_RE = re.compile(r"\b(\d{1,3})\s+(?:sample|example|dummy|blank|empty)?\s*rows?\b", re.I)

# Words that trail a column list but are not columns themselves.
_TAIL_JUNK = re.compile(
    r"^(?:please|thanks|thank you|in excel|on my sheet|for me|too|as well)\b", re.I)


def _split_columns(blob: str) -> list[str]:
    """Pull column names out of free text: 'Name, Score and Grade please'."""
    blob = re.split(r"[.?!\n]", blob)[0]
    blob = re.sub(r"\b(?:and|&)\b", ",", blob, flags=re.I)
    out = []
    for raw in blob.split(","):
        c = raw.strip().strip("'\"").strip()
        if not c or _TAIL_JUNK.match(c):
            continue
        if len(c) > 40 or len(c.split()) > 4:  # a sentence, not a heading
            continue
        if re.fullmatch(r"(?:\d+\s*)?(?:blank |empty |sample )?rows?", c, re.I):
            continue
        out.append(c[:1].upper() + c[1:])
    return out


def _columns_from(text: str) -> list[str]:
    # "with columns A, B, C and 5 blank rows" — the row count is an instruction,
    # not a heading, so take it out before splitting on commas.
    text = ROWS_RE.sub(" ", text)
    m = COLS_RE.search(text)
    if m:
        return _split_columns(m.group(1))
    # "a table of Date, Item, Cost" — no explicit "columns" word.
    m = re.search(r"\b(?:table|tracker)\b[^,]*?\b(?:of|with|for)\s+(.+)", text, re.I)
    return _split_columns(m.group(1)) if m and "," in m.group(1) else []


def handle(message: str, emit: Callable[[str, dict], None]) -> None:
    """Route one message to the engine. Emits the same events as the agent."""
    text = message.strip()
    low = text.lower()

    def tool(name: str, args: dict) -> str:
        emit("tool_start", {"name": name, "input": args})
        out, err = toolmod.run_tool(name, args)
        emit("tool_end", {"name": name, "error": err})
        return out if not err else f"__ERR__{out}"

    def say(msg: str) -> None:
        emit("text", {"text": msg})

    # ---- new spreadsheet ---------------------------------------------------
    if NEW_BOOK_RE.search(text):
        cols = _columns_from(text)
        nm = NAMED_RE.search(text)
        args = {}
        if nm:
            args["name"] = nm.group(1).strip()
        if cols:
            args["headers"] = cols
        out = tool("new_spreadsheet", args)
        if out.startswith("__ERR__"):
            say("Could not create the workbook: " + out[7:])
            return
        d = json.loads(out)
        msg = f"Created **{d['workbook']}** and opened it in Excel."
        if cols:
            msg += (" Row 1 has your headings: "
                    + ", ".join(f"**{c}**" for c in cols) + ".")
        msg += f"\n\nSaved to `{d['saved_to']}`."
        say(msg)
        return

    # ---- data table --------------------------------------------------------
    if TABLE_RE.search(text) and not NOT_A_TABLE_RE.search(text):
        cols = _columns_from(text)
        if not cols:
            say("I can build that table — tell me the column headings and I will "
                "write them in and format it as a real Excel Table. For example: "
                "*make a table with columns Date, Item, Quantity, Price*.")
            return
        rm = ROWS_RE.search(text)
        blanks = [[""] * len(cols) for _ in range(min(int(rm.group(1)), 50))] if rm else []
        m = CELL_RE.search(text)
        at = m.group(1).upper() if m and "at" in low else "A1"
        out = tool("create_table_live", {"headers": cols, "rows": blanks, "at": at})
        if out.startswith("__ERR__"):
            say("Could not build the table: " + out[7:]
                + "\n\nIf Excel is not open, say *create a new spreadsheet* first.")
            return
        d = json.loads(out)
        msg = (f"Built a {len(cols)}-column table at **{d['range']}** on "
               f"**{d['sheet']}**: " + ", ".join(f"**{c}**" for c in cols) + ".")
        if d.get("excel_table"):
            msg += (f" It is a real Excel Table (`{d['excel_table']}`), so you get "
                    "filter buttons on each heading and it grows automatically when "
                    "you type in the row underneath.")
        say(msg)
        return

    # ---- direct cell edits -------------------------------------------------
    m = SET_RE.search(text) or ASSIGN_RE.match(text)
    if m:
        cell, value = m.group(1).upper(), m.group(2).strip().strip('"').strip("'")
        emit("tool_start", {"name": "write_excel_live", "input": {cell: value}})
        try:
            from . import excel_live as X

            before = ""
            try:
                before = X.read_cell(cell).get("formula") or ""
            except X.ExcelError:
                pass
            X.write_cells({cell: value})
            after = X.read_cell(cell)
            emit("tool_end", {"name": "write_excel_live", "error": False})
            shown = after.get("formula") or after.get("value") or value
            msg = f"Set **{cell}** to `{shown}`."
            if after.get("formula", "").startswith("=") and after.get("value"):
                msg += f" Excel calculates it as **{after['value']}**."
            if before and before != shown:
                msg += f"\n\nIt was `{before}` — press ⌘Z in Excel to undo."
            say(msg)
        except Exception as e:
            emit("tool_end", {"name": "write_excel_live", "error": True})
            say(f"Could not write to {cell}: {e}")
        return

    # ---- number formatting -------------------------------------------------
    m = FORMAT_RE.search(text)
    if m:
        rng, want = m.group(1).upper(), m.group(2).strip().lower()
        code = next((v for k, v in FORMAT_WORDS.items() if k in want), None)
        if not code:
            say("I know these formats: " + ", ".join(sorted(set(FORMAT_WORDS))) + ".")
            return
        emit("tool_start", {"name": "format_excel_live", "input": {"range": rng}})
        try:
            from . import excel_live as X

            X.set_number_format(rng, code)
            emit("tool_end", {"name": "format_excel_live", "error": False})
            say(f"Formatted **{rng}** as {want} (`{code}`). ⌘Z in Excel undoes it.")
        except Exception as e:
            emit("tool_end", {"name": "format_excel_live", "error": True})
            say(f"Could not format {rng}: {e}")
        return

    # ---- auto-fix from the audit's own suggestions --------------------------
    wants_fix = any(w in low for w in (
        "fix it", "fix them", "fix the", "fix my", "apply the fix", "apply fixes",
        "apply the changes", "correct it", "correct the", "repair", "make the changes",
        "make those changes", "do it for me", "just fix",
    ))
    wants_preview = any(w in low for w in (
        "what would you", "what would change", "preview", "show me the fix",
        "what will you", "dry run",
    ))
    if wants_fix or wants_preview:
        file, label = _target(emit)
        if not file:
            say(_no_target_message("Open a workbook in Excel first, then ask me to fix it."))
            return
        from . import excel_live as X

        if not X.is_running() or X.state().error:
            say(_no_target_message("I can only apply fixes to a workbook open in Excel. Open it and ask again."))
            return

        emit("tool_start", {"name": "audit_live_workbook", "input": {"file": file}})
        changes, skipped = _fix_plan(file)
        emit("tool_end", {"name": "audit_live_workbook", "error": False})

        if not changes:
            note = ("\n\nThese need a judgement call I can't make automatically:\n"
                    + "\n".join(f"• {s}" for s in skipped[:4])) if skipped else ""
            say("Nothing I can fix automatically." + note)
            return

        preview = "\n".join(
            f"• **{c}**  `{old}` → `{new}`" for c, old, new in changes[:10])
        more = f"\n…and {len(changes) - 10} more." if len(changes) > 10 else ""

        if wants_preview and not wants_fix:
            say(f"I would change {len(changes)} cell(s) in {label}:\n\n{preview}{more}"
                "\n\nSay **fix it** and I'll apply them.")
            return

        emit("tool_start", {"name": "write_excel_live",
                            "input": {"cells": len(changes)}})
        try:
            n = _apply_changes(changes)
            emit("tool_end", {"name": "write_excel_live", "error": False})
        except Exception as e:
            emit("tool_end", {"name": "write_excel_live", "error": True})
            say(f"Could not apply the changes: {e}")
            return

        tail = ""
        if skipped:
            tail = ("\n\nStill needs your judgement:\n"
                    + "\n".join(f"• {s}" for s in skipped[:3]))
        say(f"Fixed **{n} cell(s)** in {label}:\n\n{preview}{more}"
            f"\n\nThe change is in Excel now — ⌘Z there undoes it.{tail}")
        return

    # ---- error codes ------------------------------------------------------
    m = ERROR_RE.search(text)
    if m and any(w in low for w in ("error", "mean", "why", "what", "getting", "showing", "#")):
        out = tool("explain_error", {"code": m.group(1).upper()})
        if out.startswith("__ERR__"):
            say(out[7:])
            return
        d = json.loads(out)
        parts = [f"**{d.get('title', m.group(1))}**"]
        for f in d.get("findings") or []:
            code, detail = f.get("code"), f.get("detail", "")
            if code == "ERR_PLAIN":
                parts.append("\n" + detail)
            elif code == "ERR_CAUSE":
                parts.append(f"• {detail}")
            elif code == "ERR_DEBUG":
                parts.append(f"→ {detail}")
        say("\n".join(parts[:9]))
        return

    # ---- practice ---------------------------------------------------------
    if any(w in low for w in ("practice", "exercise", "drill", "another problem",
                              "practise", "test me")):
        topic = next((v for k, v in PRACTICE_TOPICS.items() if k in low), "absolute")
        name = f"practice_{topic}.xlsx"
        out = tool("create_practice", {"topic": topic, "filename": name})
        if out.startswith("__ERR__"):
            say(out[7:])
            return
        d = json.loads(out)
        tasks = [f.get("detail", "") for f in (d.get("findings") or [])
                 if f.get("code") == "TASK"]
        say(f"Made **{name}** — {(d.get('summary') or {}).get('skill', topic)}.\n\n"
            + "\n".join(f"{i}. {t}" for i, t in enumerate(tasks[:5], 1))
            + "\n\nIt's in your files list; the answer key is on a hidden sheet.")
        return

    # ---- mock data --------------------------------------------------------
    if any(w in low for w in ("dataset", "sample data", "mock", "fake data", "test data")):
        kind = next((k for k in DATASET_KINDS if k in low), "sales")
        name = f"{kind}_data.xlsx"
        out = tool("create_dataset", {"kind": kind, "filename": name, "rows": 20})
        say(f"Created **{name}** with 20 rows of {kind} data."
            if not out.startswith("__ERR__") else out[7:])
        return

    # ---- everything below wants a workbook --------------------------------
    file, label = _target(emit)

    needs_book = any(w in low for w in (
        "audit", "check", "wrong", "mistake", "error", "review", "grade", "fix",
        "hint", "stuck", "pivot", "chart", "graph", "format", "look", "my sheet",
        "my workbook", "homework", "this", "inspect", "structure",
    ))
    if not file and needs_book:
        say(_no_target_message(
            "I can't see a workbook. Open one in Excel, or upload one with the "
            "button on the left, then ask again."))
        return

    # ---- hint on a specific cell -----------------------------------------
    if any(w in low for w in ("hint", "stuck", "help me understand", "don't understand",
                              "dont understand", "explain this cell", "why is")):
        cell = CELL_RE.search(text)
        args: dict = {"file": file} if file else {}
        if cell:
            args["cell"] = cell.group(1).upper().replace("$", "")
        else:
            topic = next((k for k in PRACTICE_TOPICS if k in low), None)
            if topic:
                args["topic"] = topic
        if "tier 1" in low or "first hint" in low:
            args["tier"] = 1
        out = tool("socratic_hint", args)
        say(_hint_text(out) if not out.startswith("__ERR__") else out[7:])
        return

    # ---- pivot ------------------------------------------------------------
    if "pivot" in low:
        out = tool("explain_pivot", {"file": file})
        if out.startswith("__ERR__"):
            say(out[7:]); return
        d = json.loads(out)
        lines = [f"**Pivot table for {label}**"]
        for f in d.get("findings") or []:
            code = f.get("code", "")
            if code.startswith("PIVOT_") and "STEP" not in code and "EMPTY" not in code:
                lines.append(f"• {f.get('message')}")
        steps = [f.get("detail") for f in (d.get("findings") or [])
                 if f.get("code") == "PIVOT_STEP"]
        if steps:
            lines.append("\n**Steps**")
            lines += [f"{i}. {s}" for i, s in enumerate(steps[:6], 1)]
        say("\n".join(lines))
        return

    # ---- charts: recommend, and actually build ------------------------------
    if any(w in low for w in ("chart", "graph", "plot", "visual")):
        file, label = _target(emit)
        if not file:
            say(_no_target_message("Open a workbook in Excel or upload one, then ask for a chart."))
            return

        wants_build = any(w in low for w in (
            "make", "create", "build", "add", "draw", "insert", "give me", "put"))
        asked_type = next((k for k in ("column", "bar", "line", "pie", "scatter",
                                       "area", "doughnut", "donut")
                           if k in low), None)

        out = tool("recommend_chart", {"file": file})
        if out.startswith("__ERR__"):
            say(out[7:])
            return
        d = json.loads(out)
        recs = [f for f in (d.get("findings") or []) if f.get("code") == "CHART_REC"]
        summary = d.get("summary") or {}

        if not recs:
            say("There isn't enough structure in that sheet to chart — I need at "
                "least one label column and one numeric column.")
            return

        best = recs[0]
        best_type = None
        for k in ("column", "bar", "line", "pie", "scatter", "area"):
            if k in best.get("message", "").lower():
                best_type = k
                break

        if not wants_build:
            lines = [f"**{best.get('message')}**", best.get("detail", "")]
            if best.get("fix"):
                lines.append(f"\n{best['fix']}")
            if len(recs) > 1:
                lines.append(f"\n_Alternative: {recs[1].get('message')}_")
            lines.append("\nSay **make that chart** and I'll build it in Excel.")
            say("\n".join(lines))
            return

        # Build it. Prefer the type they named; otherwise the analysed best fit.
        chart_type = asked_type or best_type or "column"
        if chart_type == "donut":
            chart_type = "doughnut"

        from . import excel_live as X

        if not X.is_running() or X.state().error:
            say(_no_target_message("I can build charts into a workbook open in Excel. Open it and ask again."))
            return

        rng = (RANGE_RE.search(text).group(1) if RANGE_RE.search(text)
               else (X.state().used_range or ""))
        if not rng:
            say("I couldn't work out which range to chart. Try: make a column chart of A1:C10")
            return

        title = summary.get("range", "").split("!")[-1]
        nice = {"column": "Column", "bar": "Bar", "line": "Line", "pie": "Pie",
                "scatter": "Scatter", "area": "Area", "doughnut": "Doughnut"}[chart_type]

        emit("tool_start", {"name": "create_chart_live",
                            "input": {"range": rng, "chart_type": chart_type}})
        try:
            X.add_chart(rng, chart_type, None)
            emit("tool_end", {"name": "create_chart_live", "error": False})
        except Exception as e:
            emit("tool_end", {"name": "create_chart_live", "error": True})
            say(f"Could not create the chart: {e}")
            return

        why = ("you asked for it" if asked_type
               else best.get("detail", "").split(".")[0].strip() or "it fits the data shape")
        say(f"Built a **{nice.lower()} chart** from `{rng}` — it's on your sheet in Excel now.\n\n"
            f"I chose {nice.lower()} because {why[0].lower() + why[1:] if why else 'it fits'}."
            + ("" if asked_type else
               f"\n\n_Want something else? Say e.g. \"make a pie chart of {rng}\"._"))
        return

    # ---- formatting -------------------------------------------------------
    if any(w in low for w in ("format", "currency", "percentage", "decimal",
                              "bold", "header", "align", "look")):
        out = tool("check_formatting", {"file": file})
        say(f"**Formatting — {label}**\n\n" + _findings_text(out)
            if not out.startswith("__ERR__") else out[7:])
        return

    # ---- structure --------------------------------------------------------
    if any(w in low for w in ("inspect", "structure", "what's in", "whats in",
                              "show me", "describe")):
        out = tool("inspect_workbook", {"file": file})
        say(f"**{label}**\n\n" + _findings_text(out, limit=10)
            if not out.startswith("__ERR__") else out[7:])
        return

    # ---- audit (the default for anything check-shaped) --------------------
    if needs_book or not text:
        checks = ("hardcoded" if "hardcod" in low else
                  "references" if any(w in low for w in ("absolute", "$", "anchor", "reference")) else
                  "ranges" if "range" in low else
                  "lookups" if any(w in low for w in ("vlookup", "xlookup", "lookup", "index")) else
                  "all")
        out = tool("audit_formulas", {"file": file, "checks": checks})
        say(f"**{label}**\n\n" + _findings_text(out)
            if not out.startswith("__ERR__") else out[7:])
        return

    # ---- nothing matched --------------------------------------------------
    # Without a model there is no understanding to fall back on, but there is
    # still real state to report. Saying something true about the sheet in front
    # of the student beats printing the same menu at every unmatched message.
    say(_grounded_fallback())
    return


def _grounded_fallback() -> str:
    """What we can say about the actual workbook, with no model available."""
    from . import excel_live as X

    lines: list[str] = []
    try:
        st = X.state()
    except Exception:
        st = None

    if st and st.running and st.active_workbook:
        where = f"**{st.active_workbook}**"
        if st.active_sheet:
            where += f" · sheet *{st.active_sheet}*"
        lines.append(f"I could not match that to something I can do, but here is where you are: {where}.")
        if st.used_range:
            lines.append(f"Data occupies `{st.used_range}`.")
        if st.selection:
            sel = st.selection
            try:
                cell = X.read_cell(sel.split(":")[0])
                formula, value = cell.get("formula") or "", cell.get("value")
                if formula.startswith("="):
                    lines.append(
                        f"Your selection **{sel}** holds `{formula}`"
                        + (f", which Excel calculates as **{value}**." if value else "."))
                elif value not in (None, ""):
                    lines.append(f"Your selection **{sel}** holds **{value}**.")
                else:
                    lines.append(f"Your selection **{sel}** is empty.")
            except Exception:
                lines.append(f"Your selection is **{sel}**.")
        if st.unsaved:
            lines.append("It has unsaved changes.")
    elif st and st.running:
        lines.append("Excel is open but has no workbook — open one, or ask me to "
                     "*create a new spreadsheet*.")
    else:
        lines.append("I could not match that to something I can do, and Excel does "
                     "not appear to be open.")

    lines.append(
        "\nI am running without a language model, so I match on keywords rather than "
        "understanding freely. What I can do right now:\n\n"
        "• **check / audit my work** — full formula audit\n"
        "• **fix it** — apply the audit's own corrections\n"
        "• **hint for D5** — a graded hint for one cell\n"
        "• **make a data table with columns Date, Item, Price**\n"
        "• **create a new spreadsheet called budget**\n"
        "• **make a chart of my data**\n"
        "• **B7 = SUM(B2:B6)** — write straight into a cell\n"
        "• **what does #N/A mean** — error explanations\n"
        "• **practice vlookup** — generate a practice sheet\n\n"
        "For genuine conversation, add a provider key with "
        f"`{__import__('chatbot.config', fromlist=['x']).setup_cmd()}`."
    )
    return "\n".join(lines)


def _unused_menu() -> str:
    return (
        "I'm running without a language model, so I match on keywords rather than "
        "understanding freely. Things I can do right now:\n\n"
        "• **check / audit my work** — full formula audit\n"
        "• **hint for D5** — a graded hint for one cell\n"
        "• **what does #N/A mean** — error explanations\n"
        "• **what chart should I use** — chart recommendation\n"
        "• **pivot table** — how to set one up for your data\n"
        "• **check formatting** — number formats, headers, alignment\n"
        "• **practice vlookup** — generate a practice sheet\n"
        "• **make a sales dataset** — sample data\n\n"
        "For free-form conversation, add a provider key with `setup-key.sh`."
    )
