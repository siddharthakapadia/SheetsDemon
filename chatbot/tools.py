"""Tool definitions exposed to Claude, and their dispatch into the sheetdoc engine.

Descriptions are written to be prescriptive about *when* to call each tool, not
just what it does — that is what drives correct tool selection.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from sheetdoc import audit, builder, datagen, formatting, tutor  # noqa: E402
from sheetdoc.report import Report  # noqa: E402
from sheetdoc.workbook import load  # noqa: E402

# The directory uploaded workbooks live in; set by the server at startup.
from .config import WORKSPACE_DIR

WORKSPACE = WORKSPACE_DIR


def _path(name: str) -> Path:
    """Resolve a workbook name to a path inside the workspace.

    The model supplies this name, so it is untrusted: resolve it and confirm it
    stays inside the workspace before any file operation.
    """
    p = (WORKSPACE / Path(name).name).resolve()
    if not p.is_relative_to(WORKSPACE.resolve()):
        raise ValueError(f"Path escapes the workspace: {name}")
    return p


def _out(report: Report, extra: dict | None = None) -> str:
    """Render a Report for the model: structured, compact, no terminal styling."""
    payload = {
        "title": report.title,
        "summary": report.summary,
        "counts": report.counts,
        "findings": [
            {
                k: v
                for k, v in {
                    "severity": f.severity,
                    # `code` distinguishes finding kinds (TASK vs PIVOT_STEP vs
                    # ERR_CAUSE). Callers need it to render a report as anything
                    # other than a flat list.
                    "code": f.code,
                    "cell": f.cell,
                    "sheet": f.sheet,
                    "message": f.message,
                    "detail": f.detail,
                    "fix": f.fix,
                    "formula": f.formula,
                    "affected_cells": f.count if f.count > 1 else None,
                    # Machine-actionable fix data (e.g. {"reference": "B15",
                    # "suggested": "$B$15"}) so a caller can apply the fix, not
                    # just describe it.
                    "data": f.data or None,
                }.items()
                if v
            }
            for f in report.sorted_findings()
        ],
        "notes": report.notes,
    }
    if extra:
        payload.update(extra)
    return json.dumps(payload, default=str)


# ==========================================================================
# Tool schemas
# ==========================================================================

TOOLS = [
    {
        "name": "inspect_workbook",
        "description": (
            "Read a workbook's structure: sheets, column headers and data types, "
            "where the data block sits, and every formula including which ones were "
            "filled down as a group. Call this FIRST on any workbook you have not "
            "seen before — you cannot advise on a sheet whose shape you do not know, "
            "and every other tool's output is easier to interpret once you have this."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "file": {"type": "string", "description": "Workbook filename, e.g. homework3.xlsx"},
                "sheet": {"type": "string", "description": "Sheet name. Omit for all sheets."},
            },
            "required": ["file"],
            "additionalProperties": False,
        },
    },
    {
        "name": "audit_formulas",
        "description": (
            "Run the formula audits: hardcoded numbers typed into formulas, $ anchoring "
            "and fill consistency, SUM/AVERAGE range boundaries, and VLOOKUP/XLOOKUP/"
            "INDEX+MATCH argument correctness. Call this whenever the user asks whether "
            "their work is right, why a formula is wrong, or to check/grade an assignment. "
            "Never judge a formula by reading it yourself — this parses formulas properly, "
            "simulates fills, and checks ranges against the actual data."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "file": {"type": "string"},
                "sheet": {"type": "string"},
                "checks": {
                    "type": "string",
                    "enum": ["all", "hardcoded", "references", "ranges", "lookups"],
                    "description": "Which audit to run. Default 'all'.",
                },
            },
            "required": ["file"],
            "additionalProperties": False,
        },
    },
    {
        "name": "check_formatting",
        "description": (
            "Check presentation: number formats against what each column header implies "
            "(currency, percentage, dates), plus bold headers, alignment, borders, freeze "
            "panes, merged cells and column widths. Call this when the assignment is graded "
            "on appearance, when the user asks how their sheet looks, or alongside "
            "audit_formulas for a complete review. Set fix=true to apply the structural "
            "fixes rather than just report them."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "file": {"type": "string"},
                "sheet": {"type": "string"},
                "what": {
                    "type": "string",
                    "enum": ["all", "number_formats", "structure"],
                    "description": "Default 'all'.",
                },
                "fix": {
                    "type": "boolean",
                    "description": "Apply structural fixes (bold headers, borders, freeze panes, widths). Default false.",
                },
            },
            "required": ["file"],
            "additionalProperties": False,
        },
    },
    {
        "name": "conditional_formatting",
        "description": (
            "Apply a conditional formatting rule, or simulate it. Call with dry_run=true "
            "to report exactly which cells would highlight WITHOUT changing the file — use "
            "that to check an assignment's rule before committing. Presets: negative_red, "
            "above, below, between, top_n, bottom_n, above_average, duplicates, unique, "
            "blanks, errors, text_contains, data_bar, color_scale, icon_set."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "file": {"type": "string"},
                "sheet": {"type": "string"},
                "range": {"type": "string", "description": "A1 range, e.g. D2:D21"},
                "preset": {"type": "string"},
                "value": {"type": "string", "description": "Threshold, where the preset needs one."},
                "value2": {"type": "string", "description": "Second threshold for 'between'."},
                "dry_run": {"type": "boolean", "description": "Report hits without writing. Default false."},
            },
            "required": ["file", "sheet", "range", "preset"],
            "additionalProperties": False,
        },
    },
    {
        "name": "socratic_hint",
        "description": (
            "Get a 3-tier hint ladder (concept → syntax → shape with placeholders) for a "
            "cell, a pasted formula, or a topic. The tool diagnoses the real problem first, "
            "so hints address what is actually wrong. Call this when the user says they are "
            "stuck, wants to understand rather than be told, or asks for a hint — NOT when "
            "they asked you to just fix it. Tier 3 shows the pattern with placeholder names, "
            "never the finished formula for their own cells."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "file": {"type": "string"},
                "sheet": {"type": "string"},
                "cell": {"type": "string", "description": "Cell to diagnose, e.g. D5"},
                "formula": {"type": "string", "description": "A pasted formula, if there is no file."},
                "topic": {"type": "string", "description": "Topic keyword, e.g. 'absolute references'."},
                "tier": {
                    "type": "integer",
                    "description": "0 = all tiers (default). 1-3 stops at that tier, for a user who wants to stay stuck on purpose.",
                },
            },
            "additionalProperties": False,
        },
    },
    {
        "name": "explain_error",
        "description": (
            "Explain an Excel error code in beginner terms with a concrete debugging route: "
            "#N/A, #VALUE!, #REF!, #DIV/0!, #NAME?, #NUM!, #NULL!, #SPILL!, ##### or "
            "'circular'. Call this the moment the user mentions seeing one of these."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "code": {"type": "string"},
                "formula": {"type": "string", "description": "The formula that produced it, if known."},
            },
            "required": ["code"],
            "additionalProperties": False,
        },
    },
    {
        "name": "create_practice",
        "description": (
            "Build a practice workbook: same technique as the problem the user is stuck on, "
            "different numbers, with tasks and a hidden answer key. Call this when the user "
            "wants to practise, or after they have struggled with a concept and would "
            "benefit from a parallel problem. Topics: vlookup, xlookup, index_match, "
            "absolute, sumif, if_nested, percentage, payroll, gradebook, budget."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "topic": {"type": "string"},
                "filename": {"type": "string", "description": "Output name, e.g. practice_vlookup.xlsx"},
                "rows": {"type": "integer", "description": "Rows of data. Default 15."},
            },
            "required": ["topic", "filename"],
            "additionalProperties": False,
        },
    },
    {
        "name": "create_dataset",
        "description": (
            "Generate a realistic practice dataset as a formatted worksheet. Values are "
            "internally consistent (revenue really equals units x price), so pivots and "
            "charts built on it behave like real data. Kinds: sales, payroll, grades, "
            "inventory, expenses, budget."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "kind": {"type": "string"},
                "filename": {"type": "string"},
                "rows": {"type": "integer", "description": "Default 20."},
                "append_to": {
                    "type": "string",
                    "description": "Existing workbook to add the sheet to, instead of creating a new file.",
                },
            },
            "required": ["kind", "filename"],
            "additionalProperties": False,
        },
    },
    {
        "name": "explain_pivot",
        "description": (
            "Read a sheet's actual columns and map them onto a PivotTable's Rows, Columns, "
            "Values and Filters, with click-by-click steps for Excel. Call this for any "
            "pivot table question — it reasons from the real column types rather than "
            "generic advice."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "file": {"type": "string"},
                "sheet": {"type": "string"},
                "goal": {"type": "string", "description": "What the pivot should show, e.g. 'total revenue by region'."},
            },
            "required": ["file"],
            "additionalProperties": False,
        },
    },
    {
        "name": "recommend_chart",
        "description": (
            "Analyse a range and rank chart types from the shape of the data — whether the "
            "x-axis is time, whether both variables are continuous, whether parts sum to a "
            "whole — with the reasoning for each. Call this whenever the user asks which "
            "chart to use, or before building one."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "file": {"type": "string"},
                "sheet": {"type": "string"},
                "range": {"type": "string", "description": "Optional A1 range. Omit to use the whole data block."},
            },
            "required": ["file"],
            "additionalProperties": False,
        },
    },
    {
        "name": "build_workbook",
        "description": (
            "Create a whole workbook from a specification: sheets, a formatted table with "
            "headers and formulas, number formats, a totals row, conditional formatting and "
            "charts. Call this when the user asks you to build a spreadsheet from scratch. "
            "Formats: currency, currency0, accounting, percent, percent0, percent2, number, "
            "integer, date, text, general."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "filename": {"type": "string"},
                "spec": {
                    "type": "string",
                    "description": (
                        "JSON spec. Shape: {\"sheets\": [{\"name\": \"Budget\", "
                        "\"cells\": {\"A1\": {\"v\": \"Q1 Budget\", \"bold\": true, \"size\": 14}}, "
                        "\"table\": {\"at\": \"A3\", \"headers\": [...], \"rows\": [[...]], "
                        "\"formats\": [null, \"currency\"], \"totals\": [null, \"SUM\"]}, "
                        "\"conditional\": [{\"range\": \"D4:D10\", \"preset\": \"negative_red\"}], "
                        "\"charts\": [{\"type\": \"column\", \"data\": \"B3:C10\", "
                        "\"cats\": \"A4:A10\", \"at\": \"F3\", \"title\": \"...\"}]}]}. "
                        "Put formulas in table rows as strings starting with '='."
                    ),
                },
            },
            "required": ["filename", "spec"],
            "additionalProperties": False,
        },
    },
    {
        "name": "edit_workbook",
        "description": (
            "Modify an existing workbook. Operations: 'set' writes values or formulas into "
            "cells; 'fill' replicates a formula down a column with references translated "
            "correctly; 'add_chart' inserts a chart; 'add_row' appends a table row (inserting "
            "above a totals row and extending the total formulas); 'update_row' edits an "
            "existing row. A .bak backup is made automatically before any change."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "file": {"type": "string"},
                "sheet": {"type": "string"},
                "operation": {
                    "type": "string",
                    "enum": ["set", "fill", "add_chart", "add_row", "update_row"],
                },
                "cells": {
                    "type": "string",
                    "description": "For 'set': JSON object of cell->value, e.g. {\"D2\": \"=B2*$F$1\", \"E2\": 100}",
                },
                "source": {"type": "string", "description": "For 'fill': the cell holding the formula, e.g. D2"},
                "through": {"type": "string", "description": "For 'fill': last cell to fill, e.g. D21"},
                "chart_type": {"type": "string", "description": "For 'add_chart': bar, column, line, pie, scatter"},
                "data_range": {"type": "string", "description": "For 'add_chart': data range including headers"},
                "categories": {"type": "string", "description": "For 'add_chart': category labels range"},
                "anchor": {"type": "string", "description": "For 'add_chart': where to place it, e.g. H2"},
                "title": {"type": "string", "description": "For 'add_chart': chart title"},
                "values": {
                    "type": "string",
                    "description": "For 'add_row': JSON array of values in column order; use null to keep a formula.",
                },
                "row_match": {
                    "type": "string",
                    "description": "For 'update_row': COLUMN=VALUE identifying the row, e.g. Category=Travel",
                },
                "updates": {
                    "type": "string",
                    "description": "For 'update_row': JSON object of column->new value, e.g. {\"Actual\": 999}",
                },
            },
            "required": ["file", "operation"],
            "additionalProperties": False,
        },
    },
    {
        "name": "list_workbooks",
        "description": (
            "List the workbooks currently available to work on. Call this when the user "
            "refers to a file without naming it exactly, or when you need to check whether "
            "something they mentioned has actually been uploaded."
        ),
        "input_schema": {"type": "object", "properties": {}, "additionalProperties": False},
    },

    # ---------------- live Excel ----------------
    {
        "name": "excel_status",
        "description": (
            "Check what is happening in the Microsoft Excel app right now: which workbook "
            "and sheet are active, which cell the user has selected, and whether there are "
            "unsaved changes. Call this FIRST whenever the user talks about what they are "
            "'looking at', 'this cell', 'my sheet', or working in Excel — it tells you "
            "whether live collaboration is available at all."
        ),
        "input_schema": {"type": "object", "properties": {}, "additionalProperties": False},
    },
    {
        "name": "read_excel_live",
        "description": (
            "Read from the workbook OPEN IN EXCEL RIGHT NOW, including changes the user has "
            "not saved yet. Returns both the calculated value and the underlying formula for "
            "each cell — the file-based tools can only see values Excel cached at last save, "
            "so this is the only way to see live results. Prefer this over inspect_workbook "
            "whenever the user is actively working in Excel. Pass any address you want — "
            "'H12', 'A1:D20', a whole column. The user does NOT need that cell selected, and "
            "you must never ask them to select one: if they name a cell, read it directly."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "range": {"type": "string", "description": "A1 range, e.g. A1:D20. Omit for the used range."},
                "sheet": {"type": "string", "description": "Sheet name. Omit for the active sheet."},
            },
            "additionalProperties": False,
        },
    },
    {
        "name": "write_excel_live",
        "description": (
            "Write values or formulas straight into the workbook open in Excel — the user "
            "sees the change appear immediately and Excel calculates formulas on the spot. "
            "Use this when the user asks you to fix or change something while they have the "
            "file open. Anything starting with '=' is entered as a formula. Say what you are "
            "changing before you call this; it edits their live work."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "cells": {
                    "type": "string",
                    "description": 'JSON object of cell->value, e.g. {"D2": "=C2*$B$15", "E1": "Margin"}',
                },
                "sheet": {"type": "string", "description": "Sheet name. Omit for the active sheet."},
            },
            "required": ["cells"],
            "additionalProperties": False,
        },
    },
    {
        "name": "format_excel_live",
        "description": (
            "Apply an Excel number format to a range in the live workbook, e.g. '$#,##0.00' "
            "for currency, '0.0%' for percentage, 'mm/dd/yyyy' for dates."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "range": {"type": "string"},
                "number_format": {"type": "string"},
                "sheet": {"type": "string"},
            },
            "required": ["range", "number_format"],
            "additionalProperties": False,
        },
    },
    {
        "name": "audit_live_workbook",
        "description": (
            "Run the full formula audit on the workbook currently open in Excel. This saves "
            "the workbook first (so the analysis engine can read it), then audits it. Use "
            "this when the user wants their open work checked. If they have unsaved changes "
            "they care about, tell them it will save first."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "sheet": {"type": "string"},
                "checks": {
                    "type": "string",
                    "enum": ["all", "hardcoded", "references", "ranges", "lookups", "formatting"],
                },
            },
            "additionalProperties": False,
        },
    },
    {
        "name": "create_chart_live",
        "description": (
            "Create a real chart in the workbook open in Excel — the user sees it appear "
            "on their sheet. Call this whenever they ask for a chart or graph of their "
            "data. If you are unsure which type fits, call recommend_chart first; it "
            "analyses the data shape and explains the choice. Types: column, bar, line, "
            "pie, scatter, area, doughnut."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "range": {"type": "string", "description": "Data range INCLUDING headers, e.g. A1:C10"},
                "chart_type": {"type": "string", "description": "column, bar, line, pie, scatter, area, doughnut"},
                "title": {"type": "string"},
                "sheet": {"type": "string"},
                "anchor": {"type": "string", "description": "Cell to place it near, e.g. F2"},
            },
            "required": ["range"],
            "additionalProperties": False,
        },
    },
    {
        "name": "open_in_excel",
        "description": (
            "Open one of the workspace workbooks in the Excel app so the user can work on it "
            "directly and you can then collaborate live. Use after building or generating a "
            "file when the user wants to actually use it."
        ),
        "input_schema": {
            "type": "object",
            "properties": {"file": {"type": "string"}},
            "required": ["file"],
            "additionalProperties": False,
        },
    },
    {
        "name": "create_table_live",
        "description": (
            "Build a data table on the sheet the user has open: writes the header row and "
            "the data, registers it as a real Excel Table (banded rows, filter buttons, "
            "structured references) and applies number formats. Use this whenever the user "
            "asks for a table, a list, a tracker, a log, or 'set up columns for X'. "
            "Row values starting with '=' are written as live formulas, so a Total column "
            "can be '=B2*C2'. Invent sensible sample rows only if the user asked for sample "
            "data; otherwise write the headers and leave the rows empty for them to fill."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "headers": {"type": "array", "items": {"type": "string"},
                            "description": "Column headings, left to right."},
                "rows": {"type": "array", "items": {"type": "array"},
                         "description": "Data rows. A value beginning with '=' becomes a formula."},
                "at": {"type": "string",
                       "description": "Top-left cell of the header row. Default A1."},
                "sheet": {"type": "string"},
                "number_formats": {
                    "type": "array", "items": {"type": ["string", "null"]},
                    "description": "One Excel format code per column, or null to leave alone. "
                                   "e.g. [null, '#,##0', '$#,##0.00'].",
                },
                "as_excel_table": {"type": "boolean"},
            },
            "required": ["headers"],
            "additionalProperties": False,
        },
    },
    {
        "name": "new_spreadsheet",
        "description": (
            "Create a brand-new .xlsx workbook and open it in Excel. Use when the user asks "
            "to start a new spreadsheet or a fresh file, rather than editing what is already "
            "open. Optionally seed it with a header row and data. If they want a table added "
            "to the workbook they are ALREADY in, use create_table_live instead."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "name": {"type": "string",
                         "description": "File name, e.g. 'budget.xlsx'. Auto-named if omitted."},
                "sheet_name": {"type": "string", "description": "Name of the first worksheet."},
                "headers": {"type": "array", "items": {"type": "string"}},
                "rows": {"type": "array", "items": {"type": "array"}},
            },
            "additionalProperties": False,
        },
    },
]


# ==========================================================================
# Dispatch
# ==========================================================================


def run_tool(name: str, args: dict) -> tuple[str, bool]:
    """Execute a tool. Returns (result_text, is_error)."""
    try:
        return _dispatch(name, args), False
    except FileNotFoundError as e:
        return f"File not found: {e}. Call list_workbooks to see what is available.", True
    except KeyError as e:
        return f"Not found: {e}", True
    except Exception as e:  # surfaced to the model so it can adapt (incl. ExcelError)
        return f"{type(e).__name__}: {e}", True


def _dispatch_live(name: str, a: dict) -> str:
    """Tools that talk to the running Excel app rather than files on disk."""
    from . import excel_live as X

    if name == "excel_status":
        st = X.state()
        return json.dumps({
            "excel_running": st.running,
            "active_workbook": st.active_workbook,
            "active_sheet": st.active_sheet,
            "sheets": st.sheets,
            "open_workbooks": st.workbooks,
            "selected_cell": st.selection,
            "used_range": st.used_range,
            "has_unsaved_changes": st.unsaved,
            "note": st.error or (
                "Live collaboration is available — you can read and write this workbook "
                "directly while the user works in it."
            ),
        })

    if name == "read_excel_live":
        g = X.read_grid(a.get("range"), a.get("sheet"))
        from sheetdoc.formula import col_letter

        cells = []
        for ri, row in enumerate(g.values):
            for ci, val in enumerate(row):
                formula = ""
                if ri < len(g.formulas) and ci < len(g.formulas[ri]):
                    formula = g.formulas[ri][ci]
                if not val and not formula:
                    continue
                addr = f"{col_letter(g.start_col + ci)}{g.start_row + ri}"
                entry = {"cell": addr, "value": val}
                if formula.startswith("="):
                    entry["formula"] = formula
                cells.append(entry)
        return json.dumps({
            "sheet": g.sheet, "range": g.address, "cells": cells,
            "note": "Values are live calculated results, including unsaved edits.",
        })

    if name == "write_excel_live":
        cells = json.loads(a["cells"]) if isinstance(a["cells"], str) else a["cells"]
        n = X.write_cells(cells, a.get("sheet"))
        return json.dumps({
            "written": n, "cells": list(cells),
            "note": "Applied to the live workbook. The user can see it now; not yet saved to disk.",
        })

    if name == "format_excel_live":
        r = X.set_number_format(a["range"], a["number_format"], a.get("sheet"))
        return json.dumps({"formatted": r, "number_format": a["number_format"]})

    if name == "create_chart_live":
        r = X.add_chart(a["range"], a.get("chart_type", "column"),
                        a.get("title"), a.get("sheet"), a.get("anchor"))
        return json.dumps({**r, "note": "The chart is on the user's sheet now."})

    if name == "create_table_live":
        r = X.create_table(
            a["headers"], a.get("rows"), a.get("at", "A1"), a.get("sheet"),
            a.get("as_excel_table", True), a.get("number_formats"))
        return json.dumps({**r, "note": (
            "The table is on the user's sheet now — describe what you built, do not "
            "re-list every row back at them.")})

    if name == "new_spreadsheet":
        fname = Path(str(a.get("name") or "")).name
        if fname and not fname.lower().endswith((".xlsx", ".xlsm")):
            fname += ".xlsx"
        r = X.new_workbook(
            (WORKSPACE / fname) if fname else None,
            a.get("sheet_name"), a.get("headers"), a.get("rows"))
        return json.dumps({**r, "note": (
            "A new workbook was created and is now the active one in Excel.")})

    if name == "open_in_excel":
        p = _path(a["file"])
        if not p.exists():
            raise FileNotFoundError(p)
        name = X.open_file(p)
        note = ("Now open in Excel." if not X.open_file.already_open else
                "Excel already had this workbook open, so you are looking at its "
                "in-memory copy — any unsaved edits are still there, and changes "
                "made to the file on disk since it was opened are not shown.")
        return json.dumps({"opened": name, "note": note})

    if name == "audit_live_workbook":
        path = X.active_workbook_path()
        if not path:
            X.save_active()
            path = X.active_workbook_path()
        if not path:
            return json.dumps({
                "error": "The open workbook has never been saved, so there is no file to "
                         "analyse. Ask the user to save it once in Excel first.",
            })
        X.save_active()
        book = load(Path(path))
        sheet = a.get("sheet")
        check = a.get("checks", "all")
        if check == "formatting":
            rep = Report(title="Formatting review")
            rep.extend(formatting.audit_number_formats(book, sheet))
            rep.extend(formatting.check_header_and_grid_structure(book, sheet, False))
        else:
            rep = {
                "all": lambda: audit.audit_all(book, sheet),
                "hardcoded": lambda: audit.detect_hardcoded_values(book, sheet),
                "references": lambda: audit.check_reference_types(book, sheet),
                "ranges": lambda: audit.validate_range_bounds(book, sheet),
                "lookups": lambda: audit.verify_lookup_syntax(book, sheet),
            }[check]()
        return _out(rep, {"audited_live_workbook": Path(path).name})

    return f"Unknown live tool: {name}"


def _save(book, path: Path) -> str:
    bak = path.with_suffix(path.suffix + ".bak")
    if path.exists() and not bak.exists():
        bak.write_bytes(path.read_bytes())
    book.save(path)
    return f" Saved to {path.name} (backup: {bak.name})."


def _dispatch(name: str, a: dict) -> str:  # noqa: C901 - flat dispatch table
    # ---------------- live Excel ----------------
    if name.endswith("_excel_live") or name in (
        "excel_status", "audit_live_workbook", "open_in_excel", "create_chart_live",
        "create_table_live", "new_spreadsheet",
    ):
        return _dispatch_live(name, a)

    if name == "list_workbooks":
        files = sorted(
            p.name for p in WORKSPACE.glob("*.xlsx") if not p.name.startswith("~")
        )
        return json.dumps({"workbooks": files} if files else
                          {"workbooks": [], "note": "No workbooks uploaded yet."})

    if name == "explain_error":
        return _out(tutor.explain_error_code(a["code"], a.get("formula")))

    if name == "create_practice":
        path = _path(a["filename"])
        rep = tutor.generate_practice_variant(a["topic"], str(path), a.get("rows", 15))
        return _out(rep, {"created_file": path.name})

    if name == "create_dataset":
        path = _path(a["filename"])
        append = str(_path(a["append_to"])) if a.get("append_to") else None
        rep = datagen.build_mock_dataset(
            a["kind"], str(path), a.get("rows", 20), append_to=append
        )
        return _out(rep, {"created_file": Path(append or path).name})

    if name == "build_workbook":
        path = _path(a["filename"])
        spec = json.loads(a["spec"]) if isinstance(a["spec"], str) else a["spec"]
        rep = builder.build_from_spec(spec, str(path))
        return _out(rep, {"created_file": path.name})

    if name == "socratic_hint" and not a.get("file"):
        return _out(tutor.hint_generator_socratic(
            None, None, a.get("cell"), a.get("formula"), a.get("topic"), a.get("tier", 0)
        ))

    # -- everything below needs a workbook -------------------------------
    path = _path(a["file"])
    book = load(path)
    sheet = a.get("sheet")

    if name == "inspect_workbook":
        return _out(builder.inspect(book, sheet))

    if name == "audit_formulas":
        check = a.get("checks", "all")
        rep = {
            "all": lambda: audit.audit_all(book, sheet),
            "hardcoded": lambda: audit.detect_hardcoded_values(book, sheet),
            "references": lambda: audit.check_reference_types(book, sheet),
            "ranges": lambda: audit.validate_range_bounds(book, sheet),
            "lookups": lambda: audit.verify_lookup_syntax(book, sheet),
        }[check]()
        return _out(rep)

    if name == "check_formatting":
        what = a.get("what", "all")
        fix = a.get("fix", False)
        rep = Report(title="Formatting review")
        if what in ("all", "number_formats"):
            rep.extend(formatting.audit_number_formats(book, sheet))
        if what in ("all", "structure"):
            rep.extend(formatting.check_header_and_grid_structure(book, sheet, fix))
        note = _save(book, path) if fix else ""
        return _out(rep, {"saved": note.strip()} if note else None)

    if name == "conditional_formatting":
        rep = formatting.apply_conditional_formatting_rule(
            book, sheet, a["range"], a["preset"],
            a.get("value"), a.get("value2"), a.get("dry_run", False),
        )
        note = "" if a.get("dry_run") or rep.counts["error"] else _save(book, path)
        return _out(rep, {"saved": note.strip()} if note else None)

    if name == "socratic_hint":
        return _out(tutor.hint_generator_socratic(
            book, sheet, a.get("cell"), a.get("formula"), a.get("topic"), a.get("tier", 0)
        ))

    if name == "explain_pivot":
        return _out(datagen.explain_pivot_table_setup(book, sheet, a.get("goal")))

    if name == "recommend_chart":
        return _out(datagen.chart_type_recommender(book, sheet, a.get("range")))

    if name == "edit_workbook":
        op = a["operation"]
        if op == "set":
            rep = builder.set_cells(book, sheet, json.loads(a["cells"]))
        elif op == "fill":
            rep = builder.fill_down(book, sheet, a["source"], a["through"])
        elif op == "add_chart":
            rep = builder.add_chart(
                book, sheet, a["chart_type"], a["data_range"],
                a.get("categories"), a.get("anchor", "H2"), a.get("title"),
            )
        elif op == "add_row":
            rep = builder.add_worksheet_tablerow(book, sheet, json.loads(a["values"]))
        elif op == "update_row":
            mc = mv = None
            if a.get("row_match") and "=" in a["row_match"]:
                mc, mv = a["row_match"].split("=", 1)
            rep = builder.update_worksheet_tablerow(
                book, sheet, 1, json.loads(a["updates"]), None, mc, mv
            )
        else:
            return f"Unknown operation: {op}"

        note = "" if rep.counts["error"] else _save(book, path)
        return _out(rep, {"saved": note.strip()} if note else None)

    return f"Unknown tool: {name}"
