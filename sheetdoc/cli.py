"""Command-line interface for sheetdoc."""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path

from . import audit, builder, datagen, formatting, tutor
from .report import Report, render
from .workbook import load


def _backup(path: Path) -> Path:
    """Never overwrite a student's file without leaving a copy behind."""
    bak = path.with_suffix(path.suffix + ".bak")
    n = 1
    while bak.exists():
        bak = path.with_suffix(f"{path.suffix}.bak{n}")
        n += 1
    shutil.copy2(path, bak)
    return bak


def _emit(rep: Report, args) -> int:
    print(render(rep, fmt="json" if args.json else "text", verbose=args.verbose))
    return 1 if rep.counts["error"] else 0


def _save(book, args, rep: Report) -> None:
    out = Path(args.out).expanduser() if args.out else book.path
    if out == book.path and not args.no_backup:
        bak = _backup(book.path)
        rep.notes.append(f"Backup written to {bak.name}")
    book.save(out)
    rep.notes.append(f"Saved: {out}")
    if getattr(args, "open_after", False):
        subprocess.run(["open", str(out)], check=False)


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(
        prog="sheetdoc",
        description="Excel homework auditor, tutor, and builder for intro spreadsheet courses.",
    )
    p.add_argument("--json", action="store_true", help="machine-readable output")
    p.add_argument("-v", "--verbose", action="store_true", help="include passing checks")

    # Repeated on every subcommand so `sd audit file.xlsx --json` works as
    # naturally as `sd --json audit file.xlsx`.
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--json", action="store_true", help=argparse.SUPPRESS)
    common.add_argument("-v", "--verbose", action="store_true", help=argparse.SUPPRESS)

    _sub = p.add_subparsers(dest="cmd", required=True)

    class sub:  # noqa: N801 - thin shim so every subparser inherits the common flags
        @staticmethod
        def add_parser(name, **kw):
            kw.setdefault("parents", [common])
            return _sub.add_parser(name, **kw)

    def add_file(sp, sheet=True):
        sp.add_argument("file", help="path to the .xlsx workbook")
        if sheet:
            sp.add_argument("-s", "--sheet", default=None, help="sheet name (default: all)")
        return sp

    def add_write(sp):
        sp.add_argument("--out", default=None, help="write to a new file instead of in place")
        sp.add_argument("--no-backup", action="store_true", help="skip the .bak copy")
        sp.add_argument("--open-after", action="store_true", help="open the result in Excel")
        return sp

    # ---------------- inspection ----------------
    sp = add_file(sub.add_parser("inspect", help="dump workbook structure, columns and formulas"))
    sp.add_argument("--max-rows", type=int, default=30)
    sp.add_argument("--no-formulas", action="store_true")

    # ---------------- 1. auditing ----------------
    sp = add_file(sub.add_parser("audit", help="run all four formula audits at once"))
    sp.add_argument("--strict", action="store_true", help="also flag 0/1 and structural literals")

    sp = add_file(sub.add_parser("hardcoded", help="detect_hardcoded_values"))
    sp.add_argument("--range", dest="cell_range", default=None)
    sp.add_argument("--strict", action="store_true")

    sp = add_file(sub.add_parser("refs", help="check_reference_types"))
    sp.add_argument("--min-run", type=int, default=3)

    add_file(sub.add_parser("ranges", help="validate_range_bounds"))
    add_file(sub.add_parser("lookups", help="verify_lookup_syntax"))

    # ---------------- 2. formatting ----------------
    add_file(sub.add_parser("formats", help="audit_number_formats"))

    sp = add_file(sub.add_parser("structure", help="check_header_and_grid_structure"), sheet=True)
    sp.add_argument("--fix", action="store_true", help="apply the fixes instead of reporting them")
    add_write(sp)

    sp = add_file(sub.add_parser("cf", help="apply_conditional_formatting_rule"), sheet=True)
    sp.add_argument("range", help="range to apply the rule to, e.g. D2:D21")
    sp.add_argument("preset", help="rule preset (see --list)")
    sp.add_argument("--value", default=None)
    sp.add_argument("--value2", default=None)
    sp.add_argument("--dry-run", action="store_true", help="report which cells would highlight")
    add_write(sp)

    sub.add_parser("cf-presets", help="list conditional-formatting presets")

    # ---------------- 3. tutoring ----------------
    sp = sub.add_parser("hint", help="hint_generator_socratic")
    sp.add_argument("--file", default=None)
    sp.add_argument("-s", "--sheet", default=None)
    sp.add_argument("--cell", default=None, help="cell to diagnose, e.g. D5")
    sp.add_argument("--formula", default=None, help="paste a formula instead of a file")
    sp.add_argument("--topic", default=None, help="topic keyword, e.g. vlookup, absolute")
    sp.add_argument("--tier", type=int, default=0, choices=[0, 1, 2, 3],
                    help="0 = all tiers; 1-3 = stop at that tier")

    sp = sub.add_parser("error", help="explain_error_code")
    sp.add_argument("code", help="#N/A, #VALUE!, #REF!, #DIV/0!, #NAME?, #NUM!, #SPILL!, ####, circular")
    sp.add_argument("--formula", default=None, help="the formula that produced it")

    sp = sub.add_parser("practice", help="generate_practice_variant")
    sp.add_argument("topic", help="vlookup, xlookup, index_match, absolute, sumif, if_nested, "
                                  "percentage, payroll, gradebook, budget")
    sp.add_argument("out", help="path for the new practice workbook")
    sp.add_argument("--rows", type=int, default=15)
    sp.add_argument("--seed", type=int, default=None)
    sp.add_argument("--open-after", action="store_true")

    # ---------------- 4. data and analysis ----------------
    sp = sub.add_parser("mock", help="build_mock_dataset")
    sp.add_argument("kind", help="sales, payroll, grades, inventory, expenses, budget")
    sp.add_argument("out", help="path for the workbook")
    sp.add_argument("--rows", type=int, default=20)
    sp.add_argument("--seed", type=int, default=None)
    sp.add_argument("--sheet-name", default=None)
    sp.add_argument("--append-to", default=None, help="add the sheet to an existing workbook")
    sp.add_argument("--plain", action="store_true", help="no header styling or number formats")
    sp.add_argument("--open-after", action="store_true")

    sp = add_file(sub.add_parser("pivot", help="explain_pivot_table_setup"))
    sp.add_argument("--goal", default=None, help="what you want the pivot to show")

    sp = add_file(sub.add_parser("chart", help="chart_type_recommender"))
    sp.add_argument("--range", dest="cell_range", default=None)

    # ---------------- build and edit ----------------
    sp = sub.add_parser("build", help="create a workbook from a JSON spec")
    sp.add_argument("spec", help="path to a JSON spec file, or inline JSON")
    sp.add_argument("--out", default=None)
    sp.add_argument("--open-after", action="store_true")

    sp = add_file(sub.add_parser("set", help="write values or formulas into cells"), sheet=True)
    sp.add_argument("assignments", nargs="+", help='e.g. D2="=B2*$F$1" E2=100')
    sp.add_argument("--format", dest="number_format", default=None)
    sp.add_argument("--bold", action="store_true")
    add_write(sp)

    sp = add_file(sub.add_parser("fill", help="fill a formula down a column"), sheet=True)
    sp.add_argument("source", help="cell holding the formula, e.g. D2")
    sp.add_argument("through", help="last cell to fill, e.g. D21")
    add_write(sp)

    sp = add_file(sub.add_parser("addchart", help="insert a chart"), sheet=True)
    sp.add_argument("type", help="bar, column, line, pie, scatter")
    sp.add_argument("data", help="data range including headers, e.g. B1:C21")
    sp.add_argument("--cats", default=None, help="category labels range, e.g. A2:A21")
    sp.add_argument("--at", default="H2")
    sp.add_argument("--title", default=None)
    sp.add_argument("--x-title", default=None)
    sp.add_argument("--y-title", default=None)
    add_write(sp)

    sp = add_file(sub.add_parser("addrow", help="add_worksheet_tablerow"), sheet=True)
    sp.add_argument("values", nargs="+", help="cell values in column order; use _ to keep formulas")
    sp.add_argument("--table-name", default=None)
    add_write(sp)

    sp = add_file(sub.add_parser("updaterow", help="update_worksheet_tablerow"), sheet=True)
    sp.add_argument("--row", type=int, default=None, help="1-based row within the table's data")
    sp.add_argument("--where", default=None, help="COLUMN=VALUE to locate the row")
    sp.add_argument("--set", dest="sets", nargs="+", required=True, help="COLUMN=VALUE pairs")
    sp.add_argument("--table-name", default=None)
    add_write(sp)

    sp = sub.add_parser("open", help="open a workbook in Excel")
    sp.add_argument("file")

    args = p.parse_args(argv)

    # ======================================================================
    try:
        return _dispatch(args)
    except FileNotFoundError as e:
        print(f"error: {e}", file=sys.stderr)
        return 2
    except KeyError as e:
        print(f"error: {e}", file=sys.stderr)
        return 2


def _dispatch(args) -> int:  # noqa: C901 - a dispatch table is meant to be flat
    cmd = args.cmd

    if cmd == "cf-presets":
        rep = Report(title="Conditional formatting presets")
        from .report import Finding

        for name, desc in sorted(formatting.PRESETS.items()):
            rep.add(Finding("info", "PRESET", None, None, name, desc))
        rep.notes.append("Add --dry-run to see which cells would highlight before writing.")
        return _emit(rep, args)

    if cmd == "open":
        subprocess.run(["open", str(Path(args.file).expanduser())], check=False)
        return 0

    if cmd == "error":
        return _emit(tutor.explain_error_code(args.code, args.formula), args)

    if cmd == "practice":
        rep = tutor.generate_practice_variant(args.topic, args.out, args.rows, args.seed)
        if args.open_after and rep.counts["error"] == 0:
            subprocess.run(["open", str(Path(args.out).expanduser())], check=False)
        return _emit(rep, args)

    if cmd == "mock":
        rep = datagen.build_mock_dataset(
            args.kind, args.out, args.rows, args.seed, args.sheet_name,
            formatted=not args.plain, append_to=args.append_to,
        )
        if args.open_after and rep.counts["error"] == 0:
            subprocess.run(["open", str(Path(args.append_to or args.out).expanduser())], check=False)
        return _emit(rep, args)

    if cmd == "build":
        spec = args.spec
        rep = builder.build_from_spec(spec, args.out)
        if args.open_after and rep.counts["error"] == 0:
            target = args.out or json.loads(
                Path(spec).read_text() if Path(spec).exists() else spec
            ).get("file")
            if target:
                subprocess.run(["open", str(Path(target).expanduser())], check=False)
        return _emit(rep, args)

    if cmd == "hint" and not args.file:
        return _emit(
            tutor.hint_generator_socratic(None, None, args.cell, args.formula,
                                          args.topic, args.tier),
            args,
        )

    # -- everything below needs a workbook ---------------------------------
    path = args.file
    book = load(path)

    if cmd == "inspect":
        rep = builder.inspect(book, args.sheet, not args.no_formulas, args.max_rows)
    elif cmd == "audit":
        rep = audit.audit_all(book, args.sheet, args.strict)
    elif cmd == "hardcoded":
        rep = audit.detect_hardcoded_values(book, args.sheet, args.cell_range, args.strict)
    elif cmd == "refs":
        rep = audit.check_reference_types(book, args.sheet, args.min_run)
    elif cmd == "ranges":
        rep = audit.validate_range_bounds(book, args.sheet)
    elif cmd == "lookups":
        rep = audit.verify_lookup_syntax(book, args.sheet)
    elif cmd == "formats":
        rep = formatting.audit_number_formats(book, args.sheet)
    elif cmd == "pivot":
        rep = datagen.explain_pivot_table_setup(book, args.sheet, args.goal)
    elif cmd == "chart":
        rep = datagen.chart_type_recommender(book, args.sheet, args.cell_range)
    elif cmd == "hint":
        rep = tutor.hint_generator_socratic(book, args.sheet, args.cell, args.formula,
                                            args.topic, args.tier)
    elif cmd == "structure":
        rep = formatting.check_header_and_grid_structure(book, args.sheet, args.fix)
        if args.fix:
            _save(book, args, rep)
    elif cmd == "cf":
        rep = formatting.apply_conditional_formatting_rule(
            book, args.sheet, args.range, args.preset,
            _num(args.value), _num(args.value2), args.dry_run,
        )
        if not args.dry_run and rep.counts["error"] == 0:
            _save(book, args, rep)
    elif cmd == "set":
        rep = builder.set_cells(book, args.sheet, _pairs(args.assignments),
                                args.number_format, args.bold)
        if rep.counts["error"] == 0:
            _save(book, args, rep)
    elif cmd == "fill":
        rep = builder.fill_down(book, args.sheet, args.source, args.through)
        if rep.counts["error"] == 0:
            _save(book, args, rep)
    elif cmd == "addchart":
        rep = builder.add_chart(book, args.sheet, args.type, args.data, args.cats,
                                args.at, args.title, args.x_title, args.y_title)
        if rep.counts["error"] == 0:
            _save(book, args, rep)
    elif cmd == "addrow":
        vals = [None if v == "_" else _num(v, keep_text=True) for v in args.values]
        rep = builder.add_worksheet_tablerow(book, args.sheet, vals, args.table_name)
        if rep.counts["error"] == 0:
            _save(book, args, rep)
    elif cmd == "updaterow":
        sets = _pairs(args.sets)
        mc = mv = None
        if args.where and "=" in args.where:
            mc, mv = args.where.split("=", 1)
        rep = builder.update_worksheet_tablerow(
            book, args.sheet, args.row or 1, sets, args.table_name, mc, mv
        )
        if rep.counts["error"] == 0:
            _save(book, args, rep)
    else:
        print(f"error: unhandled command {cmd}", file=sys.stderr)
        return 2

    return _emit(rep, args)


def _pairs(items: list[str]) -> dict[str, object]:
    out: dict[str, object] = {}
    for item in items:
        if "=" not in item:
            continue
        k, v = item.split("=", 1)
        out[k.strip()] = _num(v.strip().strip('"').strip("'"), keep_text=True)
    return out


def _num(v, keep_text: bool = False):
    if v is None:
        return None
    s = str(v)
    if s.startswith("="):
        return s
    try:
        return int(s)
    except ValueError:
        pass
    try:
        return float(s)
    except ValueError:
        return s if keep_text else v


if __name__ == "__main__":
    sys.exit(main())
