"""Tests for the live Excel bridge.

Skips cleanly when Excel is unavailable, so the suite still runs on a machine
without it. When Excel IS available it drives a real workbook end to end.
"""

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

PASS, FAIL, SKIP = [], [], []


def check(name, cond, detail=""):
    (PASS if cond else FAIL).append(name + (f" — {detail}" if detail and not cond else ""))


def main():
    from chatbot import excel_live as X

    # ---- pure-function tests (no Excel needed) ---------------------------
    check("escapes double quotes", X._quote('say "hi"') == 'say \\"hi\\"')
    check("escapes backslashes", X._quote("a\\b") == "a\\\\b")
    check("flattens newlines", "\n" not in X._quote("a\nb"))
    check("flattens tabs", "\t" not in X._quote("a\tb"))

    for good in ("A1", "B12", "A1:D20", "$A$1", "AA100", "A1:$C$9"):
        try:
            X._check_addr(good)
            check(f"accepts address {good}", True)
        except X.ExcelError:
            check(f"accepts address {good}", False)

    # Injection attempts must be refused before reaching AppleScript.
    for bad in ('A1" & (do shell script "id") & "', "A1; tell app", "../A1",
                "", "SELECT *", "A", "1"):
        try:
            X._check_addr(bad)
            check(f"rejects bad address {bad!r}", False, "was accepted")
        except X.ExcelError:
            check(f"rejects bad address {bad!r}", True)

    check("cleans Excel error text",
          X._clean_error("execution error: Microsoft Excel got an error: Bad thing. (-50)")
          == "Excel error: Bad thing")

    # ---- live tests ------------------------------------------------------
    if not X.is_running():
        SKIP.append("Excel is not running — live tests skipped")
        return report()

    st = X.state()
    check("state() reports running", st.running)

    if st.error:
        SKIP.append(f"No workbook open ({st.error}) — workbook tests skipped")
        # An error state must still be well-formed, not an exception.
        check("no-workbook state is structured", st.active_workbook is None)
        for tool, args in (("read_excel_live", {}), ("write_excel_live", {"cells": "{}"})):
            from chatbot import tools as T

            out, err = T.run_tool(tool, args)
            check(f"{tool} fails cleanly with no workbook", err and "Traceback" not in out,
                  out[:90])
        return report()

    import openpyxl

    fixture = Path("/tmp/sd_live_test.xlsx")
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = "T"
    ws["A1"] = "Item"; ws["B1"] = "Qty"; ws["C1"] = "Total"
    ws["A2"] = "Bolt"; ws["B2"] = 4; ws["C2"] = "=B2*10"
    wb.save(fixture)

    X.open_file(fixture, reload=True)
    st = X.state()
    check("opens a file in Excel", st.active_workbook == fixture.name, str(st.active_workbook))
    check("reports the sheet", st.active_sheet == "T", str(st.active_sheet))
    check("reports sheet list", "T" in st.sheets, str(st.sheets))
    check("reports the used range", st.used_range == "A1:C2", str(st.used_range))
    check("finds the file path", (X.active_workbook_path() or "").endswith("sd_live_test.xlsx"),
          str(X.active_workbook_path()))

    # Grid shapes — the single-cell case returns a scalar from AppleScript and
    # must not be iterated character by character.
    g = X.read_grid("A1:C2")
    check("block read shape", len(g.values) == 2 and len(g.values[0]) == 3, str(g.values))
    check("block read gets calculated value", g.values[1][2] == "40", str(g.values[1]))
    check("block read gets the formula", g.formulas[1][2] == "=B2*10", str(g.formulas[1]))

    one = X.read_grid("C2:C2")
    check("single-cell read is 1x1", one.values == [["40"]], str(one.values))
    check("single-cell formula is whole", one.formulas == [["=B2*10"]], str(one.formulas))

    row = X.read_grid("A1:C1")
    check("single-row read", row.values == [["Item", "Qty", "Total"]], str(row.values))
    col = X.read_grid("B1:B2")
    check("single-column read", len(col.values) == 2, str(col.values))

    # Live write — including a formula Excel must calculate immediately.
    X.write_cells({"D1": "Note", "B2": 7, "E2": "=B2*100"})
    check("writes a text value", X.read_cell("D1")["value"] == "Note")
    check("writes a number", X.read_cell("B2")["value"] == "7", X.read_cell("B2")["value"])
    check("dependent formula recalculated", X.read_cell("C2")["value"] == "70",
          X.read_cell("C2")["value"])
    check("new formula stored as formula", X.read_cell("E2")["formula"] == "=B2*100")
    check("new formula calculated live", X.read_cell("E2")["value"] == "700",
          X.read_cell("E2")["value"])

    check("workbook is now unsaved", X.state().unsaved)

    # A value containing quotes must survive the AppleScript boundary.
    X.write_cells({"F1": 'He said "ok" \\ done'})
    check("round-trips quotes and backslashes",
          X.read_cell("F1")["value"] == 'He said "ok" \\ done', X.read_cell("F1")["value"])

    X.set_number_format("B2:B2", "$#,##0.00")
    check("sets a number format", "$" in X.read_cell("B2")["number_format"],
          X.read_cell("B2")["number_format"])

    # Save, then confirm the file-based engine sees the same thing.
    X.save_active()
    check("save clears the unsaved flag", not X.state().unsaved)
    ws2 = openpyxl.load_workbook(fixture)["T"]
    check("live write persisted to disk", ws2["E2"].value == "=B2*100", str(ws2["E2"].value))

    # Tool layer over the bridge.
    from chatbot import tools as T

    out, err = T.run_tool("excel_status", {})
    check("excel_status tool works", not err and json.loads(out)["active_workbook"] == fixture.name)

    out, err = T.run_tool("read_excel_live", {"range": "A1:C2"})
    check("read_excel_live tool works", not err and len(json.loads(out)["cells"]) >= 5, out[:100])

    out, err = T.run_tool("audit_live_workbook", {})
    check("audit_live_workbook tool works", not err and "counts" in json.loads(out), out[:100])


    # ---- charts in the live workbook ---------------------------------------
    X.delete_charts()
    check("starts with no charts", X.list_charts() == [], str(X.list_charts()))

    r = X.add_chart("A1:C2", "column", "Test chart")
    check("creates a column chart", len(X.list_charts()) == 1, str(X.list_charts()))
    check("reports the type back", r["chart_type"] == "column", str(r))

    # Every advertised type must actually be a valid Excel term. Several obvious
    # guesses are not ("pie", "line", "xy scatter" all fail), so this guards the
    # mapping rather than trusting it.
    for kind in ("bar", "line", "pie", "scatter", "area", "doughnut"):
        before = len(X.list_charts())
        try:
            X.add_chart("A1:C2", kind)
            ok = len(X.list_charts()) == before + 1
        except X.ExcelError:
            ok = False
        check(f"chart type '{kind}' is valid in Excel", ok)

    check("aliases resolve", X.add_chart("A1:C2", "donut")["chart_type"] == "doughnut")

    try:
        X.add_chart("A1:C2", "hologram")
        check("unknown chart type is refused", False, "was accepted")
    except X.ExcelError as e:
        check("unknown chart type is refused", "don't know the chart type" in str(e), str(e)[:80])

    n = X.delete_charts()
    check("deletes charts", n > 0 and X.list_charts() == [], f"{n} deleted")

    # Through the tool layer, the way the agent calls it.
    out, err = T.run_tool("create_chart_live",
                          {"range": "A1:C2", "chart_type": "bar", "title": "Via tool"})
    check("create_chart_live tool works", not err and len(X.list_charts()) == 1, out[:120])
    X.delete_charts()


    # ---- offline edit commands (the "make changes for me" feature) ---------
    from chatbot import offline

    def ask(q):
        ev = []
        offline.handle(q, lambda k, pl: ev.append((k, pl)))
        return (" ".join(pl.get("text", "") for k, pl in ev if k == "text"),
                [pl["name"] for k, pl in ev if k == "tool_start"])

    # Direct cell write.
    t, used = ask("set A5 to Widget")
    check("edit: sets a text value", X.read_cell("A5")["value"] == "Widget",
          X.read_cell("A5")["value"])
    check("edit: reports the change", "A5" in t, t[:100])

    t, _ = ask("set C5 to =B2*3")
    check("edit: writes a formula", X.read_cell("C5")["formula"] == "=B2*3",
          X.read_cell("C5")["formula"])
    check("edit: reports the calculated result", "21" in t, t[:140])

    t, _ = ask("B2 = 9")
    check("edit: bare assignment syntax works", X.read_cell("B2")["value"] == "9",
          X.read_cell("B2")["value"])
    check("edit: mentions undo when overwriting", "⌘Z" in t, t[:140])

    t, _ = ask("format B2:B3 as currency")
    check("edit: applies a number format", "$" in X.read_cell("B2")["number_format"],
          X.read_cell("B2")["number_format"])

    t, _ = ask("format B2 as bananas")
    check("edit: unknown format is refused, not guessed",
          "know these formats" in t, t[:120])

    # Auto-fix against a workbook with a real planted bug.
    import openpyxl

    fx = Path("/tmp/sd_fix_test.xlsx")
    wb2 = openpyxl.Workbook(); w = wb2.active; w.title = "F"
    w["A1"] = "Rate"; w["B1"] = 0.1
    for r in range(3, 6):
        w[f"A{r}"] = f"i{r}"; w[f"B{r}"] = 100 * r
        w[f"C{r}"] = f"=B{r}*B1"          # unlocked rate
    w["B7"] = "=SUM(B3:B4)"                # misses B5
    wb2.save(fx)
    X.open_file(fx, reload=True)

    t, used = ask("what would you fix?")
    check("autofix: preview lists changes", "would change" in t, t[:120])
    check("autofix: preview proposes the $ lock", "$B$1" in t, t[:200])
    check("autofix: preview does NOT write yet",
          X.read_cell("C3")["formula"] == "=B3*B1", X.read_cell("C3")["formula"])

    t, used = ask("fix it")
    check("autofix: applies the $ lock", X.read_cell("C3")["formula"] == "=B3*$B$1",
          X.read_cell("C3")["formula"])
    check("autofix: fixes every cell in the run",
          X.read_cell("C5")["formula"] == "=B5*$B$1", X.read_cell("C5")["formula"])
    check("autofix: extends the short SUM",
          X.read_cell("B7")["formula"] == "=SUM(B3:B5)", X.read_cell("B7")["formula"])
    check("autofix: mentions undo", "⌘Z" in t, t[:160])

    from chatbot import tools as T2

    out, _ = T2.run_tool("audit_live_workbook", {})
    check("autofix: workbook audits clean afterwards",
          json.loads(out)["counts"]["error"] == 0, out[:160])

    t, _ = ask("fix it")
    check("autofix: nothing left to fix is stated plainly",
          "Nothing I can fix" in t, t[:120])

    try:
        X.run_script(X._tell(
            f'  close (every workbook whose name is "{fx.name}") saving no\n  return "ok"'))
    except X.ExcelError:
        pass
    fx.unlink(missing_ok=True)
    X.open_file(fixture, reload=True)

    # ---- tables and new workbooks ----------------------------------------
    # Both were added after AppleScript "save workbook as" proved unusable on
    # this machine, so they must keep working without it.
    X.write_block("A20", [["Region", "Q1", "Q2"],
                          ["North", 120, 150],
                          ["South", 95, 110]])
    check("write_block writes a whole block in one call",
          X.read_grid("A20:C22").values[1][0] == "North",
          str(X.read_grid("A20:C22").values[:2]))
    check("write_block tidies numbers",
          X.read_grid("A20:C22").values[1][1] == "120",
          str(X.read_grid("A20:C22").values[1]))

    tname = X.make_table("A20:C22")
    check("make_table creates a native Excel Table", bool(tname), str(tname))

    r = X.create_table(["Item", "Qty", "Cost", "Total"],
                       [["Bolt", 4, 2.5, "=B31*C31"], ["Nut", 9, 1.25, "=B32*C32"]],
                       at="A30", number_formats=[None, "#,##0", "$#,##0.00", "$#,##0.00"])
    check("create_table reports its range", r["range"] == "A30:D32", str(r))
    check("create_table registers an Excel Table", bool(r["excel_table"]), str(r))
    check("create_table writes real formulas, not text",
          X.read_cell("D31")["formula"] == "=B31*C31", X.read_cell("D31")["formula"])
    check("create_table formulas calculate", X.read_cell("D31")["value"] == "10",
          X.read_cell("D31")["value"])
    check("create_table applies number formats",
          "$" in X.read_cell("C31")["number_format"], X.read_cell("C31")["number_format"])

    hdr_only = X.create_table(["Alpha", "Beta"], [], at="F30")
    check("headers-only table still becomes an Excel Table",
          bool(hdr_only["excel_table"]), str(hdr_only))

    # new_workbook must not touch Excel's save-as path.
    made = X.new_workbook(Path("/tmp/sd_newbook_test.xlsx"), sheet_name="Data",
                          headers=["Trial", "Mass"], rows=[["A", 1], ["B", 2]])
    check("new_workbook writes the file", Path(made["saved_to"]).exists(), str(made))
    check("new_workbook opens it in Excel",
          X.state().active_workbook == "sd_newbook_test.xlsx", str(X.state().active_workbook))
    check("new_workbook names the sheet", made["sheet"] == "Data", str(made))
    check("new_workbook seeds the headers",
          X.read_grid("A1:B3").values[0] == ["Trial", "Mass"], str(X.read_grid("A1:B3").values))

    check("tidy_number renders integers plainly", X.tidy_number("8.3758E+4") == "83758")
    check("tidy_number leaves text alone", X.tidy_number("Excellent") == "Excellent")
    check("tidy_number leaves error codes alone", X.tidy_number("#N/A") == "#N/A")

    try:
        X.run_script(X._tell(
            '  close (every workbook whose name is "sd_newbook_test.xlsx") saving no\n'
            '  return "ok"'))
    except X.ExcelError:
        pass
    Path("/tmp/sd_newbook_test.xlsx").unlink(missing_ok=True)
    X.open_file(fixture, reload=True)

    # Cleanup: close without saving further changes.
    try:
        X.run_script(X._tell(
            f'  close (every workbook whose name is "{fixture.name}") saving no\n  return "ok"'
        ))
    except X.ExcelError:
        pass
    fixture.unlink(missing_ok=True)
    return report()


def report():
    print(f"\n{'=' * 70}")
    for s in SKIP:
        print(f"  SKIP  {s}")
    for f in FAIL:
        print(f"  FAIL  {f}")
    print(f"{'=' * 70}")
    print(f"  {len(PASS)} passed, {len(FAIL)} failed, {len(SKIP)} skipped")
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
