"""Exercise every chatbot tool the way Claude would call it, plus the HTTP layer.

Runs without an API key: the agent loop is the only part that needs one, and its
tool dispatch is what this covers.
"""

import json
import shutil
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from chatbot import tools as T  # noqa: E402
from chatbot.providers import PRESETS as T_PRESETS

PASS, FAIL = [], []


def check(name, cond, detail=""):
    (PASS if cond else FAIL).append(name + (f" — {detail}" if detail and not cond else ""))


def call(tool, **args):
    out, err = T.run_tool(tool, args)
    return out, err


def jcall(tool, **args):
    """Call a tool and parse its JSON, failing loudly if it errored."""
    out, err = call(tool, **args)
    if err:
        return {"_error": out}
    try:
        return json.loads(out)
    except Exception as e:
        return {"_error": f"bad JSON: {e}: {out[:200]}"}


def main():
    tmp = Path(tempfile.mkdtemp())
    T.WORKSPACE = tmp

    # ---- schemas are well-formed ------------------------------------------
    names = set()
    for t in T.TOOLS:
        check(f"{t['name']}: has description", len(t.get("description", "")) > 40)
        check(f"{t['name']}: schema is an object",
              t["input_schema"].get("type") == "object")
        check(f"{t['name']}: unique name", t["name"] not in names)
        names.add(t["name"])
        for prop, spec in t["input_schema"].get("properties", {}).items():
            check(f"{t['name']}.{prop}: typed", "type" in spec)
    # Too many tools degrades selection, especially on smaller free models. The
    # live-Excel group added six; past ~20 the set should be split behind a
    # discovery tool rather than grown further.
    check("tool count is manageable", len(T.TOOLS) <= 24, f"{len(T.TOOLS)} tools")

    # ---- every declared tool actually dispatches ---------------------------
    # new_spreadsheet has no required arguments, so calling it here would really
    # create a workbook and pull Excel to the foreground mid-test. It is covered
    # by test_live_excel.py instead.
    for t in T.TOOLS:
        if t["name"] == "new_spreadsheet":
            continue
        out, err = call(t["name"])  # no args: should error usefully, never crash
        check(f"{t['name']}: dispatches without crashing",
              "Unknown tool" not in out, out[:120])

    # ---- path traversal is blocked -----------------------------------------
    out, err = call("inspect_workbook", file="../../../etc/passwd")
    check("path traversal rejected", err or "escapes" in out.lower() or "not found" in out.lower(),
          out[:120])

    # ---- empty workspace ----------------------------------------------------
    r = jcall("list_workbooks")
    check("list_workbooks on empty workspace", r.get("workbooks") == [], str(r)[:120])

    # ---- generation tools ---------------------------------------------------
    r = jcall("create_dataset", kind="sales", filename="sales.xlsx", rows=12)
    check("create_dataset writes a file", (tmp / "sales.xlsx").exists(), str(r)[:200])
    check("create_dataset reports the filename", r.get("created_file") == "sales.xlsx", str(r)[:200])

    r = jcall("create_practice", topic="vlookup", filename="prac.xlsx")
    check("create_practice writes a file", (tmp / "prac.xlsx").exists(), str(r)[:200])
    check("create_practice returns tasks",
          any("Task" in (f.get("message") or "") for f in r.get("findings", [])), str(r)[:200])

    r = jcall("list_workbooks")
    check("list_workbooks finds them", set(r.get("workbooks", [])) >= {"sales.xlsx", "prac.xlsx"},
          str(r)[:150])

    # ---- read tools on generated data ---------------------------------------
    r = jcall("inspect_workbook", file="sales.xlsx")
    check("inspect returns findings", len(r.get("findings", [])) > 3, str(r)[:200])
    check("inspect names columns",
          any("Revenue" in str(f) for f in r.get("findings", [])), str(r)[:200])

    r = jcall("explain_pivot", file="sales.xlsx", goal="total revenue by region")
    msgs = " ".join(f.get("message", "") for f in r.get("findings", []))
    check("pivot maps Rows", "Rows" in msgs, msgs[:200])
    check("pivot maps Values", "Values" in msgs, msgs[:200])

    r = jcall("recommend_chart", file="sales.xlsx")
    msgs = " ".join(f.get("message", "") for f in r.get("findings", []))
    check("chart returns a best choice", "BEST CHOICE" in msgs, msgs[:200])

    r = jcall("check_formatting", file="sales.xlsx")
    check("check_formatting runs", "_error" not in r, str(r)[:200])

    # ---- audit finds planted bugs -------------------------------------------
    sys.path.insert(0, str(ROOT / "tests"))
    import make_broken

    broken = make_broken.build()
    shutil.copy(broken, tmp / "broken.xlsx")

    r = jcall("audit_formulas", file="broken.xlsx")
    msgs = " ".join(f.get("message", "") for f in r.get("findings", []))
    check("audit finds the missing $", "nothing locks it there" in msgs, msgs[:250])
    check("audit finds the VLOOKUP problem", "VLOOKUP" in msgs, msgs[:250])
    check("audit reports error counts", r.get("counts", {}).get("error", 0) >= 5, str(r.get("counts")))
    check("audit collapses repeats",
          any(f.get("affected_cells", 0) > 1 for f in r.get("findings", [])), str(r)[:250])

    for check_name in ("hardcoded", "references", "ranges", "lookups"):
        r = jcall("audit_formulas", file="broken.xlsx", checks=check_name)
        check(f"audit checks={check_name}", "_error" not in r, str(r)[:150])

    # ---- tutoring -------------------------------------------------------------
    r = jcall("socratic_hint", file="broken.xlsx", cell="F2")
    msgs = " ".join(f.get("message", "") for f in r.get("findings", []))
    check("hint gives tier 1", "Tier 1" in msgs, msgs[:200])
    check("hint gives all three tiers", msgs.count("Tier") >= 3, msgs[:200])

    r = jcall("socratic_hint", file="broken.xlsx", cell="F2", tier=1)
    msgs = " ".join(f.get("message", "") for f in r.get("findings", []))
    check("hint tier=1 withholds later tiers", "Tier 2" not in msgs, msgs[:200])

    r = jcall("socratic_hint", formula="=VLOOKUP(A2,E2:G11,3)")
    check("hint works with no file", "_error" not in r, str(r)[:200])

    r = jcall("explain_error", code="#N/A")
    check("explain_error works", "_error" not in r and len(r.get("findings", [])) > 3, str(r)[:150])

    # ---- conditional formatting dry-run must not write ------------------------
    before = (tmp / "sales.xlsx").read_bytes()
    r = jcall("conditional_formatting", file="sales.xlsx", sheet="Sales",
              range="F2:F13", preset="above", value="10", dry_run=True)
    check("cf dry_run does not modify the file", (tmp / "sales.xlsx").read_bytes() == before)
    check("cf dry_run reports hits", "would highlight" in json.dumps(r), str(r)[:200])

    r = jcall("conditional_formatting", file="sales.xlsx", sheet="Sales",
              range="F2:F13", preset="above", value="10")
    check("cf applies and saves", (tmp / "sales.xlsx.bak").exists(), str(r)[:200])

    # ---- build + edit ----------------------------------------------------------
    spec = {
        "sheets": [{
            "name": "Budget",
            "cells": {"A1": {"v": "Q1 Budget", "bold": True, "size": 14}},
            "table": {
                "at": "A3",
                "headers": ["Category", "Budget", "Actual", "Variance"],
                "rows": [["Rent", 1200, 1150, "=B4-C4"],
                         ["Travel", 800, 940, "=B5-C5"]],
                "formats": [None, "currency", "currency", "currency"],
                "totals": [None, "SUM", "SUM", "SUM"],
            },
            "conditional": [{"range": "D4:D5", "preset": "negative_red"}],
            "charts": [{"type": "column", "data": "B3:C5", "cats": "A4:A5",
                        "at": "F3", "title": "Budget vs Actual"}],
        }]
    }
    r = jcall("build_workbook", filename="budget.xlsx", spec=json.dumps(spec))
    check("build_workbook creates the file", (tmp / "budget.xlsx").exists(), str(r)[:200])

    import openpyxl

    wb = openpyxl.load_workbook(tmp / "budget.xlsx")
    ws = wb["Budget"]
    check("build writes the title", ws["A1"].value == "Q1 Budget")
    check("build writes formulas", str(ws["D4"].value).startswith("="))
    check("build adds the totals row", str(ws["B6"].value).upper().startswith("=SUM"))
    check("build adds the chart", len(ws._charts) == 1)

    r = jcall("edit_workbook", file="budget.xlsx", sheet="Budget", operation="set",
              cells=json.dumps({"E3": "Note", "E4": "=B4*2"}))
    wb = openpyxl.load_workbook(tmp / "budget.xlsx")
    check("edit set writes a value", wb["Budget"]["E3"].value == "Note")
    check("edit set writes a formula", wb["Budget"]["E4"].value == "=B4*2")
    check("edit makes a backup", (tmp / "budget.xlsx.bak").exists())

    r = jcall("edit_workbook", file="budget.xlsx", sheet="Budget", operation="fill",
              source="E4", through="E5")
    wb = openpyxl.load_workbook(tmp / "budget.xlsx")
    check("edit fill translates references", wb["Budget"]["E5"].value == "=B5*2",
          str(wb["Budget"]["E5"].value))

    r = jcall("edit_workbook", file="budget.xlsx", sheet="Budget", operation="add_row",
              values=json.dumps(["Software", 450, 512, None]))
    wb = openpyxl.load_workbook(tmp / "budget.xlsx")
    ws = wb["Budget"]
    check("add_row inserts above the totals", ws["A6"].value == "Software", str(ws["A6"].value))
    check("add_row translates the formula", ws["D6"].value == "=B6-C6", str(ws["D6"].value))
    check("add_row extends the total", ws["B7"].value == "=SUM(B4:B6)", str(ws["B7"].value))

    r = jcall("edit_workbook", file="budget.xlsx", sheet="Budget", operation="update_row",
              row_match="Category=Travel", updates=json.dumps({"Actual": 999}))
    wb = openpyxl.load_workbook(tmp / "budget.xlsx")
    check("update_row matches by column value", wb["Budget"]["C5"].value == 999,
          str(wb["Budget"]["C5"].value))

    r = jcall("edit_workbook", file="budget.xlsx", sheet="Budget", operation="add_chart",
              chart_type="bar", data_range="B3:C5", categories="A4:A5", anchor="F20")
    wb = openpyxl.load_workbook(tmp / "budget.xlsx")
    check("add_chart inserts a second chart", len(wb["Budget"]._charts) == 2)

    # ---- bad input is reported, never crashes ---------------------------------
    out, err = call("audit_formulas", file="nope.xlsx")
    check("missing file reports usefully", err and "not found" in out.lower(), out[:120])

    out, err = call("edit_workbook", file="budget.xlsx", operation="banana")
    check("unknown operation reported", "banana" in out, out[:120])

    out, err = call("build_workbook", filename="x.xlsx", spec="{not json")
    check("bad JSON spec reported as error", err, out[:120])

    # ---- HTTP layer ------------------------------------------------------------
    from fastapi.testclient import TestClient
    from chatbot import server

    server.WORKSPACE = tmp
    server.toolmod.WORKSPACE = tmp
    client = TestClient(server.app)

    r = client.get("/")
    check("index page serves", r.status_code == 200 and "SheetsDemon" in r.text)

    r = client.get("/api/status")
    check("status endpoint works", r.status_code == 200 and "model" in r.json())
    # Which model depends on the machine (a local Ollama is auto-detected), so
    # assert it reports a real model from the preset table rather than one name.
    check("status reports the model",
          r.json()["model"] in {p.model for p in T_PRESETS.values()})

    # ---- security regressions --------------------------------------------
    # A loopback bind does not stop a hostile page: DNS rebinding makes the
    # browser treat it as same-origin. Unexpected Host headers must be refused.
    r_evil = client.get("/api/status", headers={"Host": "evil.example.com"})
    check("rejects foreign Host (DNS rebinding)", r_evil.status_code == 421,
          str(r_evil.status_code))
    for good in ("localhost:8765", "127.0.0.1:8765"):
        rg = client.get("/api/status", headers={"Host": good})
        check(f"accepts Host {good}", rg.status_code == 200, str(rg.status_code))

    # Cell values and filenames are interpolated into title="..." attributes,
    # so the client-side escaper must cover quotes, not just angle brackets.
    page = (ROOT / "chatbot" / "static" / "index.html").read_text()
    esc_line = next((ln for ln in page.splitlines() if "const esc =" in ln), "")
    esc_body = page[page.index("const esc ="):page.index("const esc =") + 260]
    check("esc() escapes double quotes", "&quot;" in esc_body, esc_line[:90])
    check("esc() escapes single quotes", "&#39;" in esc_body, esc_line[:90])

    r = client.get("/api/workbooks")
    check("workbooks endpoint lists files", any(
        w["name"] == "budget.xlsx" for w in r.json()["workbooks"]))

    data = (tmp / "budget.xlsx").read_bytes()
    r = client.post("/api/upload", files={"file": ("uploaded.xlsx", data)})
    check("upload accepts .xlsx", r.status_code == 200 and (tmp / "uploaded.xlsx").exists())

    r = client.post("/api/upload", files={"file": ("evil.exe", b"MZ")})
    check("upload rejects non-xlsx", r.status_code == 400)

    r = client.post("/api/upload", files={"file": ("../../escape.xlsx", data)})
    check("upload strips path components", not (tmp.parent / "escape.xlsx").exists())

    r = client.get("/api/download/budget.xlsx")
    check("download serves the file", r.status_code == 200 and len(r.content) > 1000)

    r = client.get("/api/download/nope.xlsx")
    check("download 404s on missing file", r.status_code == 404)

    r = client.get("/api/download/..%2F..%2Fetc%2Fpasswd")
    check("download blocks traversal", r.status_code in (404, 400), str(r.status_code))

    # ---- auth failure is actionable, not a raw SDK traceback -------------------
    # The SDK raises a bare TypeError from header validation at REQUEST time when
    # no credential resolves; users must see the fix, not that.
    import os

    from chatbot import agent

    saved = {k: os.environ.pop(k, None) for k in ("ANTHROPIC_API_KEY", "ANTHROPIC_AUTH_TOKEN")}
    # A local Ollama would legitimately be picked up here; this block is about
    # having no provider at all, so pin the probe off.
    from chatbot import providers as _P
    _real_probe = _P.ollama_running
    _P.ollama_running = lambda ttl=0: False
    try:
        ev = []
        agent.run([{"role": "user", "content": "hi"}], lambda k, p: ev.append((k, p)))
        kinds = [k for k, _ in ev]
        msg = " ".join(str(p) for _, p in ev)
        # With no provider the chatbot must still WORK — it answers from the
        # engine offline rather than refusing.
        check("no-credentials answers instead of erroring",
              "text" in kinds and "error" not in kinds, str(kinds))
        check("offline reply points at setup for full conversation",
              "setup-key.sh" in msg, msg[:180])
        check("no-credentials output is not a raw traceback",
              "TypeError" not in msg and "Traceback" not in msg, msg[:180])
    finally:
        _P.ollama_running = _real_probe
        for k, v in saved.items():
            if v is not None:
                os.environ[k] = v

    # ---- config file loading (never touches a real key) ------------------------
    from chatbot import config

    envf = tmp / "env"
    envf.write_text("# comment\n\nexport ANTHROPIC_API_KEY=\"sk-ant-dummy\"\n")
    saved2 = os.environ.pop("ANTHROPIC_API_KEY", None)
    try:
        check("config loads a key from file", config.load_env_file(envf))
        check("config parses export+quotes",
              os.environ.get("ANTHROPIC_API_KEY") == "sk-ant-dummy",
              str(os.environ.get("ANTHROPIC_API_KEY")))
        check("config reports credentials available", config.has_credentials())
        # An already-set environment variable must win over the file.
        os.environ["ANTHROPIC_API_KEY"] = "sk-ant-from-env"
        config.load_env_file(envf)
        check("environment wins over the config file",
              os.environ["ANTHROPIC_API_KEY"] == "sk-ant-from-env")
    finally:
        os.environ.pop("ANTHROPIC_API_KEY", None)
        if saved2 is not None:
            os.environ["ANTHROPIC_API_KEY"] = saved2
    check("config handles a missing file", config.load_env_file(tmp / "nope") is False)

    # ---- offline mode: the chatbot must work with NO provider key ----------
    from chatbot import offline

    T.WORKSPACE = tmp
    saved3 = {k: os.environ.pop(k, None) for k in ("ANTHROPIC_API_KEY", "GEMINI_API_KEY",
                                                   "GROQ_API_KEY", "SHEETDEMON_PROVIDER")}
    try:
        def run_offline(q):
            ev = []
            offline.handle(q, lambda k, p: ev.append((k, p)))
            text = " ".join(p.get("text", "") for k, p in ev if k == "text")
            tools_used = [p["name"] for k, p in ev if k == "tool_start"]
            return text, tools_used

        t, used = run_offline("what does #N/A mean")
        check("offline explains an error code", "explain_error" in used, str(used))
        check("offline error text is substantive", "lookup" in t.lower() and len(t) > 200, t[:120])

        t, used = run_offline("check my homework for mistakes")
        check("offline routes to the audit", "audit_formulas" in used, str(used))

        t, used = run_offline("give me a hint for D5")
        check("offline routes to a hint", "socratic_hint" in used, str(used))
        check("offline hint shows tiers", "Tier 1" in t, t[:120])

        t, used = run_offline("what chart should I use")
        check("offline recommends a chart", "recommend_chart" in used, str(used))

        # Regression: the two-word forms ("column headers") used to be an
        # optional suffix on the keyword, which swallowed a first column
        # literally called Name / Title / Heading.
        from chatbot import offline as _off
        for phrase, want in [
            ("make a table with columns Name, Score, Grade", ["Name", "Score", "Grade"]),
            ("make a table with columns Title, Author", ["Title", "Author"]),
            ("make a table with columns Heading, Value", ["Heading", "Value"]),
            ("set up column headers: Student, Midterm", ["Student", "Midterm"]),
            ("create a table with field names Ref, Amount", ["Ref", "Amount"]),
        ]:
            got = _off._columns_from(phrase)
            check(f"columns parsed: {phrase[:38]}", got == want, f"{got} != {want}")

        t, used = run_offline("make me a practice sheet on vlookup")
        check("offline builds practice", "create_practice" in used, str(used))
        check("offline practice lists tasks", "1." in t, t[:150])

        t, used = run_offline("how do I set up a pivot table")
        check("offline explains pivots", "explain_pivot" in used, str(used))
        check("offline pivot gives steps", "Steps" in t or "Drag" in t, t[:150])

        t, used = run_offline("make me a sales dataset")
        check("offline generates data", "create_dataset" in used, str(used))

        t, _ = run_offline("what is the meaning of life")
        check("offline admits its limits on off-topic input",
              "language model" in t.lower() or "keyword" in t.lower(), t[:150])

        # And the agent must fall through to offline rather than erroring.
        from chatbot import providers as _P2
        _real2 = _P2.ollama_running
        _P2.ollama_running = lambda ttl=0: False
        try:
            ev = []
            agent.run([{"role": "user", "content": "check my work"}],
                      lambda k, p: ev.append((k, p)))
        finally:
            _P2.ollama_running = _real2
        kinds = [k for k, _ in ev]
        check("agent falls back to offline with no key",
              "text" in kinds and "error" not in kinds, str(kinds))
    finally:
        for k, v in saved3.items():
            if v is not None:
                os.environ[k] = v

    print(f"\n{'=' * 70}")
    for f in FAIL:
        print(f"  FAIL  {f}")
    print(f"{'=' * 70}")
    print(f"  {len(PASS)} passed, {len(FAIL)} failed")
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
