"""FastAPI server for the SheetsDemon chatbot.

Runs locally. Upload a workbook, chat about it, download the result.
"""

from __future__ import annotations

import json
import queue
import threading
import uuid
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, StreamingResponse
from pydantic import BaseModel

from . import agent
from . import tools as toolmod
from .stats import STATS

ROOT = Path(__file__).resolve().parent
WORKSPACE = ROOT.parent / "workspace"
WORKSPACE.mkdir(exist_ok=True)
toolmod.WORKSPACE = WORKSPACE

app = FastAPI(title="SheetsDemon")

# This server binds to loopback, but that alone does not keep a web page out of
# it: a hostile site can point a domain it controls at 127.0.0.1 (DNS
# rebinding), at which point the browser treats it as same-origin and can drive
# every endpoint here — read the user's sheets, write to their open workbook,
# pull files out of the workspace. Rejecting unexpected Host headers closes it.
ALLOWED_HOSTS = {"localhost", "127.0.0.1", "[::1]", "::1", "0.0.0.0", "testserver"}


@app.middleware("http")
async def _block_rebinding(request, call_next):
    host = (request.headers.get("host") or "").rsplit(":", 1)[0].strip("[]").lower()
    if host and host not in {h.strip("[]") for h in ALLOWED_HOSTS}:
        from fastapi.responses import JSONResponse

        return JSONResponse({"detail": "Host not allowed."}, status_code=421)
    return await call_next(request)

# Conversation history, keyed by session id. In-memory: this is a local
# single-user app, so there is no store to keep in sync.
SESSIONS: dict[str, list] = {}

ALLOWED_SUFFIXES = {".xlsx", ".xlsm"}
MAX_UPLOAD_BYTES = 25 * 1024 * 1024


class ChatRequest(BaseModel):
    session_id: str
    message: str


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return (ROOT / "static" / "index.html").read_text(encoding="utf-8")


def _build_stamp() -> str:
    """Newest mtime across the package — lets the UI show if it is running stale code."""
    import time

    newest = max((f.stat().st_mtime for f in ROOT.glob("*.py")), default=0)
    return time.strftime("%H:%M:%S", time.localtime(newest))


@app.get("/api/status")
def status() -> dict:
    info = agent.describe()
    return {
        **info,
        "build": _build_stamp(),
        "tools": len(toolmod.TOOLS),
        "setup_hint": None if info["credentials"] else agent.no_key_message(),
        # The UI must not hardcode an install path — it differs per checkout.
        "setup_cmd": __import__("chatbot.config", fromlist=["x"]).setup_cmd(),
        "workbooks": _list_workbooks(),
    }


def _list_workbooks() -> list[dict]:
    out = []
    for p in sorted(WORKSPACE.glob("*.xlsx")):
        if p.name.startswith("~"):
            continue
        out.append({"name": p.name, "size": p.stat().st_size,
                    "modified": int(p.stat().st_mtime)})
    return out


@app.get("/api/workbooks")
def workbooks() -> dict:
    return {"workbooks": _list_workbooks()}


@app.post("/api/upload")
async def upload(file: UploadFile = File(...)) -> dict:
    name = Path(file.filename or "upload.xlsx").name  # strip any path components
    if Path(name).suffix.lower() not in ALLOWED_SUFFIXES:
        raise HTTPException(400, f"Only {', '.join(sorted(ALLOWED_SUFFIXES))} files are supported.")

    data = await file.read()
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(400, "File is larger than 25 MB.")

    target = (WORKSPACE / name).resolve()
    if not target.is_relative_to(WORKSPACE.resolve()):
        raise HTTPException(400, "Invalid filename.")

    target.write_bytes(data)
    return {"name": name, "size": len(data)}


@app.get("/api/download/{name}")
def download(name: str) -> FileResponse:
    target = (WORKSPACE / Path(name).name).resolve()
    if not target.is_relative_to(WORKSPACE.resolve()) or not target.exists():
        raise HTTPException(404, "No such workbook.")
    return FileResponse(
        target,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=target.name,
    )


@app.post("/api/reset")
def reset(req: ChatRequest) -> dict:
    SESSIONS.pop(req.session_id, None)
    return {"ok": True}


# ==========================================================================
# Live Excel + command-center panels
# ==========================================================================


@app.get("/api/live")
def live() -> dict:
    """What Excel is doing right now, for the header strip."""
    from . import excel_live as X

    st = X.state()
    return {
        "running": st.running,
        "workbook": st.active_workbook,
        "sheet": st.active_sheet,
        "sheets": st.sheets,
        "selection": st.selection,
        "used_range": st.used_range,
        "unsaved": st.unsaved,
        "error": st.error,
    }


@app.get("/api/grid")
def grid(range: str | None = None, sheet: str | None = None) -> dict:
    """Live cell grid for the centre panel."""
    from . import excel_live as X

    try:
        g = X.read_grid(range, sheet)
    except X.ExcelError as e:
        return {"error": str(e), "rows": []}
    return {
        "sheet": g.sheet, "address": g.address,
        "start_row": g.start_row, "start_col": g.start_col,
        "values": g.values, "formulas": g.formulas,
    }


@app.post("/api/excel/open")
def excel_open(payload: dict) -> dict:
    from . import excel_live as X

    name = Path(str(payload.get("file", ""))).name
    target = (WORKSPACE / name).resolve()
    if not target.is_relative_to(WORKSPACE.resolve()) or not target.exists():
        raise HTTPException(404, "No such workbook.")
    try:
        return {"opened": X.open_file(target)}
    except X.ExcelError as e:
        raise HTTPException(400, str(e)) from e


@app.post("/api/excel/new")
def excel_new(payload: dict | None = None) -> dict:
    """Create a fresh workbook and open it in Excel."""
    from . import excel_live as X

    payload = payload or {}
    raw = str(payload.get("name") or "").strip()
    name = Path(raw).name  # strip any path components
    if name and not name.lower().endswith((".xlsx", ".xlsm")):
        name += ".xlsx"
    target = (WORKSPACE / name).resolve() if name else None
    if target is not None and not target.is_relative_to(WORKSPACE.resolve()):
        raise HTTPException(400, "Invalid filename.")
    if target is not None and target.exists():
        raise HTTPException(409, f"{target.name} already exists — pick another name.")

    headers = payload.get("headers") or None
    if headers is not None and not isinstance(headers, list):
        raise HTTPException(400, "headers must be a list of column names.")
    try:
        return X.new_workbook(target, payload.get("sheet_name"), headers)
    except X.ExcelError as e:
        raise HTTPException(400, str(e)) from e


@app.get("/api/stats")
def stats() -> dict:
    return STATS.snapshot()


QUICK_ACTIONS = {
    "hardcoded": ("Detect hardcoded values", "audit_formulas", {"checks": "hardcoded"}),
    "references": ("Check reference types", "audit_formulas", {"checks": "references"}),
    "ranges": ("Validate range bounds", "audit_formulas", {"checks": "ranges"}),
    "lookups": ("Verify lookup syntax", "audit_formulas", {"checks": "lookups"}),
    "formats": ("Audit number formats", "check_formatting", {"what": "number_formats"}),
    "structure": ("Check header & grid", "check_formatting", {"what": "structure"}),
    "audit": ("Run full audit", "audit_formulas", {"checks": "all"}),
}


@app.get("/api/quick-actions")
def quick_actions() -> dict:
    return {"actions": [{"key": k, "label": v[0]} for k, v in QUICK_ACTIONS.items()]}


@app.post("/api/quick/{action}")
def quick(action: str, payload: dict | None = None) -> dict:
    """Run an engine check directly, without going through the model.

    These are deterministic, so they work with no API key at all — the buttons
    stay useful even before a provider is configured.
    """
    if action not in QUICK_ACTIONS:
        raise HTTPException(404, "Unknown action.")
    label, tool, args = QUICK_ACTIONS[action]
    payload = payload or {}

    from . import excel_live as X

    file = payload.get("file")
    if not file:
        # Prefer whatever is open in Excel; fall back to the newest workspace file.
        path = X.active_workbook_path()
        if path:
            try:
                X.save_active()
            except X.ExcelError:
                pass
            file = Path(path).name
            if not (WORKSPACE / file).exists():
                toolmod.WORKSPACE = Path(path).parent
        else:
            books = _list_workbooks()
            if not books:
                return {"label": label, "error": "No workbook open in Excel and none uploaded."}
            file = books[-1]["name"]

    out, err = toolmod.run_tool(tool, {**args, "file": file, **(
        {"sheet": payload["sheet"]} if payload.get("sheet") else {})})
    toolmod.WORKSPACE = WORKSPACE  # restore
    if err:
        return {"label": label, "error": out}
    data = json.loads(out)
    STATS.record_audit(data)
    return {"label": label, "file": file, **data}


@app.post("/api/chat")
def chat(req: ChatRequest) -> StreamingResponse:
    """Run a turn, streaming events to the browser as Server-Sent Events."""
    history = SESSIONS.setdefault(req.session_id, [])
    history.append({"role": "user", "content": req.message})

    events: queue.Queue = queue.Queue()

    def emit(kind: str, payload: dict) -> None:
        events.put({"type": kind, **payload})

    def work() -> None:
        try:
            agent.run(history, emit)
        except Exception as e:  # never leave the stream hanging
            emit("error", {"message": f"{type(e).__name__}: {e}"})
        finally:
            events.put(None)

    threading.Thread(target=work, daemon=True).start()

    def stream():
        while True:
            item = events.get()
            if item is None:
                yield f"data: {json.dumps({'type': 'end'})}\n\n"
                return
            yield f"data: {json.dumps(item, default=str)}\n\n"

    return StreamingResponse(
        stream(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


def main() -> None:
    import uvicorn

    info = agent.describe()
    tag = " (free)" if info["free"] else ""
    print(f"\n  SheetsDemon  →  http://127.0.0.1:8765")
    print(f"  Provider: {info['provider_label']} · {info['model']}{tag}\n")
    if not info["credentials"]:
        print("  ⚠  No API key set up yet. In another terminal, run:")
        from . import config as _cfg

        print(f"     {_cfg.setup_cmd()}\n")
    uvicorn.run(app, host="127.0.0.1", port=8765, log_level="warning")


if __name__ == "__main__":
    main()
