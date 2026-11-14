"""Verify the provider layer, including the free (OpenAI-compatible) path.

No real API keys are used. The OpenAI-compatible path is exercised against a
local mock server that speaks /v1/chat/completions, which is the same wire shape
Gemini, Groq, Mistral, Cerebras, OpenRouter and Ollama expose. That proves tool
conversion, tool-call parsing, history shaping and loop termination for real.
"""

import json
import os
import sys
import tempfile
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

PASS, FAIL = [], []


def check(name, cond, detail=""):
    (PASS if cond else FAIL).append(name + (f" — {detail}" if detail and not cond else ""))


# ==========================================================================
# Mock OpenAI-compatible server
# ==========================================================================

REQUESTS: list[dict] = []
SCRIPT: list[dict] = []  # queued responses, popped per request


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):  # silence
        pass

    def do_POST(self):
        body = self.rfile.read(int(self.headers.get("Content-Length", 0)))
        REQUESTS.append(json.loads(body))
        payload = SCRIPT.pop(0) if SCRIPT else {
            "choices": [{"message": {"role": "assistant", "content": "done"},
                         "finish_reason": "stop"}]
        }
        out = json.dumps({
            "id": "chatcmpl-mock", "object": "chat.completion",
            "created": 0, "model": "mock", **payload,
        }).encode()
        self.send_response(200)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(out)))
        self.end_headers()
        self.wfile.write(out)


def start_server() -> tuple[HTTPServer, int]:
    srv = HTTPServer(("127.0.0.1", 0), Handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    return srv, srv.server_port


def main():
    tmp = Path(tempfile.mkdtemp())

    # Isolate from any real config/keys on this machine.
    from chatbot import config

    config.ENV_FILE = tmp / "env"
    for v in config.KEY_VARS + config.SETTING_VARS:
        os.environ.pop(v, None)

    from chatbot import providers

    # ---- preset table sanity ---------------------------------------------
    check("has an anthropic preset", "anthropic" in providers.PRESETS)
    check("offers free providers", len(providers.free_options()) >= 4,
          str(len(providers.free_options())))
    for key, p in providers.PRESETS.items():
        check(f"{key}: kind is known", p.kind in ("anthropic", "openai_compat"))
        check(f"{key}: has a model", bool(p.model))
        if p.kind == "openai_compat":
            check(f"{key}: has a base_url", bool(p.base_url))
        check(f"{key}: has a note", len(p.note) > 10)
    check("ollama needs no key", providers.has_key(providers.PRESETS["ollama"]))
    check("gemini needs a key", not providers.has_key(providers.PRESETS["gemini"]))

    # ---- resolution ------------------------------------------------------
    # resolve() prefers a local Ollama when one is listening, so pin the probe
    # rather than letting the result depend on what this machine has installed.
    real_probe = providers.ollama_running

    providers.ollama_running = lambda ttl=0: False
    check("defaults to anthropic with no keys", providers.resolve().key == "anthropic")

    providers.ollama_running = lambda ttl=0: True
    check("auto-detects a running ollama", providers.resolve().key == "ollama")
    check("auto-detected ollama needs no key",
          providers.has_key(providers.resolve()))

    providers.ollama_running = real_probe

    os.environ["GROQ_API_KEY"] = "test-key"
    check("auto-detects a free key", providers.resolve().key == "groq",
          providers.resolve().key)

    os.environ["SHEETDEMON_PROVIDER"] = "gemini"
    check("explicit choice wins", providers.resolve().key == "gemini")

    os.environ["SHEETDEMON_MODEL"] = "custom-model-x"
    check("model override applies", providers.resolve().model == "custom-model-x")
    os.environ.pop("SHEETDEMON_MODEL")

    os.environ["SHEETDEMON_PROVIDER"] = "nonsense-provider"
    check("unknown provider falls back safely", providers.resolve().key == "groq",
          providers.resolve().key)

    # ---- tool schema conversion -----------------------------------------
    from chatbot import tools as toolmod

    conv = providers.OpenAICompatProvider.convert_tools(toolmod.TOOLS)
    check("converts every tool", len(conv) == len(toolmod.TOOLS))
    check("uses OpenAI function shape",
          all(c["type"] == "function" and "function" in c for c in conv))
    check("carries names", {c["function"]["name"] for c in conv}
          == {t["name"] for t in toolmod.TOOLS})
    check("carries descriptions",
          all(len(c["function"]["description"]) > 40 for c in conv))
    check("maps input_schema to parameters",
          conv[0]["function"]["parameters"] == toolmod.TOOLS[0]["input_schema"])

    # ---- full loop against the mock server ------------------------------
    srv, port = start_server()
    toolmod.WORKSPACE = tmp

    os.environ["SHEETDEMON_PROVIDER"] = "groq"
    os.environ["SHEETDEMON_BASE_URL"] = f"http://127.0.0.1:{port}/v1"
    os.environ["GROQ_API_KEY"] = "test-key"

    import importlib

    from chatbot import agent

    importlib.reload(agent)

    # Script: model asks for a dataset, then answers.
    REQUESTS.clear()
    SCRIPT.clear()
    SCRIPT.append({"choices": [{"message": {
        "role": "assistant", "content": "Let me generate that.",
        "tool_calls": [{"id": "call_1", "type": "function", "function": {
            "name": "create_dataset",
            "arguments": json.dumps({"kind": "sales", "filename": "t.xlsx", "rows": 5}),
        }}],
    }, "finish_reason": "tool_calls"}]})
    SCRIPT.append({"choices": [{"message": {
        "role": "assistant", "content": "Made a 5-row sales dataset."},
        "finish_reason": "stop"}]})

    events = []
    history = [{"role": "user", "content": "make me a sales dataset"}]
    agent.run(history, lambda k, p: events.append((k, p)))
    kinds = [k for k, _ in events]

    check("free path: emits tool_start", "tool_start" in kinds, str(kinds))
    check("free path: emits tool_end", "tool_end" in kinds, str(kinds))
    check("free path: emits text", "text" in kinds, str(kinds))
    check("free path: finishes cleanly", kinds[-1] == "done", str(kinds))
    check("free path: no error events", "error" not in kinds,
          str([p for k, p in events if k == "error"]))
    check("free path: tool actually ran", (tmp / "t.xlsx").exists())
    check("free path: made two requests", len(REQUESTS) == 2, str(len(REQUESTS)))

    # The request must carry converted tools and a system message.
    req = REQUESTS[0]
    check("free path: sends tools", len(req.get("tools", [])) == len(toolmod.TOOLS))
    check("free path: sends system as a message",
          req["messages"][0]["role"] == "system", str(req["messages"][0]["role"]))
    check("free path: system mentions SheetsDemon",
          "SheetsDemon" in req["messages"][0]["content"])

    # Second request must include the assistant tool_call and the tool result.
    req2 = REQUESTS[1]
    roles = [m["role"] for m in req2["messages"]]
    check("free path: echoes assistant tool_call", "assistant" in roles, str(roles))
    check("free path: sends tool result role", "tool" in roles, str(roles))
    tool_msg = next(m for m in req2["messages"] if m["role"] == "tool")
    check("free path: tool result carries the id",
          tool_msg.get("tool_call_id") == "call_1", str(tool_msg)[:120])
    check("free path: tool result carries output",
          "created_file" in tool_msg["content"], tool_msg["content"][:120])

    # ---- malformed tool arguments are handled, not fatal ------------------
    REQUESTS.clear()
    SCRIPT.clear()
    SCRIPT.append({"choices": [{"message": {
        "role": "assistant", "content": "",
        "tool_calls": [{"id": "call_bad", "type": "function", "function": {
            "name": "list_workbooks", "arguments": "{not valid json",
        }}],
    }, "finish_reason": "tool_calls"}]})
    SCRIPT.append({"choices": [{"message": {
        "role": "assistant", "content": "Retrying."}, "finish_reason": "stop"}]})

    events = []
    agent.run([{"role": "user", "content": "list files"}],
              lambda k, p: events.append((k, p)))
    kinds = [k for k, _ in events]
    check("malformed arguments do not crash", "done" in kinds, str(kinds))
    err_flags = [p.get("error") for k, p in events if k == "tool_end"]
    check("malformed arguments reported as a tool error", any(err_flags), str(err_flags))

    # ---- provider-specific auth wording maps to AUTH -----------------------
    for status, msg, want in (
        (403, "Authorization failed", True),          # NVIDIA
        (400, "Please pass a valid API key", True),   # Gemini
        (401, "Unauthorized", True),                  # most providers
        (403, "Forbidden", True),
        (429, "Rate limit exceeded", False),          # not auth
        (500, "Internal error", False),               # not auth
        (400, "model not found", False),              # a real 400, not auth
    ):
        check(f"auth wording: {status} {msg!r} -> {want}",
              providers.is_auth_error(status, msg) is want)

    # ---- unreachable provider gives a clean message -----------------------
    os.environ["SHEETDEMON_BASE_URL"] = "http://127.0.0.1:1/v1"  # nothing listening
    importlib.reload(agent)
    events = []
    agent.run([{"role": "user", "content": "hi"}], lambda k, p: events.append((k, p)))
    msgs = " ".join(str(p) for k, p in events if k == "error")
    check("unreachable provider emits an error", bool(msgs), str(events)[:150])
    check("unreachable error names the provider", "Groq" in msgs, msgs[:160])
    check("unreachable error is not a traceback",
          "Traceback" not in msgs, msgs[:160])

    # ---- ollama gets an actionable hint ----------------------------------
    os.environ["SHEETDEMON_PROVIDER"] = "ollama"
    # Point at a port nothing is listening on, so this tests the connection
    # failure even on a machine that really does run Ollama.
    os.environ["SHEETDEMON_BASE_URL"] = "http://127.0.0.1:11333/v1"
    importlib.reload(agent)
    events = []
    agent.run([{"role": "user", "content": "hi"}], lambda k, p: events.append((k, p)))
    msgs = " ".join(str(p) for k, p in events if k == "error")
    check("ollama-not-running says how to start it",
          "ollama serve" in msgs, msgs[:200])

    # ---- missing key for a chosen free provider --------------------------
    os.environ["SHEETDEMON_PROVIDER"] = "gemini"
    os.environ.pop("GEMINI_API_KEY", None)
    importlib.reload(agent)
    events = []
    agent.run([{"role": "user", "content": "hi"}], lambda k, p: events.append((k, p)))
    kinds = [k for k, _ in events]
    text = " ".join(p.get("text", "") for k, p in events if k == "text")
    # A missing key must not break the chat — it drops to offline engine mode.
    check("missing key falls back to offline, not an error",
          "text" in kinds and "error" not in kinds, str(kinds))
    check("offline mode points at setup for the full experience",
          "setup-key.sh" in text, text[:200])

    srv.shutdown()

    print(f"\n{'=' * 70}")
    for f in FAIL:
        print(f"  FAIL  {f}")
    print(f"{'=' * 70}")
    print(f"  {len(PASS)} passed, {len(FAIL)} failed")
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
