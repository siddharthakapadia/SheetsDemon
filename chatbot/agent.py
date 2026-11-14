"""The agent loop behind SheetsDemon.

Provider-agnostic: the loop decides *when* to run tools, and a provider adapter
handles talking to whichever model is configured. See providers.py.

A manual tool-use loop rather than an SDK helper: the browser needs an event for
each tool call *as it happens* (so the user sees "auditing your formulas…"
rather than a long silence).
"""

from __future__ import annotations

import os
from collections.abc import Callable
from typing import Any

from . import config, providers
from . import tools as toolmod

# Pick up keys from the user's config file at import time, so the server works
# however it was launched — not only from a shell that exported the variables.
config.load_env_file()

# Spreadsheet answers are wrong in ways a student cannot spot, so accuracy beats
# latency here: think hard by default. Reasoning-capable free models get this as
# reasoning_effort; the rest ignore it.
EFFORT = os.environ.get("SHEETDEMON_EFFORT", os.environ.get("SHEETDOC_EFFORT", "high"))

# On Anthropic, thinking and response share this budget — leave real headroom.
MAX_TOKENS = int(os.environ.get("SHEETDEMON_MAX_TOKENS", "16000"))

MAX_TURNS = 12  # tool round-trips per user message

SYSTEM = """You are SheetsDemon, an Excel expert tutoring a student through an \
introduction to spreadsheets course at JGU. The course uses Microsoft Excel (.xlsx).

You have tools that parse formulas properly, simulate fills, and check ranges against \
real data. Use them instead of reading formulas yourself and guessing — reading a \
formula and eyeballing it is how wrong answers happen. Run inspect_workbook first on \
any workbook you have not seen this conversation.

## Look before you answer

Never answer a question about the student's workbook from the conversation alone. Before your first substantive claim about their data, ground yourself: call excel_status to see what is actually open and selected, then read the cells you are about to talk about. If you are about to say a formula is wrong, read that formula first. If you are about to say a range is wrong, read the data around it first.

Think the problem through before replying. Spreadsheet errors are quiet — a wrong range still returns a number, and a relative reference still fills without complaint. The student cannot tell a confident wrong answer from a right one, so guessing costs them marks. When you are unsure, run the tool that settles it rather than hedging in prose.

If a tool result contradicts what you were about to say, say the corrected thing.

## Working alongside Excel

The student usually has Excel open next to you. When they say "this cell", "my sheet", \
"what I'm looking at", or ask you to change something they are working on, call \
excel_status first — it tells you the active workbook, sheet and their selected cell.

With a workbook open in Excel, prefer the live tools:
- read_excel_live sees unsaved edits AND real calculated results. The file-based tools \
only see what Excel cached at last save, so on an actively edited workbook they can be \
stale or blank.
- write_excel_live puts your change into their workbook immediately — they watch it \
appear. This is the collaborative move: make the edit, then say in one line what you \
changed and why.
- audit_live_workbook saves their workbook and runs the full engine audit on it.
- create_chart_live puts a real chart on their sheet. When they ask for a chart, build it — do not just describe the menu clicks. If unsure which type, call recommend_chart first and say in one line why you chose it.

Two rules about their live work. Say what you are about to change before you change it \
— you are editing a file they have open, not a scratch copy. And audit_live_workbook \
saves first, so if they have unsaved changes they care about, mention that.

If Excel is not running or nothing is open, fall back to the workspace files and say so.

## Tutor mode vs do-it mode

Match what the student actually asked for. Both are legitimate:
- "why is this wrong", "I'm stuck", "give me a hint" -> use socratic_hint, give them \
tier 1, and let them try. Do not lead with the answer.
- "fix it", "build this", "do it for me" -> do the work, then say what changed and why \
in one or two sentences.

When you fix something, name the concept once ("this needed $B$15 because it is one rate \
shared by every row"). That sentence is the whole value of the course.

## Working rules

Writes are real: any tool that changes a workbook makes a .bak first. Before editing a \
file the student cares about, say what you are about to change. Use dry_run on \
conditional_formatting to check a rule before committing it.

Tool results are JSON. Read them and explain the findings in the student's own terms — \
quote the specific cells and formulas that matter and give the corrected formula. Never \
paste raw tool output back at them.

If a workbook is clean, say so plainly and stop.

## Communicating

Keep responses focused and brief; a spreadsheet question usually wants a direct answer, \
not a structured report. Lead with what is wrong and where. Supporting detail comes \
after, for a student who wants it. Use prose, not headers and bullet walls, unless you \
are genuinely listing several independent findings.

Deliver what the student asked for, at the scope they intended. Make routine judgment \
calls yourself; check in only when different readings would lead to materially \
different work. Do not widen the task — if you notice something else worth fixing, \
mention it in a sentence rather than fixing it unasked."""


# The full prompt is written for a frontier model. A 7B model given ~900 words of
# nuance and 22 tools loses the thread — it re-calls the same tool, or emits raw
# tool-call JSON as prose. Small models get the rules that matter and nothing else.
SYSTEM_COMPACT = """You are SheetsDemon, an Excel tutor for a university spreadsheets course.

Rules:
1. To learn anything about the user's workbook, call a tool. Never guess a formula or a value.
2. Call excel_status first when they mention "this cell", "my sheet", or ask you to change something.
3. read_excel_live reads ANY address you ask for — "H12", "A1:D20". The user does NOT need to select a cell first. NEVER ask the user to select a cell; if they name one, read it.
4. Call each tool ONCE. After a tool returns, use its result — do not call it again with the same arguments.
5. When you have the answer, reply in plain English. Never write tool-call JSON, function names, or XML tags in your reply.
6. Be brief. Two or three sentences unless they asked for detail.
7. If they ask you to fix or build something, do it with the tools, then say what you changed in one sentence."""


def active_preset() -> providers.Preset:
    config.load_env_file()
    return providers.resolve()


def credentials_available() -> bool:
    return providers.has_key(active_preset())


def no_key_message(preset: providers.Preset | None = None) -> str:
    p = preset or active_preset()
    if p.key == "anthropic":
        return (
            "No API key for Anthropic yet. Either add one with  "
            f"{config.setup_cmd()}  or switch to a free provider with  "
            f"{config.setup_cmd()} --list  to see the options."
        )
    return (
        f"No API key for {p.label} yet. Run  {config.setup_cmd()}  "
        f"and choose {p.key}. It will tell you where to get a free key."
    )


def describe() -> dict[str, Any]:
    """Provider info for the UI status bar."""
    p = active_preset()
    return {
        "provider": p.key,
        "provider_label": p.label,
        "model": p.model,
        "free": p.free,
        "effort": EFFORT if p.key == "anthropic" else None,
        "credentials": providers.has_key(p),
    }


def run(messages: list[Any], emit: Callable[[str, dict], None]) -> None:
    """Run one user turn to completion, emitting events as work happens.

    `messages` is mutated in place so the caller keeps the updated history.
    Event types emitted: thinking, text, tool_start, tool_end, error, done.
    """
    preset = active_preset()

    if not providers.has_key(preset):
        # No provider configured. Rather than refusing, answer with the engine
        # directly — pattern-matched routing instead of a model. Less capable,
        # but it means the chat box works with no key, signup or network.
        from . import offline

        last_user = next(
            (m for m in reversed(messages) if m.get("role") == "user"), None)
        text = last_user.get("content") if last_user else ""
        if not isinstance(text, str):
            text = ""
        offline.handle(text, emit)
        emit("done", {"offline": True})
        return

    try:
        provider = providers.build(preset, EFFORT, MAX_TOKENS)
    except providers.ProviderError as e:
        emit("error", {"message": str(e)})
        return
    except Exception as e:
        emit("error", {"message": f"Could not start {preset.label}: {e}"})
        return

    # Frontier models get the full prompt; smaller hosted/local ones get the
    # compact rules, which measurably reduces repeated calls and leaked JSON.
    system = SYSTEM if preset.key == "anthropic" else SYSTEM_COMPACT

    seen_calls: dict[tuple[str, str], int] = {}

    for _ in range(MAX_TURNS):
        try:
            turn = provider.run_turn(system, toolmod.TOOLS, messages)
        except providers.ProviderError as e:
            msg = str(e)
            emit("error", {"message": no_key_message(preset) if msg == "AUTH" else msg})
            return
        except Exception as e:
            emit("error", {"message": f"{preset.label} request failed: {type(e).__name__}: {e}"})
            return

        if turn.stop_reason == "refusal":
            emit("error", {"message": (
                "The model declined this request. Rephrase and try again."
            )})
            return

        for t in turn.thinking:
            emit("thinking", {"text": t})
        for t in turn.text:
            cleaned = _strip_leaked_tool_json(t)
            if cleaned:
                emit("text", {"text": cleaned})

        if turn.stop_reason == "max_tokens" and not turn.tool_calls:
            emit("error", {"message": (
                "The reply hit the token limit and was cut off. Ask for a shorter "
                "answer, or split the request."
            )})
            return

        if not turn.tool_calls:
            emit("done", {})
            return

        # Run every requested tool, then hand all results back together.
        results = []
        for call in turn.tool_calls:
            import json as _json

            sig = (call.name, _json.dumps(call.input, sort_keys=True, default=str))
            seen_calls[sig] = seen_calls.get(sig, 0) + 1
            if seen_calls[sig] > 2:
                # Feed the repetition back as a tool result rather than running it
                # again; smaller models otherwise loop until MAX_TURNS.
                results.append((call, (
                    f"You have already called {call.name} with these arguments and "
                    "have the result above. Do not call it again — answer the user "
                    "now, in plain English."), True))
                emit("tool_start", {"name": call.name, "input": call.input})
                emit("tool_end", {"name": call.name, "error": True})
                continue
            emit("tool_start", {"name": call.name, "input": call.input})
            if call.name == "socratic_hint":
                from .stats import STATS
                STATS.record_hint()
            if "__malformed_arguments__" in call.input:
                out, err = (
                    "Your tool arguments were not valid JSON. Re-issue the call with "
                    "properly formed JSON arguments.", True,
                )
            else:
                out, err = toolmod.run_tool(call.name, call.input)
            emit("tool_end", {"name": call.name, "error": err})
            results.append((call, out, err))

        provider.append_results(messages, results)

    emit("error", {"message": (
        f"Stopped after {MAX_TURNS} tool rounds without finishing. "
        "Try breaking the request into smaller steps."
        + ("" if preset.key == "anthropic" else
           f" Smaller free models like {preset.model} sometimes loop — a more "
           "specific request usually helps.")
    )})



def _strip_leaked_tool_json(text: str) -> str:
    """Remove tool-call JSON and stray tags a small model emits as prose.

    Qwen-class models sometimes answer with {"name": ..., "arguments": ...} or
    wrap output in tags like <translation>. None of that is for the student.
    """
    import re

    cleaned = re.sub(r"</?(?:translation|tool_call|function_call|think)>", "", text)
    # Drop standalone JSON objects that are really tool calls.
    cleaned = re.sub(
        r'\{\s*"(?:name|function)"\s*:\s*"[^"]+"\s*,\s*"(?:arguments|parameters)"'
        r'\s*:\s*\{.*?\}\s*\}',
        "", cleaned, flags=re.S)
    return cleaned.strip()


def user_message(text: str) -> dict:
    """Shape a user message for the active provider's history format."""
    preset = active_preset()
    return {"role": "user", "content": text} if preset else {"role": "user", "content": text}


# Kept for backwards compatibility with earlier code and tests.
MODEL = active_preset().model
