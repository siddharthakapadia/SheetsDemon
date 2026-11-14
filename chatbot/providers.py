"""Model providers for SheetsDemon.

The engine is deterministic Python; the model only decides *which* engine tool to
call. That makes the model swappable, so SheetsDemon can run on a paid Anthropic
key or on a free tier.

Two adapters cover everything: Anthropic's native API, and the OpenAI Chat
Completions shape — which Gemini, Groq, Mistral, Cerebras, OpenRouter, Together
and a local Ollama all speak. Adding a provider is usually just a PRESETS entry.
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from typing import Any

# ==========================================================================
# Provider presets
# ==========================================================================


@dataclass(frozen=True)
class Preset:
    key: str
    label: str
    kind: str  # "anthropic" | "openai_compat"
    model: str
    env_var: str
    base_url: str | None = None
    free: bool = False
    note: str = ""
    # Free tiers usually train on your prompts; say so where it applies.
    trains_on_data: bool = False
    # What a valid key looks like, so a wrong-credential paste is caught at
    # setup rather than surfacing later as an opaque API error.
    key_prefix: str = ""
    key_hint: str = ""


PRESETS: dict[str, Preset] = {
    "anthropic": Preset(
        key="anthropic", label="Anthropic Claude", kind="anthropic",
        model="claude-opus-5", env_var="ANTHROPIC_API_KEY",
        note="Best tool-use quality. Paid, billed per token.",
        key_prefix="sk-ant-", key_hint="Anthropic keys start with 'sk-ant-'.",
    ),
    "gemini": Preset(
        key="gemini", label="Google Gemini", kind="openai_compat",
        model="gemini-2.5-flash", env_var="GEMINI_API_KEY",
        base_url="https://generativelanguage.googleapis.com/v1beta/openai/",
        free=True, trains_on_data=True,
        note="Most capable free tier. Generous daily limit, solid function calling.",
        key_prefix="AIza", key_hint="Google AI Studio keys start with 'AIza'. A string starting 'AQ.' or 'ya29.' is an OAuth token, not an API key — get the key from aistudio.google.com/apikey.",
    ),
    "groq": Preset(
        key="groq", label="Groq", kind="openai_compat",
        model="llama-3.3-70b-versatile", env_var="GROQ_API_KEY",
        base_url="https://api.groq.com/openai/v1",
        free=True, trains_on_data=True,
        note="Very fast. Generous daily request limit.",
        key_prefix="gsk_", key_hint="Groq keys start with 'gsk_'.",
    ),
    "mistral": Preset(
        key="mistral", label="Mistral", kind="openai_compat",
        model="mistral-large-latest", env_var="MISTRAL_API_KEY",
        base_url="https://api.mistral.ai/v1",
        free=True, trains_on_data=True,
        note="Large monthly free quota; requires opting into data training.",
    ),
    "cerebras": Preset(
        key="cerebras", label="Cerebras", kind="openai_compat",
        model="llama-3.3-70b", env_var="CEREBRAS_API_KEY",
        base_url="https://api.cerebras.ai/v1",
        free=True, trains_on_data=True, note="Fast free tier.",
    ),
    "openrouter": Preset(
        key="openrouter", label="OpenRouter", kind="openai_compat",
        model="meta-llama/llama-3.3-70b-instruct:free", env_var="OPENROUTER_API_KEY",
        base_url="https://openrouter.ai/api/v1",
        free=True, trains_on_data=True,
        note="Router with some ':free' models. Quality varies by model.",
        key_prefix="sk-or-", key_hint="OpenRouter keys start with 'sk-or-'.",
    ),
    "nvidia": Preset(
        key="nvidia", label="NVIDIA NIM", kind="openai_compat",
        model="meta/llama-3.3-70b-instruct", env_var="NVIDIA_API_KEY",
        base_url="https://integrate.api.nvidia.com/v1",
        free=True, trains_on_data=True,
        note="Free credits, 100+ models, OpenAI-compatible. Good tool calling.",
        key_prefix="nvapi-", key_hint="NVIDIA keys start with 'nvapi-'. Get one at build.nvidia.com.",
    ),
    "ollama": Preset(
        key="ollama", label="Ollama (local)", kind="openai_compat",
        model="qwen2.5:7b", env_var="",  # no key needed
        base_url="http://localhost:11434/v1",
        free=True,
        note="Runs on this Mac. No key, no limits, works offline. Weakest tool use.",
    ),
}


# Result of the last local-Ollama probe, with the time it was taken. Probing is
# a socket connect to localhost, but resolve() runs on every request, so cache it.
_OLLAMA_PROBE: tuple[float, bool] = (0.0, False)
_OLLAMA_TTL = 20.0


def ollama_running(ttl: float = _OLLAMA_TTL) -> bool:
    """Is a local Ollama server accepting connections?

    Ollama needs no key and no account, so if it is up it is a better answer
    than telling the student to go and sign up for something.
    """
    global _OLLAMA_PROBE
    import socket
    import time
    from urllib.parse import urlparse

    now = time.monotonic()
    checked_at, result = _OLLAMA_PROBE
    if now - checked_at < ttl:
        return result

    url = urlparse(PRESETS["ollama"].base_url)
    try:
        with socket.create_connection((url.hostname or "127.0.0.1",
                                       url.port or 11434), timeout=0.35):
            result = True
    except OSError:
        result = False
    _OLLAMA_PROBE = (now, result)
    return result


def resolve() -> Preset:
    """Pick the active provider: explicit choice, else the first with a key."""
    chosen = os.environ.get("SHEETDEMON_PROVIDER", "").strip().lower()
    if chosen in PRESETS:
        p = PRESETS[chosen]
        return _with_overrides(p)
    # Auto-detect: prefer Anthropic, then any free provider whose key is present.
    for name in ("anthropic", "gemini", "groq", "nvidia", "mistral",
                 "cerebras", "openrouter"):
        if os.environ.get(PRESETS[name].env_var):
            return _with_overrides(PRESETS[name])
    # No key anywhere. A local Ollama beats falling back to keyword matching.
    if ollama_running():
        return _with_overrides(PRESETS["ollama"])
    return _with_overrides(PRESETS["anthropic"])


def _with_overrides(p: Preset) -> Preset:
    """Let SHEETDEMON_MODEL / SHEETDEMON_BASE_URL override a preset."""
    model = os.environ.get("SHEETDEMON_MODEL", "").strip() or p.model
    base = os.environ.get("SHEETDEMON_BASE_URL", "").strip() or p.base_url
    if model == p.model and base == p.base_url:
        return p
    return Preset(**{**p.__dict__, "model": model, "base_url": base})


def has_key(p: Preset) -> bool:
    """Ollama needs no key; everything else does."""
    return True if not p.env_var else bool(os.environ.get(p.env_var))


def free_options() -> list[Preset]:
    return [p for p in PRESETS.values() if p.free]


# ==========================================================================
# Turn result
# ==========================================================================


@dataclass
class ToolCall:
    id: str
    name: str
    input: dict[str, Any]


@dataclass
class Turn:
    text: list[str] = field(default_factory=list)
    thinking: list[str] = field(default_factory=list)
    tool_calls: list[ToolCall] = field(default_factory=list)
    stop_reason: str = "end_turn"  # end_turn | tool_use | max_tokens | refusal


class ProviderError(Exception):
    """Carries a message already written for the end user."""


# Providers disagree on how they report a bad key: Gemini answers 400 "Please
# pass a valid API key", NVIDIA 403 "Authorization failed", most others 401.
# Any of them should reach the same actionable message, never a raw status code.
_AUTH_PHRASES = (
    "api key", "api_key", "unauthorized", "invalid authentication",
    "authorization failed", "forbidden", "invalid api",
)


def is_auth_error(status: int | None, text: str) -> bool:
    blob = (text or "").lower()
    return status in (400, 401, 403) and any(p in blob for p in _AUTH_PHRASES)


# ==========================================================================
# Anthropic
# ==========================================================================


class AnthropicProvider:
    kind = "anthropic"

    def __init__(self, preset: Preset, effort: str, max_tokens: int):
        self.preset = preset
        self.effort = effort
        self.max_tokens = max_tokens
        import anthropic

        self._sdk = anthropic
        self._client = anthropic.Anthropic()

    def user_message(self, text: str) -> dict:
        return {"role": "user", "content": text}

    def run_turn(self, system: str, tools: list[dict], history: list) -> Turn:
        a = self._sdk
        try:
            r = self._client.messages.create(
                model=self.preset.model,
                max_tokens=self.max_tokens,
                system=[{"type": "text", "text": system,
                         "cache_control": {"type": "ephemeral"}}],
                thinking={"type": "adaptive", "display": "summarized"},
                output_config={"effort": self.effort},
                tools=tools,
                messages=history,
            )
        except a.AuthenticationError as e:
            raise ProviderError("AUTH") from e
        except a.RateLimitError as e:
            raise ProviderError("Rate limited by Anthropic. Wait a moment and retry.") from e
        except a.APIStatusError as e:
            raise ProviderError(f"Anthropic API error {e.status_code}: {e.message}") from e
        except a.APIConnectionError as e:
            raise ProviderError("Could not reach the Anthropic API. Check your connection.") from e
        except TypeError as e:
            if "authentication" in str(e).lower():
                raise ProviderError("AUTH") from e
            raise ProviderError(f"Request rejected: {e}") from e

        turn = Turn()
        if r.stop_reason == "refusal":
            turn.stop_reason = "refusal"
            return turn

        for b in r.content:
            if b.type == "thinking" and b.thinking:
                turn.thinking.append(b.thinking)
            elif b.type == "text" and b.text:
                turn.text.append(b.text)
            elif b.type == "tool_use":
                turn.tool_calls.append(ToolCall(b.id, b.name, dict(b.input)))

        history.append({"role": "assistant", "content": r.content})
        turn.stop_reason = (
            "tool_use" if r.stop_reason == "tool_use"
            else "max_tokens" if r.stop_reason == "max_tokens"
            else "end_turn"
        )
        return turn

    def append_results(self, history: list, results: list[tuple[ToolCall, str, bool]]) -> None:
        history.append({"role": "user", "content": [
            {"type": "tool_result", "tool_use_id": c.id,
             "content": out, "is_error": err}
            for c, out, err in results
        ]})


# ==========================================================================
# OpenAI-compatible (Gemini, Groq, Mistral, Cerebras, OpenRouter, Ollama…)
# ==========================================================================


class OpenAICompatProvider:
    kind = "openai_compat"

    def __init__(self, preset: Preset, effort: str, max_tokens: int):
        self.preset = preset
        self.max_tokens = max_tokens
        self.effort = effort
        # Set once a provider rejects reasoning parameters, so the retry only
        # happens on the first turn rather than every message.
        self._no_reasoning = False
        import openai

        self._sdk = openai
        self._client = openai.OpenAI(
            # Ollama accepts any non-empty key.
            api_key=os.environ.get(preset.env_var) or "not-needed",
            base_url=preset.base_url,
            timeout=120.0,
        )

    def user_message(self, text: str) -> dict:
        return {"role": "user", "content": text}

    @staticmethod
    def convert_tools(tools: list[dict]) -> list[dict]:
        """Anthropic tool schema -> OpenAI function schema."""
        return [{
            "type": "function",
            "function": {
                "name": t["name"],
                "description": t["description"],
                "parameters": t["input_schema"],
            },
        } for t in tools]

    def run_turn(self, system: str, tools: list[dict], history: list) -> Turn:
        o = self._sdk
        messages = [{"role": "system", "content": system}, *history]
        kwargs: dict[str, Any] = {
            "model": self.preset.model,
            "max_tokens": self.max_tokens,
            "messages": messages,
            "tools": self.convert_tools(tools),
        }
        # Most OpenAI-compatible reasoning models take reasoning_effort; the ones
        # that do not answer 400. Try it, and remember a refusal so we only pay
        # for the round trip once.
        if not self._no_reasoning:
            kwargs["reasoning_effort"] = self.effort
        try:
            try:
                r = self._client.chat.completions.create(**kwargs)
            except o.BadRequestError:
                if self._no_reasoning:
                    raise
                self._no_reasoning = True
                kwargs.pop("reasoning_effort", None)
                r = self._client.chat.completions.create(**kwargs)
        except o.AuthenticationError as e:
            raise ProviderError("AUTH") from e
        except o.RateLimitError as e:
            raise ProviderError(
                f"{self.preset.label} rate limit reached (free tiers cap daily requests). "
                "Wait a while, or switch provider."
            ) from e
        except o.APIStatusError as e:
            # Providers disagree on the status for a bad key: Gemini answers 400
            # with "pass a valid API key" where others send 401. Treat any
            # key-shaped complaint as an auth failure so the user gets the fix
            # rather than a raw status code.
            if is_auth_error(e.status_code, f"{e.message} {getattr(e, 'body', '')}"):
                raise ProviderError("AUTH") from e
            # Ollama is installed and running but the model was never downloaded.
            # The raw 404 tells a student nothing; the pull command tells them
            # everything.
            if (self.preset.key == "ollama" and e.status_code == 404
                    and "not found" in f"{e.message}".lower()):
                raise ProviderError(
                    f"Ollama is running, but the model {self.preset.model} has not been "
                    f"downloaded yet. In a terminal run:\n\n"
                    f"    ollama pull {self.preset.model}\n\n"
                    "It is a few GB and only needs doing once. Then just ask me again."
                ) from e
            raise ProviderError(f"{self.preset.label} API error {e.status_code}: {e.message}") from e
        except o.APIConnectionError as e:
            hint = (
                " Is Ollama running? Start it with: ollama serve"
                if self.preset.key == "ollama" else ""
            )
            raise ProviderError(f"Could not reach {self.preset.label}.{hint}") from e

        turn = Turn()
        if not r.choices:
            raise ProviderError(f"{self.preset.label} returned no choices.")
        choice = r.choices[0]
        msg = choice.message

        # Providers name the reasoning trace differently: reasoning_content on
        # NVIDIA/DeepSeek, reasoning on OpenRouter.
        for attr in ("reasoning_content", "reasoning"):
            trace = getattr(msg, attr, None)
            if isinstance(trace, str) and trace.strip():
                turn.thinking.append(trace.strip())
                break

        if getattr(msg, "content", None):
            turn.text.append(msg.content)

        calls = getattr(msg, "tool_calls", None) or []
        for c in calls:
            raw = c.function.arguments or "{}"
            try:
                args = json.loads(raw)
            except json.JSONDecodeError:
                # Smaller models sometimes emit malformed JSON. Surface it as a
                # tool error rather than crashing the turn.
                args = {"__malformed_arguments__": raw}
            if not isinstance(args, dict):
                args = {"__malformed_arguments__": raw}
            turn.tool_calls.append(ToolCall(c.id, c.function.name, args))

        # Echo the assistant turn back verbatim-ish so the model sees its own calls.
        entry: dict[str, Any] = {"role": "assistant", "content": msg.content or ""}
        if calls:
            entry["tool_calls"] = [{
                "id": c.id, "type": "function",
                "function": {"name": c.function.name, "arguments": c.function.arguments or "{}"},
            } for c in calls]
        history.append(entry)

        turn.stop_reason = (
            "tool_use" if calls
            else "max_tokens" if choice.finish_reason == "length"
            else "end_turn"
        )
        return turn

    def append_results(self, history: list, results: list[tuple[ToolCall, str, bool]]) -> None:
        for c, out, err in results:
            history.append({
                "role": "tool",
                "tool_call_id": c.id,
                "content": f"ERROR: {out}" if err else out,
            })


def build(preset: Preset, effort: str, max_tokens: int):
    if preset.kind == "anthropic":
        return AnthropicProvider(preset, effort, max_tokens)
    return OpenAICompatProvider(preset, effort, max_tokens)
