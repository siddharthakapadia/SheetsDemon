"""Credential loading.

The chatbot needs an Anthropic API key. Requiring it to be exported in the same
terminal that launches the server is fragile — the server is often started by
something else (an IDE, a launch config, a double-clicked script) that does not
inherit that shell's environment. So the key is read from a file the user owns,
and the environment still wins if it is set.
"""

from __future__ import annotations

import os
from pathlib import Path

# Everything that needs to know where the project lives derives it from this
# file's own location. Hardcoding an install path (it used to assume
# ~/Desktop/sheets-doctor) breaks on any machine that checked the repo out
# somewhere else — and silently writes new workbooks into a folder the user
# never created.
PROJECT_ROOT = Path(__file__).resolve().parent.parent
WORKSPACE_DIR = PROJECT_ROOT / "workspace"
SETUP_SCRIPT = PROJECT_ROOT / "setup-key.sh"


def setup_cmd() -> str:
    """The exact command this user should run, for whatever path they cloned to."""
    try:
        return f"~/{SETUP_SCRIPT.relative_to(Path.home())}"
    except ValueError:
        return str(SETUP_SCRIPT)


ENV_FILE = Path.home() / ".config" / "sheets-doctor" / "env"

KEY_VARS = (
    "ANTHROPIC_API_KEY", "ANTHROPIC_AUTH_TOKEN",
    "GEMINI_API_KEY", "GROQ_API_KEY", "MISTRAL_API_KEY",
    "CEREBRAS_API_KEY", "OPENROUTER_API_KEY", "NVIDIA_API_KEY",
)

# Settings (not secrets) that may also live in the config file.
SETTING_VARS = (
    "SHEETDEMON_PROVIDER", "SHEETDEMON_MODEL", "SHEETDEMON_BASE_URL",
    "SHEETDEMON_EFFORT", "SHEETDEMON_MAX_TOKENS",
)


def load_env_file(path: Path = ENV_FILE) -> bool:
    """Load KEY=VALUE lines into the environment. Returns True if a key was found.

    An already-set environment variable is never overwritten.
    """
    if not path.is_file():
        return False
    try:
        for raw in path.read_text(encoding="utf-8").splitlines():
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            # Tolerate `export FOO=...` and surrounding quotes.
            if key.startswith("export "):
                key = key[len("export "):].strip()
            value = value.strip().strip('"').strip("'")
            if key and value and not os.environ.get(key):
                os.environ[key] = value
    except OSError:
        return False
    return any(os.environ.get(v) for v in KEY_VARS)


def has_credentials() -> bool:
    """Is a credential available for any supported provider?"""
    if any(os.environ.get(v) for v in KEY_VARS):
        return True
    cfg = Path(os.environ.get("ANTHROPIC_CONFIG_DIR", Path.home() / ".config" / "anthropic"))
    return (cfg / "credentials").is_dir()


def write_setting(name: str, value: str, path: Path = ENV_FILE) -> None:
    """Persist one KEY=VALUE into the config file, replacing any existing line.

    Used for non-secret settings (e.g. provider choice). Secrets are written by
    setup-key.sh so their values never pass through here.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    path.parent.chmod(0o700)
    lines = path.read_text(encoding="utf-8").splitlines() if path.is_file() else []
    kept = [ln for ln in lines if not ln.strip().startswith(f"{name}=")]
    kept.append(f"{name}={value}")
    path.write_text("\n".join(kept) + "\n", encoding="utf-8")
    path.chmod(0o600)
    os.environ[name] = value
