#!/bin/bash
# Configure a model provider for SheetsDemon.
# Keys are read silently (never echoed, never written to shell history) and
# saved to a file only you can read.

set -u
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
CONFIG_DIR="$HOME/.config/sheets-doctor"
ENV_FILE="$CONFIG_DIR/env"
PY="$DIR/.venv/bin/python"

show_list() {
  cd "$DIR" || exit 1
  "$PY" - <<'PY'
import sys
sys.path.insert(0, ".")
from chatbot import providers
rows = list(providers.PRESETS.values())
w = max(len(p.key) for p in rows)
print()
print("  Providers SheetsDemon can use")
print("  " + "─" * 62)
for p in rows:
    tag = "FREE" if p.free else "paid"
    print(f"  {p.key.ljust(w)}  [{tag}]  {p.label} — {p.model}")
    print(f"  {' ' * w}          {p.note}")
    if p.trains_on_data:
        print(f"  {' ' * w}          NOTE: free tier trains on your prompts.")
print()
PY
}

KEY_URLS_gemini="https://aistudio.google.com/apikey"
KEY_URLS_groq="https://console.groq.com/keys"
KEY_URLS_mistral="https://console.mistral.ai/api-keys"
KEY_URLS_cerebras="https://cloud.cerebras.ai"
KEY_URLS_openrouter="https://openrouter.ai/keys"
KEY_URLS_anthropic="https://console.anthropic.com"
KEY_URLS_nvidia="https://build.nvidia.com"

if [ "${1:-}" = "--list" ] || [ "${1:-}" = "-l" ]; then
  show_list
  exit 0
fi

PROVIDER="${1:-}"
if [ -z "$PROVIDER" ]; then
  show_list
  printf "  Which provider? (e.g. gemini, groq, ollama, anthropic): "
  read -r PROVIDER
fi
PROVIDER="$(echo "$PROVIDER" | tr '[:upper:]' '[:lower:]' | tr -d '[:space:]')"

# Validate against the real preset list and fetch its env var name.
cd "$DIR" || exit 1
INFO=$("$PY" - "$PROVIDER" <<'PY'
import sys
sys.path.insert(0, ".")
from chatbot import providers
key = sys.argv[1]
p = providers.PRESETS.get(key)
if not p:
    print("INVALID")
else:
    print(f"{p.env_var}|{p.label}|{p.model}|{p.key_prefix}|{p.key_hint}")
PY
)

if [ "$INFO" = "INVALID" ]; then
  echo
  echo "  '$PROVIDER' is not a known provider. Run with --list to see the options."
  exit 1
fi

ENV_VAR="$(printf '%s' "$INFO" | cut -d'|' -f1)"
LABEL="$(printf  '%s' "$INFO" | cut -d'|' -f2)"
MODEL="$(printf  '%s' "$INFO" | cut -d'|' -f3)"
PREFIX="$(printf '%s' "$INFO" | cut -d'|' -f4)"
HINT="$(printf   '%s' "$INFO" | cut -d'|' -f5-)"

mkdir -p "$CONFIG_DIR"; chmod 700 "$CONFIG_DIR"
touch "$ENV_FILE"; chmod 600 "$ENV_FILE"

echo
echo "  Setting up: $LABEL ($MODEL)"
echo

if [ -z "$ENV_VAR" ]; then
  # Ollama — no key, but it must be installed and running.
  echo "  Ollama needs no API key, but it must be installed and running:"
  echo "    1. Install:  brew install ollama    (or download from ollama.com)"
  echo "    2. Start:    ollama serve"
  echo "    3. Pull the model:  ollama pull $MODEL"
  echo
else
  URL_VAR="KEY_URLS_${PROVIDER}"
  URL="${!URL_VAR:-}"
  [ -n "$URL" ] && echo "  Get a key at: $URL" && echo
  echo "  Paste the key below. It will not appear on screen."
  echo
  printf "  API key: "
  read -rs KEY
  echo; echo

  if [ -z "$KEY" ]; then
    echo "  Nothing entered. Nothing was saved."
    exit 1
  fi

  # Catch a wrong-credential paste here rather than as an opaque API error later.
  if [ -n "$PREFIX" ]; then
    case "$KEY" in
      "$PREFIX"*) ;;
      *)
        echo "  That doesn't look like a $LABEL API key."
        [ -n "$HINT" ] && echo "  $HINT"
        echo
        echo "  Nothing was saved. Copy the key again and re-run this."
        exit 1
        ;;
    esac
  fi

  # Write the key, replacing any previous value for this variable.
  TMP="$(mktemp)"; chmod 600 "$TMP"
  grep -v "^${ENV_VAR}=" "$ENV_FILE" 2>/dev/null > "$TMP" || true
  printf '%s=%s\n' "$ENV_VAR" "$KEY" >> "$TMP"
  mv "$TMP" "$ENV_FILE"; chmod 600 "$ENV_FILE"
  unset KEY
  echo "  Key saved to $ENV_FILE (readable only by you)."
fi

# Record the provider choice.
"$PY" - "$PROVIDER" <<'PY'
import sys
sys.path.insert(0, ".")
from chatbot import config
config.write_setting("SHEETDEMON_PROVIDER", sys.argv[1])
PY
echo "  Provider set to '$PROVIDER'."
echo
printf "  Testing it... "

RESULT=$("$PY" - <<'PY' 2>&1
import sys
sys.path.insert(0, ".")
from chatbot import config
config.load_env_file()
from chatbot import agent, providers
p = agent.active_preset()
try:
    prov = providers.build(p, "medium", 64)
    hist = [{"role": "user", "content": "Reply with the single word OK."}]
    turn = prov.run_turn("You are a test harness. Reply with exactly: OK", [], hist)
    print("OK" if turn.text else "EMPTY")
except providers.ProviderError as e:
    print(f"FAIL:{e}")
except Exception as e:
    print(f"FAIL:{type(e).__name__}: {e}")
PY
)

case "$RESULT" in
  OK)
    echo "works."
    echo
    echo "  Start SheetsDemon:  $DIR/run-chatbot.sh"
    ;;
  EMPTY)
    echo "responded but empty."
    echo "  Probably fine — start SheetsDemon and try a real question."
    ;;
  FAIL:AUTH)
    echo "key rejected."
    echo
    echo "  The key was saved but $LABEL says it's invalid. Check you copied all of"
    echo "  it, or create a fresh key and run this again."
    ;;
  *)
    echo "couldn't verify."
    echo "  ${RESULT#FAIL:}"
    ;;
esac
echo
