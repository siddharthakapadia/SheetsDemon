#!/bin/bash
# SheetsDemon chatbot launcher.
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
if [ ! -f "$HOME/.config/sheets-doctor/env" ] \
   && [ -z "${ANTHROPIC_API_KEY:-}" ] && [ -z "${ANTHROPIC_AUTH_TOKEN:-}" ]; then
  echo
  echo "  No API key set up yet. Run this first:"
  echo "    $DIR/setup-key.sh"
  echo
fi
cd "$DIR" || exit 1
# --reload so editing the code takes effect without a manual restart; running a
# stale build is otherwise invisible and looks exactly like "the feature doesn't work".
exec "$DIR/.venv/bin/uvicorn" chatbot.server:app \
  --host 127.0.0.1 --port 8765 --log-level warning \
  --reload --reload-dir "$DIR/chatbot"
