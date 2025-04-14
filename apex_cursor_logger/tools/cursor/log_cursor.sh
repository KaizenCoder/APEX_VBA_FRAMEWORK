#!/bin/bash
if [ "$#" -lt 2 ]; then
  echo "Usage: ./log_cursor.sh \"Prompt ici\" \"Réponse ici\" [Agent] [Note]"
  exit 1
fi

PROMPT="$1"
RESPONSE="$2"
AGENT=${3:-GPT-4}
NOTE=${4:-+}

SCRIPT_DIR="$(dirname "$0")"
python3 "$SCRIPT_DIR/cursor-autolog.py" "$PROMPT" "$RESPONSE" "$AGENT" "$NOTE"
