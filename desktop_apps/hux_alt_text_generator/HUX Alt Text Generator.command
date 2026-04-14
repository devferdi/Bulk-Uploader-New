#!/bin/zsh

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"

PYTHON_BIN=""

for candidate in \
  "$REPO_ROOT/.venv/bin/python" \
  "$REPO_ROOT/venv/bin/python" \
  "$REPO_ROOT/backend/.venv/bin/python"
do
  if [ -x "$candidate" ]; then
    PYTHON_BIN="$candidate"
    break
  fi
done

if [ -z "$PYTHON_BIN" ]; then
  PYTHON_BIN="$(command -v python3)"
fi

cd "$SCRIPT_DIR"
exec "$PYTHON_BIN" "$SCRIPT_DIR/app.py"
