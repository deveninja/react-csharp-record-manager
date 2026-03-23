#!/usr/bin/env bash
set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
VENV_PYTHON="$PROJECT_ROOT/.venv/bin/python"

if [[ ! -x "$VENV_PYTHON" ]]; then
  echo "Error: missing virtual environment at $VENV_PYTHON"
  echo "Create it first: /usr/local/opt/python@3.9/bin/python3.9 -m venv .venv"
  exit 1
fi

exec "$VENV_PYTHON" "$PROJECT_ROOT/main.py"
