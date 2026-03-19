#!/usr/bin/env bash
set -euo pipefail

PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
DIST_DIR="$PROJECT_ROOT/dist"
PARENT_DIR="$(dirname "$PROJECT_ROOT")"
OUT_DIR="$PARENT_DIR/PITC-mailer"

cd "$PROJECT_ROOT"

echo "Building PITC-VA-Mailer.exe..."
python -m PyInstaller --onefile --windowed --name "PITC-VA-Mailer" --hidden-import tzdata main.py

if [[ ! -d "$DIST_DIR" ]]; then
  echo "Error: dist folder not found after build: $DIST_DIR"
  exit 1
fi

mkdir -p "$OUT_DIR"

# Copy only executable(s)
shopt -s nullglob
exe_files=("$DIST_DIR"/*.exe)
if (( ${#exe_files[@]} == 0 )); then
  echo "Warning: no .exe files found in $DIST_DIR"
else
  cp -f "$DIST_DIR"/*.exe "$OUT_DIR"/
fi

# Copy only sample env file
if [[ -f "$DIST_DIR/.env-sample" ]]; then
  cp -f "$DIST_DIR/.env-sample" "$OUT_DIR"/
else
  echo "Warning: .env-sample not found in $DIST_DIR"
fi

echo "Done. Copied files to: $OUT_DIR"
ls -la "$OUT_DIR"
