#!/usr/bin/env bash
set -euo pipefail

APP_NAME="PITC-VA-Mailer"
PROJECT_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
DIST_DIR="$PROJECT_ROOT/dist"
PARENT_DIR="$(dirname "$PROJECT_ROOT")"
OUT_DIR="$PARENT_DIR/PITC-mailer"

cd "$PROJECT_ROOT"

echo "Building $APP_NAME..."
VENV_PYTHON="$PROJECT_ROOT/.venv/bin/python"
PYTHON_BIN="${PYTHON_BIN:-python}"

if [[ -x "$VENV_PYTHON" ]]; then
  PYTHON_BIN="$VENV_PYTHON"
elif ! command -v "$PYTHON_BIN" >/dev/null 2>&1; then
  if command -v python3 >/dev/null 2>&1; then
    PYTHON_BIN="python3"
  else
    echo "Error: python/python3 not found in PATH. Activate your virtual environment first."
    exit 1
  fi
fi

pyinstaller_args=(--noconfirm --windowed --name "$APP_NAME" --hidden-import tzdata)
if [[ "$(uname -s)" != "Darwin" ]]; then
  pyinstaller_args+=(--onefile)
fi

"$PYTHON_BIN" -m PyInstaller "${pyinstaller_args[@]}" main.py

if [[ ! -d "$DIST_DIR" ]]; then
  echo "Error: dist folder not found after build: $DIST_DIR"
  exit 1
fi

mkdir -p "$OUT_DIR"
rm -rf "$OUT_DIR/$APP_NAME.app"
rm -f "$OUT_DIR/$APP_NAME" "$OUT_DIR/$APP_NAME.exe"

# Copy platform-specific build artifacts.
shopt -s nullglob
exe_files=("$DIST_DIR"/*.exe)
app_bundles=("$DIST_DIR"/*.app)
binary_files=("$DIST_DIR"/"$APP_NAME")
copied_any=0

if (( ${#exe_files[@]} > 0 )); then
  cp -f "$DIST_DIR"/*.exe "$OUT_DIR"/
  copied_any=1
fi

if (( ${#app_bundles[@]} > 0 )); then
  for app_bundle in "${app_bundles[@]}"; do
    cp -R "$app_bundle" "$OUT_DIR"/
  done
  copied_any=1
fi

if [[ -f "${binary_files[0]}" ]]; then
  cp -f "${binary_files[0]}" "$OUT_DIR"/
  copied_any=1
fi

if (( copied_any == 0 )); then
  echo "Warning: no app artifacts found in $DIST_DIR (expected .exe, .app, or $APP_NAME)"
fi

# Copy only sample env file
if [[ -f "$DIST_DIR/.env-sample" ]]; then
  cp -f "$DIST_DIR/.env-sample" "$OUT_DIR"/
else
  echo "Warning: .env-sample not found in $DIST_DIR"
fi

echo "Done. Copied files to: $OUT_DIR"
ls -la "$OUT_DIR"
