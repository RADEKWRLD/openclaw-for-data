#!/usr/bin/env bash
set -euo pipefail

SKILL_DIR="/data/workspace/skills/engineering-price-collector"
VENV_DIR="/data/workspace/skill-data/.venv/engineering-price-collector"
DATA_DIR="/data/workspace/skill-data"

# Allow override for local development
if [ ! -d "$SKILL_DIR" ]; then
  SKILL_DIR="$(cd "$(dirname "$0")" && pwd)"
fi

mkdir -p "$DATA_DIR"

# Install CJK fonts if missing (for chart rendering)
if command -v fc-list &>/dev/null && ! fc-list | grep -qi "cjk\|noto.*sc\|simhei"; then
  echo "[run.sh] Installing CJK fonts for chart rendering..."
  apt-get update -qq && apt-get install -y -qq fonts-noto-cjk 2>/dev/null || true
fi

# Create/reuse virtualenv in writable location
if [ ! -d "$VENV_DIR" ]; then
  echo "[run.sh] Creating virtual environment at $VENV_DIR..."
  if command -v uv &>/dev/null; then
    uv venv "$VENV_DIR"
    uv pip install --python "$VENV_DIR/bin/python" -r "$SKILL_DIR/requirements.txt"
  else
    python3 -m venv "$VENV_DIR"
    "$VENV_DIR/bin/pip" install -r "$SKILL_DIR/requirements.txt"
  fi
fi

exec "$VENV_DIR/bin/python" -m src.main "$@"
