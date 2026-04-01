#!/usr/bin/env bash
set -e
PYTHON=${PYTHON:-$(command -v python3.10 || command -v python3.11 || command -v python3.9 || command -v python3 || command -v python)}
if [ -z "$PYTHON" ]; then
  echo "ERROR: No se encontró un intérprete de Python (3.10+ preferido)." >&2
  exit 1
fi

echo "Usando intérprete: $PYTHON"
$PYTHON -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
pip install -e .
$PYTHON -m retail_roi_model.cli "$@"
