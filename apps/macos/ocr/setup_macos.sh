#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PYTHON_BIN="${VISUALTEX_OCR_BOOTSTRAP_PYTHON:-/usr/bin/python3}"
VENV_DIR="$SCRIPT_DIR/.venv"

if [[ ! -x "$PYTHON_BIN" ]]; then
  echo "Python not found: $PYTHON_BIN" >&2
  exit 1
fi

"$PYTHON_BIN" -m venv "$VENV_DIR"
"$VENV_DIR/bin/python" -m pip install --upgrade pip setuptools wheel
"$VENV_DIR/bin/python" -m pip install -r "$SCRIPT_DIR/requirements.txt"

export PADDLE_PDX_MODEL_SOURCE="${PADDLE_PDX_MODEL_SOURCE:-BOS}"
"$VENV_DIR/bin/python" -c "import paddle, paddleocr; print('paddle', paddle.__version__); print('paddleocr', paddleocr.__version__)"
"$VENV_DIR/bin/python" -m unittest discover -s "$SCRIPT_DIR/tests" -p 'test_*.py'

cat <<EOF
VisualTeX OCR environment is ready:
  $VENV_DIR/bin/python

The PP-FormulaNet model is downloaded lazily on the first warmup or recognition request.
EOF
