# VisualTeX Formula OCR Sidecar

This directory contains the local PaddleOCR PP-FormulaNet runtime used by the Tauri application.

## Setup on macOS

```bash
cd /Users/lpj/devspace/workspaces/visualtex
chmod +x ocr/setup_macos.sh
./ocr/setup_macos.sh
```

## Protocol smoke test

```bash
printf '%s\n' '{"id":"ping-1","action":"ping"}' | ocr/.venv/bin/python ocr/formula_ocr_server.py
```

Expected stdout:

```json
{"id":"ping-1","ok":true,"status":"ready","loaded_models":[]}
```

Logs are written to stderr. Formula models are downloaded lazily when `warmup` or `recognize` is called. Formula decoding also requires `tokenizers==0.19.1`, `imagesize`, `ftfy`, and `Wand`; these are included in `requirements.txt` and in the VisualTeX in-app installer.

## Recognize a formula image

```bash
printf '%s\n' '{"id":"rec-1","action":"recognize","image_path":"/absolute/path/formula.png","model":"PP-FormulaNet_plus-M","device":"cpu"}' \
  | ocr/.venv/bin/python ocr/formula_ocr_server.py
```

Set `PADDLE_PDX_MODEL_SOURCE=BOS` when Hugging Face is unavailable.
