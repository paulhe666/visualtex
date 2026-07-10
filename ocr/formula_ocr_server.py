from __future__ import annotations

import json
import os
import sys
import time
import traceback
from contextlib import redirect_stdout
from pathlib import Path
from typing import Any, Dict, Iterable, Optional, Tuple

from PIL import Image, UnidentifiedImageError

from normalize_latex import normalize_formula_latex
from protocol import OcrRequest, ProtocolError, parse_request

os.environ.setdefault("PADDLE_PDX_MODEL_SOURCE", "BOS")

_PROTOCOL_STDOUT = sys.stdout
_MODELS: Dict[Tuple[str, str], Any] = {}
_MAX_INPUT_BYTES = 12 * 1024 * 1024


def log(message: str) -> None:
    print(message, file=sys.stderr, flush=True)


def emit(payload: Dict[str, Any]) -> None:
    print(
        json.dumps(payload, ensure_ascii=False, separators=(",", ":")),
        file=_PROTOCOL_STDOUT,
        flush=True,
    )


def error_payload(
    request_id: str,
    code: str,
    message: str,
    detail: Optional[str] = None,
) -> Dict[str, Any]:
    error: Dict[str, Any] = {"code": code, "message": message}
    if detail:
        error["detail"] = detail
    return {"id": request_id, "ok": False, "error": error}


def get_model(model_name: str, device: str) -> Any:
    key = (model_name, device)
    if key in _MODELS:
        return _MODELS[key]

    try:
        # PaddleOCR and PaddleX can print progress information. Redirect it to
        # stderr so stdout remains a strict JSON Lines protocol.
        with redirect_stdout(sys.stderr):
            from paddleocr import FormulaRecognition

            model = FormulaRecognition(model_name=model_name, device=device)
    except Exception as exc:  # noqa: BLE001 - converted into a stable protocol error
        raise ProtocolError(
            "MODEL_LOAD_FAILED",
            "Unable to load the formula recognition model",
            str(exc),
        ) from exc

    _MODELS[key] = model
    return model


def validate_image(path: Path) -> None:
    if not path.exists() or not path.is_file():
        raise ProtocolError("IMAGE_NOT_FOUND", "Formula image does not exist")
    if path.stat().st_size > _MAX_INPUT_BYTES:
        raise ProtocolError("IMAGE_TOO_LARGE", "Formula image exceeds 12 MiB")

    try:
        with Image.open(path) as image:
            image.verify()
    except (UnidentifiedImageError, OSError, ValueError) as exc:
        raise ProtocolError(
            "UNSUPPORTED_IMAGE",
            "Input is not a valid PNG, JPEG, or WebP image",
            str(exc),
        ) from exc


def result_to_mapping(result: Any) -> Dict[str, Any]:
    raw = getattr(result, "json", result)
    if callable(raw):
        raw = raw()
    if isinstance(raw, str):
        raw = json.loads(raw)
    if not isinstance(raw, dict):
        raise ProtocolError("INFERENCE_FAILED", "Unexpected PaddleOCR result type")
    return raw


def extract_formula(results: Iterable[Any]) -> Tuple[str, str]:
    first = next(iter(results), None)
    if first is None:
        raise ProtocolError("EMPTY_RESULT", "The model returned no result")

    mapping = result_to_mapping(first)
    body = mapping.get("res", mapping)
    if not isinstance(body, dict):
        raise ProtocolError("INFERENCE_FAILED", "Unexpected PaddleOCR result payload")

    raw_latex = body.get("rec_formula")
    if not isinstance(raw_latex, str) or not raw_latex.strip():
        raise ProtocolError("EMPTY_RESULT", "No formula was recognized in the image")

    latex = normalize_formula_latex(raw_latex)
    if not latex:
        raise ProtocolError("EMPTY_RESULT", "The recognized formula is empty")
    return raw_latex, latex


def handle_request(request: OcrRequest) -> Dict[str, Any]:
    started = time.perf_counter()

    if request.action == "ping":
        return {
            "id": request.request_id,
            "ok": True,
            "status": "ready",
            "loaded_models": [model for model, _device in _MODELS.keys()],
        }

    if request.action == "warmup":
        get_model(request.model, request.device)
        return {
            "id": request.request_id,
            "ok": True,
            "status": "model-ready",
            "model": request.model,
            "elapsed_ms": round((time.perf_counter() - started) * 1000),
        }

    if request.image_path is None:
        raise ProtocolError("INVALID_REQUEST", "image_path is required")

    validate_image(request.image_path)
    model = get_model(request.model, request.device)

    try:
        with redirect_stdout(sys.stderr):
            results = model.predict(input=str(request.image_path), batch_size=1)
        raw_latex, latex = extract_formula(results)
    except ProtocolError:
        raise
    except Exception as exc:  # noqa: BLE001 - converted into stable response
        raise ProtocolError(
            "INFERENCE_FAILED",
            "Formula recognition failed",
            str(exc),
        ) from exc

    return {
        "id": request.request_id,
        "ok": True,
        "latex": latex,
        "raw_latex": raw_latex,
        "model": request.model,
        "elapsed_ms": round((time.perf_counter() - started) * 1000),
        "warnings": [],
    }


def process_line(line: str) -> None:
    request_id = "unknown"
    try:
        payload = json.loads(line)
        if isinstance(payload, dict) and isinstance(payload.get("id"), str):
            request_id = payload["id"]
        request = parse_request(payload)
        emit(handle_request(request))
    except json.JSONDecodeError as exc:
        emit(error_payload(request_id, "INVALID_REQUEST", "Invalid JSON", str(exc)))
    except ProtocolError as exc:
        emit(error_payload(request_id, exc.code, exc.message, exc.detail))
    except Exception as exc:  # noqa: BLE001 - final process safety net
        log(traceback.format_exc())
        emit(error_payload(request_id, "INFERENCE_FAILED", "Unexpected OCR error", str(exc)))


def main() -> int:
    log("VisualTeX formula OCR sidecar started")
    for raw_line in sys.stdin:
        line = raw_line.strip()
        if line:
            process_line(line)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
