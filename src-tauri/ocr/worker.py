#!/usr/bin/env python3
"""Persistent PaddleOCR PP-FormulaNet worker for VisualTeX.

Protocol: one JSON request per stdin line, one JSON response per stdout line.
All Paddle/PaddleOCR logs are redirected to stderr so stdout remains machine-readable.
"""

from __future__ import annotations

import contextlib
import gc
import json
import os
import sys
import threading
import time
import traceback
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# BOS is generally more reliable in regions where Hugging Face access is limited.
os.environ.setdefault("PADDLE_PDX_MODEL_SOURCE", "BOS")
os.environ.setdefault("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "True")
os.environ.setdefault("GLOG_minloglevel", "2")

# The worker protocol is UTF-8 on every platform. Windows may otherwise inherit
# a legacy console code page such as CP936/GBK.
if hasattr(sys.stdin, "reconfigure"):
    sys.stdin.reconfigure(encoding="utf-8", errors="strict")
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="strict", line_buffering=True)
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(
        encoding="utf-8",
        errors="backslashreplace",
        line_buffering=True,
    )

_PROTOCOL_STDOUT = sys.stdout
_CURRENT_MODEL: Any = None
_CURRENT_MODEL_NAME: Optional[str] = None
_CURRENT_DEVICE: Optional[str] = None
_MODEL_DOWNLOAD_MB = {
    "PP-FormulaNet_plus-S": 259.6,
    "PP-FormulaNet_plus-M": 620.5,
    "PP-FormulaNet_plus-L": 731.5,
    "PP-FormulaNet-S": 234.4,
    "PP-FormulaNet-L": 728.8,
}


def _emit(payload: Dict[str, Any]) -> None:
    # ASCII-safe JSON makes the pipe protocol independent of the user's system
    # locale even if a dependency changes a stream encoding at runtime.
    _PROTOCOL_STDOUT.write(json.dumps(payload, ensure_ascii=True, separators=(",", ":")) + "\n")
    _PROTOCOL_STDOUT.flush()


def _log(message: str) -> None:
    print(f"[VisualTeX OCR] {message}", file=sys.stderr, flush=True)


def _emit_progress(
    request_id: str,
    stage: str,
    message: str,
    model_name: str,
) -> None:
    _emit(
        {
            "event": "progress",
            "id": request_id,
            "stage": stage,
            "message": message,
            "model": model_name,
        }
    )


def _watch_parent_process() -> None:
    raw_parent_pid = os.environ.get("VISUALTEX_PARENT_PID", "").strip()
    if not raw_parent_pid:
        return
    try:
        parent_pid = int(raw_parent_pid)
    except ValueError:
        return

    if os.name == "nt":
        # On Windows, os.kill(pid, 0) is not a harmless existence check: Python
        # routes non-console signals through TerminateProcess. Wait on a process
        # handle instead so the OCR worker exits only after VisualTeX exits.
        import ctypes
        from ctypes import wintypes

        synchronize = 0x00100000
        infinite = 0xFFFFFFFF
        wait_object_0 = 0x00000000
        wait_failed = 0xFFFFFFFF
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        kernel32.OpenProcess.restype = wintypes.HANDLE
        kernel32.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
        kernel32.WaitForSingleObject.restype = wintypes.DWORD
        kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
        kernel32.CloseHandle.restype = wintypes.BOOL

        handle = kernel32.OpenProcess(synchronize, False, parent_pid)
        if not handle:
            _log(
                "Unable to open the VisualTeX parent process handle on Windows; "
                f"parent monitoring is disabled (error={ctypes.get_last_error()})"
            )
            return
        try:
            wait_result = kernel32.WaitForSingleObject(handle, infinite)
        finally:
            kernel32.CloseHandle(handle)

        if wait_result == wait_object_0:
            os._exit(0)
        if wait_result == wait_failed:
            _log(
                "Waiting for the VisualTeX parent process failed on Windows; "
                f"parent monitoring is disabled (error={ctypes.get_last_error()})"
            )
        else:
            _log(
                "Unexpected Windows parent wait result; parent monitoring is "
                f"disabled (result={wait_result})"
            )
        return

    while True:
        time.sleep(2.0)
        try:
            os.kill(parent_pid, 0)
        except OSError:
            os._exit(0)


def _model_is_cached(model_name: str) -> bool:
    candidates = [Path.home() / ".paddlex" / "official_models" / model_name]
    cache_home = os.environ.get("PADDLE_PDX_CACHE_HOME", "").strip()
    if cache_home:
        candidates.insert(0, Path(cache_home) / "official_models" / model_name)
    return any(
        (directory / "inference.json").exists()
        and (directory / "inference.pdiparams").exists()
        for directory in candidates
    )


def _strip_outer_delimiters(value: str) -> str:
    latex = value.strip()
    pairs = (("$$", "$$"), ("\\[", "\\]"), ("\\(", "\\)"))
    for left, right in pairs:
        if latex.startswith(left) and latex.endswith(right) and len(latex) >= len(left) + len(right):
            return latex[len(left) : -len(right)].strip()
    return latex


def _collect_formulas(value: Any, output: List[str]) -> None:
    if isinstance(value, dict):
        formula = value.get("rec_formula")
        if isinstance(formula, str):
            cleaned = _strip_outer_delimiters(formula)
            if cleaned and cleaned not in output:
                output.append(cleaned)
        for nested in value.values():
            _collect_formulas(nested, output)
    elif isinstance(value, (list, tuple)):
        for nested in value:
            _collect_formulas(nested, output)


def _result_payload(result: Any) -> Any:
    payload = getattr(result, "json", None)
    if callable(payload):
        payload = payload()
    if payload is not None:
        return payload
    if isinstance(result, (dict, list, tuple)):
        return result
    return {"repr": repr(result)}


def _preprocess_image(
    input_path: str,
    processed_path: str,
) -> Tuple[str, Tuple[int, int], Dict[str, Any]]:
    import numpy as np
    from PIL import Image, ImageChops, ImageOps

    source = Path(input_path)
    target = Path(processed_path)
    target.parent.mkdir(parents=True, exist_ok=True)

    with Image.open(source) as opened:
        image = ImageOps.exif_transpose(opened)

        # Transparent formula assets may contain either dark or light glyphs.
        # Composite onto the opposite tone so white glyphs are not lost on a
        # white canvas before polarity detection.
        if image.mode in ("RGBA", "LA") or "transparency" in image.info:
            rgba = image.convert("RGBA")
            rgba_array = np.asarray(rgba, dtype=np.uint8)
            alpha = rgba_array[:, :, 3]
            visible = alpha > 16
            if np.any(visible):
                rgb = rgba_array[:, :, :3].astype(np.float32)
                luminance = (
                    0.2126 * rgb[:, :, 0]
                    + 0.7152 * rgb[:, :, 1]
                    + 0.0722 * rgb[:, :, 2]
                )
                visible_luminance = float(np.median(luminance[visible]))
            else:
                visible_luminance = 0.0
            matte_value = 0 if visible_luminance >= 160 else 255
            matte = Image.new(
                "RGBA",
                rgba.size,
                (matte_value, matte_value, matte_value, 255),
            )
            image = Image.alpha_composite(matte, rgba).convert("RGB")
        else:
            image = image.convert("RGB")

        grayscale = ImageOps.grayscale(image)
        grayscale_array = np.asarray(grayscale, dtype=np.uint8)
        height, width = grayscale_array.shape
        border_width = max(1, min(width, height) // 20)
        border_pixels = np.concatenate(
            [
                grayscale_array[:border_width, :].reshape(-1),
                grayscale_array[-border_width:, :].reshape(-1),
                grayscale_array[:, :border_width].reshape(-1),
                grayscale_array[:, -border_width:].reshape(-1),
            ]
        )
        border_median = float(np.median(border_pixels))
        image_median = float(np.median(grayscale_array))
        border_dark_ratio = float(np.mean(border_pixels < 96))
        border_bright_ratio = float(np.mean(border_pixels > 160))

        # Formula models and their internal cropper expect dark glyphs on a
        # light canvas. A dark majority plus a dark border is a strong signal
        # for a dark-theme screenshot; invert only in that case.
        dark_background = (
            image_median < 128
            and border_median < 160
            and border_dark_ratio > border_bright_ratio
        )
        if dark_background:
            grayscale = ImageOps.invert(grayscale)

        # Preserve antialiasing while stretching weak gray-on-gray contrast.
        # Hard binarization is deliberately avoided because it damages thin
        # fraction bars, radicals, superscripts, and small punctuation.
        grayscale = ImageOps.autocontrast(grayscale, cutoff=1)

        # After polarity normalization the background is white, so crop only
        # meaningful dark pixels and then add a fresh white safety margin.
        white = Image.new("L", grayscale.size, 255)
        difference = ImageChops.difference(grayscale, white)
        mask = difference.point(lambda pixel: 255 if pixel > 14 else 0)
        bbox = mask.getbbox()
        if bbox:
            cropped = grayscale.crop(bbox)
            safety = max(8, min(grayscale.size) // 50)
            padded = Image.new(
                "L",
                (cropped.width + 2 * safety, cropped.height + 2 * safety),
                255,
            )
            padded.paste(cropped, (safety, safety))
            grayscale = padded

        # Short formula screenshots often lose superscript/subscript detail.
        # Upscale conservatively while preserving aspect ratio.
        minimum_height = 96
        if 0 < grayscale.height < minimum_height:
            scale = min(4.0, minimum_height / grayscale.height)
            target_size = (
                max(1, round(grayscale.width * scale)),
                max(1, round(grayscale.height * scale)),
            )
            grayscale = grayscale.resize(target_size, Image.Resampling.LANCZOS)

        # Keep an upper bound to avoid accidental huge-memory inputs.
        max_side = 4096
        if max(grayscale.size) > max_side:
            scale = max_side / max(grayscale.size)
            grayscale = grayscale.resize(
                (
                    max(1, round(grayscale.width * scale)),
                    max(1, round(grayscale.height * scale)),
                ),
                Image.Resampling.LANCZOS,
            )

        image = grayscale.convert("RGB")
        image.save(target, format="PNG", optimize=True)
        return (
            str(target),
            image.size,
            {
                "background_inverted": dark_background,
                "background_luminance": round(border_median, 1),
            },
        )


def _load_model(model_name: str, device: str) -> Any:
    global _CURRENT_MODEL, _CURRENT_MODEL_NAME, _CURRENT_DEVICE

    if os.environ.get("VISUALTEX_OFFLINE_OCR") == "1" and not _model_is_cached(model_name):
        raise RuntimeError(
            f"The offline model pack for {model_name} is not installed. "
            "The bundled M model works without a network connection; install the optional S or L pack before selecting it."
        )

    if (
        _CURRENT_MODEL is not None
        and _CURRENT_MODEL_NAME == model_name
        and _CURRENT_DEVICE == device
    ):
        return _CURRENT_MODEL

    if _CURRENT_MODEL is not None:
        _log(f"Releasing model {_CURRENT_MODEL_NAME}")
        _CURRENT_MODEL = None
        _CURRENT_MODEL_NAME = None
        _CURRENT_DEVICE = None
        gc.collect()

    _log(f"Loading {model_name} on {device}")
    with contextlib.redirect_stdout(sys.stderr):
        from paddleocr import FormulaRecognition

        model = FormulaRecognition(model_name=model_name, device=device)

    _CURRENT_MODEL = model
    _CURRENT_MODEL_NAME = model_name
    _CURRENT_DEVICE = device
    return model


def _runtime_versions() -> Dict[str, str]:
    import platform
    from importlib.metadata import version

    import paddle

    return {
        "python_version": platform.python_version(),
        "paddle_version": paddle.__version__,
        "paddleocr_version": version("paddleocr"),
    }


def _warmup(request: Dict[str, Any]) -> Dict[str, Any]:
    started = time.perf_counter()
    request_id = str(request.get("id", ""))
    model_name = str(request.get("model") or "PP-FormulaNet_plus-M")
    device = str(request.get("device") or "cpu")
    _load_model(model_name, device)
    return {
        "id": request_id,
        "ok": True,
        "event": "model-ready",
        "model": model_name,
        "device": device,
        "elapsed_ms": round((time.perf_counter() - started) * 1000),
        **_runtime_versions(),
    }


def _recognize(request: Dict[str, Any]) -> Dict[str, Any]:
    started = time.perf_counter()
    request_id = str(request.get("id", ""))
    image_path = str(request["image_path"])
    processed_path = str(request["processed_path"])
    model_name = str(request.get("model") or "PP-FormulaNet_plus-M")
    device = str(request.get("device") or "cpu")

    _emit_progress(request_id, "preprocess", "正在分析公式图片背景与对比度", model_name)
    normalized_path, image_size, preprocessing = _preprocess_image(
        image_path,
        processed_path,
    )
    if preprocessing["background_inverted"]:
        preprocess_message = "检测到深色背景，已自动反色并统一为白底"
    else:
        preprocess_message = "已统一对比度、裁剪公式区域并补充白边"
    _emit_progress(
        request_id,
        "preprocess-complete",
        preprocess_message,
        model_name,
    )

    download_mb = _MODEL_DOWNLOAD_MB.get(model_name)
    if _CURRENT_MODEL_NAME == model_name and _CURRENT_DEVICE == device:
        model_message = "正在复用已加载的模型"
    elif _model_is_cached(model_name):
        model_message = f"正在从本地缓存加载 {model_name}"
    elif os.environ.get("VISUALTEX_OFFLINE_OCR") == "1":
        model_message = f"正在检查 {model_name} 离线模型包"
    elif download_mb is not None:
        model_message = (
            f"正在下载并加载 {model_name}；下载量约 {download_mb:.1f} MB"
        )
    else:
        model_message = f"正在准备 {model_name}"
    _emit_progress(request_id, "model", model_message, model_name)
    model = _load_model(model_name, device)

    _emit_progress(request_id, "inference", "模型已就绪，正在识别公式", model_name)
    with contextlib.redirect_stdout(sys.stderr):
        predictions = model.predict(input=normalized_path, batch_size=1)
        formulas: List[str] = []
        for prediction in predictions:
            _collect_formulas(_result_payload(prediction), formulas)

    if not formulas:
        raise RuntimeError(
            "PP-FormulaNet did not return rec_formula. Try a tighter crop, a clearer image, or another model."
        )

    elapsed_ms = round((time.perf_counter() - started) * 1000)
    return {
        "id": request_id,
        "ok": True,
        "model": model_name,
        "device": device,
        "elapsed_ms": elapsed_ms,
        "processed_width": image_size[0],
        "processed_height": image_size[1],
        "background_inverted": preprocessing["background_inverted"],
        "background_luminance": preprocessing["background_luminance"],
        "formulas": [{"latex": formula} for formula in formulas],
    }


def _handle(request: Dict[str, Any]) -> Dict[str, Any]:
    action = request.get("action")
    request_id = str(request.get("id", ""))

    if action == "ping":
        return {"id": request_id, "ok": True, "event": "pong"}
    if action == "shutdown":
        return {"id": request_id, "ok": True, "event": "shutdown"}
    if action == "warmup":
        return _warmup(request)
    if action == "recognize":
        return _recognize(request)
    raise ValueError(f"Unsupported action: {action!r}")


def main() -> int:
    threading.Thread(target=_watch_parent_process, daemon=True).start()
    _emit({"event": "ready", "ok": True, "protocol": 1})

    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        request: Optional[Dict[str, Any]] = None
        request_id = ""
        try:
            request = json.loads(line)
            if not isinstance(request, dict):
                raise ValueError("Request must be a JSON object")
            request_id = str(request.get("id", ""))
            response = _handle(request)
        except Exception as exc:  # noqa: BLE001 - worker must serialize all failures
            details = traceback.format_exc()
            _log(details)
            response = {
                "id": request_id,
                "ok": False,
                "error": str(exc) or exc.__class__.__name__,
                "details": details,
            }
        _emit(response)
        if request is not None and request.get("action") == "shutdown":
            return 0

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
