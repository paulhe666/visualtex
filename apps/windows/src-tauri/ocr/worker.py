#!/usr/bin/env python3
"""Persistent PaddleOCR PP-FormulaNet worker for VisualTeX.

Protocol: one JSON request per stdin line, one JSON response per stdout line.
All Paddle/PaddleOCR logs are redirected to stderr so stdout remains machine-readable.
"""

from __future__ import annotations

import contextlib
import gc
import hashlib
import json
import os
import subprocess
import sys
import tempfile
import threading
import time
import traceback
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# VisualTeX owns model acquisition. The worker must never invoke PaddleX model
# discovery or network downloads; it only loads an already verified model_dir.
os.environ.setdefault("VISUALTEX_OFFLINE_OCR", "1")
os.environ.setdefault("HF_HUB_OFFLINE", "1")
os.environ.setdefault("TRANSFORMERS_OFFLINE", "1")
os.environ.setdefault("MODELSCOPE_OFFLINE", "1")
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
_MODEL_FILE_SHA256 = {
    "PP-FormulaNet_plus-S": {
        "inference.json": "01238434e33df83588e2627f350559b576e34551d2b2ffea148345032de56c00",
        "inference.pdiparams": "e464f94412feaa98f8791eacc84684f887b3569e30e80c52b8112e9cf7d4069b",
        "inference.yml": "96062655d94c21d39274328dbc82c1a487e66addb8425f5a7fd5b7dfb2421ec3",
    },
    "PP-FormulaNet_plus-M": {
        "inference.json": "8333a7f650766a748e273c550d278601dd19dfeee1c4b01038ff632f134d9884",
        "inference.pdiparams": "f16ef9b5c8227da70d3ec969a5195f4d62c1154427b883f4d6cff07633654041",
        "inference.yml": "87b5f3d7f2b2fe553627d77b37f496608ca150ebd0ef62d362591edca47b5538",
    },
    "PP-FormulaNet_plus-L": {
        "inference.json": "ad259c4b896d99aa3479336b9121112fb40ff1ababfbf8765a3428a3b86df582",
        "inference.pdiparams": "4245c39c181d1d21e472bc85c7434df9b23f177be46552c0542bf153addbc355",
        "inference.yml": "afc92a2737268da0499c37b0b6741da268c369fd7424667fcfeb8fa6c7b22d30",
    },
}
_MODEL_VALIDATION_CACHE: Dict[str, Tuple[Tuple[str, int, int, int], ...]] = {}


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


def _sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while True:
            chunk = handle.read(1024 * 1024)
            if not chunk:
                break
            digest.update(chunk)
    return digest.hexdigest()


def _validate_model_integrity(directory: Path, model_name: str) -> None:
    expected = _MODEL_FILE_SHA256.get(model_name)
    if expected is None:
        raise RuntimeError(
            f"OCR_MODEL_INTEGRITY_ERROR: VisualTeX has no pinned integrity manifest for {model_name}."
        )
    if not directory.is_dir():
        raise RuntimeError(
            f"OCR_MODEL_INTEGRITY_ERROR: The model directory does not exist: {directory}"
        )

    actual_files = {entry.name for entry in directory.iterdir() if entry.is_file()}
    expected_files = set(expected)
    if actual_files != expected_files:
        missing = sorted(expected_files - actual_files)
        unexpected = sorted(actual_files - expected_files)
        raise RuntimeError(
            "OCR_MODEL_INTEGRITY_ERROR: "
            f"{model_name} has an incomplete or unexpected file set at {directory}; "
            f"missing={missing}, unexpected={unexpected}. Reinstall or re-import this model."
        )

    signature: List[Tuple[str, int, int, int]] = []
    for filename in sorted(expected):
        path = directory / filename
        stat = path.stat()
        if stat.st_size <= 0:
            raise RuntimeError(
                "OCR_MODEL_INTEGRITY_ERROR: "
                f"{model_name}/{filename} is empty at {path}. Reinstall or re-import this model."
            )
        signature.append((filename, stat.st_size, stat.st_mtime_ns, stat.st_ctime_ns))

    cache_key = str(directory.resolve())
    signature_tuple = tuple(signature)
    if _MODEL_VALIDATION_CACHE.get(cache_key) == signature_tuple:
        return

    json_path = directory / "inference.json"
    try:
        with json_path.open("r", encoding="utf-8") as handle:
            parsed_json = json.load(handle)
        if not isinstance(parsed_json, (dict, list)):
            raise ValueError("top-level JSON value is not an object or array")
    except Exception as exc:
        raise RuntimeError(
            "OCR_MODEL_INTEGRITY_ERROR: "
            f"{model_name}/inference.json is unreadable or invalid at {json_path}: {exc}. "
            "Reinstall or re-import this model."
        ) from exc

    for filename, expected_hash in expected.items():
        path = directory / filename
        actual_hash = _sha256_file(path)
        if actual_hash.lower() != expected_hash.lower():
            raise RuntimeError(
                "OCR_MODEL_INTEGRITY_ERROR: "
                f"{model_name}/{filename} failed SHA-256 verification at {path}; "
                f"expected={expected_hash}, actual={actual_hash}. Reinstall or re-import this model."
            )

    _MODEL_VALIDATION_CACHE[cache_key] = signature_tuple


def _cached_model_dir(model_name: str) -> Optional[Path]:
    cache_home = os.environ.get("PADDLE_PDX_CACHE_HOME", "").strip()
    if not cache_home:
        return None
    directory = Path(cache_home) / "official_models" / model_name
    if not directory.exists():
        return None
    _validate_model_integrity(directory, model_name)
    return directory


def _windows_short_path(path: Path) -> Optional[Path]:
    if os.name != "nt":
        return None
    try:
        import ctypes
        from ctypes import wintypes

        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        get_short_path = kernel32.GetShortPathNameW
        get_short_path.argtypes = [
            wintypes.LPCWSTR,
            wintypes.LPWSTR,
            wintypes.DWORD,
        ]
        get_short_path.restype = wintypes.DWORD
        required = get_short_path(str(path), None, 0)
        if required == 0:
            return None
        buffer = ctypes.create_unicode_buffer(required + 1)
        written = get_short_path(str(path), buffer, len(buffer))
        if written == 0 or written >= len(buffer):
            return None
        short = Path(buffer.value)
        if not str(short).isascii() or not short.exists():
            return None
        return short
    except Exception:
        return None


def _ascii_bridge_root() -> Path:
    candidates = [
        Path(tempfile.gettempdir()) / "VisualTeX-OCR-bridge",
        Path(os.environ.get("LOCALAPPDATA", tempfile.gettempdir()))
        / "VisualTeX"
        / "ocr-path-bridge",
    ]
    system_root = os.environ.get("SystemRoot", "").strip()
    if system_root:
        candidates.append(Path(system_root) / "Temp" / "VisualTeX-OCR-bridge")

    failures: List[str] = []
    for candidate in candidates:
        try:
            candidate.mkdir(parents=True, exist_ok=True)
            if str(candidate).isascii():
                return candidate
            short = _windows_short_path(candidate)
            if short is not None and str(short).isascii():
                return short
            failures.append(f"{candidate}: no ASCII short path")
        except Exception as exc:
            failures.append(f"{candidate}: {exc}")

    raise RuntimeError(
        "OCR_MODEL_PATH_COMPATIBILITY_ERROR: Unable to create an ASCII compatibility root for Paddle Inference. "
        + " | ".join(failures)
    )


def _junction_model_dir(directory: Path) -> Path:
    bridge_root = _ascii_bridge_root()
    identity = hashlib.sha256(str(directory.resolve()).encode("utf-8")).hexdigest()[:20]
    alias = bridge_root / f"model-{identity}"
    if not str(alias).isascii():
        raise RuntimeError(
            f"OCR_MODEL_PATH_COMPATIBILITY_ERROR: Compatibility alias is not ASCII: {alias}"
        )

    if alias.exists():
        try:
            if os.path.samefile(alias, directory):
                return alias
        except OSError:
            pass
        raise RuntimeError(
            f"OCR_MODEL_PATH_COMPATIBILITY_ERROR: Existing compatibility alias points elsewhere: {alias}"
        )

    command = [
        "powershell.exe",
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-Command",
        "& { param([string]$Path, [string]$Target) $ErrorActionPreference='Stop'; New-Item -ItemType Junction -Path $Path -Target $Target | Out-Null }",
        str(alias),
        str(directory),
    ]
    completed = subprocess.run(
        command,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        creationflags=0x08000000,
        check=False,
    )
    if completed.returncode != 0 or not alias.is_dir():
        detail = (completed.stderr or completed.stdout).strip()
        raise RuntimeError(
            "OCR_MODEL_PATH_COMPATIBILITY_ERROR: Paddle Inference cannot read the Unicode model path, "
            f"and VisualTeX could not create an ASCII junction for {directory}. "
            f"PowerShell exit={completed.returncode}; detail={detail or 'no output'}"
        )
    try:
        if not os.path.samefile(alias, directory):
            raise RuntimeError(
                f"OCR_MODEL_PATH_COMPATIBILITY_ERROR: Junction verification failed: {alias} -> {directory}"
            )
    except OSError as exc:
        raise RuntimeError(
            f"OCR_MODEL_PATH_COMPATIBILITY_ERROR: Junction verification failed for {alias}: {exc}"
        ) from exc
    return alias


def _paddle_compatible_model_dir(directory: Path) -> Path:
    if os.name != "nt" or str(directory).isascii():
        return directory
    short = _windows_short_path(directory)
    if short is not None:
        _log(f"Using Windows short path for Paddle model_dir: {short}")
        return short
    alias = _junction_model_dir(directory)
    _log(f"Using ASCII junction for Paddle model_dir: {alias} -> {directory}")
    return alias


def _model_is_cached(model_name: str) -> bool:
    return _cached_model_dir(model_name) is not None


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

    if not _model_is_cached(model_name):
        raise RuntimeError(
            f"The verified local model {model_name} is not installed. "
            "Import a .vtxocrmodel package or explicitly download it in VisualTeX before warmup or recognition."
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

    cached_model_dir = _cached_model_dir(model_name)
    if cached_model_dir is None:
        raise RuntimeError(f"Verified local model_dir is unavailable for {model_name}")
    paddle_model_dir = _paddle_compatible_model_dir(cached_model_dir)
    _log(
        f"Loading {model_name} directly from local model_dir "
        f"{paddle_model_dir} on {device}"
    )
    with contextlib.redirect_stdout(sys.stderr):
        from paddleocr import FormulaRecognition

        # Passing model_dir is mandatory. This bypasses PaddleX host checks,
        # registry metadata and every implicit model download path.
        try:
            model = FormulaRecognition(
                model_name=model_name,
                model_dir=str(paddle_model_dir),
                device=device,
            )
        except Exception as exc:
            raise RuntimeError(
                "OCR_MODEL_LOAD_ERROR: Paddle FormulaRecognition failed while initializing "
                f"{model_name} from {paddle_model_dir} on {device}: {exc}"
            ) from exc

    _CURRENT_MODEL = model
    _CURRENT_MODEL_NAME = model_name
    _CURRENT_DEVICE = device
    return model


def _warmup(request: Dict[str, Any]) -> Dict[str, Any]:
    started = time.perf_counter()
    request_id = str(request.get("id", ""))
    model_name = str(request.get("model") or "PP-FormulaNet_plus-M")
    device = str(request.get("device") or "cpu")
    _load_model(model_name, device)
    elapsed_ms = round((time.perf_counter() - started) * 1000)
    _log(f"Warmed {model_name} on {device} in {elapsed_ms} ms")
    return {
        "id": request_id,
        "ok": True,
        "event": "warmed",
        "model": model_name,
        "device": device,
        "elapsed_ms": elapsed_ms,
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
    else:
        model_message = f"正在检查 {model_name} 的本地已校验模型目录"
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
