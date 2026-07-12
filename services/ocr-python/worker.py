#!/usr/bin/env python3
"""VisualTeX offline OCR sidecar.

The worker speaks newline-delimited JSON and deliberately does not download models.
A production formula backend becomes available only when VISUALTEX_FORMULA_MODEL_DIR
(or the legacy VISUALTEX_OCR_MODEL_DIR) points to an existing local model package.

The input parser accepts both the current `{id, method, params}` envelope and the
older flat `{id, action, ...}` shape so packaged runtimes can be upgraded without
breaking an older desktop binary.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import threading
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence


FORMULA_KEYS = (
    "rec_formula",
    "formula",
    "latex",
    "text",
    "rec_text",
    "prediction",
)
SCORE_KEYS = ("rec_score", "score", "confidence", "probability")
MODEL_VERSION_KEYS = ("model_version", "modelVersion", "model_name", "modelName")


@dataclass(frozen=True)
class FormulaCandidate:
    latex: str
    confidence: float
    backend: str

    def to_json(self) -> dict[str, Any]:
        return {
            "latex": self.latex,
            "confidence": max(0.0, min(1.0, float(self.confidence))),
            "backend": self.backend,
        }


class FormulaBackend:
    """Lazily loads a local PaddleOCR formula recognition module."""

    def __init__(self, *, mock: bool = False) -> None:
        self._mock = mock
        self._model: Any | None = None
        self._load_error: str | None = None
        self._lock = threading.Lock()
        configured = (
            os.environ.get("VISUALTEX_FORMULA_MODEL_DIR")
            or os.environ.get("VISUALTEX_OCR_MODEL_DIR")
            or ""
        )
        self._model_dir = Path(configured).expanduser() if configured else None
        self._model_name = os.environ.get("VISUALTEX_FORMULA_MODEL_NAME", "local-formula-model")
        self._device = os.environ.get("VISUALTEX_OCR_DEVICE", "cpu")

        # PaddleX performs source checks in some releases. The worker must remain
        # offline, so disable model-source probing before importing PaddleOCR.
        os.environ.setdefault("PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK", "True")
        os.environ.setdefault("PADDLEOCR_DISABLE_MODEL_SOURCE_CHECK", "True")

    @property
    def backend_name(self) -> str:
        return "mock" if self._mock else "paddleocr-formula"

    @property
    def model_version(self) -> str:
        if self._mock:
            return "mock-1"
        if self._model_dir is None:
            return "unconfigured"
        manifest = self._read_manifest()
        value = manifest.get("version") or manifest.get("modelVersion") or manifest.get("name")
        return str(value) if value else self._model_name

    def health(self) -> dict[str, Any]:
        if self._mock:
            return {
                "available": True,
                "backend": self.backend_name,
                "modelVersion": self.model_version,
                "detail": None,
            }
        if self._model_dir is None:
            return {
                "available": False,
                "backend": self.backend_name,
                "modelVersion": self.model_version,
                "detail": (
                    "Set VISUALTEX_FORMULA_MODEL_DIR to an installed local "
                    "PaddleOCR formula model package."
                ),
            }
        if not self._model_dir.is_dir():
            return {
                "available": False,
                "backend": self.backend_name,
                "modelVersion": self.model_version,
                "detail": f"Formula model directory does not exist: {self._model_dir}",
            }
        if self._load_error is not None:
            return {
                "available": False,
                "backend": self.backend_name,
                "modelVersion": self.model_version,
                "detail": self._load_error,
            }
        return {
            "available": True,
            "backend": self.backend_name,
            "modelVersion": self.model_version,
            "detail": "Local model package found; model will load on first recognition.",
        }

    def recognize(self, image_path: Path) -> dict[str, Any]:
        if self._mock:
            return {
                "candidates": [
                    FormulaCandidate(
                        r"\int_0^1 x^2\,\mathrm{d}x=\frac{1}{3}", 0.99, "mock"
                    ).to_json(),
                    FormulaCandidate(r"\int_0^1 x^2 dx=1/3", 0.72, "mock").to_json(),
                ],
                "modelVersion": self.model_version,
                "warnings": ["Mock OCR result; do not use for production recognition."],
            }

        health = self.health()
        if not health["available"]:
            raise RuntimeError(str(health["detail"]))
        if not image_path.is_file():
            raise FileNotFoundError(f"OCR input image does not exist: {image_path}")

        model = self._ensure_model()
        raw_outputs = self._predict(model, image_path)
        candidates: list[FormulaCandidate] = []
        warnings: list[str] = []
        for raw in raw_outputs:
            candidate = extract_candidate(raw, self.backend_name)
            if candidate is not None and candidate.latex.strip():
                candidates.append(candidate)
        candidates = deduplicate_candidates(candidates)
        if not candidates:
            warnings.append("The local formula model returned no usable LaTeX candidate.")
        return {
            "candidates": [candidate.to_json() for candidate in candidates],
            "modelVersion": self.model_version,
            "warnings": warnings,
        }

    def _ensure_model(self) -> Any:
        if self._model is not None:
            return self._model
        with self._lock:
            if self._model is not None:
                return self._model
            try:
                from paddleocr import FormulaRecognition  # type: ignore

                kwargs: dict[str, Any] = {
                    "model_dir": str(self._model_dir),
                }
                # PaddleOCR/PaddleX releases differ in device keyword support.
                try:
                    self._model = FormulaRecognition(device=self._device, **kwargs)
                except TypeError:
                    self._model = FormulaRecognition(**kwargs)
            except Exception as exc:  # pragma: no cover - depends on optional runtime
                self._load_error = format_exception(exc)
                raise RuntimeError(self._load_error) from exc
        return self._model

    @staticmethod
    def _predict(model: Any, image_path: Path) -> list[Any]:
        predictor = getattr(model, "predict", None)
        if predictor is None:
            raise RuntimeError("FormulaRecognition backend has no predict() method")
        try:
            output = predictor(input=str(image_path), batch_size=1)
        except TypeError:
            try:
                output = predictor(str(image_path), batch_size=1)
            except TypeError:
                output = predictor(str(image_path))
        return list(output) if is_iterable_output(output) else [output]

    def _read_manifest(self) -> Mapping[str, Any]:
        if self._model_dir is None:
            return {}
        for name in ("visualtex-model.json", "model.json", "inference.json"):
            path = self._model_dir / name
            if not path.is_file():
                continue
            try:
                parsed = json.loads(path.read_text(encoding="utf-8"))
                if isinstance(parsed, Mapping):
                    return parsed
            except Exception:
                continue
        return {}


class DocumentBackend:
    """Lazily loads a fully local PP-StructureV3 pipeline configuration."""

    def __init__(self, *, mock: bool = False) -> None:
        self._mock = mock
        configured = os.environ.get("VISUALTEX_DOCUMENT_PIPELINE_CONFIG", "")
        package_root = os.environ.get("VISUALTEX_DOCUMENT_PACKAGE_ROOT", "")
        self._config_path = Path(configured).expanduser() if configured else None
        self._package_root = Path(package_root).expanduser() if package_root else None
        self._model_name = os.environ.get(
            "VISUALTEX_DOCUMENT_MODEL_NAME", "local-pp-structure-v3"
        )
        self._device = os.environ.get("VISUALTEX_OCR_DEVICE", "cpu")
        self._pipeline: Any | None = None
        self._resolved_config: Mapping[str, Any] | None = None
        self._load_error: str | None = None
        self._lock = threading.Lock()

    @property
    def backend_name(self) -> str:
        return "mock-document" if self._mock else "paddleocr-ppstructurev3"

    @property
    def model_version(self) -> str:
        return "mock-document-1" if self._mock else self._model_name

    def health(self) -> dict[str, Any]:
        if self._mock:
            return {
                "available": True,
                "backend": self.backend_name,
                "modelVersion": self.model_version,
                "detail": None,
            }
        if self._config_path is None or self._package_root is None:
            return {
                "available": False,
                "backend": self.backend_name,
                "modelVersion": "unconfigured",
                "detail": (
                    "Install and activate a layout_ocr package whose entrypoint is a "
                    "fully local PP-StructureV3 YAML configuration."
                ),
            }
        if self._load_error is not None:
            return {
                "available": False,
                "backend": self.backend_name,
                "modelVersion": self.model_version,
                "detail": self._load_error,
            }
        try:
            self._load_local_config()
        except Exception as exc:
            self._load_error = format_exception(exc)
            return {
                "available": False,
                "backend": self.backend_name,
                "modelVersion": self.model_version,
                "detail": self._load_error,
            }
        return {
            "available": True,
            "backend": self.backend_name,
            "modelVersion": self.model_version,
            "detail": "Fully local PP-StructureV3 configuration validated.",
        }

    def recognize(self, image_path: Path) -> dict[str, Any]:
        if self._mock:
            return {
                "imagePath": None,
                "pageWidth": 600,
                "pageHeight": 800,
                "regions": [
                    {
                        "kind": "document_title",
                        "x": 40.0,
                        "y": 24.0,
                        "width": 520.0,
                        "height": 48.0,
                        "text": "Mock Document Title",
                        "latex": None,
                        "confidence": 0.98,
                    },
                    {
                        "kind": "text",
                        "x": 40.0,
                        "y": 92.0,
                        "width": 520.0,
                        "height": 110.0,
                        "text": "Mock paragraph in restored reading order.",
                        "latex": None,
                        "confidence": 0.95,
                    },
                    {
                        "kind": "formula",
                        "x": 180.0,
                        "y": 224.0,
                        "width": 240.0,
                        "height": 48.0,
                        "text": None,
                        "latex": r"E=mc^2",
                        "confidence": 0.93,
                    },
                ],
                "readingOrder": [0, 1, 2],
                "modelVersion": self.model_version,
                "warnings": ["Mock document OCR result; do not use for production."],
            }

        health = self.health()
        if not health["available"]:
            raise RuntimeError(str(health["detail"]))
        if not image_path.is_file():
            raise FileNotFoundError(f"Document OCR input does not exist: {image_path}")
        pipeline = self._ensure_pipeline()
        raw_outputs = self._predict(pipeline, image_path)
        if not raw_outputs:
            return {
                "imagePath": None,
                "pageWidth": 0,
                "pageHeight": 0,
                "regions": [],
                "readingOrder": [],
                "modelVersion": self.model_version,
                "warnings": ["PP-StructureV3 returned no page result."],
            }
        result = normalize_document_result(raw_outputs[0], self.model_version)
        if len(raw_outputs) > 1:
            result["warnings"].append(
                "The image request returned multiple pages; only the first page was imported."
            )
        return result

    def _load_local_config(self) -> Mapping[str, Any]:
        if self._resolved_config is not None:
            return self._resolved_config
        if self._config_path is None or self._package_root is None:
            raise RuntimeError("Document OCR pipeline is not configured")
        config_path = self._config_path.resolve(strict=True)
        package_root = self._package_root.resolve(strict=True)
        if not config_path.is_file() or not config_path.is_relative_to(package_root):
            raise RuntimeError("PP-StructureV3 config must be a file inside its model package")
        if config_path.suffix.lower() not in {".yaml", ".yml"}:
            raise RuntimeError("PP-StructureV3 package entrypoint must be a YAML file")
        try:
            import yaml  # type: ignore
        except Exception as exc:
            raise RuntimeError("PyYAML is required by the local document OCR backend") from exc
        parsed = yaml.safe_load(config_path.read_text(encoding="utf-8"))
        if not isinstance(parsed, Mapping):
            raise RuntimeError("PP-StructureV3 YAML root must be a mapping")
        resolved, model_count = resolve_local_model_dirs(parsed, package_root, config_path.parent)
        if model_count == 0:
            raise RuntimeError("PP-StructureV3 YAML contains no local model_dir entries")
        self._resolved_config = resolved
        return resolved

    def _ensure_pipeline(self) -> Any:
        if self._pipeline is not None:
            return self._pipeline
        with self._lock:
            if self._pipeline is not None:
                return self._pipeline
            try:
                from paddleocr import PPStructureV3  # type: ignore

                config = self._load_local_config()
                try:
                    self._pipeline = PPStructureV3(
                        paddlex_config=config,
                        device=self._device,
                    )
                except TypeError:
                    self._pipeline = PPStructureV3(paddlex_config=config)
            except Exception as exc:  # pragma: no cover - optional production runtime
                self._load_error = format_exception(exc)
                raise RuntimeError(self._load_error) from exc
        return self._pipeline

    @staticmethod
    def _predict(pipeline: Any, image_path: Path) -> list[Any]:
        predictor = getattr(pipeline, "predict", None)
        if predictor is None:
            raise RuntimeError("PPStructureV3 backend has no predict() method")
        try:
            output = predictor(
                input=str(image_path),
                use_doc_orientation_classify=False,
                use_doc_unwarping=False,
                use_textline_orientation=False,
                use_table_recognition=True,
                use_formula_recognition=True,
                format_block_content=True,
            )
        except TypeError:
            output = predictor(input=str(image_path))
        return list(output) if is_iterable_output(output) else [output]


def resolve_local_model_dirs(
    value: Any,
    package_root: Path,
    config_parent: Path,
) -> tuple[Any, int]:
    if isinstance(value, Mapping):
        resolved: dict[str, Any] = {}
        model_count = 0
        for key, item in value.items():
            if str(key) == "model_dir":
                if not isinstance(item, str) or not item.strip():
                    raise RuntimeError(
                        "Every model_dir in the offline PP-StructureV3 config must be set"
                    )
                candidate = Path(item).expanduser()
                if not candidate.is_absolute():
                    candidate = config_parent / candidate
                candidate = candidate.resolve(strict=True)
                if not candidate.is_dir() or not candidate.is_relative_to(package_root):
                    raise RuntimeError(
                        f"Model directory escapes the installed package: {candidate}"
                    )
                resolved[str(key)] = str(candidate)
                model_count += 1
            else:
                next_value, nested_count = resolve_local_model_dirs(
                    item, package_root, config_parent
                )
                resolved[str(key)] = next_value
                model_count += nested_count
        return resolved, model_count
    if isinstance(value, list):
        resolved_items = []
        model_count = 0
        for item in value:
            next_value, nested_count = resolve_local_model_dirs(
                item, package_root, config_parent
            )
            resolved_items.append(next_value)
            model_count += nested_count
        return resolved_items, model_count
    if isinstance(value, str) and value.lower().startswith(("http://", "https://")):
        raise RuntimeError("Remote URLs are forbidden in offline OCR pipeline configuration")
    return value, 0


def plain_value(value: Any) -> Any:
    converter = getattr(value, "tolist", None)
    if callable(converter):
        try:
            return converter()
        except Exception:
            pass
    return value


def bbox_xywh(value: Any) -> tuple[float, float, float, float] | None:
    value = plain_value(value)
    if not isinstance(value, Sequence) or isinstance(value, (str, bytes)):
        return None
    if len(value) == 4 and all(isinstance(item, (int, float)) for item in value):
        left, top, right, bottom = (float(item) for item in value)
        return left, top, max(0.0, right - left), max(0.0, bottom - top)
    points = []
    for item in value:
        item = plain_value(item)
        if (
            isinstance(item, Sequence)
            and not isinstance(item, (str, bytes))
            and len(item) >= 2
            and isinstance(item[0], (int, float))
            and isinstance(item[1], (int, float))
        ):
            points.append((float(item[0]), float(item[1])))
    if not points:
        return None
    left = min(point[0] for point in points)
    top = min(point[1] for point in points)
    right = max(point[0] for point in points)
    bottom = max(point[1] for point in points)
    return left, top, max(0.0, right - left), max(0.0, bottom - top)


def first_present(mapping: Mapping[str, Any], keys: Sequence[str]) -> Any | None:
    for key in keys:
        if key in mapping and mapping[key] is not None:
            return mapping[key]
    return None


def normalized_confidence(value: Any) -> float:
    try:
        return max(0.0, min(1.0, float(value)))
    except (TypeError, ValueError):
        return 0.0


def content_text(value: Any) -> str | None:
    value = plain_value(value)
    if value is None:
        return None
    if isinstance(value, str):
        return value.strip() or None
    try:
        return json.dumps(value, ensure_ascii=False, separators=(",", ":"))
    except TypeError:
        return str(value)


def normalize_document_result(raw: Any, model_version: str) -> dict[str, Any]:
    mapping = object_to_mapping(raw)
    parsing_value = first_present(mapping, ("parsing_res_list", "parsingResList"))
    parsing = parsing_value if parsing_value is not None else []
    if not isinstance(parsing, Sequence) or isinstance(parsing, (str, bytes)):
        parsing = []
    ordered_regions: list[tuple[float, int, dict[str, Any]]] = []
    for index, raw_block in enumerate(parsing):
        block = object_to_mapping(raw_block)
        bbox = bbox_xywh(first_present(block, ("block_bbox", "blockBbox")))
        if bbox is None:
            continue
        kind = str(first_present(block, ("block_label", "blockLabel")) or "unknown")
        content = content_text(first_present(block, ("block_content", "blockContent")))
        score = normalized_confidence(
            first_present(block, ("block_score", "score", "confidence"))
        )
        is_formula = "formula" in kind.lower()
        order_value = first_present(block, ("block_order", "blockOrder"))
        try:
            order = float(order_value) if order_value is not None else float(index)
        except (TypeError, ValueError):
            order = float(index)
        x, y, width, height = bbox
        ordered_regions.append(
            (
                order,
                index,
                {
                    "kind": kind,
                    "x": x,
                    "y": y,
                    "width": width,
                    "height": height,
                    "text": None if is_formula else content,
                    "latex": content if is_formula else None,
                    "confidence": score,
                },
            )
        )
    ordered_regions.sort(key=lambda item: (item[0], item[1]))
    regions = [item[2] for item in ordered_regions]

    if not regions:
        overall_value = first_present(mapping, ("overall_ocr_res", "overallOcrRes"))
        overall = object_to_mapping(overall_value if overall_value is not None else {})
        boxes_value = first_present(overall, ("rec_boxes", "recBoxes"))
        texts_value = first_present(overall, ("rec_texts", "recTexts"))
        scores_value = first_present(overall, ("rec_scores", "recScores"))
        boxes = plain_value(boxes_value if boxes_value is not None else [])
        texts = plain_value(texts_value if texts_value is not None else [])
        scores = plain_value(scores_value if scores_value is not None else [])
        if isinstance(boxes, Sequence) and isinstance(texts, Sequence):
            for index, raw_box in enumerate(boxes):
                bbox = bbox_xywh(raw_box)
                if bbox is None:
                    continue
                x, y, width, height = bbox
                text = content_text(texts[index] if index < len(texts) else None)
                score = normalized_confidence(
                    scores[index] if isinstance(scores, Sequence) and index < len(scores) else None
                )
                regions.append(
                    {
                        "kind": "text",
                        "x": x,
                        "y": y,
                        "width": width,
                        "height": height,
                        "text": text,
                        "latex": None,
                        "confidence": score,
                    }
                )

    warnings = []
    if not regions:
        warnings.append("PP-StructureV3 returned no usable layout or OCR region.")
    elif any(region["confidence"] == 0.0 for region in regions):
        warnings.append(
            "Some layout blocks do not expose confidence scores; those scores are reported as 0."
        )
    try:
        page_width = int(mapping.get("width") or 0)
        page_height = int(mapping.get("height") or 0)
    except (TypeError, ValueError):
        page_width = 0
        page_height = 0
    return {
        "imagePath": None,
        "pageWidth": max(0, page_width),
        "pageHeight": max(0, page_height),
        "regions": regions,
        "readingOrder": list(range(len(regions))),
        "modelVersion": model_version,
        "warnings": warnings,
    }


def is_iterable_output(value: Any) -> bool:
    return isinstance(value, Iterable) and not isinstance(value, (str, bytes, Mapping))


def object_to_mapping(value: Any) -> Mapping[str, Any]:
    if isinstance(value, Mapping):
        return value
    for attribute in ("json", "res", "result", "data"):
        candidate = getattr(value, attribute, None)
        if callable(candidate):
            try:
                candidate = candidate()
            except TypeError:
                continue
        if isinstance(candidate, str):
            try:
                candidate = json.loads(candidate)
            except json.JSONDecodeError:
                continue
        if isinstance(candidate, Mapping):
            return candidate
    try:
        attributes = vars(value)
    except TypeError:
        return {}
    return attributes if isinstance(attributes, Mapping) else {}


def find_nested(mapping: Mapping[str, Any], keys: Sequence[str]) -> Any | None:
    for key in keys:
        value = mapping.get(key)
        if value is not None:
            return value
    for value in mapping.values():
        if isinstance(value, Mapping):
            nested = find_nested(value, keys)
            if nested is not None:
                return nested
        elif isinstance(value, Sequence) and not isinstance(value, (str, bytes)):
            for item in value:
                if isinstance(item, Mapping):
                    nested = find_nested(item, keys)
                    if nested is not None:
                        return nested
    return None


def extract_candidate(raw: Any, backend: str) -> FormulaCandidate | None:
    mapping = object_to_mapping(raw)
    if not mapping and isinstance(raw, str):
        return FormulaCandidate(raw, 1.0, backend)
    latex = find_nested(mapping, FORMULA_KEYS)
    if isinstance(latex, Sequence) and not isinstance(latex, (str, bytes)):
        latex = latex[0] if latex else None
    if latex is None:
        return None
    score = find_nested(mapping, SCORE_KEYS)
    if isinstance(score, Sequence) and not isinstance(score, (str, bytes)):
        score = score[0] if score else None
    try:
        confidence = float(score) if score is not None else 1.0
    except (TypeError, ValueError):
        confidence = 1.0
    return FormulaCandidate(str(latex).strip(), confidence, backend)


def deduplicate_candidates(candidates: Sequence[FormulaCandidate]) -> list[FormulaCandidate]:
    best: dict[str, FormulaCandidate] = {}
    for candidate in candidates:
        normalized = " ".join(candidate.latex.split())
        existing = best.get(normalized)
        if existing is None or candidate.confidence > existing.confidence:
            best[normalized] = candidate
    return sorted(best.values(), key=lambda item: item.confidence, reverse=True)


def format_exception(exc: BaseException) -> str:
    return f"{type(exc).__name__}: {exc}"


def request_method(message: Mapping[str, Any]) -> str:
    value = message.get("method") or message.get("action") or message.get("kind") or ""
    return str(value).replace("_", ".").lower()


def request_params(message: Mapping[str, Any]) -> Mapping[str, Any]:
    params = message.get("params")
    return params if isinstance(params, Mapping) else message


def image_path_from(params: Mapping[str, Any]) -> Path:
    value = (
        params.get("imagePath")
        or params.get("image_path")
        or params.get("path")
        or params.get("input")
    )
    if not isinstance(value, str) or not value:
        raise ValueError("formula OCR request is missing imagePath")
    return Path(value).expanduser().resolve(strict=False)


def success_response(request_id: Any, payload: Mapping[str, Any]) -> dict[str, Any]:
    # Include both the typed flat response and a JSON-RPC-like result for backward
    # and forward compatibility. Serde ignores unknown fields by default.
    return {
        "id": request_id,
        "ok": True,
        **payload,
        "result": dict(payload),
    }


def error_response(request_id: Any, exc: BaseException) -> dict[str, Any]:
    detail = format_exception(exc)
    return {
        "id": request_id,
        "ok": False,
        "detail": detail,
        "traceback": traceback.format_exc(limit=8),
        "error": {
            "code": -32001,
            "message": detail,
            "data": {"exceptionType": type(exc).__name__},
        },
    }


def handle(
    message: Mapping[str, Any],
    formula_backend: FormulaBackend,
    document_backend: DocumentBackend,
) -> dict[str, Any]:
    request_id = message.get("id")
    method = request_method(message)
    params = request_params(message)
    if method in {"health", "ocr.health", "formula.health"}:
        return success_response(request_id, formula_backend.health())
    if method in {"document.health", "ocr.document.health"}:
        return success_response(request_id, document_backend.health())
    if method in {"capabilities", "ocr.capabilities", "worker.capabilities"}:
        formula_health = formula_backend.health()
        document_health = document_backend.health()
        return success_response(
            request_id,
            {
                "formulaRecognition": {
                    "available": formula_health["available"],
                    "backend": formula_health["backend"],
                    "modelVersion": formula_health["modelVersion"],
                },
                "documentRecognition": {
                    "available": document_health["available"],
                    "backend": document_health["backend"],
                    "modelVersion": document_health["modelVersion"],
                },
                "offlineOnly": True,
            },
        )
    if method in {
        "formula",
        "ocr.formula",
        "formula.recognize",
        "recognize.formula",
        "recognizeformula",
    }:
        return success_response(
            request_id, formula_backend.recognize(image_path_from(params))
        )
    if method in {
        "document",
        "ocr.document",
        "document.recognize",
        "recognize.document",
        "recognizedocument",
    }:
        return success_response(
            request_id, document_backend.recognize(image_path_from(params))
        )
    if method in {"shutdown", "worker.shutdown"}:
        return success_response(request_id, {"shutdown": True})
    raise ValueError(f"unsupported OCR worker method: {method or '<empty>'}")


def run(mock: bool) -> int:
    formula_backend = FormulaBackend(mock=mock)
    document_backend = DocumentBackend(mock=mock)
    for raw_line in sys.stdin:
        line = raw_line.strip()
        if not line:
            continue
        request_id: Any = None
        should_stop = False
        try:
            message = json.loads(line)
            if not isinstance(message, Mapping):
                raise ValueError("OCR request must be a JSON object")
            request_id = message.get("id")
            response = handle(message, formula_backend, document_backend)
            should_stop = bool(response.get("shutdown")) or bool(
                isinstance(response.get("result"), Mapping)
                and response["result"].get("shutdown")
            )
        except Exception as exc:
            response = error_response(request_id, exc)
        sys.stdout.write(json.dumps(response, ensure_ascii=False, separators=(",", ":")) + "\n")
        sys.stdout.flush()
        if should_stop:
            break
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="VisualTeX offline OCR worker")
    parser.add_argument("--mock", action="store_true", help="Use deterministic test output")
    args = parser.parse_args()
    return run(mock=args.mock or os.environ.get("VISUALTEX_OCR_MOCK") == "1")


if __name__ == "__main__":
    raise SystemExit(main())
