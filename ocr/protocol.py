from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

SUPPORTED_MODELS = {
    "PP-FormulaNet_plus-M",
    "PP-FormulaNet_plus-L",
    "PP-FormulaNet-S",
}
SUPPORTED_DEVICES = {"cpu"}


class ProtocolError(ValueError):
    def __init__(self, code: str, message: str, detail: Optional[str] = None) -> None:
        super().__init__(message)
        self.code = code
        self.message = message
        self.detail = detail


@dataclass(frozen=True)
class OcrRequest:
    request_id: str
    action: str
    image_path: Optional[Path] = None
    model: str = "PP-FormulaNet_plus-M"
    device: str = "cpu"


def parse_request(payload: Any) -> OcrRequest:
    if not isinstance(payload, dict):
        raise ProtocolError("INVALID_REQUEST", "Request must be a JSON object")

    request_id = payload.get("id")
    action = payload.get("action")
    if not isinstance(request_id, str) or not request_id.strip():
        raise ProtocolError("INVALID_REQUEST", "Request id is required")
    if action not in {"ping", "warmup", "recognize"}:
        raise ProtocolError("INVALID_REQUEST", "Unsupported OCR action")

    model = payload.get("model") or "PP-FormulaNet_plus-M"
    device = payload.get("device") or "cpu"
    if model not in SUPPORTED_MODELS:
        raise ProtocolError("INVALID_REQUEST", f"Unsupported model: {model}")
    if device not in SUPPORTED_DEVICES:
        raise ProtocolError("INVALID_REQUEST", f"Unsupported device: {device}")

    image_path: Optional[Path] = None
    if action == "recognize":
        raw_path = payload.get("image_path")
        if not isinstance(raw_path, str) or not raw_path:
            raise ProtocolError("INVALID_REQUEST", "image_path is required")
        image_path = Path(raw_path)
        if not image_path.is_absolute():
            raise ProtocolError("INVALID_REQUEST", "image_path must be absolute")

    return OcrRequest(
        request_id=request_id,
        action=action,
        image_path=image_path,
        model=model,
        device=device,
    )
