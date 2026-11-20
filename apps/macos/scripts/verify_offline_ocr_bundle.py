#!/usr/bin/env python3
"""Verify VisualTeX's self-contained macOS arm64 OCR bundle."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import platform
import shutil
import subprocess
import sys
import tarfile
import tempfile
from pathlib import Path, PurePosixPath
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_BUNDLE = ROOT / "dist-ocr/macos-arm64"
EXPECTED_MODEL = "PP-FormulaNet_plus-M"
EXPECTED_RUNTIME_LAYOUT_VERSION = 3
EXPECTED_VERSIONS = {
    "python": "3.10.20",
    "paddlepaddle": "3.3.1",
    "paddleocr": "3.7.0",
    "paddlex": "3.7.2",
    "tokenizers": "0.19.1",
}
OPTIONAL_MODEL_HASHES = {
    "PP-FormulaNet_plus-S": {
        "inference.json": "01238434e33df83588e2627f350559b576e34551d2b2ffea148345032de56c00",
        "inference.pdiparams": "e464f94412feaa98f8791eacc84684f887b3569e30e80c52b8112e9cf7d4069b",
        "inference.yml": "96062655d94c21d39274328dbc82c1a487e66addb8425f5a7fd5b7dfb2421ec3",
    },
    "PP-FormulaNet_plus-L": {
        "inference.json": "ad259c4b896d99aa3479336b9121112fb40ff1ababfbf8765a3428a3b86df582",
        "inference.pdiparams": "4245c39c181d1d21e472bc85c7434df9b23f177be46552c0542bf153addbc355",
        "inference.yml": "afc92a2737268da0499c37b0b6741da268c369fd7424667fcfeb8fa6c7b22d30",
    },
}


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load_manifest(root: Path) -> dict[str, Any]:
    path = root / "manifest.json"
    if not path.is_file():
        raise RuntimeError(f"Missing manifest: {path}")
    manifest = json.loads(path.read_text("utf-8"))
    if manifest.get("schemaVersion") != 1:
        raise RuntimeError("Unsupported OCR manifest schema")
    if manifest.get("platform") != "macos" or manifest.get("architecture") != "arm64":
        raise RuntimeError("Offline OCR manifest target must be macOS arm64")
    if manifest.get("runtimeLayoutVersion") != EXPECTED_RUNTIME_LAYOUT_VERSION:
        raise RuntimeError("Offline OCR runtime layout version is invalid")
    if manifest.get("defaultModel", {}).get("name") != EXPECTED_MODEL:
        raise RuntimeError("Offline OCR manifest does not contain the default M model")
    return manifest


def verify_record(root: Path, record: dict[str, Any]) -> Path:
    name = record.get("name")
    if not isinstance(name, str) or not name or "/" in name or "\\" in name:
        raise RuntimeError(f"Unsafe archive name: {name!r}")
    path = root / name
    if not path.is_file():
        raise RuntimeError(f"Missing archive: {path}")
    if path.stat().st_size != record.get("size"):
        raise RuntimeError(f"Archive size mismatch: {path}")
    digest = sha256_file(path)
    if digest != record.get("sha256"):
        raise RuntimeError(f"Archive SHA-256 mismatch: {path}")
    return path


def normalize_link(member_path: PurePosixPath, link: PurePosixPath) -> PurePosixPath:
    if link.is_absolute():
        raise RuntimeError(f"Absolute archive link: {member_path} -> {link}")
    parts = list(member_path.parent.parts)
    for part in link.parts:
        if part in ("", "."):
            continue
        if part == "..":
            if not parts:
                raise RuntimeError(f"Archive link escapes root: {member_path} -> {link}")
            parts.pop()
        else:
            parts.append(part)
    return PurePosixPath(*parts)


def verify_archive_layout(path: Path, required: set[str]) -> set[str]:
    names: set[str] = set()
    with tarfile.open(path, "r:gz") as archive:
        for member in archive.getmembers():
            normalized = PurePosixPath(member.name)
            if normalized.is_absolute() or ".." in normalized.parts:
                raise RuntimeError(f"Unsafe path in {path.name}: {member.name}")
            if member.ischr() or member.isblk() or member.isfifo():
                raise RuntimeError(f"Special file in {path.name}: {member.name}")
            if member.issym() or member.islnk():
                normalize_link(normalized, PurePosixPath(member.linkname))
            names.add(normalized.as_posix().rstrip("/"))
    missing = sorted(required - names)
    if missing:
        raise RuntimeError(f"{path.name} is missing entries: {missing}")
    return names


def safe_extract(path: Path, destination: Path) -> None:
    root = destination.resolve()
    destination.mkdir(parents=True, exist_ok=True)
    with tarfile.open(path, "r:gz") as archive:
        for member in archive.getmembers():
            relative = PurePosixPath(member.name)
            if relative.is_absolute() or ".." in relative.parts:
                raise RuntimeError(f"Unsafe extraction path: {member.name}")
            target = (destination / Path(*relative.parts)).resolve()
            if target != root and root not in target.parents:
                raise RuntimeError(f"Archive entry escapes destination: {member.name}")
            if member.issym() or member.islnk():
                normalize_link(relative, PurePosixPath(member.linkname))
        archive.extractall(destination)


def verify_licenses(root: Path, manifest: dict[str, Any]) -> None:
    licenses = root / "licenses"
    required = [
        licenses / "CPYTHON-LICENSE.txt",
        licenses / "PYTHON-BUILD-STANDALONE-LICENSE.txt",
        licenses / "THIRD_PARTY_NOTICES.json",
    ]
    for path in required:
        if not path.is_file() or path.stat().st_size == 0:
            raise RuntimeError(f"Missing OCR license file: {path}")
    notices = json.loads((licenses / "THIRD_PARTY_NOTICES.json").read_text("utf-8"))
    if len(notices) != manifest.get("thirdPartyPackageCount"):
        raise RuntimeError("Third-party notice count does not match the manifest")
    by_name = {str(item.get("name", "")).lower(): item for item in notices}
    for package in ("paddlepaddle", "paddleocr", "paddlex", "pillow"):
        if package not in by_name:
            raise RuntimeError(f"Missing third-party notice for {package}")
        if not by_name[package].get("licenseFiles"):
            raise RuntimeError(f"No bundled license file recorded for {package}")


def hash_stream(stream: Any) -> tuple[int, str]:
    digest = hashlib.sha256()
    size = 0
    for chunk in iter(lambda: stream.read(1024 * 1024), b""):
        size += len(chunk)
        digest.update(chunk)
    return size, digest.hexdigest()


def verify_optional_model_pack(path: Path) -> str:
    if path.suffix != ".vtxocrmodel" or not path.is_file():
        raise RuntimeError(f"Invalid optional OCR model package: {path}")
    if path.stat().st_size <= 0 or path.stat().st_size > 1024 * 1024 * 1024:
        raise RuntimeError(f"Optional OCR model package size is invalid: {path}")

    manifest_name = "visualtex-model-pack/pack-manifest.json"
    regular_files: dict[str, tarfile.TarInfo] = {}
    with tarfile.open(path, "r:gz") as archive:
        for member in archive.getmembers():
            normalized = PurePosixPath(member.name)
            if normalized.is_absolute() or ".." in normalized.parts:
                raise RuntimeError(f"Unsafe path in {path.name}: {member.name}")
            if member.issym() or member.islnk() or member.ischr() or member.isblk() or member.isfifo():
                raise RuntimeError(f"Unsupported entry in {path.name}: {member.name}")
            if member.isfile():
                regular_files[normalized.as_posix()] = member
        manifest_member = regular_files.get(manifest_name)
        if manifest_member is None:
            raise RuntimeError(f"{path.name} has no pack-manifest.json")
        manifest_stream = archive.extractfile(manifest_member)
        if manifest_stream is None:
            raise RuntimeError(f"Unable to read {manifest_name} from {path.name}")
        manifest = json.loads(manifest_stream.read().decode("utf-8"))
        model = manifest.get("model")
        if (
            manifest.get("schemaVersion") != 1
            or manifest.get("platform") != "macos"
            or manifest.get("architecture") != "arm64"
            or model not in OPTIONAL_MODEL_HASHES
        ):
            raise RuntimeError(f"Invalid optional OCR model manifest in {path.name}")

        expected_hashes = OPTIONAL_MODEL_HASHES[model]
        expected_files = {manifest_name}
        records = manifest.get("files")
        if not isinstance(records, dict) or set(records) != set(expected_hashes):
            raise RuntimeError(f"Unexpected file records in {path.name}")
        for name, expected_hash in expected_hashes.items():
            archive_name = (
                f"visualtex-model-pack/paddlex/official_models/{model}/{name}"
            )
            expected_files.add(archive_name)
            member = regular_files.get(archive_name)
            if member is None:
                raise RuntimeError(f"{path.name} is missing {archive_name}")
            record = records[name]
            if record.get("name") != name or record.get("sha256") != expected_hash:
                raise RuntimeError(f"Invalid manifest checksum for {archive_name}")
            stream = archive.extractfile(member)
            if stream is None:
                raise RuntimeError(f"Unable to read {archive_name}")
            size, digest = hash_stream(stream)
            if size != record.get("size") or digest != expected_hash:
                raise RuntimeError(f"Optional model verification failed: {archive_name}")
        if set(regular_files) != expected_files:
            extras = sorted(set(regular_files) - expected_files)
            missing = sorted(expected_files - set(regular_files))
            raise RuntimeError(
                f"Unexpected optional model contents in {path.name}; extras={extras}, missing={missing}"
            )
    return str(model)


def run_full_check(runtime_archive: Path, model_archive: Path) -> None:
    with tempfile.TemporaryDirectory(prefix="visualtex-ocr-verify-") as temporary:
        root = Path(temporary)
        safe_extract(runtime_archive, root)
        safe_extract(model_archive, root / "cache")
        python = root / "python/bin/python3"
        if not python.is_file():
            raise RuntimeError("Extracted OCR runtime has no Python executable")
        python.chmod(python.stat().st_mode | 0o111)

        forbidden_markers = [
            str(ROOT).encode(),
            b".cache/visualtex-ocr",
            b"/Users/lpj/devspace/workspaces/visualtex",
        ]
        violations: list[str] = []
        for path in sorted((root / "python").rglob("*"), key=lambda item: item.as_posix()):
            if not path.is_file() or path.is_symlink() or path.stat().st_size > 5 * 1024 * 1024:
                continue
            data = path.read_bytes()
            if any(marker in data for marker in forbidden_markers):
                violations.append(path.relative_to(root).as_posix())
                if len(violations) >= 20:
                    break
        if violations:
            raise RuntimeError(
                "Extracted OCR runtime contains non-relocatable build paths: "
                + ", ".join(violations)
            )

        file_output = subprocess.check_output(["/usr/bin/file", str(python)], text=True)
        if "arm64" not in file_output and "Mach-O 64-bit executable arm64" not in file_output:
            raise RuntimeError(f"Bundled Python is not arm64: {file_output.strip()}")

        environment = os.environ.copy()
        environment.update(
            {
                "PYTHONNOUSERSITE": "1",
                "PYTHONPATH": "",
                "VISUALTEX_OFFLINE_OCR": "1",
                "PADDLE_PDX_CACHE_HOME": str(root / "cache/paddlex"),
                "PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK": "True",
                "HF_HUB_OFFLINE": "1",
                "TRANSFORMERS_OFFLINE": "1",
                "MODELSCOPE_OFFLINE": "1",
                "HTTP_PROXY": "http://127.0.0.1:9",
                "HTTPS_PROXY": "http://127.0.0.1:9",
                "ALL_PROXY": "http://127.0.0.1:9",
                "NO_PROXY": "",
            }
        )
        script = f"""
import json, platform
import paddle, paddleocr, paddlex, tokenizers
from paddleocr import FormulaRecognition
model = FormulaRecognition(model_name={EXPECTED_MODEL!r}, device='cpu')
print(json.dumps({{
  'python': platform.python_version(),
  'architecture': platform.machine(),
  'paddlepaddle': paddle.__version__,
  'paddleocr': paddleocr.__version__,
  'paddlex': paddlex.__version__,
  'tokenizers': tokenizers.__version__,
  'modelClass': model.__class__.__name__,
}}, sort_keys=True))
"""
        output = subprocess.check_output(
            [str(python), "-c", script],
            env=environment,
            text=True,
            stderr=subprocess.STDOUT,
            timeout=180,
        )
        values = json.loads(output.strip().splitlines()[-1])
        for key, expected in EXPECTED_VERSIONS.items():
            if values.get(key) != expected:
                raise RuntimeError(
                    f"Extracted runtime version mismatch for {key}: {values.get(key)!r}"
                )
        if values.get("architecture") != "arm64":
            raise RuntimeError("Extracted Python did not report arm64")
        if not values.get("modelClass"):
            raise RuntimeError("Default M model could not be instantiated offline")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--bundle", type=Path, default=DEFAULT_BUNDLE)
    parser.add_argument("--full", action="store_true")
    parser.add_argument("--model-packs", type=Path)
    args = parser.parse_args()

    root = args.bundle.resolve()
    manifest = load_manifest(root)
    runtime_archive = verify_record(root, manifest["archives"]["runtime"])
    model_archive = verify_record(root, manifest["archives"]["defaultModel"])
    verify_archive_layout(
        runtime_archive,
        {
            "python/bin/python3",
            "python/lib/python3.10/site-packages/paddle/__init__.py",
            "python/lib/python3.10/site-packages/paddleocr/__init__.py",
            "python/lib/python3.10/site-packages/paddlex/__init__.py",
        },
    )
    verify_archive_layout(
        model_archive,
        {
            f"paddlex/official_models/{EXPECTED_MODEL}/inference.json",
            f"paddlex/official_models/{EXPECTED_MODEL}/inference.pdiparams",
            f"paddlex/official_models/{EXPECTED_MODEL}/inference.yml",
        },
    )
    verify_licenses(root, manifest)

    if args.full:
        if platform.system() != "Darwin" or platform.machine() != "arm64":
            raise RuntimeError("Full OCR verification must run on Apple Silicon macOS")
        run_full_check(runtime_archive, model_archive)

    if args.model_packs is not None:
        model_pack_root = args.model_packs.resolve()
        expected_paths = {
            model_pack_root
            / f"VisualTeX-{model}-macos-arm64.vtxocrmodel"
            for model in OPTIONAL_MODEL_HASHES
        }
        actual_paths = {path for path in model_pack_root.iterdir() if path.is_file()}
        if actual_paths != expected_paths:
            extras = sorted(path.name for path in actual_paths - expected_paths)
            missing = sorted(path.name for path in expected_paths - actual_paths)
            raise RuntimeError(
                f"Optional model pack directory is not clean; extras={extras}, missing={missing}"
            )
        verified_models = {verify_optional_model_pack(path) for path in sorted(actual_paths)}
        if verified_models != set(OPTIONAL_MODEL_HASHES):
            raise RuntimeError(
                f"Expected S and L optional model packs, found {sorted(verified_models)}"
            )

    print(
        "Offline OCR bundle verification passed"
        + (" (full extraction and offline model load)" if args.full else "")
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(f"Offline OCR verification failed: {error}", file=sys.stderr)
        raise SystemExit(1)
