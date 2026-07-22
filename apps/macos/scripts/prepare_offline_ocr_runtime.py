#!/usr/bin/env python3
"""Build the self-contained Apple Silicon OCR bundle used by VisualTeX.

All network access happens at build time. The shipped application extracts the
verified archives locally and never invokes pip or downloads a model at runtime.
"""

from __future__ import annotations

import argparse
import email
import fcntl
import gzip
import hashlib
import json
import os
import platform
import shutil
import subprocess
import sys
import tarfile
import tempfile
import urllib.request
from pathlib import Path
from typing import Any, Iterable

ROOT = Path(__file__).resolve().parents[1]
LOCK_FILE = ROOT / "src-tauri/ocr/offline/requirements-macos-arm64-py310.lock"
CACHE_ROOT = ROOT / ".cache/visualtex-ocr"
WORK_ROOT = CACHE_ROOT / "work-macos-arm64-py310"
DOWNLOAD_ROOT = CACHE_ROOT / "downloads"
OUTPUT_ROOT = ROOT / "dist-ocr/macos-arm64"
MODEL_PACK_ROOT = ROOT / "dist-ocr/model-packs/macos-arm64"
BUILD_LOCK = CACHE_ROOT / "offline-build.lock"

PYTHON_VERSION = "3.10.20"
PYTHON_BUILD_TAG = "20260623"
PYTHON_ASSET_NAME = (
    "cpython-3.10.20+20260623-aarch64-apple-darwin-"
    "install_only_stripped.tar.gz"
)
PYTHON_ASSET_URL = (
    "https://github.com/astral-sh/python-build-standalone/releases/download/"
    "20260623/cpython-3.10.20%2B20260623-aarch64-apple-darwin-"
    "install_only_stripped.tar.gz"
)
PYTHON_ASSET_SHA256 = "1b966fa4f23c0b9a68eafa1c5720c86a622d5fbfedd49a38fe17e18987149caf"
PYTHON_BUILD_LICENSE_URL = (
    "https://raw.githubusercontent.com/astral-sh/python-build-standalone/"
    "20260623/LICENSE"
)
OFFLINE_BUNDLE_FORMAT_VERSION = 2
RUNTIME_LAYOUT_VERSION = 3

DEFAULT_MODEL = "PP-FormulaNet_plus-M"
OPTIONAL_MODELS = ("PP-FormulaNet_plus-S", "PP-FormulaNet_plus-L")
MODEL_HASHES: dict[str, dict[str, str]] = {
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


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def file_record(path: Path) -> dict[str, Any]:
    return {
        "name": path.name,
        "size": path.stat().st_size,
        "sha256": sha256_file(path),
    }


def download(url: str, destination: Path, expected_sha256: str | None = None) -> Path:
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.is_file():
        current = sha256_file(destination)
        if expected_sha256 is None or current == expected_sha256:
            return destination
        destination.unlink()

    temporary = destination.with_suffix(destination.suffix + ".tmp")
    temporary.unlink(missing_ok=True)
    request = urllib.request.Request(url, headers={"User-Agent": "VisualTeX-offline-builder"})
    with urllib.request.urlopen(request, timeout=120) as response, temporary.open("wb") as output:
        shutil.copyfileobj(response, output, 1024 * 1024)
    current = sha256_file(temporary)
    if expected_sha256 is not None and current != expected_sha256:
        temporary.unlink(missing_ok=True)
        raise RuntimeError(
            f"SHA-256 mismatch for {url}: expected {expected_sha256}, got {current}"
        )
    temporary.replace(destination)
    return destination


def safe_extract(archive: Path, destination: Path) -> None:
    destination.mkdir(parents=True, exist_ok=True)
    root = destination.resolve()
    with tarfile.open(archive, "r:gz") as tar:
        for member in tar.getmembers():
            target = (destination / member.name).resolve()
            if root != target and root not in target.parents:
                raise RuntimeError(f"Unsafe archive path: {member.name}")
            if member.issym() or member.islnk():
                link_target = (target.parent / member.linkname).resolve()
                if root != link_target and root not in link_target.parents:
                    raise RuntimeError(
                        f"Unsafe archive link: {member.name} -> {member.linkname}"
                    )
        tar.extractall(destination)


def run(command: list[str], *, env: dict[str, str] | None = None, cwd: Path | None = None) -> None:
    print("+", " ".join(command), flush=True)
    subprocess.run(command, check=True, cwd=cwd, env=env)


def clean_python_tree(python_root: Path) -> None:
    for directory in sorted(python_root.rglob("__pycache__"), reverse=True):
        shutil.rmtree(directory, ignore_errors=True)
    for file in python_root.rglob("*.pyc"):
        file.unlink(missing_ok=True)
    for directory in [
        python_root / "share/doc",
        python_root / "share/man",
    ]:
        shutil.rmtree(directory, ignore_errors=True)

    bin_directory = python_root / "bin"
    if bin_directory.is_dir():
        for script in sorted(bin_directory.iterdir(), key=lambda path: path.name):
            if not script.is_file() or script.is_symlink():
                continue
            data = script.read_bytes()
            if not data.startswith(b"#!"):
                continue
            newline = data.find(b"\n")
            if newline < 0:
                continue
            normalized = b"#!/usr/bin/env python3\n" + data[newline + 1 :]
            if normalized != data:
                mode = script.stat().st_mode
                script.write_bytes(normalized)
                script.chmod(mode)


def verify_relocatable_tree(python_root: Path) -> None:
    forbidden = {
        str(ROOT).encode(),
        str(CACHE_ROOT).encode(),
        b".cache/visualtex-ocr",
    }
    violations: list[str] = []
    for path in sorted(python_root.rglob("*"), key=lambda item: item.as_posix()):
        if not path.is_file() or path.is_symlink() or path.stat().st_size > 5 * 1024 * 1024:
            continue
        data = path.read_bytes()
        if any(marker in data for marker in forbidden):
            violations.append(path.relative_to(python_root).as_posix())
            if len(violations) >= 20:
                break
    if violations:
        raise RuntimeError(
            "Bundled Python contains non-relocatable build paths: "
            + ", ".join(violations)
        )


def verify_python_runtime(python: Path) -> dict[str, str]:
    script = """
import json, platform, sys
import paddle, paddleocr, paddlex, numpy, PIL, tokenizers
print(json.dumps({
  'python': platform.python_version(),
  'architecture': platform.machine(),
  'paddlepaddle': paddle.__version__,
  'paddleocr': paddleocr.__version__,
  'paddlex': paddlex.__version__,
  'numpy': numpy.__version__,
  'pillow': PIL.__version__,
  'tokenizers': tokenizers.__version__,
  'executable': sys.executable,
}, sort_keys=True))
"""
    output = subprocess.check_output([str(python), "-c", script], text=True)
    values = json.loads(output.strip().splitlines()[-1])
    expected = {
        "python": PYTHON_VERSION,
        "architecture": "arm64",
        "paddlepaddle": "3.3.1",
        "paddleocr": "3.7.0",
        "paddlex": "3.7.2",
        "tokenizers": "0.19.1",
    }
    for key, value in expected.items():
        if values.get(key) != value:
            raise RuntimeError(f"Unexpected {key}: {values.get(key)!r}; expected {value!r}")
    return {key: str(value) for key, value in values.items()}


def prepare_python_runtime(force: bool) -> tuple[Path, dict[str, str]]:
    runtime_stage = WORK_ROOT / "runtime-stage"
    python_root = runtime_stage / "python"
    python = python_root / "bin/python3"
    marker = runtime_stage / ".visualtex-ready.json"
    lock_sha = sha256_file(LOCK_FILE)
    expected_marker = {
        "pythonAssetSha256": PYTHON_ASSET_SHA256,
        "requirementsSha256": lock_sha,
        "runtimeLayoutVersion": RUNTIME_LAYOUT_VERSION,
    }
    if not force and marker.is_file() and python.is_file():
        try:
            current = json.loads(marker.read_text("utf-8"))
            source_inputs_match = all(
                current.get(key) == value
                for key, value in expected_marker.items()
                if key != "runtimeLayoutVersion"
            )
            if source_inputs_match:
                versions = verify_python_runtime(python)
                clean_python_tree(python_root)
                verify_relocatable_tree(python_root)
                marker.write_text(
                    json.dumps(
                        {**expected_marker, "versions": versions},
                        indent=2,
                        sort_keys=True,
                    )
                    + "\n",
                    "utf-8",
                )
                return runtime_stage, versions
        except Exception:
            pass

    shutil.rmtree(runtime_stage, ignore_errors=True)
    extract_root = WORK_ROOT / "python-extract"
    shutil.rmtree(extract_root, ignore_errors=True)
    archive = download(
        PYTHON_ASSET_URL,
        DOWNLOAD_ROOT / PYTHON_ASSET_NAME,
        PYTHON_ASSET_SHA256,
    )
    safe_extract(archive, extract_root)
    extracted_python = extract_root / "python"
    if not (extracted_python / "bin/python3").is_file():
        raise RuntimeError("python-build-standalone archive has no python/bin/python3")
    runtime_stage.mkdir(parents=True, exist_ok=True)
    shutil.move(str(extracted_python), str(python_root))
    shutil.rmtree(extract_root, ignore_errors=True)

    environment = os.environ.copy()
    environment.update(
        {
            "PIP_DISABLE_PIP_VERSION_CHECK": "1",
            "PIP_NO_INPUT": "1",
            "PYTHONNOUSERSITE": "1",
        }
    )
    run(
        [
            str(python),
            "-m",
            "pip",
            "install",
            "--no-cache-dir",
            "--only-binary=:all:",
            "--requirement",
            str(LOCK_FILE),
        ],
        env=environment,
    )
    versions = verify_python_runtime(python)
    clean_python_tree(python_root)
    verify_relocatable_tree(python_root)
    marker.write_text(
        json.dumps({**expected_marker, "versions": versions}, indent=2, sort_keys=True) + "\n",
        "utf-8",
    )
    return runtime_stage, versions


def model_candidates(model_name: str) -> Iterable[Path]:
    configured = os.environ.get("VISUALTEX_OCR_MODEL_SOURCE_DIR", "").strip()
    if configured:
        configured_path = Path(configured).expanduser()
        yield configured_path / model_name if configured_path.name != model_name else configured_path
    home = Path.home()
    yield (
        home
        / "Library/Application Support/com.visualtex.studio/ocr-runtime/cache/paddlex/official_models"
        / model_name
    )
    yield home / ".paddlex/official_models" / model_name
    yield WORK_ROOT / "model-download/paddlex/official_models" / model_name


def verify_model(directory: Path, model_name: str) -> dict[str, dict[str, Any]]:
    expected = MODEL_HASHES[model_name]
    records: dict[str, dict[str, Any]] = {}
    for name, digest in expected.items():
        file = directory / name
        if not file.is_file():
            raise RuntimeError(f"Missing {model_name}/{name}")
        actual = sha256_file(file)
        if actual != digest:
            raise RuntimeError(
                f"SHA-256 mismatch for {model_name}/{name}: expected {digest}, got {actual}"
            )
        records[name] = file_record(file)
    return records


def locate_or_download_model(
    model_name: str,
    runtime_stage: Path,
) -> tuple[Path, dict[str, dict[str, Any]]]:
    for candidate in model_candidates(model_name):
        if candidate.is_dir():
            try:
                return candidate, verify_model(candidate, model_name)
            except RuntimeError:
                continue

    cache_home = WORK_ROOT / "model-download/paddlex"
    python = runtime_stage / "python/bin/python3"
    environment = os.environ.copy()
    environment.update(
        {
            "PADDLE_PDX_CACHE_HOME": str(cache_home),
            "PADDLE_PDX_MODEL_SOURCE": "BOS",
            "PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK": "True",
            "PYTHONNOUSERSITE": "1",
        }
    )
    run(
        [
            str(python),
            "-c",
            (
                "from paddleocr import FormulaRecognition; "
                f"FormulaRecognition(model_name={model_name!r}, device='cpu')"
            ),
        ],
        env=environment,
    )
    downloaded = cache_home / "official_models" / model_name
    return downloaded, verify_model(downloaded, model_name)


def deterministic_archive(source: Path, output: Path, arcname: str) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_name(f".{output.name}.{os.getpid()}.tmp")
    temporary.unlink(missing_ok=True)
    with temporary.open("wb") as raw:
        with gzip.GzipFile(filename="", mode="wb", fileobj=raw, mtime=0, compresslevel=9) as gz:
            with tarfile.open(fileobj=gz, mode="w", format=tarfile.PAX_FORMAT) as tar:
                paths = [source, *sorted(source.rglob("*"), key=lambda item: item.as_posix())]
                for path in paths:
                    relative = path.relative_to(source)
                    name = arcname if relative == Path(".") else f"{arcname}/{relative.as_posix()}"
                    info = tar.gettarinfo(str(path), arcname=name)
                    info.uid = 0
                    info.gid = 0
                    info.uname = ""
                    info.gname = ""
                    info.mtime = 0
                    if info.isfile():
                        with path.open("rb") as stream:
                            tar.addfile(info, stream)
                    else:
                        tar.addfile(info)
    temporary.replace(output)


def atomic_write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    temporary.unlink(missing_ok=True)
    payload = json.dumps(value, indent=2, ensure_ascii=False, sort_keys=True) + "\n"
    with temporary.open("w", encoding="utf-8") as stream:
        stream.write(payload)
        stream.flush()
        os.fsync(stream.fileno())
    temporary.replace(path)


def cleanup_stale_output_temporaries() -> None:
    if not OUTPUT_ROOT.is_dir():
        return
    stale = [
        path
        for path in OUTPUT_ROOT.iterdir()
        if path.is_file()
        and path.name.startswith(".")
        and path.name.endswith(".tmp")
    ]
    if not stale:
        return
    total_bytes = sum(path.stat().st_size for path in stale)
    for path in stale:
        path.unlink(missing_ok=True)
    print(
        f"Removed {len(stale)} stale OCR build temporary files "
        f"({total_bytes / 1024 / 1024:.1f} MB)."
    )


def verify_manifest_outputs(manifest: dict[str, Any]) -> None:
    archives = manifest.get("archives")
    if not isinstance(archives, dict) or not archives:
        raise RuntimeError("Offline OCR manifest has no archive records")
    for label, record in archives.items():
        if not isinstance(record, dict):
            raise RuntimeError(f"Offline OCR manifest archive {label} is invalid")
        name = record.get("name")
        if not isinstance(name, str) or not name:
            raise RuntimeError(f"Offline OCR manifest archive {label} has no file name")
        path = OUTPUT_ROOT / name
        actual = file_record(path)
        if actual["size"] != record.get("size") or actual["sha256"] != record.get("sha256"):
            raise RuntimeError(
                f"Offline OCR manifest mismatch for {name}: expected "
                f"size={record.get('size')} sha256={record.get('sha256')}, "
                f"actual size={actual['size']} sha256={actual['sha256']}"
            )


def collect_licenses(runtime_stage: Path, destination: Path) -> list[dict[str, Any]]:
    shutil.rmtree(destination, ignore_errors=True)
    destination.mkdir(parents=True, exist_ok=True)
    python_root = runtime_stage / "python"
    python_license = next(
        (
            path
            for path in [
                python_root / "LICENSE",
                python_root / "LICENSE.txt",
                python_root / "lib/python3.10/LICENSE.txt",
            ]
            if path.is_file()
        ),
        None,
    )
    if python_license:
        shutil.copy2(python_license, destination / "CPYTHON-LICENSE.txt")

    standalone_license = download(
        PYTHON_BUILD_LICENSE_URL,
        DOWNLOAD_ROOT / "python-build-standalone-LICENSE.txt",
    )
    shutil.copy2(standalone_license, destination / "PYTHON-BUILD-STANDALONE-LICENSE.txt")

    site_packages = python_root / "lib/python3.10/site-packages"
    packages: list[dict[str, Any]] = []
    for dist_info in sorted(site_packages.glob("*.dist-info"), key=lambda path: path.name.lower()):
        metadata_file = dist_info / "METADATA"
        if not metadata_file.is_file():
            continue
        message = email.message_from_bytes(metadata_file.read_bytes())
        name = str(message.get("Name", dist_info.name)).strip()
        version = str(message.get("Version", "")).strip()
        license_value = str(
            message.get("License-Expression") or message.get("License") or ""
        ).strip()
        package_dir = destination / "python-packages" / f"{name}-{version}"
        copied: list[str] = []
        for file in sorted(dist_info.rglob("*"), key=lambda path: path.as_posix()):
            if not file.is_file():
                continue
            lowered = file.name.lower()
            relative_lower = file.relative_to(dist_info).as_posix().lower()
            if (
                lowered.startswith(("license", "copying", "notice", "authors"))
                or "/licenses/" in f"/{relative_lower}"
                or relative_lower.startswith("licenses/")
            ):
                target = package_dir / file.relative_to(dist_info)
                target.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(file, target)
                copied.append(target.relative_to(destination).as_posix())
        packages.append(
            {
                "name": name,
                "version": version,
                "license": license_value,
                "licenseFiles": copied,
            }
        )
    (destination / "THIRD_PARTY_NOTICES.json").write_text(
        json.dumps(packages, indent=2, ensure_ascii=False, sort_keys=True) + "\n",
        "utf-8",
    )
    return packages


def build_fingerprint() -> str:
    digest = hashlib.sha256()
    digest.update(f"offline-bundle-format:{OFFLINE_BUNDLE_FORMAT_VERSION}".encode())
    digest.update(f"runtime-layout:{RUNTIME_LAYOUT_VERSION}".encode())
    digest.update(LOCK_FILE.read_bytes())
    digest.update(PYTHON_ASSET_SHA256.encode())
    digest.update(json.dumps(MODEL_HASHES, sort_keys=True).encode())
    return digest.hexdigest()


def archive_record_is_valid(manifest: dict[str, Any], label: str) -> bool:
    try:
        record = manifest["archives"][label]
        path = OUTPUT_ROOT / record["name"]
        return (
            path.is_file()
            and path.stat().st_size == record.get("size")
            and sha256_file(path) == record.get("sha256")
        )
    except (KeyError, TypeError, OSError):
        return False


def runtime_archive_is_reusable(manifest: dict[str, Any]) -> bool:
    return (
        manifest.get("platform") == "macos"
        and manifest.get("architecture") == "arm64"
        and manifest.get("python", {}).get("sourceSha256") == PYTHON_ASSET_SHA256
        and manifest.get("runtimeLayoutVersion") == RUNTIME_LAYOUT_VERSION
        and manifest.get("requirements", {}).get("sha256") == sha256_file(LOCK_FILE)
        and archive_record_is_valid(manifest, "runtime")
    )


def default_model_archive_is_reusable(manifest: dict[str, Any]) -> bool:
    try:
        if manifest.get("platform") != "macos" or manifest.get("architecture") != "arm64":
            return False
        if manifest.get("defaultModel", {}).get("name") != DEFAULT_MODEL:
            return False
        records = manifest.get("defaultModel", {}).get("files", {})
        if set(records) != set(MODEL_HASHES[DEFAULT_MODEL]):
            return False
        if any(
            records[name].get("sha256") != digest
            for name, digest in MODEL_HASHES[DEFAULT_MODEL].items()
        ):
            return False
        return archive_record_is_valid(manifest, "defaultModel")
    except (KeyError, TypeError):
        return False


def build_default_bundle(force: bool) -> None:
    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)
    cleanup_stale_output_temporaries()
    fingerprint = build_fingerprint()
    manifest_path = OUTPUT_ROOT / "manifest.json"
    current: dict[str, Any] | None = None
    if manifest_path.is_file():
        try:
            current = json.loads(manifest_path.read_text("utf-8"))
        except (json.JSONDecodeError, OSError):
            current = None

    runtime_reusable = (
        not force and current is not None and runtime_archive_is_reusable(current)
    )
    model_reusable = (
        not force and current is not None and default_model_archive_is_reusable(current)
    )
    archives_reusable = runtime_reusable and model_reusable
    licenses_complete = all(
        path.is_file() and path.stat().st_size > 0
        for path in [
            OUTPUT_ROOT / "licenses/CPYTHON-LICENSE.txt",
            OUTPUT_ROOT / "licenses/PYTHON-BUILD-STANDALONE-LICENSE.txt",
            OUTPUT_ROOT / "licenses/THIRD_PARTY_NOTICES.json",
        ]
    )
    if (
        archives_reusable
        and current is not None
        and current.get("buildFingerprint") == fingerprint
        and licenses_complete
    ):
        verify_manifest_outputs(current)
        print("Offline OCR bundle is already up to date.")
        return

    manifest_path.unlink(missing_ok=True)
    runtime_stage, versions = prepare_python_runtime(force)
    model_source, model_files = locate_or_download_model(DEFAULT_MODEL, runtime_stage)
    runtime_archive = OUTPUT_ROOT / "runtime-python310-macos-arm64.tar.gz"
    model_archive = OUTPUT_ROOT / f"model-{DEFAULT_MODEL}.tar.gz"
    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)

    if runtime_reusable:
        print("Reusing verified OCR runtime archive.")
    else:
        deterministic_archive(runtime_stage / "python", runtime_archive, "python")

    if model_reusable:
        print("Reusing verified default M model archive.")
    else:
        model_stage = WORK_ROOT / "default-model-stage/paddlex/official_models" / DEFAULT_MODEL
        shutil.rmtree(model_stage.parent.parent.parent, ignore_errors=True)
        model_stage.parent.mkdir(parents=True, exist_ok=True)
        shutil.copytree(model_source, model_stage)
        verify_model(model_stage, DEFAULT_MODEL)
        deterministic_archive(model_stage.parent.parent, model_archive, "paddlex")

    packages = collect_licenses(runtime_stage, OUTPUT_ROOT / "licenses")

    manifest = {
        "schemaVersion": 1,
        "platform": "macos",
        "architecture": "arm64",
        "buildFingerprint": fingerprint,
        "runtimeLayoutVersion": RUNTIME_LAYOUT_VERSION,
        "python": {
            "version": PYTHON_VERSION,
            "buildTag": PYTHON_BUILD_TAG,
            "sourceUrl": PYTHON_ASSET_URL,
            "sourceSha256": PYTHON_ASSET_SHA256,
        },
        "versions": versions,
        "requirements": {
            "name": LOCK_FILE.name,
            "sha256": sha256_file(LOCK_FILE),
        },
        "archives": {
            "runtime": file_record(runtime_archive),
            "defaultModel": file_record(model_archive),
        },
        "defaultModel": {
            "name": DEFAULT_MODEL,
            "files": model_files,
        },
        "thirdPartyPackageCount": len(packages),
    }
    verify_manifest_outputs(manifest)
    atomic_write_json(manifest_path, manifest)
    verify_manifest_outputs(json.loads(manifest_path.read_text("utf-8")))
    print(
        f"Prepared offline OCR bundle: {runtime_archive.stat().st_size / 1024 / 1024:.1f} MB runtime, "
        f"{model_archive.stat().st_size / 1024 / 1024:.1f} MB model"
    )


def build_optional_model_packs(force: bool) -> None:
    runtime_stage, _ = prepare_python_runtime(False)
    MODEL_PACK_ROOT.mkdir(parents=True, exist_ok=True)
    for legacy in MODEL_PACK_ROOT.glob("VisualTeX-*.json"):
        legacy.unlink(missing_ok=True)
    for model_name in OPTIONAL_MODELS:
        source, files = locate_or_download_model(model_name, runtime_stage)
        pack_root = WORK_ROOT / "optional-model-stage" / model_name / "visualtex-model-pack"
        shutil.rmtree(pack_root.parent, ignore_errors=True)
        target = pack_root / "paddlex/official_models" / model_name
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copytree(source, target)
        verify_model(target, model_name)
        pack_manifest = {
            "schemaVersion": 1,
            "platform": "macos",
            "architecture": "arm64",
            "model": model_name,
            "files": files,
        }
        atomic_write_json(pack_root / "pack-manifest.json", pack_manifest)
        archive = MODEL_PACK_ROOT / f"VisualTeX-{model_name}-macos-arm64.vtxocrmodel"
        if force or not archive.is_file():
            deterministic_archive(pack_root, archive, "visualtex-model-pack")
        print(
            f"Prepared optional model pack {model_name}: "
            f"{archive.stat().st_size / 1024 / 1024:.1f} MB"
        )


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--force", action="store_true", help="rebuild all generated assets")
    parser.add_argument(
        "--with-optional-model-packs",
        action="store_true",
        help="also generate standalone S and L model archives",
    )
    args = parser.parse_args()

    if platform.system() != "Darwin" or platform.machine() != "arm64":
        raise SystemExit(
            "The current offline OCR bundle target is macOS arm64; run this script on Apple Silicon."
        )
    if not LOCK_FILE.is_file():
        raise SystemExit(f"Missing lock file: {LOCK_FILE}")

    DOWNLOAD_ROOT.mkdir(parents=True, exist_ok=True)
    WORK_ROOT.mkdir(parents=True, exist_ok=True)
    BUILD_LOCK.parent.mkdir(parents=True, exist_ok=True)
    with BUILD_LOCK.open("a+b") as lock:
        fcntl.flock(lock.fileno(), fcntl.LOCK_EX)
        build_default_bundle(args.force)
        if args.with_optional_model_packs:
            build_optional_model_packs(args.force)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
