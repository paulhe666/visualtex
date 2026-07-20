#!/usr/bin/env python3
"""Smoke-test the VisualTeX OCR worker protocol and parent-process watcher."""

from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from pathlib import Path


def read_json_line(process: subprocess.Popen[str], label: str) -> dict[str, object]:
    assert process.stdout is not None
    line = process.stdout.readline()
    if not line:
        stderr = process.stderr.read() if process.stderr is not None else ""
        raise RuntimeError(
            f"OCR worker closed before {label}; returncode={process.poll()}; stderr={stderr}"
        )
    try:
        payload = json.loads(line)
    except json.JSONDecodeError as error:
        raise RuntimeError(f"Invalid JSON while waiting for {label}: {line!r}") from error
    if not isinstance(payload, dict):
        raise RuntimeError(f"Unexpected {label} payload: {payload!r}")
    return payload


def main() -> int:
    root = Path(__file__).resolve().parents[1]
    worker_path = root / "src-tauri" / "ocr" / "worker.py"
    environment = os.environ.copy()
    environment.update(
        {
            "VISUALTEX_PARENT_PID": str(os.getpid()),
            "PYTHONUTF8": "1",
            "PYTHONIOENCODING": "utf-8",
        }
    )

    process = subprocess.Popen(
        [sys.executable, str(worker_path)],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        env=environment,
    )

    try:
        ready = read_json_line(process, "ready signal")
        if ready.get("event") != "ready":
            raise RuntimeError(f"Unexpected ready response: {ready!r}")

        # The old Windows implementation called os.kill(parent_pid, 0) every
        # two seconds. Waiting longer than that makes this a regression test for
        # the parent-process watcher rather than only a protocol test.
        time.sleep(3.2)
        if process.poll() is not None:
            stderr = process.stderr.read() if process.stderr is not None else ""
            raise RuntimeError(
                f"OCR worker exited during parent watch; returncode={process.returncode}; stderr={stderr}"
            )

        assert process.stdin is not None
        process.stdin.write('{"id":"smoke","action":"ping"}\n')
        process.stdin.flush()
        pong = read_json_line(process, "ping response")
        if pong.get("event") != "pong" or pong.get("ok") is not True:
            raise RuntimeError(f"Unexpected ping response: {pong!r}")

        process.stdin.write('{"id":"stop","action":"shutdown"}\n')
        process.stdin.flush()
        shutdown = read_json_line(process, "shutdown response")
        if shutdown.get("event") != "shutdown" or shutdown.get("ok") is not True:
            raise RuntimeError(f"Unexpected shutdown response: {shutdown!r}")

        returncode = process.wait(timeout=10)
        if returncode != 0:
            stderr = process.stderr.read() if process.stderr is not None else ""
            raise RuntimeError(
                f"OCR worker returned {returncode} after shutdown; stderr={stderr}"
            )
    finally:
        if process.poll() is None:
            process.kill()
            process.wait(timeout=10)

    print("OCR worker lifecycle smoke test passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
