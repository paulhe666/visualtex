"""VisualTeX private OCR Python isolation guard.

This runtime is application-owned. User-level Python packages must never
participate in dependency resolution or imports.
"""

from __future__ import annotations

import os
import site
import sys

os.environ["PYTHONNOUSERSITE"] = "1"

# PaddlePaddle's Windows CPU runtime depends on Microsoft's OpenMP runtime.
# Preload the app-local copy beside python.exe so an older or incompatible
# system-wide vcomp140.dll cannot win DLL resolution.
_visualtex_vcomp140 = None
if os.name == "nt":
    _runtime_root = os.path.dirname(sys.executable)
    _vcomp140_path = os.path.join(_runtime_root, "vcomp140.dll")
    if os.path.isfile(_vcomp140_path):
        import ctypes

        _visualtex_vcomp140 = ctypes.WinDLL(_vcomp140_path)

try:
    configured = site.getusersitepackages()
    user_sites = [configured] if isinstance(configured, str) else list(configured)
except Exception:
    user_sites = []


def _normalized(value: str) -> str:
    return os.path.normcase(os.path.abspath(value))


normalized_user_sites = [_normalized(path) for path in user_sites if path]
sys.path[:] = [
    path
    for path in sys.path
    if not any(
        _normalized(path) == user_site
        or _normalized(path).startswith(user_site + os.sep)
        for user_site in normalized_user_sites
    )
]
site.ENABLE_USER_SITE = False
