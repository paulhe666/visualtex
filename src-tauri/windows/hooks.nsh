; VisualTeX Windows installer prerequisite check.
; The editor itself works without Python, so an incompatible environment warns
; the user instead of silently failing later or blocking installation outright.

!macro VisualTeXProbeLauncher SELECTOR
  nsExec::ExecToStack `"py.exe" ${SELECTOR} -c "import platform,sys;sys.exit(0 if (3,9) <= sys.version_info[:2] <= (3,13) and platform.machine().lower() in ('amd64','x86_64','x64') else 1)"`
  Pop $0
  Pop $1
  StrCmp $0 "0" visualtex_python_ok
!macroend

!macro VisualTeXProbeCommand PROGRAM
  nsExec::ExecToStack `"${PROGRAM}" -c "import platform,sys;sys.exit(0 if (3,9) <= sys.version_info[:2] <= (3,13) and platform.machine().lower() in ('amd64','x86_64','x64') else 1)"`
  Pop $0
  Pop $1
  StrCmp $0 "0" visualtex_python_ok
!macroend

!macro NSIS_HOOK_PREINSTALL
  DetailPrint "Checking the Python environment required by VisualTeX OCR..."

  ; Probe every supported runtime through both the new Python Install Manager
  ; selector and the legacy py launcher selector. A default Python 3.14 must not
  ; hide a compatible side-by-side installation.
  !insertmacro VisualTeXProbeLauncher "-V:3.13"
  !insertmacro VisualTeXProbeLauncher "-3.13"
  !insertmacro VisualTeXProbeLauncher "-V:3.12"
  !insertmacro VisualTeXProbeLauncher "-3.12"
  !insertmacro VisualTeXProbeLauncher "-V:3.11"
  !insertmacro VisualTeXProbeLauncher "-3.11"
  !insertmacro VisualTeXProbeLauncher "-V:3.10"
  !insertmacro VisualTeXProbeLauncher "-3.10"
  !insertmacro VisualTeXProbeLauncher "-V:3.9"
  !insertmacro VisualTeXProbeLauncher "-3.9"

  ; Fall back to interpreters exposed directly on PATH.
  !insertmacro VisualTeXProbeCommand "python.exe"
  !insertmacro VisualTeXProbeCommand "python"

  MessageBox MB_ICONEXCLAMATION|MB_YESNO "未检测到可用于 OCR 的 64 位 Python 3.9–3.13。$\r$\n$\r$\nVisualTeX 编辑器仍可正常安装和使用，但图片公式 OCR 将不可用。请安装 x64 Python 3.13，并启用 Python Launcher。仅安装默认 Python 3.14 不兼容当前 OCR 运行环境。$\r$\n$\r$\nNo compatible 64-bit Python 3.9–3.13 installation was detected. The editor can still be installed, but formula OCR will remain unavailable until a compatible Python runtime is installed.$\r$\n$\r$\n是否继续安装？ / Continue installation?" IDYES visualtex_python_continue

  Abort "VisualTeX installation cancelled because the OCR Python prerequisite is missing."

visualtex_python_continue:
  DetailPrint "Continuing without a compatible OCR Python runtime."
  Goto visualtex_python_check_done

visualtex_python_ok:
  DetailPrint "Compatible Python 3.9–3.13 x64 runtime detected."

visualtex_python_check_done:
!macroend
