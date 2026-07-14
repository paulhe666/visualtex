; VisualTeX Windows installer prerequisite check and per-user OLE choice.
; VSTO remains a deferred development path and is not shipped by this installer.

Var VisualTeXOfficeChoice
Var VisualTeXOfficeOnlyRadio
Var VisualTeXOfficeOleRadio

Page custom VisualTeXOfficePageCreate VisualTeXOfficePageLeave

Function VisualTeXOfficePageCreate
  nsDialogs::Create 1018
  Pop $0
  ${If} $0 == error
    Abort
  ${EndIf}

  ${NSD_CreateLabel} 0 0 100% 30u "请选择是否启用 Windows Office OLE 集成 / Choose Windows Office OLE integration"
  Pop $0

  ${NSD_CreateRadioButton} 0 38u 100% 16u "仅 VisualTeX（不安装 Office 插件） / VisualTeX only"
  Pop $VisualTeXOfficeOnlyRadio

  ${NSD_CreateRadioButton} 0 62u 100% 16u "VisualTeX + OLE Office 集成（安装简单）"
  Pop $VisualTeXOfficeOleRadio
  ${NSD_Check} $VisualTeXOfficeOleRadio

  ${NSD_CreateLabel} 0 92u 100% 42u "OLE 使用 Office.js Ribbon 与当前用户命名管道。安装程序会清理或禁用旧版 VisualTeX VSTO 加载项，避免重复按钮。"
  Pop $0

  nsDialogs::Show
FunctionEnd

Function VisualTeXOfficePageLeave
  ${NSD_GetState} $VisualTeXOfficeOleRadio $0
  ${If} $0 == ${BST_CHECKED}
    StrCpy $VisualTeXOfficeChoice "ole"
    Return
  ${EndIf}
  StrCpy $VisualTeXOfficeChoice "none"
FunctionEnd

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

!macro NSIS_HOOK_POSTINSTALL
  DetailPrint "Applying the selected VisualTeX Office integration mode: $VisualTeXOfficeChoice"
  ${If} $VisualTeXOfficeChoice == "ole"
    MessageBox MB_ICONINFORMATION|MB_OK "接下来安装程序会自动打开 Word 和 PowerPoint 的 Office 加载项窗口，用于添加 VisualTeX。$\r$\n$\r$\n请不要关闭、切换或操作这些窗口；安装程序完成配置后会自动将它们关闭。整个过程通常需要约 1 分钟。$\r$\n$\r$\nVisualTeX will temporarily open the Office Add-ins dialogs in Word and PowerPoint. Do not close or interact with them; setup will close them automatically."
    DetailPrint "Automatically configuring Word and PowerPoint. Keep the Office Add-ins windows open until setup closes them."
    ; Best-effort removal of legacy VisualTeX VSTO MSI instances. The OLE
    ; installer also forces any surviving add-in LoadBehavior values to zero.
    IfFileExists "$INSTDIR\scripts\uninstall_windows_vsto.ps1" 0 +3
    nsExec::ExecToLog `powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\uninstall_windows_vsto.ps1"`
    Pop $0
    IfFileExists "$INSTDIR\scripts\install_windows_ole.ps1" 0 visualtex_office_missing
    nsExec::ExecToLog `powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\install_windows_ole.ps1"`
    Pop $0
    StrCmp $0 "0" visualtex_office_done
    SetDetailsView show
    MessageBox MB_ICONEXCLAMATION "VisualTeX 主程序已安装，但 Office 集成配置未完成。$\r$\n$\r$\n如果你关闭了刚才自动打开的 Word 或 PowerPoint 加载项窗口，配置会被中断。请关闭所有 Office 窗口，然后在 VisualTeX 设置中点击修复；修复期间不要关闭或操作自动打开的 Office 窗口。$\r$\n$\r$\n详细错误已显示在安装日志区域。"
    Goto visualtex_office_done
  ${Else}
    IfFileExists "$INSTDIR\scripts\uninstall_windows_ole.ps1" 0 +3
    nsExec::ExecToLog `powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\uninstall_windows_ole.ps1"`
    Pop $0
    IfFileExists "$INSTDIR\scripts\uninstall_windows_vsto.ps1" 0 visualtex_office_done
    nsExec::ExecToLog `powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\uninstall_windows_vsto.ps1"`
    Pop $0
    Goto visualtex_office_done
  ${EndIf}

visualtex_office_missing:
  MessageBox MB_ICONEXCLAMATION "Windows OLE Office 安装资源缺失。VisualTeX 主程序已正常安装。"
  Goto visualtex_office_done

visualtex_office_done:
!macroend

!macro NSIS_HOOK_PREUNINSTALL
  IfFileExists "$INSTDIR\scripts\uninstall_windows_ole.ps1" 0 +3
  nsExec::ExecToLog `powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\uninstall_windows_ole.ps1"`
  Pop $0
  IfFileExists "$INSTDIR\scripts\uninstall_windows_vsto.ps1" 0 +3
  nsExec::ExecToLog `powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\uninstall_windows_vsto.ps1"`
  Pop $0
  IfFileExists "$INSTDIR\scripts\remove_windows_office_certificate.ps1" 0 visualtex_preuninstall_done
  nsExec::ExecToLog `powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\remove_windows_office_certificate.ps1"`
  Pop $0
visualtex_preuninstall_done:
!macroend
