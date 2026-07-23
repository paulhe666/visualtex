; VisualTeX Windows installer prerequisite check and per-user native Office OLE choice.
; The production path installs the per-user VSTO add-ins and ATL OLE LocalServer
; without opening Word/PowerPoint or driving Office UI Automation.

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

  ${NSD_CreateRadioButton} 0 62u 100% 16u "VisualTeX + 原生 Office OLE 集成（推荐）"
  Pop $VisualTeXOfficeOleRadio
  ${NSD_Check} $VisualTeXOfficeOleRadio

  ${NSD_CreateLabel} 0 92u 100% 42u "原生模式使用 Word/PowerPoint VSTO Ribbon 与 ATL OLE LocalServer。安装过程不会启动或操作 Office，并会清理旧版 Office.js 集成以避免重复按钮。"
  Pop $0

  nsDialogs::Show
FunctionEnd

Function VisualTeXOfficePageLeave
  ${NSD_GetState} $VisualTeXOfficeOleRadio $0
  ${If} $0 == ${BST_CHECKED}
    nsExec::ExecToStack `powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -Command "if (Get-Process WINWORD,POWERPNT,EXCEL,OUTLOOK,ONENOTE,MSACCESS,MSPUB,VISIO,MSPROJECT -ErrorAction SilentlyContinue) { exit 1 }; exit 0"`
    Pop $1
    Pop $2
    ${If} $1 != "0"
      MessageBox MB_ICONEXCLAMATION|MB_YESNO "检测到 Microsoft Office 仍在运行。强制关闭会立即结束 Word、PowerPoint、Excel、Outlook、OneNote、Access、Publisher、Visio 和 Project；未保存的 Office 文档可能丢失。$\r$\n$\r$\n是否强制关闭所有这些 Office 进程并继续安装？选择“否”将返回上一页，由您自行保存并关闭 Office。$\r$\n$\r$\nMicrosoft Office is still running. Force closing will terminate all common Office apps immediately and may discard unsaved work.$\r$\n$\r$\nForce close all Office processes and continue? Choose No to go back and close Office yourself." IDYES visualtex_force_close_office IDNO visualtex_office_close_declined

visualtex_force_close_office:
      nsExec::ExecToStack `powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -Command "Get-Process WINWORD,POWERPNT,EXCEL,OUTLOOK,ONENOTE,MSACCESS,MSPUB,VISIO,MSPROJECT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue; Start-Sleep -Milliseconds 800; if (Get-Process WINWORD,POWERPNT,EXCEL,OUTLOOK,ONENOTE,MSACCESS,MSPUB,VISIO,MSPROJECT -ErrorAction SilentlyContinue) { exit 1 }; exit 0"`
      Pop $1
      Pop $2
      ${If} $1 != "0"
        MessageBox MB_ICONSTOP "无法完全关闭所有 Office 进程。请保存工作并在任务管理器中关闭残留的 Office 进程后重试。$\r$\n$\r$\nThe installer could not close every Office process. Save your work, close the remaining Office processes in Task Manager, and try again."
        Abort
      ${EndIf}
      Goto visualtex_office_process_check_done

visualtex_office_close_declined:
      Abort

visualtex_office_process_check_done:
    ${EndIf}
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
  ; Custom pages are skipped by NSIS /S. Preserve an explicit interactive
  ; choice, but default unattended installs to the recommended Office mode.
  ${If} $VisualTeXOfficeChoice == ""
    StrCpy $VisualTeXOfficeChoice "ole"
  ${EndIf}

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
    DetailPrint "Installing the per-user VisualTeX VSTO add-ins and native Formula OLE LocalServer."
    IfFileExists "$INSTDIR\scripts\install_windows_vsto.ps1" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\windows-office\VisualTeX-WindowsOffice-VSTO-x64.msi" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\windows-office\VisualTeX-WindowsOffice-VSTO-x64.sha256.json" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\windows-office\VisualTeX-WindowsOffice-VSTO-x86.msi" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\windows-office\VisualTeX-WindowsOffice-VSTO-x86.sha256.json" 0 visualtex_office_missing
    nsExec::ExecToLog `powershell.exe -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\install_windows_vsto.ps1" -PackageDirectory "$INSTDIR\windows-office"`
    Pop $0
    StrCmp $0 "0" visualtex_office_done
    SetDetailsView show
    MessageBox MB_ICONEXCLAMATION "VisualTeX 主程序已安装，但原生 Office VSTO + OLE 集成安装失败。安装器不会自动关闭或重新启动 Word/PowerPoint。请查看安装详情，并检查 %LOCALAPPDATA%\VisualTeX\office\install-logs 中最新的 vsto-bootstrap 日志。"
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
  MessageBox MB_ICONEXCLAMATION "Windows 原生 Office OLE 安装资源缺失。VisualTeX 主程序已正常安装。"
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
