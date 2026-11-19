; VisualTeX Windows installer prerequisite check and per-user native Office choice.
; The production path installs the per-user Ribbon COM add-ins and ATL OLE
; LocalServer. Legacy Office.js Trusted Catalog resources are not installed.

!define VISUALTEX_INSTALLER_VERSION "1.2.7"
!define VISUALTEX_RUNTIME_GUARD_SOURCE "${__FILEDIR__}\..\..\scripts\manage_private_mathtype_runtime.ps1"
!define VISUALTEX_OFFICE_PROCESS_GUARD_SOURCE "${__FILEDIR__}\..\..\scripts\office_process_guard.ps1"

Var VisualTeXOfficeChoice
Var VisualTeXMsOfficeCheckbox
Var VisualTeXWpsCheckbox
Var VisualTeXImeCheckbox
; "1" / "0": the add-ins run in Microsoft Office / WPS Office (OfficeHostSelection), the input method
Var VisualTeXMsOfficeChoice
Var VisualTeXWpsChoice
Var VisualTeXImeChoice
Var VisualTeXOcrChoice
Var VisualTeXOcrCheckbox
Var VisualTeXOcrResourcePrefix
Var VisualTeXAcceptanceMode

; The generated Tauri PageReinstall function is patched after bundling so the
; same-version maintenance page defaults to "Uninstall VisualTeX" directly at
; control creation time. Do not use a GUI timer here: it races the generated
; page and does not reliably change the checked radio button.

!macro VisualTeXRunOfficeProcessGuard MODE
  InitPluginsDir
  File "/oname=$PLUGINSDIR\visualtex-office-process-guard.ps1" "${VISUALTEX_OFFICE_PROCESS_GUARD_SOURCE}"
  nsExec::ExecToStack /TIMEOUT=20000 `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$PLUGINSDIR\visualtex-office-process-guard.ps1" -Mode ${MODE}`
  Pop $0
  Pop $1
  DetailPrint "Office process guard ${MODE}: ExitCode=$0 Output=$1"
!macroend

Function VisualTeXEnsureOfficeProcessesClosed
visualtex_office_guard_check:
  !insertmacro VisualTeXRunOfficeProcessGuard Check
  ${If} $0 == "0"
    Push "0"
    Return
  ${EndIf}
  ${If} $0 != "10"
    IfSilent visualtex_office_guard_silent_failure 0
    MessageBox MB_ICONEXCLAMATION|MB_RETRYCANCEL "无法检查当前 Office 进程，尚未开始安装 Office 插件。$\r$\n$\r$\n$1$\r$\n$\r$\n可重试，或取消本次 Office 集成。$\r$\n$\r$\nOffice process inspection failed before Office integration was changed. Retry or cancel." IDRETRY visualtex_office_guard_check IDCANCEL visualtex_office_guard_declined
  ${EndIf}

  IfSilent visualtex_office_guard_silent_blocked 0
  MessageBox MB_ICONQUESTION|MB_YESNO|MB_DEFBUTTON2 "检测到 Microsoft Office 进程仍在运行：$\r$\n$\r$\n$1$\r$\n$\r$\n其中可能包含没有任何可见窗口的后台 COM/OLE 进程（例如 POWERPNT.EXE -Embedding）。强制关闭会立即结束这些 Office 进程，未保存的 Office 文档可能丢失。$\r$\n$\r$\n是否强制关闭这些 Office 进程并继续安装？选择“否”不会关闭任何进程。$\r$\n$\r$\nMicrosoft Office processes are still running, including possible hidden COM/OLE instances. Force close them and continue?" IDYES visualtex_office_guard_force_close IDNO visualtex_office_guard_declined

visualtex_office_guard_force_close:
  !insertmacro VisualTeXRunOfficeProcessGuard Stop
  ${If} $0 == "0"
    Push "0"
    Return
  ${EndIf}
  IfSilent visualtex_office_guard_silent_failure 0
  MessageBox MB_ICONEXCLAMATION|MB_RETRYCANCEL "无法完全关闭所有 Office 进程：$\r$\n$\r$\n$1$\r$\n$\r$\n可重试，或取消本次 Office 集成。$\r$\n$\r$\nThe installer could not close every Office process. Retry or cancel Office integration." IDRETRY visualtex_office_guard_check IDCANCEL visualtex_office_guard_declined

visualtex_office_guard_declined:
  Push "1"
  Return
visualtex_office_guard_silent_blocked:
  DetailPrint "Silent Office integration refused to force-close running Office processes: $1"
  Push "2"
  Return
visualtex_office_guard_silent_failure:
  DetailPrint "Office process guard failed: ExitCode=$0 Output=$1"
  Push "3"
FunctionEnd

Function VisualTeXRepairMainUninstallRegistration
  ; Older 1.2.3 builds could leave the remembered install directory while the
  ; standard Add/Remove Programs key was missing. Reconstruct the key before
  ; Tauri's maintenance page tries to launch uninstall.exe.
  ReadRegStr $0 HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\VisualTeX" "UninstallString"
  ${If} $0 != ""
    Return
  ${EndIf}

  ReadRegStr $1 HKCU "Software\visualtex\VisualTeX" ""
  ${If} $1 == ""
    IfFileExists "$LOCALAPPDATA\VisualTeX\uninstall.exe" 0 +2
      StrCpy $1 "$LOCALAPPDATA\VisualTeX"
  ${EndIf}
  ${If} $1 == ""
    IfFileExists "$PROFILE\AppData\VisualTeX\uninstall.exe" 0 +2
      StrCpy $1 "$PROFILE\AppData\VisualTeX"
  ${EndIf}
  ${If} $1 == ""
    Return
  ${EndIf}
  IfFileExists "$1\uninstall.exe" 0 visualtex_repair_uninstall_done

  WriteRegStr HKCU "Software\visualtex\VisualTeX" "" "$1"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\VisualTeX" "DisplayName" "VisualTeX"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\VisualTeX" "DisplayIcon" '$\"$1\visualtex.exe$\"'
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\VisualTeX" "DisplayVersion" "${VISUALTEX_INSTALLER_VERSION}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\VisualTeX" "Publisher" "visualtex"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\VisualTeX" "InstallLocation" '$\"$1$\"'
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\VisualTeX" "UninstallString" '$\"$1\uninstall.exe$\"'
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\VisualTeX" "MainBinaryName" "visualtex.exe"
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\VisualTeX" "NoModify" 1
  WriteRegDWORD HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\VisualTeX" "NoRepair" 1

visualtex_repair_uninstall_done:
FunctionEnd

; The components page: every optional part of VisualTeX as a check box. Microsoft Office and WPS Office
; load the same COM add-ins and formula renderer, so either box installs them; the add-ins stay inert in a
; host whose box is clear (HKCU\Software\VisualTeX\OfficeIntegration MicrosoftOffice / Wps).
Function VisualTeXOfficePageCreate
  Call VisualTeXRepairMainUninstallRegistration
  nsDialogs::Create 1018
  Pop $0
  ${If} $0 == error
    Abort
  ${EndIf}

  !insertmacro MUI_HEADER_TEXT "选择组件 / Components" "勾选要安装的组件，VisualTeX 主程序总会安装"

  ${NSD_CreateLabel} 0 0 100% 12u "Office 插件 / Office add-ins"
  Pop $0
  ${NSD_CreateCheckbox} 8u 13u 100% 13u "Microsoft Office 插件（Word、PowerPoint）"
  Pop $VisualTeXMsOfficeCheckbox
  ${If} $VisualTeXMsOfficeChoice != "0"
    ${NSD_Check} $VisualTeXMsOfficeCheckbox
  ${EndIf}
  ${NSD_CreateCheckbox} 8u 27u 100% 13u "WPS Office 插件（WPS 文字、WPS 演示）"
  Pop $VisualTeXWpsCheckbox
  ${If} $VisualTeXWpsChoice != "0"
    ${NSD_Check} $VisualTeXWpsCheckbox
  ${EndIf}

  ${NSD_CreateLabel} 0 46u 100% 12u "其他组件 / Other components"
  Pop $0
  ${NSD_CreateCheckbox} 8u 59u 100% 13u "VisualTeX 输入法（在任意软件中输入拼音、英文和公式）"
  Pop $VisualTeXImeCheckbox
  ${If} $VisualTeXImeChoice != "0"
    ${NSD_Check} $VisualTeXImeCheckbox
  ${EndIf}
!ifndef VISUALTEX_NO_OCR_BUNDLE
  ${NSD_CreateCheckbox} 8u 73u 100% 13u "OCR 离线资源（截图识别公式；S/M/L 模型另行下载）"
  Pop $VisualTeXOcrCheckbox
  ${If} $VisualTeXOcrChoice != "none"
    ${NSD_Check} $VisualTeXOcrCheckbox
  ${EndIf}
!endif

  ${NSD_CreateLabel} 0 96u 100% 36u "安装 Office 或 WPS 插件、输入法时可能出现管理员权限确认。未勾选的组件不会安装；已安装过的会被移除。以后可以重新运行安装包更改选择。"
  Pop $0

  nsDialogs::Show
FunctionEnd

Function VisualTeXOfficePageLeave
  ${NSD_GetState} $VisualTeXMsOfficeCheckbox $0
  ${NSD_GetState} $VisualTeXWpsCheckbox $1
  StrCpy $VisualTeXMsOfficeChoice "0"
  StrCpy $VisualTeXWpsChoice "0"
  ${If} $0 == ${BST_CHECKED}
    StrCpy $VisualTeXMsOfficeChoice "1"
  ${EndIf}
  ${If} $1 == ${BST_CHECKED}
    StrCpy $VisualTeXWpsChoice "1"
  ${EndIf}
  ${NSD_GetState} $VisualTeXImeCheckbox $2
  ${If} $2 == ${BST_CHECKED}
    StrCpy $VisualTeXImeChoice "1"
  ${Else}
    StrCpy $VisualTeXImeChoice "0"
  ${EndIf}
!ifndef VISUALTEX_NO_OCR_BUNDLE
  ${NSD_GetState} $VisualTeXOcrCheckbox $2
  ${If} $2 == ${BST_CHECKED}
    StrCpy $VisualTeXOcrChoice "install"
  ${Else}
    StrCpy $VisualTeXOcrChoice "none"
  ${EndIf}
!endif

  ${If} $VisualTeXMsOfficeChoice == "1"
  ${OrIf} $VisualTeXWpsChoice == "1"
    ; First interactive guard. This catches both visible Office applications and
    ; hidden COM/OLE servers such as POWERPNT.EXE -Embedding before payload work.
    ; Choosing No must leave the user on this page and must never terminate Office.
    Call VisualTeXEnsureOfficeProcessesClosed
    Pop $3
    ${If} $3 != "0"
      Abort
    ${EndIf}
    StrCpy $VisualTeXOfficeChoice "native"
  ${Else}
    StrCpy $VisualTeXOfficeChoice "none"
  ${EndIf}
FunctionEnd

; OCR uses the bundled private Python 3.12.10 x64 runtime and a fixed local
; wheelhouse. The installer must never probe or depend on system Python.
;
; Tauri normally writes every configured resource unconditionally. The custom
; NSIS template routes every resource directory/file through these macros so
; the user's OCR choice is applied before extraction. OCR resources remain
; embedded in the installer so the default checked mode stays fully offline,
; but choosing "none" never writes ocr*, wheel or private-Python payloads into
; the installation directory.
!macro VisualTeXCreateBundledResourceDirectory DESTINATION
  StrCpy $VisualTeXOcrResourcePrefix "${DESTINATION}" 3
  ${If} $VisualTeXOcrChoice == "none"
  ${AndIf} $VisualTeXOcrResourcePrefix == "ocr"
    DetailPrint "Skipping optional OCR resource directory: ${DESTINATION}"
  ${Else}
    CreateDirectory "$INSTDIR\\${DESTINATION}"
  ${EndIf}
!macroend

!macro VisualTeXInstallBundledResource DESTINATION SOURCE
  StrCpy $VisualTeXOcrResourcePrefix "${DESTINATION}" 3
  ${If} $VisualTeXOcrChoice == "none"
  ${AndIf} $VisualTeXOcrResourcePrefix == "ocr"
    DetailPrint "Skipping optional OCR resource: ${DESTINATION}"
  ${Else}
    File /a "/oname=${DESTINATION}" "${SOURCE}"
  ${EndIf}
!macroend

; Keep service ownership and legacy-registration maintenance in a bundled
; script: short -File invocations avoid command-line truncation and ensure an
; old or partially removed installation cannot supply stale maintenance logic.
; Current runtime maintenance never enumerates MathType.exe processes.
!macro VisualTeXRunPrivateRuntimeGuard MODE
  InitPluginsDir
  File "/oname=$PLUGINSDIR\visualtex-runtime-maintenance.ps1" "${VISUALTEX_RUNTIME_GUARD_SOURCE}"
  ; Stop may require one UAC elevation. Allow time for the user to approve it;
  ; the script itself has a bounded 15-second SCM stop wait after elevation.
  nsExec::ExecToStack /TIMEOUT=120000 `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$PLUGINSDIR\visualtex-runtime-maintenance.ps1" -InstallRoot "$INSTDIR" -Mode ${MODE}`
  Pop $0
  Pop $1
  DetailPrint "VisualTeX native runtime ${MODE}: ExitCode=$0 Output=$1"
!macroend

!macro VisualTeXEnsureNativeRendererStoppedBody PREFIX
visualtex_native_renderer_stop_retry:
  ; The renderer is an SCM-owned VisualTeX service with a fixed account and
  ; executable path. Stop only that exact identity. Its Job Object owns every
  ; Session 0 MathType descendant, so no MathType.exe process enumeration or
  ; name-based termination is needed or permitted here.
  !insertmacro VisualTeXRunPrivateRuntimeGuard Stop
  ${If} $0 == "0"
    Return
  ${EndIf}
  SetDetailsView show
  IfSilent visualtex_native_renderer_stop_abort 0
  MessageBox MB_ICONEXCLAMATION|MB_RETRYCANCEL "无法停止 VisualTeX 私有公式渲染服务，尚未继续覆盖文件。$\r$\n返回值：$0$\r$\n$1$\r$\n$\r$\n该操作只管理 VisualTeX.NativeRender 服务，不会枚举、关闭或修改独立安装的 MathType。$\r$\n$\r$\nUnable to stop the VisualTeX native-render service. Retry or cancel without replacing files. Standalone MathType is never enumerated or terminated." IDRETRY visualtex_native_renderer_stop_retry IDCANCEL visualtex_native_renderer_stop_abort
visualtex_native_renderer_stop_abort:
  SetErrorLevel 1
  Abort
!macroend

Function VisualTeXEnsureNativeRendererStopped
  !insertmacro VisualTeXEnsureNativeRendererStoppedBody ""
FunctionEnd

Function un.VisualTeXEnsureNativeRendererStopped
  !insertmacro VisualTeXEnsureNativeRendererStoppedBody "un."
FunctionEnd

; An AI agent's MCP client (Codex, Claude, ...) keeps visualtex-mcp.exe running as its VisualTeX server, and Windows
; will not replace or remove a running executable. Stop it before the files are written or removed, as the app itself
; is stopped; the client starts it again when it next needs VisualTeX.
!macro VisualTeXStopAgentLauncher
  nsExec::ExecToStack `"$SYSDIR\taskkill.exe" /F /IM visualtex-mcp.exe`
  Pop $0
  Pop $1
  DetailPrint "visualtex-mcp: taskkill ExitCode=$0"
  ${If} $0 == "0"
    Sleep 500
  ${EndIf}
!macroend

!macro NSIS_HOOK_PREINSTALL
  ; Normalize the two legacy 1.2.3 install locations back to Tauri's canonical
  ; current-user directory. Preserve genuinely custom directories.
  ${If} $INSTDIR == "$PROFILE\AppData\VisualTeX"
    StrCpy $INSTDIR "$LOCALAPPDATA\VisualTeX"
  ${ElseIf} $INSTDIR == "$APPDATA\VisualTeX"
    StrCpy $INSTDIR "$LOCALAPPDATA\VisualTeX"
  ${EndIf}

  !insertmacro VisualTeXStopAgentLauncher

  ; Builds before 1.2.7 could leave MathType's canonical per-user COM class
  ; pointing at VisualTeX's reduced renderer after a native sidecar crash. That
  ; HKCU value shadows a later machine-wide MathType installation. Remove only
  ; registrations whose executable is inside a recognized VisualTeX runtime.
  !insertmacro VisualTeXRunPrivateRuntimeGuard RepairRegistration
  ${If} $0 != "0"
    SetDetailsView show
    Abort "Unable to repair a legacy VisualTeX-owned MathType registration. ExitCode=$0 Output=$1"
  ${EndIf}

  ; Stop only the explicitly owned native-render service immediately before
  ; payload replacement. Its Job Object releases the private Session 0 MathType
  ; tree; standalone MathType is outside this ownership boundary.
  Call VisualTeXEnsureNativeRendererStopped

  ; Custom pages are skipped by NSIS /S. A release acceptance install may use
  ; /VISUALTEXOFFICE=skip to leave the machine's existing Office integration
  ; untouched while testing the exact installed desktop executable. Interactive
  ; installs retain the page choice; ordinary unattended installs default to
  ; the recommended native Office mode.
  ${GetParameters} $0
  ClearErrors
  ${GetOptions} $0 "/VISUALTEXOFFICE=" $1
  ${IfNot} ${Errors}
    ${If} $1 == "native"
      StrCpy $VisualTeXOfficeChoice "native"
    ${ElseIf} $1 == "none"
      StrCpy $VisualTeXOfficeChoice "none"
    ${ElseIf} $1 == "skip"
      StrCpy $VisualTeXOfficeChoice "skip"
      StrCpy $VisualTeXAcceptanceMode "1"
    ${Else}
      Abort "Unsupported /VISUALTEXOFFICE value: $1"
    ${EndIf}
  ${EndIf}
  ${If} $VisualTeXOfficeChoice == ""
    StrCpy $VisualTeXOfficeChoice "native"
  ${EndIf}
  ; /VISUALTEXMSOFFICE=0|1, /VISUALTEXWPS=0|1, /VISUALTEXIME=0|1 (unattended installs; the components page
  ; sets them interactively). Office or WPS alone still installs the shared add-ins.
  ${If} $VisualTeXMsOfficeChoice == ""
    StrCpy $VisualTeXMsOfficeChoice "1"
    ClearErrors
    ${GetOptions} $0 "/VISUALTEXMSOFFICE=" $1
    ${IfNot} ${Errors}
      StrCpy $VisualTeXMsOfficeChoice $1
    ${EndIf}
  ${EndIf}
  ${If} $VisualTeXWpsChoice == ""
    StrCpy $VisualTeXWpsChoice "1"
    ClearErrors
    ${GetOptions} $0 "/VISUALTEXWPS=" $1
    ${IfNot} ${Errors}
      StrCpy $VisualTeXWpsChoice $1
    ${EndIf}
  ${EndIf}
  ${If} $VisualTeXImeChoice == ""
    StrCpy $VisualTeXImeChoice "1"
    ClearErrors
    ${GetOptions} $0 "/VISUALTEXIME=" $1
    ${IfNot} ${Errors}
      StrCpy $VisualTeXImeChoice $1
    ${EndIf}
  ${EndIf}
  ${If} $VisualTeXOfficeChoice == "native"
  ${AndIf} $VisualTeXMsOfficeChoice == "0"
  ${AndIf} $VisualTeXWpsChoice == "0"
    StrCpy $VisualTeXOfficeChoice "none"
  ${EndIf}

  !ifdef VISUALTEX_NO_OCR_BUNDLE
    ; The lightweight installer is compiled without worker/Python/wheel/model
    ; resources. Keep the runtime choice pinned to none even for unattended
    ; installs so logs and resource hooks reflect the actual package contents.
    StrCpy $VisualTeXOcrChoice "none"
  !else
    ClearErrors
    ${GetOptions} $0 "/VISUALTEXOCR=" $1
    ${IfNot} ${Errors}
      ${If} $1 == "install"
        StrCpy $VisualTeXOcrChoice "install"
      ${ElseIf} $1 == "none"
        StrCpy $VisualTeXOcrChoice "none"
      ${Else}
        Abort "Unsupported /VISUALTEXOCR value: $1"
      ${EndIf}
    ${EndIf}
    ${If} $VisualTeXOcrChoice == ""
      StrCpy $VisualTeXOcrChoice "install"
    ${EndIf}
  !endif

  ${If} $VisualTeXAcceptanceMode == "1"
    DetailPrint "Installed-release acceptance mode: preserving existing Office integration and skipping machine prerequisite prompts."
  ${EndIf}

  ${If} $VisualTeXOcrChoice == "install"
    DetailPrint "Installing VisualTeX OCR offline resources: private Python 3.12.10 x64 runtime, fixed wheelhouse, worker and model catalog."
  ${Else}
    DetailPrint "OCR offline resources were disabled by the user; no ocr*, wheel or private-Python resources will be written to the installation directory."
  ${EndIf}
!macroend

; "Open with" for Markdown files (.md, .markdown): VisualTeX is listed for them. The program that opens them by
; default is not touched (no default value of the extension's key, no UserChoice): the user chooses it.
!macro VisualTeXRegisterMarkdownOpenWith
  WriteRegStr HKCU "Software\Classes\VisualTeX.Markdown" "" "Markdown"
  WriteRegStr HKCU "Software\Classes\VisualTeX.Markdown\DefaultIcon" "" "$INSTDIR\${MAINBINARYNAME}.exe,0"
  WriteRegStr HKCU "Software\Classes\VisualTeX.Markdown\shell\open\command" "" '"$INSTDIR\${MAINBINARYNAME}.exe" "%1"'
  WriteRegStr HKCU "Software\Classes\.md\OpenWithProgids" "VisualTeX.Markdown" ""
  WriteRegStr HKCU "Software\Classes\.markdown\OpenWithProgids" "VisualTeX.Markdown" ""
  WriteRegStr HKCU "Software\Classes\Applications\${MAINBINARYNAME}.exe" "FriendlyAppName" "VisualTeX"
  WriteRegStr HKCU "Software\Classes\Applications\${MAINBINARYNAME}.exe\shell\open\command" "" '"$INSTDIR\${MAINBINARYNAME}.exe" "%1"'
  WriteRegStr HKCU "Software\Classes\Applications\${MAINBINARYNAME}.exe\SupportedTypes" ".md" ""
  WriteRegStr HKCU "Software\Classes\Applications\${MAINBINARYNAME}.exe\SupportedTypes" ".markdown" ""
  ; SHCNE_ASSOCCHANGED: Explorer reads the "Open with" lists again
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend

!macro VisualTeXUnregisterMarkdownOpenWith
  DeleteRegKey HKCU "Software\Classes\VisualTeX.Markdown"
  DeleteRegValue HKCU "Software\Classes\.md\OpenWithProgids" "VisualTeX.Markdown"
  DeleteRegValue HKCU "Software\Classes\.markdown\OpenWithProgids" "VisualTeX.Markdown"
  DeleteRegKey HKCU "Software\Classes\Applications\${MAINBINARYNAME}.exe"
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend

!macro NSIS_HOOK_POSTINSTALL
  !insertmacro VisualTeXRegisterMarkdownOpenWith
  DetailPrint "Applying the selected VisualTeX Office integration mode: $VisualTeXOfficeChoice"
  ${If} $VisualTeXOfficeChoice == "native"
    ; the hosts the add-ins run in (read by the add-ins at startup)
    ${If} $VisualTeXMsOfficeChoice == "0"
      WriteRegDWORD HKCU "Software\VisualTeX\OfficeIntegration" "MicrosoftOffice" 0
    ${Else}
      WriteRegDWORD HKCU "Software\VisualTeX\OfficeIntegration" "MicrosoftOffice" 1
    ${EndIf}
    ${If} $VisualTeXWpsChoice == "0"
      WriteRegDWORD HKCU "Software\VisualTeX\OfficeIntegration" "Wps" 0
    ${Else}
      WriteRegDWORD HKCU "Software\VisualTeX\OfficeIntegration" "Wps" 1
    ${EndIf}
    DetailPrint "Office add-in hosts: Microsoft Office=$VisualTeXMsOfficeChoice WPS=$VisualTeXWpsChoice"
    DetailPrint "Installing the machine-wide VisualTeX Ribbon add-ins and native Formula OLE LocalServer. A UAC prompt may appear."
    IfFileExists "$INSTDIR\${MAINBINARYNAME}.exe" 0 visualtex_main_binary_missing
    IfFileExists "$INSTDIR\scripts\ensure_windows_office_certificate.ps1" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\scripts\install_windows_vsto.ps1" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\scripts\manage_windows_native_render.ps1" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\windows-office\VisualTeX-NativeRender-x64.msi" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\windows-office\VisualTeX-NativeRender-x64.sha256.json" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\scripts\install_windows_vsto_runtime.ps1" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\scripts\test_windows_office_runtime.ps1" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\windows-office\VisualTeX-WindowsOffice-VSTO-x64.msi" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\windows-office\VisualTeX-WindowsOffice-VSTO-x64.sha256.json" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\windows-office\VisualTeX-WindowsOffice-VSTO-x86.msi" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\windows-office\VisualTeX-WindowsOffice-VSTO-x86.sha256.json" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\windows-office\vstor_redist.exe" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\windows-office\vstor_redist.sha256.json" 0 visualtex_office_missing

    DetailPrint "Checking Microsoft Visual Studio Tools for Office Runtime..."
    nsExec::ExecToStack `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\install_windows_vsto_runtime.ps1" -RuntimeInstallerPath "$INSTDIR\windows-office\vstor_redist.exe" -ManifestPath "$INSTDIR\windows-office\vstor_redist.sha256.json" -CheckOnly`
    Pop $1
    Pop $2
    StrCmp $1 "0" visualtex_vsto_runtime_ready 0
    DetailPrint "Microsoft VSTO Runtime is missing. Detection output: $2"
    IfSilent visualtex_vsto_runtime_install 0
    MessageBox MB_ICONQUESTION|MB_YESNO "此电脑尚未安装 Microsoft Visual Studio Tools for Office Runtime。VisualTeX 的 Word/PowerPoint 原生 Ribbon 插件必须依赖该微软组件。$\r$\n$\r$\n安装包已内置微软官方、数字签名有效的 VSTO Runtime。是否现在安装并继续配置 Office 集成？安装过程中会出现 Windows UAC 管理员权限确认。$\r$\n$\r$\n选择“否”仍会保留 VisualTeX 主程序，但会跳过 Office 插件安装。" IDYES visualtex_vsto_runtime_install IDNO visualtex_vsto_runtime_declined

visualtex_vsto_runtime_declined:
    DetailPrint "The user declined Microsoft VSTO Runtime installation. VisualTeX Office integration was skipped."
    Goto visualtex_office_done

visualtex_vsto_runtime_install:
    DetailPrint "Installing the bundled Microsoft VSTO Runtime. A UAC elevation prompt may appear."
    nsExec::ExecToLog `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\install_windows_vsto_runtime.ps1" -RuntimeInstallerPath "$INSTDIR\windows-office\vstor_redist.exe" -ManifestPath "$INSTDIR\windows-office\vstor_redist.sha256.json"`
    Pop $1
    StrCmp $1 "0" visualtex_vsto_runtime_ready visualtex_vsto_runtime_failed

visualtex_vsto_runtime_failed:
    DetailPrint "Microsoft VSTO Runtime installation failed or the UAC prompt was cancelled. Native Office integration was skipped."
    IfSilent 0 visualtex_vsto_runtime_failed_interactive
    ; An unattended native-Office install must never report success when the
    ; requested Office integration was not installed. Preserve the desktop
    ; payload, but return a non-zero process exit code to the caller.
    SetErrorLevel 1
    Goto visualtex_office_done
visualtex_vsto_runtime_failed_interactive:
    MessageBox MB_ICONEXCLAMATION "Microsoft VSTO Runtime 安装失败，或管理员权限确认被取消。VisualTeX 主程序已经安装，但本次将跳过 Word/PowerPoint 原生插件。$\r$\n$\r$\n请查看 %LOCALAPPDATA%\VisualTeX\office\install-logs 中最新的 vsto-runtime 日志。"
    Goto visualtex_office_done

visualtex_vsto_runtime_ready:
    DetailPrint "Microsoft VSTO Runtime is installed and verified."

    nsExec::ExecToLog `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\ensure_windows_office_certificate.ps1" -VisualTeXPath "$INSTDIR\${MAINBINARYNAME}.exe"`
    Pop $0
    StrCmp $0 "0" 0 visualtex_office_failed

    ; Recheck at the last safe moment immediately before MSI/COM registration.
    ; Office can be activated by COM/OLE after the earlier custom page (including
    ; a windowless POWERPNT.EXE -Embedding), so the early check alone has a race.
    ; Interactive installs ask again before any forced termination. Silent installs
    ; fail Office integration rather than killing user processes without consent.
    Call VisualTeXEnsureOfficeProcessesClosed
    Pop $0
    ${If} $0 != "0"
      DetailPrint "Office integration was not started because running Office processes were not closed. GuardResult=$0"
      ${If} $0 == "1"
        IfSilent 0 visualtex_office_late_guard_declined_interactive
        SetErrorLevel 1
        Goto visualtex_office_done
visualtex_office_late_guard_declined_interactive:
        MessageBox MB_ICONINFORMATION "已取消关闭正在运行的 Office 进程。VisualTeX 主程序已安装，但本次未修改 Word/PowerPoint Office 集成。关闭相关 Office 进程后可重新运行安装包。$\r$\n$\r$\nNo Office process was terminated. The main VisualTeX application is installed, but Office integration was skipped."
        Goto visualtex_office_done
      ${EndIf}
      SetErrorLevel 1
      Goto visualtex_office_done
    ${EndIf}

    nsExec::ExecToLog `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\install_windows_vsto.ps1" -PackageDirectory "$INSTDIR\windows-office" -VisualTeXPath "$INSTDIR\${MAINBINARYNAME}.exe"`
    Pop $0
    StrCmp $0 "0" visualtex_office_static_installed visualtex_office_failed

visualtex_office_static_installed:
    DetailPrint "Machine-wide Office files and registrations passed. Office bootstrap completed without leaving a resident VisualTeX process."
    WriteRegDWORD HKCU "Software\VisualTeX\OfficeIntegration" "RuntimeVerificationPending" 1
    DetailPrint "Companion and Word/PowerPoint connection verification are deferred until VisualTeX is launched from Finish or by the user."
    IfSilent visualtex_office_done 0
    MessageBox MB_ICONINFORMATION "Office 集成及独立公式渲染服务已安装。$\r$\n$\r$\n请重新打开 Word 或 PowerPoint。"
    Goto visualtex_office_done

visualtex_office_failed:
    SetDetailsView show
    DetailPrint "VisualTeX main application installed, but the machine-wide Office files, registry entries, COM classes or OLE server failed static installation verification. See the newest vsto-bootstrap and vsto-diagnostic reports under %LOCALAPPDATA%\VisualTeX\office\install-logs."
    IfSilent 0 visualtex_office_failed_interactive
    ; /S with native Office integration is used by deployment/acceptance too.
    ; Returning 0 here previously made a stale VSTO DLL look like a successful
    ; installer update whenever Word/PowerPoint was still running.
    SetErrorLevel 1
    Goto visualtex_office_done
visualtex_office_failed_interactive:
    MessageBox MB_ICONEXCLAMATION "VisualTeX 主程序已安装，但 Office 插件的文件、注册信息、COM 类或 OLE 服务未通过静态安装验证。请查看安装详情，以及 %LOCALAPPDATA%\VisualTeX\office\install-logs 中最新的 vsto-bootstrap 和 vsto-diagnostic 报告。"
    Goto visualtex_office_done
  ${ElseIf} $VisualTeXOfficeChoice == "none"
    IfFileExists "$INSTDIR\scripts\uninstall_windows_vsto.ps1" 0 visualtex_office_done
    nsExec::ExecToLog `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\uninstall_windows_vsto.ps1"`
    Pop $0
    Goto visualtex_office_done
  ${Else}
    DetailPrint "Skipping Office integration changes for installed-release acceptance. Existing Office files, registrations, certificates and companion configuration are untouched."
    Goto visualtex_office_done
  ${EndIf}

visualtex_main_binary_missing:
  DetailPrint "The VisualTeX main executable is missing after installation. Windows Security or another antivirus may have quarantined it. Office integration was skipped."
  IfSilent visualtex_office_done 0
  MessageBox MB_ICONEXCLAMATION "VisualTeX 主程序在安装后立即丢失，Windows 安全中心或其他安全软件可能已将 visualtex.exe 隔离。Office 插件安装已跳过。$\r$\n$\r$\n请打开 Windows 安全中心 → 病毒和威胁防护 → 保护历史记录，检查 Behavior:Win32/Persistence.A!ml 等记录。"
  Goto visualtex_office_done

visualtex_office_missing:
  DetailPrint "Windows native Office installation resources are missing. The VisualTeX main application was installed without Office integration."
  IfSilent visualtex_office_done 0
  MessageBox MB_ICONEXCLAMATION "Windows 原生 Office 安装资源缺失。VisualTeX 主程序已正常安装。"
  Goto visualtex_office_done

visualtex_office_done:
  ; PREINSTALL stopped the renderer service (MathType equations in Word, PowerPoint and WPS)
  ; before replacing files. Reinstalling the renderer starts it again; every other way out of
  ; the Office step (skipped, declined, failed, acceptance) must not leave it stopped.
  !insertmacro VisualTeXRunPrivateRuntimeGuard Start
  ; VisualTeX's input method (VisualTeX Anywhere: pinyin, English and formulas in any
  ; application) is a machine-wide text service: registering it may show one UAC prompt.
  ; The script leaves a current registration alone. Acceptance installs skip it.
  ${If} $VisualTeXAcceptanceMode != "1"
  ${AndIf} $VisualTeXImeChoice == "0"
    ; left out on the components page: an earlier installation's input method goes (UAC only if registered)
    IfFileExists "$INSTDIR\scripts\install_windows_input_method.ps1" 0 visualtex_input_method_done
    DetailPrint "The VisualTeX input method was not selected; removing an earlier registration if there is one."
    System::Call 'kernel32::SetEnvironmentVariable(t "VISUALTEX_REINSTALL", p 0)'
    nsExec::ExecToLog `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\install_windows_input_method.ps1" -Uninstall`
    Pop $0
    Goto visualtex_input_method_done
  ${EndIf}
  ${If} $VisualTeXAcceptanceMode != "1"
    IfFileExists "$INSTDIR\input-method\visualtex_tip.dll" 0 visualtex_input_method_done
    IfFileExists "$INSTDIR\scripts\install_windows_input_method.ps1" 0 visualtex_input_method_done
    DetailPrint "Installing the VisualTeX input method (pinyin, English and formulas in any application). A UAC prompt may appear."
    nsExec::ExecToLog `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\install_windows_input_method.ps1" -SourceDirectory "$INSTDIR\input-method"`
    Pop $0
    ${If} $0 != "0"
      DetailPrint "The VisualTeX input method was not installed (ExitCode=$0). VisualTeX works without it; typing formulas in other applications needs it."
    ${EndIf}
  ${EndIf}
visualtex_input_method_done:
  ${If} $VisualTeXAcceptanceMode == "1"
    DetailPrint "Installed-release acceptance mode: legacy install roots are untouched."
    Goto visualtex_postinstall_cleanup_done
  ${EndIf}

  ; Remove only recognized legacy installation roots after the canonical
  ; installation has completed. User data lives under the application bundle
  ; directories, not these installer roots.
  ${If} $INSTDIR != "$PROFILE\AppData\VisualTeX"
    IfFileExists "$PROFILE\AppData\VisualTeX\uninstall.exe" visualtex_remove_direct_appdata 0
    IfFileExists "$PROFILE\AppData\VisualTeX\visualtex.exe" visualtex_remove_direct_appdata visualtex_direct_appdata_done
visualtex_remove_direct_appdata:
    RMDir /r "$PROFILE\AppData\VisualTeX"
visualtex_direct_appdata_done:
  ${EndIf}
  ${If} $INSTDIR != "$APPDATA\VisualTeX"
    IfFileExists "$APPDATA\VisualTeX\uninstall.exe" visualtex_remove_roaming_legacy 0
    IfFileExists "$APPDATA\VisualTeX\visualtex.exe" visualtex_remove_roaming_legacy visualtex_roaming_legacy_done
visualtex_remove_roaming_legacy:
    ; Remove only known legacy application payloads. Preserve
    ; %APPDATA%\VisualTeX\ocr-storage.json, OfficeSessions, logs, and any
    ; unknown user data so a later VisualTeX installation can reuse the
    ; independently stored OCR environment without reinstalling it.
    Delete "$APPDATA\VisualTeX\visualtex.exe"
    Delete "$APPDATA\VisualTeX\VisualTeX.exe"
    Delete "$APPDATA\VisualTeX\uninstall.exe"
    Delete "$APPDATA\VisualTeX\visualtex-windows-office-bridge.exe"
    RMDir /r "$APPDATA\VisualTeX\ocr"
    RMDir /r "$APPDATA\VisualTeX\ocr-models"
    RMDir /r "$APPDATA\VisualTeX\ocr-python"
    RMDir /r "$APPDATA\VisualTeX\office"
    RMDir /r "$APPDATA\VisualTeX\scripts"
    RMDir /r "$APPDATA\VisualTeX\windows-office"
    RMDir "$APPDATA\VisualTeX"
visualtex_roaming_legacy_done:
  ${EndIf}
visualtex_postinstall_cleanup_done:
!macroend

!macro NSIS_HOOK_PREUNINSTALL
  !insertmacro VisualTeXStopAgentLauncher

  ; Stop only the explicitly owned renderer service before removing its MSI.
  ; Service shutdown closes the owned Job Object and all private descendants.
  Call un.VisualTeXEnsureNativeRendererStopped

  ; Restore the public MathType COM namespace before removing the private files,
  ; including installations upgraded from the legacy temporary registration.
  !insertmacro VisualTeXRunPrivateRuntimeGuard RepairRegistration
  ${If} $0 != "0"
    SetDetailsView show
    Abort "Unable to repair a legacy VisualTeX-owned MathType registration. ExitCode=$0 Output=$1"
  ${EndIf}
  ; VisualTeX's input method: off the user's keyboards, unregistered (one UAC prompt);
  ; left in place when a new installer runs this uninstaller (VISUALTEX_REINSTALL).
  IfFileExists "$INSTDIR\scripts\install_windows_input_method.ps1" 0 visualtex_input_method_uninstall_done
  nsExec::ExecToLog `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\install_windows_input_method.ps1" -Uninstall`
  Pop $0
visualtex_input_method_uninstall_done:
  IfFileExists "$INSTDIR\scripts\uninstall_windows_vsto.ps1" 0 visualtex_preuninstall_done
  nsExec::ExecToStack `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\uninstall_windows_vsto.ps1"`
  Pop $0
  Pop $1
  ${If} $0 != "0"
    SetDetailsView show
    DetailPrint "VisualTeX Office integration uninstall failed. ExitCode=$0 Output=$1"
    MessageBox MB_ICONSTOP "无法卸载 VisualTeX Office 集成或 HTTPS 证书。主程序尚未删除，您可以关闭 Word 和 PowerPoint 后重试。$\r$\n$\r$\n请查看 %LOCALAPPDATA%\VisualTeX\office\install-logs 中最新的 vsto-uninstall-bootstrap 和 certificate-remove 日志。"
    SetErrorLevel 1
    Quit
  ${EndIf}
visualtex_preuninstall_done:
!macroend

!macro NSIS_HOOK_POSTUNINSTALL
  ; The generated maintenance flow runs uninstall.exe directly from $INSTDIR
  ; with _?=$INSTDIR. The running uninstaller therefore cannot delete itself.
  ; Remove everything that is currently deletable, then launch a detached
  ; cleanup process that waits for this uninstaller PID to exit before deleting
  ; the final uninstall.exe and empty installation root.
  ;
  ; Preserve %APPDATA%\VisualTeX because it stores OfficeSessions user data,
  ; and preserve %APPDATA%\com.visualtex.studio.
  DeleteRegKey HKCU "Software\visualtex\VisualTeX"
  !insertmacro VisualTeXUnregisterMarkdownOpenWith
  RMDir /r "$PROFILE\AppData\VisualTeX"
  RMDir /r "$INSTDIR"

  System::Call 'kernel32::GetCurrentProcessId() i .r0'
  Exec `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -Command "Wait-Process -Id $0 -ErrorAction SilentlyContinue; Start-Sleep -Milliseconds 300; Remove-Item -LiteralPath '$INSTDIR' -Recurse -Force -ErrorAction SilentlyContinue"`
!macroend
