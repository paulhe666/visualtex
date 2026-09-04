; VisualTeX Windows installer prerequisite check and per-user native Office choice.
; The production path installs the per-user Ribbon COM add-ins and ATL OLE
; LocalServer. Legacy Office.js Trusted Catalog resources are not installed.

!define VISUALTEX_INSTALLER_VERSION "1.2.6"

Var VisualTeXOfficeChoice
Var VisualTeXOfficeOnlyRadio
Var VisualTeXOfficeNativeRadio
Var VisualTeXOcrChoice
Var VisualTeXOcrCheckbox
Var VisualTeXOcrResourcePrefix
Var VisualTeXAcceptanceMode

; The generated Tauri PageReinstall function is patched after bundling so the
; same-version maintenance page defaults to "Uninstall VisualTeX" directly at
; control creation time. Do not use a GUI timer here: it races the generated
; page and does not reliably change the checked radio button.

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

Function VisualTeXOfficePageCreate
  Call VisualTeXRepairMainUninstallRegistration
  nsDialogs::Create 1018
  Pop $0
  ${If} $0 == error
    Abort
  ${EndIf}

  !insertmacro MUI_HEADER_TEXT "Office 集成" "选择是否安装 VisualTeX 的 Word / PowerPoint 原生集成"
  ${NSD_CreateLabel} 0 0 100% 24u "请选择是否启用 Windows 原生 Office 集成 / Choose Windows native Office integration"
  Pop $0

  ${NSD_CreateRadioButton} 0 34u 100% 16u "仅 VisualTeX（不安装 Office 插件） / VisualTeX only"
  Pop $VisualTeXOfficeOnlyRadio

  ${NSD_CreateRadioButton} 0 58u 100% 16u "VisualTeX + 原生 Office 集成（推荐）"
  Pop $VisualTeXOfficeNativeRadio
  ${NSD_Check} $VisualTeXOfficeNativeRadio

  ${NSD_CreateLabel} 0 88u 100% 44u "原生模式统一使用 Word/PowerPoint Ribbon COM 加载项与 ATL OLE LocalServer。安装过程不会启动 Office，并会清理旧 Office.js Trusted Catalog 残留。"
  Pop $0

  nsDialogs::Show
FunctionEnd

Function VisualTeXOfficePageLeave
  ${NSD_GetState} $VisualTeXOfficeNativeRadio $0
  ${If} $0 == ${BST_CHECKED}
    nsExec::ExecToStack `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -Command "if (Get-Process WINWORD,POWERPNT,EXCEL,OUTLOOK,ONENOTE,MSACCESS,MSPUB,VISIO,MSPROJECT -ErrorAction SilentlyContinue) { exit 1 }; exit 0"`
    Pop $1
    Pop $2
    ${If} $1 != "0"
      MessageBox MB_ICONEXCLAMATION|MB_YESNO "检测到 Microsoft Office 仍在运行。强制关闭会立即结束 Word、PowerPoint、Excel、Outlook、OneNote、Access、Publisher、Visio 和 Project；未保存的 Office 文档可能丢失。$\r$\n$\r$\n是否强制关闭所有这些 Office 进程并继续安装？选择“否”将返回上一页，由您自行保存并关闭 Office。$\r$\n$\r$\nMicrosoft Office is still running. Force closing will terminate all common Office apps immediately and may discard unsaved work.$\r$\n$\r$\nForce close all Office processes and continue? Choose No to go back and close Office yourself." IDYES visualtex_force_close_office IDNO visualtex_office_close_declined

visualtex_force_close_office:
      nsExec::ExecToStack `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -Command "Get-Process WINWORD,POWERPNT,EXCEL,OUTLOOK,ONENOTE,MSACCESS,MSPUB,VISIO,MSPROJECT -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue; Start-Sleep -Milliseconds 800; if (Get-Process WINWORD,POWERPNT,EXCEL,OUTLOOK,ONENOTE,MSACCESS,MSPUB,VISIO,MSPROJECT -ErrorAction SilentlyContinue) { exit 1 }; exit 0"`
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
    StrCpy $VisualTeXOfficeChoice "native"
    Return
  ${EndIf}
  StrCpy $VisualTeXOfficeChoice "none"
FunctionEnd

Function VisualTeXOcrPageCreate
  nsDialogs::Create 1018
  Pop $0
  ${If} $0 == error
    Abort
  ${EndIf}

  !insertmacro MUI_HEADER_TEXT "OCR 离线资源" "选择是否随 VisualTeX 安装本地 OCR 运行资源"
  ${NSD_CreateLabel} 0 0 100% 28u "VisualTeX 的 OCR 为可选组件。默认安装后可在应用内配置 OCR 环境；如果暂时不需要 OCR，可以取消下面的勾选以节省磁盘空间。"
  Pop $0

  ${NSD_CreateCheckbox} 0 38u 100% 18u "安装 OCR 离线资源（推荐） / Install offline OCR resources (recommended)"
  Pop $VisualTeXOcrCheckbox
  ${If} $VisualTeXOcrChoice != "none"
    ${NSD_Check} $VisualTeXOcrCheckbox
  ${EndIf}

  ${NSD_CreateLabel} 0 66u 100% 48u "包含 VisualTeX 私有 Python 3.12.10、离线 wheel 依赖、OCR worker 与模型目录索引。S/M/L OCR 模型仍按需单独下载。"
  Pop $0

  ${NSD_CreateLabel} 0 116u 100% 20u "取消勾选后这些 OCR 文件不会写入安装目录；以后可重新运行安装包并勾选此项补装。"
  Pop $0

  nsDialogs::Show
FunctionEnd

Function VisualTeXOcrPageLeave
  ${NSD_GetState} $VisualTeXOcrCheckbox $0
  ${If} $0 == ${BST_CHECKED}
    StrCpy $VisualTeXOcrChoice "install"
  ${Else}
    StrCpy $VisualTeXOcrChoice "none"
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

!macro NSIS_HOOK_PREINSTALL
  ; Normalize the two legacy 1.2.3 install locations back to Tauri's canonical
  ; current-user directory. Preserve genuinely custom directories.
  ${If} $INSTDIR == "$PROFILE\AppData\VisualTeX"
    StrCpy $INSTDIR "$LOCALAPPDATA\VisualTeX"
  ${ElseIf} $INSTDIR == "$APPDATA\VisualTeX"
    StrCpy $INSTDIR "$LOCALAPPDATA\VisualTeX"
  ${EndIf}

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

  ${If} $VisualTeXAcceptanceMode == "1"
    DetailPrint "Installed-release acceptance mode: preserving existing Office integration and skipping machine prerequisite prompts."
  ${EndIf}

  ${If} $VisualTeXOcrChoice == "install"
    DetailPrint "Installing VisualTeX OCR offline resources: private Python 3.12.10 x64 runtime, fixed wheelhouse, worker and model catalog."
  ${Else}
    DetailPrint "OCR offline resources were disabled by the user; no ocr*, wheel or private-Python resources will be written to the installation directory."
  ${EndIf}
!macroend

!macro NSIS_HOOK_POSTINSTALL
  DetailPrint "Applying the selected VisualTeX Office integration mode: $VisualTeXOfficeChoice"
  ${If} $VisualTeXOfficeChoice == "native"
    DetailPrint "Installing the machine-wide VisualTeX Ribbon add-ins and native Formula OLE LocalServer. A UAC prompt may appear."
    IfFileExists "$INSTDIR\${MAINBINARYNAME}.exe" 0 visualtex_main_binary_missing
    IfFileExists "$INSTDIR\scripts\ensure_windows_office_certificate.ps1" 0 visualtex_office_missing
    IfFileExists "$INSTDIR\scripts\install_windows_vsto.ps1" 0 visualtex_office_missing
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

    nsExec::ExecToLog `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File "$INSTDIR\scripts\install_windows_vsto.ps1" -PackageDirectory "$INSTDIR\windows-office" -VisualTeXPath "$INSTDIR\${MAINBINARYNAME}.exe"`
    Pop $0
    StrCmp $0 "0" visualtex_office_static_installed visualtex_office_failed

visualtex_office_static_installed:
    DetailPrint "Machine-wide Office files and registrations passed. Office bootstrap completed without leaving a resident VisualTeX process."
    WriteRegDWORD HKCU "Software\VisualTeX\OfficeIntegration" "RuntimeVerificationPending" 1
    DetailPrint "Companion and Word/PowerPoint connection verification are deferred until VisualTeX is launched from Finish or by the user."
    IfSilent visualtex_office_done 0
    MessageBox MB_ICONINFORMATION "Office 集成的文件、注册信息、证书、COM 类和 OLE 服务已安装并完成静态验证。安装阶段不会启动常驻后台进程，也不会创建任何 WebView。$\r$\n$\r$\n点击“完成”启动 VisualTeX 后，本地 companion 才会按正常运行模式启动；也可以稍后在“设置 → Office 集成”中验证 Word 和 PowerPoint 连接。"
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
  RMDir /r "$PROFILE\AppData\VisualTeX"
  RMDir /r "$INSTDIR"

  System::Call 'kernel32::GetCurrentProcessId() i .r0'
  Exec `"$WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -Command "Wait-Process -Id $0 -ErrorAction SilentlyContinue; Start-Sleep -Milliseconds 300; Remove-Item -LiteralPath '$INSTDIR' -Recurse -Force -ErrorAction SilentlyContinue"`
!macroend
