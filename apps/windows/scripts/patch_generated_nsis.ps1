[CmdletBinding()]
param(
    [string]$InstallerNsi,
    [string]$OutputInstaller,
    [string]$AppVersion = "1.2.6"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
if ([string]::IsNullOrWhiteSpace($InstallerNsi)) {
    $InstallerNsi = Join-Path $root "src-tauri\target\release\nsis\x64\installer.nsi"
}
if ([string]::IsNullOrWhiteSpace($OutputInstaller)) {
    $OutputInstaller = Join-Path $root "src-tauri\target\release\bundle\nsis\VisualTeX_${AppVersion}_x64-setup.exe"
}
if (-not (Test-Path -LiteralPath $InstallerNsi -PathType Leaf)) {
    throw "Generated Tauri NSIS source is missing: $InstallerNsi"
}

function ConvertTo-Lf([string]$Value) {
    return $Value.Replace("`r`n", "`n").Replace("`r", "`n")
}

$source = ConvertTo-Lf (Get-Content -LiteralPath $InstallerNsi -Raw)
$functionPattern = '(?m)^Function PageReinstall\n'
$acceptanceStart = @"
Function PageReinstall
  ; Installed-release acceptance uses a clean custom directory and must not
  ; enter maintenance mode for another installed VisualTeX copy.
  `$`{GetParameters} `$R7
  ClearErrors
  `$`{GetOptions} `$R7 "/VISUALTEXACCEPTANCE" `$R8
  `$`{IfNot} `$`{Errors}
    Abort
  `$`{EndIf}
"@
$acceptanceStart = ConvertTo-Lf $acceptanceStart
if (-not [regex]::IsMatch($source, $functionPattern)) {
    throw "Generated PageReinstall function was not found."
}
$source = [regex]::Replace($source, $functionPattern, $acceptanceStart, 1)

$oldSelection = ConvertTo-Lf @'
    ; Check the first radio button if this the first time
    ; we enter this page or if the second button wasn't
    ; selected the last time we were on this page
    ${If} $ReinstallPageCheck <> 2
      SendMessage $R2 ${BM_SETCHECK} ${BST_CHECKED} 0
    ${Else}
      SendMessage $R3 ${BM_SETCHECK} ${BST_CHECKED} 0
    ${EndIf}

    ${NSD_SetFocus} $R2
'@
$newSelection = ConvertTo-Lf @'
    ; Same-version maintenance defaults to the second option, "Uninstall
    ; VisualTeX". Preserve an explicit user selection when navigating back.
    ; Upgrade/downgrade pages keep Tauri's original first-option default.
    ${If} $R0 = 0
      ${If} $ReinstallPageCheck = 1
        SendMessage $R2 ${BM_SETCHECK} ${BST_CHECKED} 0
        ${NSD_SetFocus} $R2
      ${Else}
        SendMessage $R3 ${BM_SETCHECK} ${BST_CHECKED} 0
        StrCpy $ReinstallPageCheck 2
        ${NSD_SetFocus} $R3
      ${EndIf}
    ${Else}
      ${If} $ReinstallPageCheck <> 2
        SendMessage $R2 ${BM_SETCHECK} ${BST_CHECKED} 0
        ${NSD_SetFocus} $R2
      ${Else}
        SendMessage $R3 ${BM_SETCHECK} ${BST_CHECKED} 0
        ${NSD_SetFocus} $R3
      ${EndIf}
    ${EndIf}
'@
if (-not $source.Contains($oldSelection)) {
    throw "Generated PageReinstall selection block did not match the expected Tauri template. Refusing to build an installer with an unverified maintenance default."
}
$source = $source.Replace($oldSelection, $newSelection)

$oldSameVersionLeave = ConvertTo-Lf @'
  ${If} $R0 = 0 ; Same version, proceed
    ${If} $R1 = 1              ; User chose to add/reinstall
      Goto reinst_done
    ${Else}                    ; User chose to uninstall
      Goto reinst_uninstall
    ${EndIf}
'@
$newSameVersionLeave = ConvertTo-Lf @'
  ${If} $R0 = 0 ; Same version: always remove the installed payload before reinstalling.
    ; Tauri's in-place same-version path can leave the previous EXE/resources
    ; untouched while still running post-install hooks from the stale install.
    ; Force a real uninstall so the File commands below install this package's
    ; exact payload rather than silently reusing the old 1.2.6 files.
    Goto reinst_uninstall
'@
if (-not $source.Contains($oldSameVersionLeave)) {
    throw "Generated PageLeaveReinstall same-version block did not match the expected Tauri template. Refusing to build an installer that can retain stale payload files."
}
$source = $source.Replace($oldSameVersionLeave, $newSameVersionLeave)

$outputDirectory = Split-Path -Parent $OutputInstaller
New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null
$outputForNsis = [IO.Path]::GetFullPath($OutputInstaller).Replace('\', '\\')
$oldOutFile = '!define OUTFILE "nsis-output.exe"'
$newOutFile = '!define OUTFILE "' + $outputForNsis + '"'
if (-not $source.Contains($oldOutFile)) {
    throw "Generated NSIS OUTFILE definition was not found."
}
$source = $source.Replace($oldOutFile, $newOutFile)

$encoding = New-Object System.Text.UTF8Encoding($false)
[IO.File]::WriteAllText($InstallerNsi, $source, $encoding)

$makensisCandidates = @(
    (Join-Path $env:LOCALAPPDATA "tauri\NSIS\makensis.exe"),
    (Join-Path $env:LOCALAPPDATA "tauri\NSIS\Bin\makensis.exe")
)
$makensis = $makensisCandidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
if ([string]::IsNullOrWhiteSpace($makensis)) {
    throw "Tauri's makensis.exe was not found below $env:LOCALAPPDATA\tauri\NSIS."
}

Remove-Item -LiteralPath $OutputInstaller -Force -ErrorAction SilentlyContinue
Write-Host "Rebuilding patched NSIS installer with $makensis"
& $makensis /V3 $InstallerNsi
if ($LASTEXITCODE -ne 0) {
    throw "makensis failed with exit code $LASTEXITCODE"
}
if (-not (Test-Path -LiteralPath $OutputInstaller -PathType Leaf)) {
    throw "Patched NSIS installer was not produced: $OutputInstaller"
}

$patched = Get-Content -LiteralPath $InstallerNsi -Raw
if ($patched -notmatch 'Same-version maintenance defaults to the second option' -or
    $patched -notmatch 'StrCpy \$ReinstallPageCheck 2' -or
    $patched -notmatch 'Same version: always remove the installed payload before reinstalling' -or
    $patched -notmatch '/VISUALTEXACCEPTANCE') {
    throw "Generated NSIS source does not contain the verified same-version forced reinstall flow."
}

$item = Get-Item -LiteralPath $OutputInstaller
$hash = (Get-FileHash -LiteralPath $OutputInstaller -Algorithm SHA256).Hash
Write-Host "Patched NSIS installer: $($item.FullName)"
Write-Host "Size: $($item.Length) bytes"
Write-Host "SHA-256: $hash"
