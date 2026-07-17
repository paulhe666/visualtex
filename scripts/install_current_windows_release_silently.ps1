[CmdletBinding()]
param([string]$Version = "1.2.0")

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$installer = Join-Path $root "src-tauri\target\release\bundle\nsis\VisualTeX_${Version}_x64-setup.exe"
if (-not (Test-Path -LiteralPath $installer)) {
    throw "Installer does not exist: $installer"
}
if (Get-Process WINWORD, POWERPNT -ErrorAction SilentlyContinue) {
    throw "Word or PowerPoint is running; refusing to install."
}
if (Get-Process VisualTeX -ErrorAction SilentlyContinue) {
    throw "VisualTeX is running; refusing to install."
}

$process = Start-Process -FilePath $installer -ArgumentList "/S" -PassThru -Wait
Write-Host ("Installer exit code: {0}" -f $process.ExitCode)
if ($process.ExitCode -ne 0) {
    throw "VisualTeX installer failed with exit code $($process.ExitCode)."
}
