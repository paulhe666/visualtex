param(
    [string]$Prefix = $(if ($env:VISUALTEX_TEXSTUDIO_HOME) { $env:VISUALTEX_TEXSTUDIO_HOME } else { Join-Path $env:LOCALAPPDATA "VisualTeX\texstudio" })
)
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$marker = Join-Path $Prefix ".visualtex-texstudio-adapter"
if (-not (Test-Path $marker)) {
    throw "Refusing to remove unmarked directory $Prefix"
}
$files = @(
    "VisualTeX-Common.ps1",
    "VisualTeX-BridgeStart.ps1",
    "VisualTeX-BridgeStop.ps1",
    "VisualTeX-BridgeStatus.ps1",
    "VisualTeX-Open.ps1",
    "VisualTeX-Compile.ps1",
    "VisualTeX-ForwardSearch.ps1",
    "VisualTeX-InverseSearch.ps1"
)
foreach ($name in $files) {
    Remove-Item -Force -ErrorAction SilentlyContinue (Join-Path (Join-Path $Prefix "bin") $name)
}
Remove-Item -Force -ErrorAction SilentlyContinue (Join-Path $Prefix "README.md")
Remove-Item -Force $marker
Remove-Item -Force -ErrorAction SilentlyContinue (Join-Path $Prefix "bin")
Remove-Item -Force -ErrorAction SilentlyContinue $Prefix
Write-Host "VisualTeX TeXstudio adapters removed from $Prefix"
