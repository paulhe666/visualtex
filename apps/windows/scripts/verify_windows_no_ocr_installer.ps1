[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InstallerPath,
    [string]$ExpectedAppVersion = "1.2.7"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$installer = Get-Item -LiteralPath $InstallerPath -ErrorAction Stop
if ($installer.Length -le 0) {
    throw "No-OCR installer is empty: $($installer.FullName)"
}

$baseConfigPath = Join-Path $root "src-tauri\tauri.conf.json"
$ocrOverlayPath = Join-Path $root "src-tauri\tauri.ocr-bundle.conf.json"
$noOcrConfigPath = Join-Path $root "src-tauri\tauri.no-ocr.conf.json"
$noOcrTemplatePath = Join-Path $root "src-tauri\target-no-ocr\nsis-template\visualtex-installer.nsi"
$hooksPath = Join-Path $root "src-tauri\windows\hooks.nsh"

foreach ($path in @($baseConfigPath, $ocrOverlayPath, $noOcrConfigPath, $noOcrTemplatePath, $hooksPath)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "No-OCR verification input is missing: $path"
    }
}

$baseConfig = Get-Content -LiteralPath $baseConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ([string]$baseConfig.version -ne $ExpectedAppVersion) {
    throw "Unexpected app version in tauri.conf.json: $($baseConfig.version)"
}
$baseResources = @($baseConfig.bundle.resources.PSObject.Properties)
$forbiddenBaseResources = @($baseResources | Where-Object {
    [string]$_.Name -match '(?i)(^|[\\/])ocr([\\/-]|$)' -or
    [string]$_.Value -match '(?i)(^|[\\/])ocr([\\/-]|$)'
})
if ($forbiddenBaseResources.Count -ne 0) {
    throw "Base Tauri configuration still embeds OCR resources: $($forbiddenBaseResources.Name -join ', ')"
}

$ocrOverlay = Get-Content -LiteralPath $ocrOverlayPath -Raw -Encoding UTF8 | ConvertFrom-Json
$overlayResources = @($ocrOverlay.bundle.resources.PSObject.Properties)
$expectedOcrSources = @(
    "ocr/worker.py",
    "resources/ocr-python/windows-x64",
    "resources/ocr-models/windows-x64"
)
if ($overlayResources.Count -ne $expectedOcrSources.Count) {
    throw "Full-installer OCR overlay must contain exactly the three optional OCR resource roots."
}
foreach ($source in $expectedOcrSources) {
    if ($overlayResources.Name -notcontains $source) {
        throw "Full-installer OCR overlay is missing: $source"
    }
}

$noOcrConfig = Get-Content -LiteralPath $noOcrConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ([string]$noOcrConfig.bundle.windows.nsis.template -ne "./target-no-ocr/nsis-template/visualtex-installer.nsi") {
    throw "No-OCR Tauri flavor does not point at its isolated NSIS template."
}
if ($null -ne $noOcrConfig.bundle.PSObject.Properties["resources"]) {
    throw "No-OCR Tauri flavor must not add any bundle resources of its own."
}

$template = Get-Content -LiteralPath $noOcrTemplatePath -Raw -Encoding UTF8
if (-not $template.Contains("!define VISUALTEX_NO_OCR_BUNDLE")) {
    throw "No-OCR NSIS template is missing the compile-time OCR exclusion marker."
}
if ($template.Contains("Page custom VisualTeXOcrPageCreate VisualTeXOcrPageLeave")) {
    throw "No-OCR installer must not show the offline OCR resource selection page."
}
$hooks = Get-Content -LiteralPath $hooksPath -Raw -Encoding UTF8
if (-not $hooks.Contains("!ifdef VISUALTEX_NO_OCR_BUNDLE") -or
    -not $hooks.Contains('StrCpy $VisualTeXOcrChoice "none"')) {
    throw "NSIS hooks do not pin the no-OCR flavor to an empty OCR payload."
}

$fullInstallerPath = Join-Path $root "src-tauri\target\release\bundle\nsis\VisualTeX_${ExpectedAppVersion}_x64-setup.exe"
if (Test-Path -LiteralPath $fullInstallerPath -PathType Leaf) {
    $fullInstaller = Get-Item -LiteralPath $fullInstallerPath
    if ($fullInstaller.FullName -eq $installer.FullName) {
        throw "No-OCR build reused the full installer's output path."
    }
    $savedBytes = [int64]$fullInstaller.Length - [int64]$installer.Length
    if ($savedBytes -lt 50MB) {
        throw "No-OCR installer is not materially smaller than the full installer. Full=$($fullInstaller.Length), NoOCR=$($installer.Length), Saved=$savedBytes."
    }
    Write-Host ("Full installer: {0} bytes" -f $fullInstaller.Length)
    Write-Host ("No-OCR installer: {0} bytes; reduced by {1} bytes" -f $installer.Length, $savedBytes)
}

$hash = (Get-FileHash -LiteralPath $installer.FullName -Algorithm SHA256).Hash
Write-Host ("No-OCR installer verified: {0} | {1} bytes | SHA256 {2}" -f $installer.FullName, $installer.Length, $hash)
Write-Host "No bundled OCR worker, private Python, wheelhouse or OCR model catalog is configured for this flavor."
