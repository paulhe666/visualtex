[CmdletBinding()]
param(
    [switch]$EnableBackgroundStart,
    [string]$VisualTeXPath,
    [string]$PackageDirectory,
    [ValidateSet("auto", "x86", "x64")]
    [string]$OfficePlatform = "auto"
)

$ErrorActionPreference = "Stop"
Write-Warning "install_windows_ole.ps1 is retired. VisualTeX no longer installs Office.js Trusted Catalog manifests; forwarding to the native Ribbon + OLE LocalServer installer."

$installer = Join-Path $PSScriptRoot "install_windows_vsto.ps1"
if (-not (Test-Path -LiteralPath $installer -PathType Leaf)) {
    throw "The native Office installer is missing: $installer"
}

$arguments = @{
    OfficePlatform = $OfficePlatform
}
if (-not [string]::IsNullOrWhiteSpace($VisualTeXPath)) {
    $arguments.VisualTeXPath = $VisualTeXPath
}
if (-not [string]::IsNullOrWhiteSpace($PackageDirectory)) {
    $arguments.PackageDirectory = $PackageDirectory
}
& $installer @arguments
