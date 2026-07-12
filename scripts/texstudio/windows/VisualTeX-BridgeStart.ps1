param([Parameter(Mandatory = $true)][string]$ProjectRoot)
. (Join-Path $PSScriptRoot "VisualTeX-Common.ps1")
Ensure-VisualTeXBridge -ProjectRoot $ProjectRoot
Invoke-VisualTeX -Arguments @("bridge-status", $ProjectRoot)
