param([Parameter(Mandatory = $true)][string]$ProjectRoot)
. (Join-Path $PSScriptRoot "VisualTeX-Common.ps1")
Invoke-VisualTeX -Arguments @("bridge-shutdown", $ProjectRoot)
