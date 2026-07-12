param([Parameter(Mandatory = $true)][string]$ProjectRoot)
. (Join-Path $PSScriptRoot "VisualTeX-Common.ps1")
Invoke-VisualTeX -Arguments @("bridge-request", $ProjectRoot, "initialize", "--params", "{}", "--result-only")
Invoke-VisualTeX -Arguments @("bridge-status", $ProjectRoot)
