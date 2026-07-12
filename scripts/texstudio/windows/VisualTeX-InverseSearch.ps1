param(
    [Parameter(Mandatory = $true)][string]$ProjectRoot,
    [Parameter(Mandatory = $true)][string]$PdfPath,
    [Parameter(Mandatory = $true)][int]$Page,
    [Parameter(Mandatory = $true)][double]$X,
    [Parameter(Mandatory = $true)][double]$Y,
    [Parameter(Mandatory = $true)][string]$OutputJson
)
. (Join-Path $PSScriptRoot "VisualTeX-Common.ps1")
Ensure-VisualTeXBridge -ProjectRoot $ProjectRoot
Invoke-VisualTeXJsonToFile -Arguments @(
    "bridge-inverse-search",
    $ProjectRoot,
    $PdfPath,
    $Page,
    $X,
    $Y
) -OutputJson $OutputJson
