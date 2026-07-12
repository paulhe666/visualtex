param(
    [Parameter(Mandatory = $true)][string]$ProjectRoot,
    [Parameter(Mandatory = $true)][string]$SourceFile,
    [Parameter(Mandatory = $true)][int]$Line,
    [Parameter(Mandatory = $true)][int]$Column,
    [Parameter(Mandatory = $true)][string]$PdfPath,
    [Parameter(Mandatory = $true)][string]$OutputJson
)
. (Join-Path $PSScriptRoot "VisualTeX-Common.ps1")
Ensure-VisualTeXBridge -ProjectRoot $ProjectRoot
Invoke-VisualTeXJsonToFile -Arguments @(
    "bridge-forward-search",
    $ProjectRoot,
    $SourceFile,
    $Line,
    $Column,
    $PdfPath
) -OutputJson $OutputJson
