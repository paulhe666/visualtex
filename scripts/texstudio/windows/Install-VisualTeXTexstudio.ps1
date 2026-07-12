param(
    [string]$Prefix = $(if ($env:VISUALTEX_TEXSTUDIO_HOME) { $env:VISUALTEX_TEXSTUDIO_HOME } else { Join-Path $env:LOCALAPPDATA "VisualTeX\texstudio" }),
    [switch]$Force
)
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$sourceDirectory = $PSScriptRoot
$destinationDirectory = Join-Path $Prefix "bin"
New-Item -ItemType Directory -Force -Path $destinationDirectory | Out-Null
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
    $source = Join-Path $sourceDirectory $name
    $destination = Join-Path $destinationDirectory $name
    if ((Test-Path $destination) -and -not $Force) {
        $sourceHash = (Get-FileHash -Algorithm SHA256 $source).Hash
        $destinationHash = (Get-FileHash -Algorithm SHA256 $destination).Hash
        if ($sourceHash -eq $destinationHash) { continue }
        throw "Refusing to overwrite existing adapter $destination. Use -Force only for an older VisualTeX adapter."
    }
    $temporary = $destination + ".tmp." + $PID
    Copy-Item -Force $source $temporary
    Move-Item -Force $temporary $destination
}

Set-Content -Encoding UTF8 -Path (Join-Path $Prefix ".visualtex-texstudio-adapter") -Value "VisualTeX TeXstudio adapter v1"
Copy-Item -Force (Join-Path (Split-Path $sourceDirectory -Parent) "README.md") (Join-Path $Prefix "README.md")
Write-Host "VisualTeX TeXstudio adapters installed in $destinationDirectory"
Write-Host "TeXstudio settings were not modified. See $(Join-Path $Prefix 'README.md')."
