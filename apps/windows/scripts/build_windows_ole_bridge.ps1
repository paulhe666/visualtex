[CmdletBinding()]
param(
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$project = Join-Path $root "src-windows\VisualTeX.WindowsOleBridge\VisualTeX.WindowsOleBridge.csproj"
$publish = Join-Path $root "src-windows\artifacts\windows-ole-bridge"
$destinationDirectory = Join-Path $root "src-tauri\binaries"
$destination = Join-Path $destinationDirectory "visualtex-windows-office-bridge-x86_64-pc-windows-msvc.exe"

$dotnet = Get-Command dotnet.exe -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty Source
if (-not $dotnet) {
    foreach ($candidate in @(
        (Join-Path $env:LOCALAPPDATA "Microsoft\dotnet\dotnet.exe"),
        (Join-Path $env:USERPROFILE ".dotnet\dotnet.exe"),
        (Join-Path $env:ProgramFiles "dotnet\dotnet.exe")
    )) {
        if (Test-Path $candidate) { $dotnet = $candidate; break }
    }
}
if (-not $dotnet) {
    throw ".NET 8 SDK is required to build the VisualTeX Windows OLE Bridge."
}

Remove-Item $publish -Recurse -Force -ErrorAction SilentlyContinue
New-Item $publish -ItemType Directory -Force | Out-Null
New-Item $destinationDirectory -ItemType Directory -Force | Out-Null

& $dotnet publish $project `
    --configuration $Configuration `
    --runtime win-x64 `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:PublishTrimmed=false `
    -p:DebugType=None `
    --output $publish
if ($LASTEXITCODE -ne 0) { throw "Windows OLE Bridge publish failed." }

$source = Join-Path $publish "visualtex-windows-office-bridge.exe"
if (-not (Test-Path $source)) {
    throw "Published Windows OLE Bridge executable is missing: $source"
}
Copy-Item $source $destination -Force

$info = Get-Item $destination
if ($info.Length -lt 1MB) {
    throw "Windows OLE Bridge sidecar is unexpectedly small: $($info.Length) bytes"
}
Write-Host "Windows OLE Bridge sidecar: $destination"
