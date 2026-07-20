[CmdletBinding()]
param(
    [string]$Configuration = "Release",
    [switch]$SkipTests
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$solution = Join-Path $root "src-windows\VisualTeX.WindowsOffice.sln"
$tests = Join-Path $root "src-windows\VisualTeX.WindowsOffice.Tests\VisualTeX.WindowsOffice.Tests.csproj"
$wordProject = Join-Path $root "src-windows\VisualTeX.WordVsto\VisualTeX.WordVsto.csproj"
$powerPointProject = Join-Path $root "src-windows\VisualTeX.PowerPointVsto\VisualTeX.PowerPointVsto.csproj"
$installerProject = Join-Path $root "src-windows\VisualTeX.WindowsOffice.Installer\VisualTeX.WindowsOffice.Installer.wixproj"
$resourceRoot = Join-Path $root "src-tauri\resources\windows-office"
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
if (-not $dotnet) { throw ".NET 8 SDK is required to build and test Windows Office integration." }

if (-not $SkipTests) {
    & $dotnet test $tests --configuration $Configuration
    if ($LASTEXITCODE -ne 0) { throw "Windows Office tests failed." }
}

& (Join-Path $PSScriptRoot "build_windows_ole_bridge.ps1") -Configuration $Configuration

$msbuild = Get-Command msbuild.exe -ErrorAction SilentlyContinue | Select-Object -First 1 -ExpandProperty Source
if (-not $msbuild) {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswhere) {
        $installationPath = & $vswhere -latest -products * -requires Microsoft.Component.MSBuild -property installationPath
        if ($installationPath) {
            $candidate = Join-Path $installationPath "MSBuild\Current\Bin\amd64\MSBuild.exe"
            if (Test-Path $candidate) { $msbuild = $candidate }
        }
    }
}
if (-not $msbuild) {
    throw "MSBuild from Visual Studio Build Tools is required for Word/PowerPoint native add-ins and WiX MSI."
}
$sdkRoot = Join-Path (Split-Path -Parent $dotnet) "sdk"
$sdk = Get-ChildItem $sdkRoot -Directory | Sort-Object { [version]$_.Name } -Descending | Select-Object -First 1
if (-not $sdk) { throw ".NET SDK directory was not found below $sdkRoot." }
$env:DOTNET_ROOT = Split-Path -Parent $dotnet
$env:MSBuildSDKsPath = Join-Path $sdk.FullName "Sdks"
$env:MSBuildEnableWorkloadResolver = "false"
$referenceRoot = Join-Path $env:USERPROFILE ".nuget\packages\microsoft.netframework.referenceassemblies.net48\1.0.3\build"
if (-not (Test-Path $referenceRoot)) { throw ".NET Framework 4.8 reference assemblies package is missing." }

foreach ($project in @($wordProject, $powerPointProject)) {
    & $dotnet restore $project --ignore-failed-sources
    if ($LASTEXITCODE -ne 0) { throw "NuGet restore failed: $project" }
    & $msbuild $project /m /p:Configuration=$Configuration /p:Platform=x64 /p:TargetFrameworkRootPath=$referenceRoot
    if ($LASTEXITCODE -ne 0) { throw "x64 VSTO build failed: $project" }
}
& $dotnet msbuild $installerProject /p:Configuration=$Configuration /p:Platform=x64 /p:BuildProjectReferences=false /p:SuppressValidation=true
if ($LASTEXITCODE -ne 0) { throw "Windows Office MSI build failed." }

$msi = Get-ChildItem (Join-Path $root "src-windows\VisualTeX.WindowsOffice.Installer\bin") `
    -Filter "*.msi" -Recurse -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1
if (-not $msi) { throw "VisualTeX VSTO MSI was not produced." }
New-Item $resourceRoot -ItemType Directory -Force | Out-Null
Copy-Item $msi.FullName (Join-Path $resourceRoot "VisualTeX-WindowsOffice-VSTO.msi") -Force
$wordOutput = Join-Path $root "src-windows\VisualTeX.WordVsto\bin\x64\$Configuration\net48\VisualTeX.WordVsto.dll"
$powerPointOutput = Join-Path $root "src-windows\VisualTeX.PowerPointVsto\bin\x64\$Configuration\net48\VisualTeX.PowerPointVsto.dll"
if (-not (Test-Path $wordOutput) -or -not (Test-Path $powerPointOutput)) {
    throw "The x64 VSTO build outputs required for SHA-256 verification are missing."
}
$hashManifest = [ordered]@{
    architecture = "x64"
    word = [ordered]@{
        file = "VisualTeX.WordVsto.dll"
        sha256 = (Get-FileHash $wordOutput -Algorithm SHA256).Hash
    }
    powerPoint = [ordered]@{
        file = "VisualTeX.PowerPointVsto.dll"
        sha256 = (Get-FileHash $powerPointOutput -Algorithm SHA256).Hash
    }
}
$hashManifest | ConvertTo-Json -Depth 4 | Set-Content `
    (Join-Path $resourceRoot "VisualTeX-WindowsOffice-VSTO.sha256.json") `
    -Encoding UTF8
Write-Host "Windows Office packages are ready for the Tauri/NSIS bundle."
