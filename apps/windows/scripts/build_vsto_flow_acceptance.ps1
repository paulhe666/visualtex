[CmdletBinding()]
param(
    [string]$Configuration = "Release",
    [ValidateSet("x64", "x86")]
    [string]$Platform = "x64"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$project = Join-Path $root "src-windows\VisualTeX.VstoFlowAcceptance\VisualTeX.VstoFlowAcceptance.csproj"

$dotnet = @(
    (Join-Path $env:LOCALAPPDATA "Microsoft\dotnet\dotnet.exe"),
    (Join-Path $env:USERPROFILE ".dotnet\dotnet.exe"),
    (Join-Path $env:ProgramFiles "dotnet\dotnet.exe")
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $dotnet) { throw ".NET 8 SDK is required." }

$vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
$installationPath = & $vswhere -latest -version "[17.0,18.0)" -products * -requires Microsoft.Component.MSBuild -property installationPath
if (-not $installationPath) { throw "Visual Studio MSBuild was not found." }
$msbuild = Join-Path $installationPath "MSBuild\Current\Bin\amd64\MSBuild.exe"
if (-not (Test-Path $msbuild)) { throw "MSBuild was not found: $msbuild" }

$sdkRoot = Join-Path (Split-Path -Parent $dotnet) "sdk"
$sdk = Get-ChildItem $sdkRoot -Directory | Sort-Object { [version]$_.Name } -Descending | Select-Object -First 1
if (-not $sdk) { throw ".NET SDK directory was not found below $sdkRoot." }

$env:DOTNET_ROOT = Split-Path -Parent $dotnet
$env:DOTNET_HOST_PATH = $dotnet
$env:DOTNET_MSBUILD_SDK_RESOLVER_CLI_DIR = Split-Path -Parent $dotnet
$env:MSBuildSDKsPath = Join-Path $sdk.FullName "Sdks"
$env:MSBuildEnableWorkloadResolver = "false"
$referenceRoot48 = Join-Path $env:USERPROFILE ".nuget\packages\microsoft.netframework.referenceassemblies.net48\1.0.3\build"
$referenceRoot472 = Join-Path $env:USERPROFILE ".nuget\packages\microsoft.netframework.referenceassemblies.net472\1.0.3\build"
if (-not (Test-Path $referenceRoot48)) { throw ".NET Framework 4.8 reference assemblies are missing." }
if (-not (Test-Path $referenceRoot472)) { throw ".NET Framework 4.7.2 reference assemblies are missing." }

$runtimeIdentifier = "win7-$Platform"
$wordProject = Join-Path $root "src-windows\VisualTeX.WordVsto\VisualTeX.WordVsto.csproj"
$powerPointProject = Join-Path $root "src-windows\VisualTeX.PowerPointVsto\VisualTeX.PowerPointVsto.csproj"
foreach ($dependencyProject in @($wordProject, $powerPointProject)) {
    & $dotnet restore $dependencyProject --ignore-failed-sources -p:Platform=$Platform -r $runtimeIdentifier
    if ($LASTEXITCODE -ne 0) { throw "NuGet restore failed: $dependencyProject" }
    & $msbuild $dependencyProject /m /p:Configuration=$Configuration /p:Platform=$Platform /p:TargetFrameworkRootPath=$referenceRoot472
    if ($LASTEXITCODE -ne 0) { throw "VSTO dependency build failed: $dependencyProject" }
}

& $dotnet restore $project --ignore-failed-sources -p:Platform=$Platform -r $runtimeIdentifier
if ($LASTEXITCODE -ne 0) { throw "NuGet restore failed." }
& $msbuild $project /m /p:Configuration=$Configuration /p:Platform=$Platform /p:TargetFrameworkRootPath=$referenceRoot48 /p:BuildProjectReferences=false
if ($LASTEXITCODE -ne 0) { throw "VSTO flow acceptance build failed." }
