[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",
    [string]$LogDirectory
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$root = Split-Path -Parent $PSScriptRoot
if (-not $LogDirectory) {
    $LogDirectory = Join-Path $root "src-windows\artifacts\test-logs"
}
New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
$timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$logPath = Join-Path $LogDirectory "formula-ole-server-$timestamp.log"

$vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
if (-not (Test-Path $vswhere)) {
    throw "Visual Studio Installer vswhere.exe was not found."
}
$visualStudio = & $vswhere -latest -version "[17.0,18.0)" -products * -requires Microsoft.Component.MSBuild -property installationPath
if (-not $visualStudio) {
    throw "Visual Studio 2022 with MSBuild was not found."
}
$msbuild = Join-Path $visualStudio "MSBuild\Current\Bin\MSBuild.exe"
if (-not (Test-Path $msbuild)) {
    throw "MSBuild.exe was not found at $msbuild"
}

$testProject = Join-Path $root "src-windows\VisualTeX.FormulaOleServer.Tests\VisualTeX.FormulaOleServer.Tests.vcxproj"
$serverArtifacts = Join-Path $root "src-windows\artifacts\formula-ole-server"
$testArtifacts = Join-Path $root "src-windows\artifacts\formula-ole-server-tests"
$clsid = "{8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B}"
$iid = "{6C672AF0-7321-4D21-B325-868CB34592C2}"
$libid = "{DF66EC66-3B3A-4675-A7BE-30456A04EB96}"

function Assert-RegistryEntryRemoved {
    param(
        [Parameter(Mandatory = $true)][string]$Key,
        [Parameter(Mandatory = $true)][ValidateSet("32", "64")][string]$RegistryView
    )

    $query = Start-Process `
        -FilePath "reg.exe" `
        -ArgumentList @("query", $Key, "/reg:$RegistryView") `
        -Wait `
        -PassThru `
        -WindowStyle Hidden
    if ($query.ExitCode -eq 0) {
        throw "Native OLE smoke test left a registration behind: $Key (registry view $RegistryView)"
    }
}

Start-Transcript -Path $logPath -Force | Out-Null
try {
    Write-Host "MSBuild: $msbuild"
    Write-Host "Configuration: $Configuration"

    foreach ($target in @(
        @{ Platform = "x64"; RegistryView = "64" },
        @{ Platform = "Win32"; RegistryView = "32" }
    )) {
        $platform = $target.Platform
        Write-Host "=== Rebuild and test $platform ==="
        & $msbuild $testProject `
            /m `
            /t:Rebuild `
            "/p:Configuration=$Configuration" `
            "/p:Platform=$platform" `
            /v:minimal
        if ($LASTEXITCODE -ne 0) {
            throw "Native Formula OLE $platform build failed with exit code $LASTEXITCODE."
        }

        $server = Join-Path $serverArtifacts "$platform\$Configuration\VisualTeX.FormulaOleServer.exe"
        $test = Join-Path $testArtifacts "$platform\$Configuration\VisualTeX.FormulaOleServer.Tests.exe"
        if (-not (Test-Path $server)) {
            throw "Native Formula OLE LocalServer is missing: $server"
        }
        if (-not (Test-Path $test)) {
            throw "Native Formula OLE smoke test is missing: $test"
        }

        & $test $server
        if ($LASTEXITCODE -ne 0) {
            throw "Native Formula OLE $platform smoke test failed with exit code $LASTEXITCODE."
        }

        Assert-RegistryEntryRemoved -Key "HKCU\Software\Classes\CLSID\$clsid" -RegistryView $target.RegistryView
        Assert-RegistryEntryRemoved -Key "HKCU\Software\Classes\Interface\$iid" -RegistryView $target.RegistryView
        Assert-RegistryEntryRemoved -Key "HKCU\Software\Classes\TypeLib\$libid" -RegistryView $target.RegistryView
    }

    Write-Host "VisualTeX Formula OLE LocalServer x64 and Win32 acceptance passed."
    Write-Host "Log: $logPath"
}
finally {
    Stop-Transcript | Out-Null
}
