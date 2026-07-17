[CmdletBinding()]
param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",
    [ValidateRange(30, 900)]
    [int]$TimeoutSeconds = 240,
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
$artifactRoot = Join-Path $LogDirectory "native-office-ole-$timestamp"
New-Item -ItemType Directory -Path $artifactRoot -Force | Out-Null
$logPath = Join-Path $artifactRoot "acceptance-wrapper.log"
$stdoutPath = Join-Path $artifactRoot "acceptance.stdout.log"
$stderrPath = Join-Path $artifactRoot "acceptance.stderr.log"

$clsid = "{8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B}"
$iid = "{6C672AF0-7321-4D21-B325-868CB34592C2}"
$libid = "{DF66EC66-3B3A-4675-A7BE-30456A04EB96}"
$serverProject = Join-Path $root "src-windows\VisualTeX.FormulaOleServer\VisualTeX.FormulaOleServer.vcxproj"
$acceptanceProject = Join-Path $root "src-windows\VisualTeX.NativeOfficeOleAcceptance\VisualTeX.NativeOfficeOleAcceptance.csproj"
$markerPath = Join-Path $env:LOCALAPPDATA "VisualTeX\office\ole-server-trace.enabled"
$globalTracePath = Join-Path $env:LOCALAPPDATA "VisualTeX\office\ole-server-trace.log"
$serverPath = $null
$acceptanceProcess = $null

function Resolve-DotNet {
    foreach ($candidate in @(
        (Join-Path $env:LOCALAPPDATA "Microsoft\dotnet\dotnet.exe"),
        (Join-Path $env:USERPROFILE ".dotnet\dotnet.exe"),
        (Join-Path $env:ProgramFiles "dotnet\dotnet.exe")
    )) {
        if (Test-Path $candidate) { return $candidate }
    }
    return Get-Command dotnet.exe -ErrorAction SilentlyContinue |
        Select-Object -First 1 -ExpandProperty Source
}

function Resolve-MSBuild {
    $command = Get-Command msbuild.exe -ErrorAction SilentlyContinue |
        Select-Object -First 1 -ExpandProperty Source
    if ($command) { return $command }

    $vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (-not (Test-Path $vswhere)) {
        throw "Visual Studio Installer vswhere.exe was not found."
    }
    $installationPath = & $vswhere -latest -version "[17.0,18.0)" -products * -requires Microsoft.Component.MSBuild -property installationPath
    if (-not $installationPath) {
        throw "Visual Studio 2022 with MSBuild was not found."
    }
    foreach ($relative in @(
        "MSBuild\Current\Bin\amd64\MSBuild.exe",
        "MSBuild\Current\Bin\MSBuild.exe"
    )) {
        $candidate = Join-Path $installationPath $relative
        if (Test-Path $candidate) { return $candidate }
    }
    throw "MSBuild.exe was not found below $installationPath."
}

function Resolve-OfficePlatform {
    foreach ($registryPath in @(
        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\ClickToRun\Configuration"
    )) {
        try {
            $value = (Get-ItemProperty -LiteralPath $registryPath -Name Platform -ErrorAction Stop).Platform
            if ($value -in @("x64", "x86")) { return $value }
        }
        catch { }
    }
    throw "Microsoft Office Click-to-Run platform could not be detected."
}

function Assert-NoOfficeProcesses {
    $running = Get-Process WINWORD, POWERPNT -ErrorAction SilentlyContinue
    if ($running) {
        $details = ($running | ForEach-Object { "$($_.ProcessName) PID=$($_.Id)" }) -join ", "
        throw "Close Word and PowerPoint before native OLE acceptance: $details"
    }
}

function Test-RegistryEntry {
    param(
        [Parameter(Mandatory = $true)][string]$Key,
        [Parameter(Mandatory = $true)][ValidateSet("32", "64")][string]$RegistryView
    )
    $query = Start-Process -FilePath "reg.exe" -ArgumentList @("query", $Key, "/reg:$RegistryView") -Wait -PassThru -WindowStyle Hidden
    return $query.ExitCode -eq 0
}

function Assert-NoVisualTeXRegistration {
    foreach ($view in @("32", "64")) {
        foreach ($key in @(
            "HKCU\Software\Classes\CLSID\$clsid",
            "HKCU\Software\Classes\Interface\$iid",
            "HKCU\Software\Classes\TypeLib\$libid"
        )) {
            if (Test-RegistryEntry -Key $key -RegistryView $view) {
                throw "Native OLE acceptance left a registration behind: $key (registry view $view)"
            }
        }
    }
}

function Stop-TestProcesses {
    Get-Process "VisualTeX.NativeOfficeOleAcceptance", "VisualTeX.FormulaOleServer" -ErrorAction SilentlyContinue |
        Stop-Process -Force -ErrorAction SilentlyContinue

    Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -in @("WINWORD.EXE", "POWERPNT.EXE") -and
            ($_.CommandLine -match "(?i)/(Automation|Embedding)\b")
        } |
        ForEach-Object {
            Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
        }
}

Start-Transcript -Path $logPath -Force | Out-Null
try {
    Assert-NoOfficeProcesses
    Assert-NoVisualTeXRegistration

    $officePlatform = Resolve-OfficePlatform
    $olePlatform = if ($officePlatform -eq "x64") { "x64" } else { "Win32" }

    $dotnet = Resolve-DotNet
    if (-not $dotnet) { throw ".NET 8 SDK is required." }
    $dotnetVersion = & $dotnet --version
    if ($LASTEXITCODE -ne 0 -or -not $dotnetVersion.StartsWith("8.")) {
        throw ".NET 8 SDK is required; resolved dotnet reports '$dotnetVersion'."
    }
    $msbuild = Resolve-MSBuild

    $sdkRoot = Join-Path (Split-Path -Parent $dotnet) "sdk"
    $sdk = Get-ChildItem $sdkRoot -Directory |
        Sort-Object { [version]$_.Name } -Descending |
        Select-Object -First 1
    if (-not $sdk) { throw ".NET SDK directory was not found below $sdkRoot." }
    $env:DOTNET_ROOT = Split-Path -Parent $dotnet
    $env:DOTNET_HOST_PATH = $dotnet
    $env:DOTNET_MSBUILD_SDK_RESOLVER_CLI_DIR = Split-Path -Parent $dotnet
    $env:MSBuildSDKsPath = Join-Path $sdk.FullName "Sdks"
    $env:MSBuildEnableWorkloadResolver = "false"

    $referenceRoot = Join-Path $env:USERPROFILE ".nuget\packages\microsoft.netframework.referenceassemblies.net48\1.0.3\build"
    if (-not (Test-Path $referenceRoot)) {
        throw ".NET Framework 4.8 reference assemblies package is missing."
    }

    Write-Host "Office platform: $officePlatform"
    Write-Host "MSBuild: $msbuild"
    Write-Host "dotnet: $dotnet ($dotnetVersion)"

    & $msbuild $serverProject /m /t:Rebuild "/p:Configuration=$Configuration" "/p:Platform=$olePlatform" /v:minimal
    if ($LASTEXITCODE -ne 0) {
        throw "Native Formula OLE $olePlatform build failed with exit code $LASTEXITCODE."
    }

    & $dotnet restore $acceptanceProject --ignore-failed-sources
    if ($LASTEXITCODE -ne 0) {
        throw "Native Office OLE acceptance restore failed with exit code $LASTEXITCODE."
    }
    & $msbuild $acceptanceProject `
        /m `
        /t:Rebuild `
        "/p:Configuration=$Configuration" `
        "/p:Platform=$officePlatform" `
        "/p:TargetFrameworkRootPath=$referenceRoot" `
        /v:minimal
    if ($LASTEXITCODE -ne 0) {
        throw "Native Office OLE acceptance build failed with exit code $LASTEXITCODE."
    }

    $serverPath = Join-Path $root "src-windows\artifacts\formula-ole-server\$olePlatform\$Configuration\VisualTeX.FormulaOleServer.exe"
    $acceptancePath = Join-Path $root "src-windows\VisualTeX.NativeOfficeOleAcceptance\bin\$officePlatform\$Configuration\net48\VisualTeX.NativeOfficeOleAcceptance.exe"
    if (-not (Test-Path $serverPath)) { throw "Native Formula OLE server is missing: $serverPath" }
    if (-not (Test-Path $acceptancePath)) { throw "Native Office OLE acceptance executable is missing: $acceptancePath" }

    New-Item -ItemType Directory -Path (Split-Path -Parent $markerPath) -Force | Out-Null
    Set-Content -LiteralPath $markerPath -Value "enabled" -Encoding Ascii
    Remove-Item -LiteralPath $globalTracePath -Force -ErrorAction SilentlyContinue

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $acceptancePath
    $startInfo.Arguments = "`"$serverPath`" `"$artifactRoot`""
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $acceptanceProcess = New-Object System.Diagnostics.Process
    $acceptanceProcess.StartInfo = $startInfo
    if (-not $acceptanceProcess.Start()) {
        throw "Native Office OLE acceptance process could not be started."
    }
    $stdoutTask = $acceptanceProcess.StandardOutput.ReadToEndAsync()
    $stderrTask = $acceptanceProcess.StandardError.ReadToEndAsync()

    if (-not $acceptanceProcess.WaitForExit($TimeoutSeconds * 1000)) {
        Stop-Process -Id $acceptanceProcess.Id -Force -ErrorAction SilentlyContinue
        throw "Native Office OLE acceptance exceeded $TimeoutSeconds seconds."
    }
    $acceptanceProcess.WaitForExit()
    $stdout = $stdoutTask.GetAwaiter().GetResult()
    $stderr = $stderrTask.GetAwaiter().GetResult()
    Set-Content -LiteralPath $stdoutPath -Value $stdout -Encoding UTF8
    Set-Content -LiteralPath $stderrPath -Value $stderr -Encoding UTF8
    if ($stdout) { Write-Host $stdout.TrimEnd() }
    if ($stderr) { Write-Host $stderr.TrimEnd() }
    $exitCode = $acceptanceProcess.ExitCode
    if ($exitCode -ne 0) {
        throw "Native Office OLE acceptance failed with exit code $exitCode."
    }

    Assert-NoVisualTeXRegistration
    Write-Host "VisualTeX real Word/PowerPoint native OLE acceptance passed for $officePlatform Office."
    Write-Host "Artifacts: $artifactRoot"
}
finally {
    try {
        if ($acceptanceProcess) {
            try {
                if (-not $acceptanceProcess.HasExited) {
                    Stop-Process -Id $acceptanceProcess.Id -Force -ErrorAction SilentlyContinue
                }
            }
            catch { }
        }
        Stop-TestProcesses
        if ($serverPath -and (Test-Path $serverPath)) {
            try { & $serverPath /UnregServerPerUser | Out-Null } catch { }
        }
        Remove-Item -LiteralPath $markerPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $globalTracePath -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 300
        Stop-TestProcesses
        Assert-NoVisualTeXRegistration
    }
    finally {
        Stop-Transcript -ErrorAction SilentlyContinue | Out-Null
    }
}
