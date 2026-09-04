[CmdletBinding()]
param(
    [string]$InstallerPath,
    [string]$ExpectedAppVersion = "1.2.6",
    [switch]$VerifyOfficeRuntime,
    [switch]$VerifyOfficeReadonly,
    [switch]$ForceCloseOffice,
    [switch]$KeepInstall
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$windowsRoot = if (-not [string]::IsNullOrWhiteSpace($env:SystemRoot)) { $env:SystemRoot } else { $env:WINDIR }
if ([string]::IsNullOrWhiteSpace($windowsRoot)) {
    throw "SystemRoot/WINDIR is unavailable; cannot locate Windows PowerShell"
}
$powerShellExe = Join-Path $windowsRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
if (-not (Test-Path -LiteralPath $powerShellExe -PathType Leaf)) {
    throw "Windows PowerShell is missing: $powerShellExe"
}

$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
if ([string]::IsNullOrWhiteSpace($InstallerPath)) {
    $InstallerPath = Join-Path $root "src-tauri\target\release\bundle\nsis\VisualTeX_${ExpectedAppVersion}_x64-setup.exe"
}
if (-not (Test-Path -LiteralPath $InstallerPath -PathType Leaf)) {
    throw "Final NSIS installer is missing: $InstallerPath"
}
$InstallerPath = (Resolve-Path -LiteralPath $InstallerPath).Path

$runKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$runValueName = "VisualTeXOffice"
$registryKeys = @(
    "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\VisualTeX",
    "HKCU\Software\visualtex\VisualTeX",
    "HKCU\Software\VisualTeX\OfficeIntegration"
)
$workRoot = Join-Path $env:TEMP ("visualtex-installed-release-" + [guid]::NewGuid().ToString("N"))
$installRoot = Join-Path $workRoot "app"
$registryBackupRoot = Join-Path $workRoot "registry"
$lifecycleLog = Join-Path $env:LOCALAPPDATA "VisualTeX\logs\app-lifecycle.log"
$reportPath = Join-Path $root "build-logs\windows-installed-release.json"
$startedProcesses = New-Object System.Collections.Generic.List[System.Diagnostics.Process]
$runValueExisted = $false
$runValue = $null
$officeIntegrationStatusBackup = Get-ItemProperty `
    -LiteralPath "HKCU:\Software\VisualTeX\OfficeIntegration" `
    -ErrorAction SilentlyContinue
$officeIntegrationStatusNames = @(
    "Mode",
    "NativeOleEnabled",
    "FilesAndRegistryVerified",
    "OfficeRuntimeVerified",
    "WordConnected",
    "PowerPointConnected",
    "OfficeConnectionVerificationAttempted",
    "OfficePlatform",
    "RuntimeVerificationPending",
    "LastRuntimeReport",
    "LastDiagnosticReport",
    "LastRuntimeError"
)
$logOffset = 0L

Add-Type @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class VisualTeXWindowProbe {
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowTextLength(IntPtr hWnd);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr hWnd, uint message, IntPtr wParam, IntPtr lParam);
    const uint WM_CLOSE = 0x0010;

    public static IntPtr FindMainWindow(uint targetProcessId) {
        IntPtr found = IntPtr.Zero;
        EnumWindows((window, _) => {
            uint processId;
            GetWindowThreadProcessId(window, out processId);
            if (processId != targetProcessId || !IsWindowVisible(window)) return true;
            int length = GetWindowTextLength(window);
            if (length <= 0) return true;
            var text = new StringBuilder(length + 1);
            GetWindowText(window, text, text.Capacity);
            if (String.Equals(text.ToString(), "VisualTeX", StringComparison.Ordinal)) {
                found = window;
                return false;
            }
            return true;
        }, IntPtr.Zero);
        return found;
    }

    public static bool CloseWindow(IntPtr window) {
        return window != IntPtr.Zero && PostMessage(window, WM_CLOSE, IntPtr.Zero, IntPtr.Zero);
    }
}
'@

function Stop-AllVisualTeXProcesses {
    $processes = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -ieq "visualtex.exe" -or $_.Name -ieq "VisualTeX.exe"
    }
    foreach ($process in $processes) {
        Stop-Process -Id ([int]$process.ProcessId) -Force -ErrorAction SilentlyContinue
    }
    if ($processes) { Start-Sleep -Milliseconds 700 }
}

function Invoke-Reg([string]$Arguments, [switch]$AllowFailure) {
    $regExe = Join-Path $env:SystemRoot "System32\reg.exe"
    $process = Start-Process `
        -FilePath $regExe `
        -ArgumentList $Arguments `
        -WindowStyle Hidden `
        -Wait `
        -PassThru
    try {
        if ($process.ExitCode -ne 0 -and -not $AllowFailure) {
            throw "reg.exe $Arguments failed with exit code $($process.ExitCode)"
        }
        return $process.ExitCode
    } finally {
        $process.Dispose()
    }
}

function Export-RegistryKey([string]$Key, [string]$Destination) {
    $queryExit = Invoke-Reg ('query "{0}"' -f $Key) -AllowFailure
    if ($queryExit -ne 0) { return $false }
    [void](Invoke-Reg ('export "{0}" "{1}" /y' -f $Key, $Destination))
    return $true
}

function Delete-RegistryKey([string]$Key) {
    [void](Invoke-Reg ('delete "{0}" /f' -f $Key) -AllowFailure)
}

function Restore-RegistryKey([string]$Key, [string]$Backup, [bool]$Existed) {
    Delete-RegistryKey $Key
    if ($Existed) {
        [void](Invoke-Reg ('import "{0}"' -f $Backup))
    }
}

function Get-NewLifecycleLog {
    if (-not (Test-Path -LiteralPath $lifecycleLog -PathType Leaf)) { return "" }
    $stream = [IO.File]::Open($lifecycleLog, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try {
        if ($logOffset -gt $stream.Length) { $script:logOffset = 0 }
        [void]$stream.Seek($logOffset, [IO.SeekOrigin]::Begin)
        $reader = New-Object IO.StreamReader($stream, [Text.Encoding]::UTF8, $true, 4096, $true)
        try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally { $stream.Dispose() }
}

function Wait-Until([scriptblock]$Condition, [string]$Description, [int]$TimeoutSeconds = 20) {
    $deadline = [DateTimeOffset]::UtcNow.AddSeconds($TimeoutSeconds)
    do {
        $value = & $Condition
        if ($value) { return $value }
        Start-Sleep -Milliseconds 150
    } while ([DateTimeOffset]::UtcNow -lt $deadline)
    throw "Timed out waiting for $Description"
}

function Start-ExactVisualTeX([string[]]$Arguments) {
    $parameters = @{
        FilePath = $script:installedExe
        PassThru = $true
    }
    if ($Arguments.Count -gt 0) { $parameters.ArgumentList = $Arguments }
    $process = Start-Process @parameters
    [void]$startedProcesses.Add($process)
    return $process
}

function Find-MainWindow([int]$ProcessId) {
    return [VisualTeXWindowProbe]::FindMainWindow([uint32]$ProcessId)
}

New-Item -ItemType Directory -Path $registryBackupRoot -Force | Out-Null
New-Item -ItemType Directory -Path (Split-Path -Parent $reportPath) -Force | Out-Null
$registryBackups = @()
foreach ($key in $registryKeys) {
    $safeName = ($key -replace '[^A-Za-z0-9]+', '-') + ".reg"
    $backup = Join-Path $registryBackupRoot $safeName
    $registryBackups += [pscustomobject]@{
        Key = $key
        Backup = $backup
        Existed = Export-RegistryKey $key $backup
    }
}
if (Test-Path -LiteralPath $runKey) {
    $property = Get-ItemProperty -LiteralPath $runKey -Name $runValueName -ErrorAction SilentlyContinue
    if ($null -ne $property) {
        $runValueExisted = $true
        $runValue = [string]$property.$runValueName
    }
}
if (Test-Path -LiteralPath $lifecycleLog -PathType Leaf) {
    $logOffset = (Get-Item -LiteralPath $lifecycleLog).Length
}

$report = [ordered]@{
    schemaVersion = 1
    installer = $InstallerPath
    installRoot = $installRoot
    bootstrapExited = $false
    bootstrapResidualProcess = $false
    backgroundRequestedMainAssets = $false
    backgroundCreatedWindow = $false
    finishCreatedMainWindow = $false
    embeddedMainAssetsVerified = $false
    closeRetainedCompanion = $false
    secondLaunchRecreatedMainWindow = $false
    desktopCloseExited = $false
    ocrWarmupScheduled = $false
    ocrResourcesOmitted = $false
    officeEditorCreatedBeforeRequest = $false
    officeRuntimeVerified = $false
    officePayloadHashesVerified = $false
    lifecycleLogTail = ""
}

try {
    Stop-AllVisualTeXProcesses
    foreach ($entry in $registryBackups) { Delete-RegistryKey $entry.Key }
    Remove-Item -LiteralPath $installRoot -Recurse -Force -ErrorAction SilentlyContinue
    New-Item -ItemType Directory -Path $installRoot -Force | Out-Null

    $arguments = @(
        "/S",
        "/NS",
        "/VISUALTEXACCEPTANCE",
        "/VISUALTEXOFFICE=skip",
        "/VISUALTEXOCR=none",
        "/D=$installRoot"
    )
    $installer = Start-Process -FilePath $InstallerPath -ArgumentList $arguments -PassThru
    if (-not $installer.WaitForExit(180000)) {
        Stop-Process -Id $installer.Id -Force -ErrorAction SilentlyContinue
        throw "Final NSIS installer did not exit within 180 seconds"
    }
    if ($installer.ExitCode -ne 0) {
        throw "Final NSIS installer failed with exit code $($installer.ExitCode)"
    }
    $installer.Dispose()

    $installedCandidates = @(
        (Join-Path $installRoot "visualtex.exe"),
        (Join-Path $installRoot "VisualTeX.exe")
    )
    $script:installedExe = $installedCandidates | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1
    if ([string]::IsNullOrWhiteSpace($script:installedExe)) {
        throw "The exact installed VisualTeX executable is missing below $installRoot"
    }
    $script:installedExe = (Resolve-Path -LiteralPath $installedExe).Path

    foreach ($optionalOcrPath in @(
        (Join-Path $installRoot "ocr-python"),
        (Join-Path $installRoot "ocr-models"),
        (Join-Path $installRoot "ocr")
    )) {
        if (Test-Path -LiteralPath $optionalOcrPath) {
            throw "Installer /VISUALTEXOCR=none unexpectedly extracted optional OCR resources: $optionalOcrPath"
        }
    }
    $report.ocrResourcesOmitted = $true

    # Acceptance mode intentionally skips machine-wide Office registration, but it
    # still extracts the exact Office payload embedded in this NSIS. Verify that
    # payload against the resources used for this build so a stale same-version
    # installer can never pass merely because the desktop EXE starts correctly.
    $sourceOfficeRoot = Join-Path $root "src-tauri\resources\windows-office"
    $installedOfficeRoot = Join-Path $installRoot "windows-office"
    foreach ($architecture in @("x64", "x86")) {
        $manifestName = "VisualTeX-WindowsOffice-VSTO-$architecture.sha256.json"
        $msiName = "VisualTeX-WindowsOffice-VSTO-$architecture.msi"
        $sourceManifestPath = Join-Path $sourceOfficeRoot $manifestName
        $installedManifestPath = Join-Path $installedOfficeRoot $manifestName
        $sourceMsiPath = Join-Path $sourceOfficeRoot $msiName
        $installedMsiPath = Join-Path $installedOfficeRoot $msiName
        foreach ($requiredPath in @($sourceManifestPath, $installedManifestPath, $sourceMsiPath, $installedMsiPath)) {
            if (-not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
                throw "Installed-release Office payload verification is missing: $requiredPath"
            }
        }
        $sourceManifest = Get-Content -LiteralPath $sourceManifestPath -Raw | ConvertFrom-Json
        $installedManifest = Get-Content -LiteralPath $installedManifestPath -Raw | ConvertFrom-Json
        if ($installedManifest.package.sha256 -ne $sourceManifest.package.sha256 -or
            $installedManifest.word.sha256 -ne $sourceManifest.word.sha256 -or
            $installedManifest.powerPoint.sha256 -ne $sourceManifest.powerPoint.sha256 -or
            $installedManifest.formulaOleServer.sha256 -ne $sourceManifest.formulaOleServer.sha256) {
            throw "Installed-release $architecture Office manifest does not match the current build resources."
        }
        $sourceMsiHash = (Get-FileHash -LiteralPath $sourceMsiPath -Algorithm SHA256).Hash
        $installedMsiHash = (Get-FileHash -LiteralPath $installedMsiPath -Algorithm SHA256).Hash
        if ($sourceMsiHash -ne $installedMsiHash -or $installedMsiHash -ne $sourceManifest.package.sha256) {
            throw "Installed-release $architecture Office MSI payload is stale: installed=$installedMsiHash source=$sourceMsiHash manifest=$($sourceManifest.package.sha256)"
        }
    }
    $report.officePayloadHashesVerified = $true

    $bootstrap = Start-ExactVisualTeX @("--office-bootstrap")
    if (-not $bootstrap.WaitForExit(30000)) {
        Stop-Process -Id $bootstrap.Id -Force -ErrorAction SilentlyContinue
        throw "Installed --office-bootstrap did not exit within 30 seconds"
    }
    if ($bootstrap.ExitCode -ne 0) { throw "Installed --office-bootstrap failed with exit code $($bootstrap.ExitCode)" }
    Start-Sleep -Milliseconds 500
    $bootstrapResidual = Get-CimInstance Win32_Process -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -ieq "visualtex.exe" -and
        [string]::Equals([string]$_.ExecutablePath, $script:installedExe, [StringComparison]::OrdinalIgnoreCase)
    }
    $report.bootstrapExited = $true
    $report.bootstrapResidualProcess = [bool]$bootstrapResidual
    if ($bootstrapResidual) { throw "--office-bootstrap left a residual installed process" }

    # Acceptance mode deliberately asks the installer not to reinstall or
    # uninstall the machine-wide Office packages. Bootstrap rewrites only the
    # exact companion executable/configuration values for the temporary EXE, so
    # restore the pre-existing installation verification markers without
    # overwriting ExecutablePath, AppDataRoot, certificate, port or protocol.
    if ($null -ne $officeIntegrationStatusBackup) {
        $integrationPath = "HKCU:\Software\VisualTeX\OfficeIntegration"
        New-Item -Path $integrationPath -Force | Out-Null
        foreach ($name in $officeIntegrationStatusNames) {
            $property = $officeIntegrationStatusBackup.PSObject.Properties[$name]
            if ($null -eq $property -or $null -eq $property.Value) { continue }
            $propertyType = if ($property.Value -is [int] -or $property.Value -is [long]) {
                "DWord"
            } else {
                "String"
            }
            New-ItemProperty `
                -LiteralPath $integrationPath `
                -Name $name `
                -PropertyType $propertyType `
                -Value $property.Value `
                -Force | Out-Null
        }
    }

    New-Item -Path $runKey -Force | Out-Null
    New-ItemProperty -LiteralPath $runKey -Name $runValueName -PropertyType String -Value ('"' + $script:installedExe + '" --office-background') -Force | Out-Null

    $background = Start-ExactVisualTeX @("--office-background")
    Wait-Until { -not $background.HasExited } "installed Office background process" 10 | Out-Null
    Start-Sleep -Seconds 2
    $background.Refresh()
    if ($background.HasExited) { throw "Installed Office background process exited during startup with code $($background.ExitCode)" }
    $backgroundWindow = Find-MainWindow $background.Id
    $backgroundLog = Get-NewLifecycleLog
    $report.backgroundCreatedWindow = $backgroundWindow -ne [IntPtr]::Zero
    $report.backgroundRequestedMainAssets = $backgroundLog.Contains("main-asset diagnostic")
    if ($report.backgroundCreatedWindow) { throw "--office-background created a main window" }
    if ($report.backgroundRequestedMainAssets) { throw "--office-background requested the embedded main index.html" }

    $finishLaunch = Start-ExactVisualTeX @()
    [void]$finishLaunch.WaitForExit(15000)
    $firstMain = Wait-Until { Find-MainWindow $background.Id } "Finish/ordinary launch to create the main window" 25
    $report.finishCreatedMainWindow = $firstMain -ne [IntPtr]::Zero
    $afterFinishLog = Get-NewLifecycleLog
    $report.embeddedMainAssetsVerified =
        $afterFinishLog.Contains("Tauri embedded main assets verified") -and
        $afterFinishLog.Contains("new main window created successfully") -and
        -not $afterFinishLog.Contains("main window creation blocked by embedded asset failure")
    if (-not $report.embeddedMainAssetsVerified) {
        throw "The exact installed EXE did not record a successful embedded index.html/JS/CSS preflight"
    }

    if (-not [VisualTeXWindowProbe]::CloseWindow($firstMain)) {
        throw "Unable to send WM_CLOSE to the installed main window"
    }
    Wait-Until { (Find-MainWindow $background.Id) -eq [IntPtr]::Zero } "main window destruction while retaining companion" 15 | Out-Null
    $background.Refresh()
    $report.closeRetainedCompanion = -not $background.HasExited
    if (-not $report.closeRetainedCompanion) { throw "Closing main window terminated the configured background companion" }

    $secondLaunch = Start-ExactVisualTeX @()
    [void]$secondLaunch.WaitForExit(15000)
    $secondMain = Wait-Until { Find-MainWindow $background.Id } "second launch to recreate the main window" 25
    $report.secondLaunchRecreatedMainWindow = $secondMain -ne [IntPtr]::Zero

    Remove-ItemProperty -LiteralPath $runKey -Name $runValueName -ErrorAction SilentlyContinue
    if (-not [VisualTeXWindowProbe]::CloseWindow($secondMain)) {
        throw "Unable to close the recreated installed main window"
    }
    Wait-Until { $background.Refresh(); $background.HasExited } "complete desktop shutdown with background disabled" 30 | Out-Null
    $report.desktopCloseExited = $true

    $finalLog = Get-NewLifecycleLog
    $report.ocrWarmupScheduled = $finalLog.Contains("OCR startup warmup scheduled")
    $report.officeEditorCreatedBeforeRequest = $finalLog.Contains("creating Office editor WebView on demand")
    if (-not $report.ocrWarmupScheduled) { throw "OCR startup warmup was not scheduled" }
    if ($report.officeEditorCreatedBeforeRequest) { throw "Office editor WebView was created before an actual Office edit request" }
    if ($finalLog.Contains("asset not found: index.html")) { throw "Installed lifecycle log contains asset not found: index.html" }

    if ($VerifyOfficeRuntime) {
        $runtimeScript = Join-Path $installRoot "scripts\test_windows_office_runtime.ps1"
        if (-not (Test-Path -LiteralPath $runtimeScript -PathType Leaf)) {
            throw "Installed Office runtime verifier is missing: $runtimeScript"
        }
        $runtimeArguments = @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", $runtimeScript,
            "-VisualTeXPath", $script:installedExe
        )
        if ($ForceCloseOffice) { $runtimeArguments += "-ForceCloseOffice" }
        & $powerShellExe @runtimeArguments
        if ($LASTEXITCODE -ne 0) { throw "Installed Office runtime verification failed with exit code $LASTEXITCODE" }
        $report.officeRuntimeVerified = $true
    } elseif ($VerifyOfficeReadonly) {
        $runtimeScript = Join-Path $PSScriptRoot "test_windows_office_connection_readonly.ps1"
        if (-not (Test-Path -LiteralPath $runtimeScript -PathType Leaf)) {
            throw "Non-destructive Office connection verifier is missing: $runtimeScript"
        }
        & $powerShellExe -NoProfile -ExecutionPolicy Bypass -File $runtimeScript -VisualTeXPath $script:installedExe
        if ($LASTEXITCODE -ne 0) { throw "Installed Office connection verification failed with exit code $LASTEXITCODE" }
        $report.officeRuntimeVerified = $true
    }

    $report.lifecycleLogTail = Get-NewLifecycleLog
    $report | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $reportPath -Encoding UTF8
    Write-Host "Installed release acceptance passed."
    Write-Host "Exact installed EXE: $script:installedExe"
    Write-Host "Report: $reportPath"
} finally {
    Stop-AllVisualTeXProcesses
    foreach ($process in $startedProcesses) {
        try { $process.Dispose() } catch { }
    }
    if ($runValueExisted) {
        New-Item -Path $runKey -Force | Out-Null
        New-ItemProperty -LiteralPath $runKey -Name $runValueName -PropertyType String -Value $runValue -Force | Out-Null
    } else {
        Remove-ItemProperty -LiteralPath $runKey -Name $runValueName -ErrorAction SilentlyContinue
    }
    foreach ($entry in $registryBackups) {
        Restore-RegistryKey $entry.Key $entry.Backup $entry.Existed
    }
    if (-not $KeepInstall) {
        Remove-Item -LiteralPath $workRoot -Recurse -Force -ErrorAction SilentlyContinue
    } else {
        Write-Host "Kept installed acceptance directory: $installRoot"
    }
}
