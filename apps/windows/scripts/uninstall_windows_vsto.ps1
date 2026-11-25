[CmdletBinding()]
param(
    [string]$LogPath,
    # A file that receives the reason when the uninstall fails (the VisualTeX uninstaller shows it).
    [string]$ReasonPath,
    [switch]$Elevated,
    [switch]$ArchitectureRelaunched
)

$ErrorActionPreference = "Stop"
$wpsCompanionScript = Join-Path $PSScriptRoot "windows_wps_companion.ps1"
if (Test-Path -LiteralPath $wpsCompanionScript -PathType Leaf) { . $wpsCompanionScript }
# Precise causes for failures (Windows Installer's own error, a stuck service).
$diagnosisScript = Join-Path $PSScriptRoot "windows_office_diagnosis.ps1"
if (Test-Path -LiteralPath $diagnosisScript -PathType Leaf) { . $diagnosisScript }

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Quote-ProcessArgument([string]$Value) {
    return '"' + $Value.Replace('"', '\"') + '"'
}

function Resolve-ForwardedPath([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Value }
    $trimmed = $Value.Trim().Trim('"')
    if ([IO.Path]::IsPathRooted($trimmed)) { return [IO.Path]::GetFullPath($trimmed) }
    return [IO.Path]::GetFullPath((Join-Path (Get-Location).Path $trimmed))
}

function Resolve-PowerShellExecutable([ValidateSet("x86", "x64")][string]$TargetPlatform) {
    $windowsRoot = if ([string]::IsNullOrWhiteSpace($env:WINDIR)) { "C:\Windows" } else { $env:WINDIR }
    if (-not [Environment]::Is64BitOperatingSystem) {
        return Join-Path $windowsRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
    }
    if ($TargetPlatform -eq "x86") {
        return Join-Path $windowsRoot "SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    }
    if (-not [Environment]::Is64BitProcess) {
        $sysnative = Join-Path $windowsRoot "Sysnative\WindowsPowerShell\v1.0\powershell.exe"
        if (Test-Path -LiteralPath $sysnative -PathType Leaf) { return $sysnative }
    }
    return Join-Path $windowsRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
}

# The reason this run failed, as one line: written to -ReasonPath for the installer, and printed as a
# "REASON: " line for VisualTeX. The output of a copy started elevated reaches nobody, so the process
# that started it prints the reason again (Invoke-SelfProcess).
function Write-FailureReason([string]$Reason) {
    $line = ($Reason -replace '[\r\n]+', ' ').Trim()
    if (-not [string]::IsNullOrWhiteSpace($ReasonPath)) {
        try { [IO.File]::WriteAllText($ReasonPath, $line, [Text.Encoding]::Default) } catch { }
    }
    return $line
}

function Stop-WithReason([string]$Reason, [int]$ExitCode = 1) {
    Write-Output ("REASON: " + (Write-FailureReason $Reason))
    exit $ExitCode
}

# Windows reports the declined administrator prompt as ERROR_CANCELLED (1223).
function Test-ElevationDeclined([Exception]$Exception) {
    for ($current = $Exception; $null -ne $current; $current = $current.InnerException) {
        if ($current -is [ComponentModel.Win32Exception] -and $current.NativeErrorCode -eq 1223) { return $true }
    }
    return $false
}

function Invoke-SelfProcess {
    param(
        [ValidateSet("x86", "x64")][string]$TargetPlatform,
        [bool]$RunAsAdministrator,
        [bool]$MarkArchitectureRelaunched
    )

    $reasonFile = $ReasonPath
    if ([string]::IsNullOrWhiteSpace($reasonFile)) {
        $reasonFile = Join-Path ([IO.Path]::GetTempPath()) ("visualtex-office-uninstall-" + [Guid]::NewGuid().ToString("N") + ".txt")
    }
    $arguments = New-Object System.Collections.Generic.List[string]
    foreach ($value in @(
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        (Quote-ProcessArgument $PSCommandPath)
    )) {
        [void]$arguments.Add($value)
    }
    if ($RunAsAdministrator) { [void]$arguments.Add("-Elevated") }
    if ($MarkArchitectureRelaunched) { [void]$arguments.Add("-ArchitectureRelaunched") }
    if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
        [void]$arguments.Add("-LogPath")
        [void]$arguments.Add((Quote-ProcessArgument (Resolve-ForwardedPath $LogPath)))
    }
    [void]$arguments.Add("-ReasonPath")
    [void]$arguments.Add((Quote-ProcessArgument (Resolve-ForwardedPath $reasonFile)))

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = Resolve-PowerShellExecutable $TargetPlatform
    $startInfo.Arguments = $arguments -join " "
    $startInfo.UseShellExecute = $true
    if ($RunAsAdministrator) { $startInfo.Verb = "runas" }
    try {
        $process = [System.Diagnostics.Process]::Start($startInfo)
    } catch {
        if (Test-ElevationDeclined $_.Exception) {
            Stop-WithReason "Windows did not grant administrator permission: the User Account Control prompt was declined." 4
        }
        Stop-WithReason ("Unable to start PowerShell: " + $_.Exception.Message)
    }
    try {
        # Wait only for the direct PowerShell child. Do not wait for descendant
        # processes such as Windows Installer service helpers.
        $process.WaitForExit()
        $exitCode = $process.ExitCode
    } finally {
        $process.Dispose()
    }
    if ($exitCode -ne 0) {
        $reason = ""
        try { $reason = [IO.File]::ReadAllText($reasonFile, [Text.Encoding]::Default).Trim() } catch { }
        if ([string]::IsNullOrWhiteSpace($reason)) {
            $reason = "VisualTeX Office integration uninstall failed with exit code $exitCode."
            [void](Write-FailureReason $reason)
        }
        Write-Output ("REASON: " + $reason)
    }
    if ($reasonFile -ne $ReasonPath) { Remove-Item -LiteralPath $reasonFile -Force -ErrorAction SilentlyContinue }
    exit $exitCode
}

if (-not (Test-IsAdministrator)) {
    if ($Elevated) {
        Stop-WithReason "VisualTeX Office integration uninstall requires administrator privileges."
    }
    $currentPlatform = if ([Environment]::Is64BitProcess) { "x64" } else { "x86" }
    Invoke-SelfProcess $currentPlatform $true $false
}

if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    if ($ArchitectureRelaunched) {
        Stop-WithReason "Unable to relaunch VisualTeX Office uninstall in 64-bit PowerShell."
    }
    Invoke-SelfProcess "x64" $false $true
}

$logRoot = Join-Path $env:LOCALAPPDATA "VisualTeX\office\install-logs"
New-Item -Path $logRoot -ItemType Directory -Force | Out-Null
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Join-Path $logRoot "vsto-uninstall-bootstrap-$stamp.log"
} else {
    $LogPath = Resolve-ForwardedPath $LogPath
}
$transcriptStarted = $false

function Write-UninstallLog([string]$Message) {
    $line = "[{0:O}] {1}" -f (Get-Date), $Message
    Write-Host $line
    Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
}

function Wait-DirectProcess {
    param(
        [string]$FilePath,
        [string[]]$ArgumentList
    )
    $process = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -PassThru
    try {
        $process.WaitForExit()
        return $process.ExitCode
    } finally {
        $process.Dispose()
    }
}

# The ProductCode of an MSI file; empty when it cannot be read.
function Get-MsiProductCode([string]$Path) {
    $installer = $null
    $database = $null
    $view = $null
    $record = $null
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $database = $installer.GetType().InvokeMember("OpenDatabase", [Reflection.BindingFlags]::InvokeMethod, $null, $installer, @($Path, 0))
        $view = $database.GetType().InvokeMember("OpenView", [Reflection.BindingFlags]::InvokeMethod, $null, $database, @("SELECT Value FROM Property WHERE Property = 'ProductCode'"))
        [void]$view.GetType().InvokeMember("Execute", [Reflection.BindingFlags]::InvokeMethod, $null, $view, $null)
        $record = $view.GetType().InvokeMember("Fetch", [Reflection.BindingFlags]::InvokeMethod, $null, $view, $null)
        if ($null -eq $record) { return "" }
        return [string]$record.GetType().InvokeMember("StringData", [Reflection.BindingFlags]::GetProperty, $null, $record, @(1))
    } catch {
        return ""
    } finally {
        foreach ($comObject in @($record, $view, $database, $installer)) {
            if ($null -ne $comObject -and [Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject)
            }
        }
    }
}

# Windows Installer removes a product with the copy of its package it kept at installation, and refuses
# (1612) when a cleaner has deleted that copy. The Office packages this installation carries have the
# same ProductCode, so one of them is cached again first by a small reinstall from it.
function Restore-CachedOfficePackage([string]$ProductCode) {
    $localPackage = ""
    $installer = New-Object -ComObject WindowsInstaller.Installer
    try {
        try { $localPackage = [string]$installer.ProductInfo($ProductCode, "LocalPackage") } catch { }
    } finally {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($installer)
    }
    if (-not [string]::IsNullOrWhiteSpace($localPackage) -and (Test-Path -LiteralPath $localPackage -PathType Leaf)) { return }
    $packageDirectory = Join-Path (Split-Path -Parent $PSScriptRoot) "windows-office"
    foreach ($package in @(Get-ChildItem -LiteralPath $packageDirectory -Filter "VisualTeX-WindowsOffice-VSTO-*.msi" -ErrorAction SilentlyContinue)) {
        if ((Get-MsiProductCode $package.FullName) -ine $ProductCode) { continue }
        $recacheLog = Join-Path $logRoot ("vsto-recache-{0}-{1}.log" -f $ProductCode.Trim('{}'), $stamp)
        Write-UninstallLog "Windows Installer no longer has the package of $ProductCode ('$localPackage'); caching $($package.Name) again. MSI log: $recacheLog"
        $exitCode = Wait-DirectProcess "msiexec.exe" @(
            "/i", (Quote-ProcessArgument $package.FullName),
            "REINSTALL=ALL",
            "REINSTALLMODE=vomus",
            "/qn",
            "/norestart",
            "/L*v", (Quote-ProcessArgument $recacheLog)
        )
        Write-UninstallLog "MSI recache exit code for ${ProductCode}: $exitCode"
        return
    }
    Write-UninstallLog "Windows Installer no longer has the package of $ProductCode ('$localPackage') and $packageDirectory has no package with that ProductCode."
}

function Test-ProductInstalled([string]$ProductCode) {
    $installer = New-Object -ComObject WindowsInstaller.Installer
    try {
        try {
            $null = $installer.ProductInfo($ProductCode, "VersionString")
            return $true
        } catch {
            return $false
        }
    } finally {
        if ([Runtime.InteropServices.Marshal]::IsComObject($installer)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($installer)
        }
    }
}

try {
    Start-Transcript -LiteralPath ($LogPath + ".transcript.txt") -Force | Out-Null
    $transcriptStarted = $true

    Write-UninstallLog "Starting machine-wide VisualTeX Office integration uninstall. Process64=$([Environment]::Is64BitProcess); Admin=$(Test-IsAdministrator)."

    $runningOffice = @(Get-Process WINWORD, POWERPNT -ErrorAction SilentlyContinue)
    if ($runningOffice.Count -gt 0) {
        $details = ($runningOffice | ForEach-Object { "$($_.ProcessName):$($_.Id)" }) -join ", "
        throw "Close Word and PowerPoint before uninstalling VisualTeX Office integration. Running: $details"
    }

    foreach ($visualTeXProcess in @(Get-Process visualtex -ErrorAction SilentlyContinue)) {
        $path = ""
        try { $path = [string]$visualTeXProcess.Path } catch { }
        Write-UninstallLog "Stopping VisualTeX process PID=$($visualTeXProcess.Id) Path=$path before uninstall."
        Stop-Process -Id $visualTeXProcess.Id -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Milliseconds 500
    $remainingVisualTeX = @(Get-Process visualtex -ErrorAction SilentlyContinue)
    if ($remainingVisualTeX.Count -gt 0) {
        throw "Unable to stop all VisualTeX processes before uninstall. Remaining PIDs: $($remainingVisualTeX.Id -join ', ')"
    }

    foreach ($server in @(Get-Process VisualTeX.FormulaOleServer -ErrorAction SilentlyContinue)) {
        Write-UninstallLog "Stopping Formula OLE Server PID=$($server.Id)."
        Stop-Process -Id $server.Id -Force -ErrorAction Stop
    }

    $knownProductCodes = @(
        "{8BF4D9CB-320D-4AEB-929F-7E04812795AF}",
        "{48ABC5AF-2963-4BE6-86E3-F03950ECD270}"
    )
    $productCodes = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
    foreach ($productCode in $knownProductCodes) {
        if (Test-ProductInstalled $productCode) { [void]$productCodes.Add($productCode) }
    }

    foreach ($root in @(
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKCU:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )) {
        if (-not (Test-Path -LiteralPath $root)) { continue }
        foreach ($key in Get-ChildItem -LiteralPath $root -ErrorAction SilentlyContinue) {
            $item = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction SilentlyContinue
            if ([string]$item.DisplayName -ne "VisualTeX Windows Office Integration") { continue }
            if ($key.PSChildName -match '^\{[0-9A-Fa-f-]{36}\}$') {
                [void]$productCodes.Add($key.PSChildName)
            }
        }
    }

    foreach ($productCode in @($productCodes)) {
        Restore-CachedOfficePackage $productCode
        $msiLog = Join-Path $logRoot ("vsto-uninstall-{0}-{1}.log" -f $productCode.Trim('{}'), $stamp)
        Write-UninstallLog "Uninstalling Office MSI $productCode. MSI log: $msiLog"
        $exitCode = Wait-DirectProcess "msiexec.exe" @(
            "/x", $productCode,
            "/qn",
            "/norestart",
            "/L*v", (Quote-ProcessArgument $msiLog)
        )
        Write-UninstallLog "MSI uninstall exit code for ${productCode}: $exitCode"
        if ($exitCode -notin @(0, 1605, 3010)) {
            $cause = if (Get-Command Get-MsiFailureCause -ErrorAction SilentlyContinue) { Get-MsiFailureCause $exitCode $msiLog } else { "exit code $exitCode" }
            throw "Removing the VisualTeX Office add-in package $productCode failed: $cause. MSI log: $msiLog"
        }
        if (Test-ProductInstalled $productCode) {
            throw "VisualTeX Office MSI $productCode is still registered after uninstall."
        }
    }

    & (Join-Path $PSScriptRoot 'manage_windows_native_render.ps1') -Action Uninstall

    # The add-ins registered for a WPS Office of the other bitness.
    if ((Get-Command Remove-WpsCompanionRegistration -ErrorAction SilentlyContinue) -and (Remove-WpsCompanionRegistration)) {
        Write-UninstallLog "Removed the add-ins registered for WPS Office of the other bitness."
    }

    foreach ($key in @(
        "HKLM:\Software\Microsoft\Office\Word\Addins\VisualTeX.WordVsto",
        "HKLM:\Software\Microsoft\Office\PowerPoint\Addins\VisualTeX.PowerPointVsto",
        "HKLM:\Software\Classes\CLSID\{F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1}",
        "HKLM:\Software\Classes\CLSID\{7E586D2D-57B0-4D14-AB24-EBA9021A5E6D}",
        "HKLM:\Software\Classes\VisualTeX.WordVsto",
        "HKLM:\Software\Classes\VisualTeX.PowerPointVsto",
        "HKLM:\Software\Classes\VisualTeX.Formula.1",
        "HKLM:\Software\Classes\VisualTeX.Formula",
        "HKLM:\Software\Classes\CLSID\{8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B}",
        "HKLM:\Software\Classes\Interface\{6C672AF0-7321-4D21-B325-868CB34592C2}",
        "HKLM:\Software\Classes\Interface\{A59B7798-6F24-4CF0-B378-E951BFFAFB3A}",
        "HKLM:\Software\Classes\TypeLib\{DF66EC66-3B3A-4675-A7BE-30456A04EB96}",
        "HKLM:\Software\Classes\AppID\{3C72FF7F-B04A-4FD0-AA7D-61D110D8B3C1}",
        "HKLM:\Software\Classes\AppID\VisualTeX.FormulaOleServer.exe",
        "HKCU:\Software\Microsoft\Office\Word\Addins\VisualTeX.WordVsto",
        "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\VisualTeX.PowerPointVsto",
        "HKCU:\Software\Classes\VisualTeX.WordVsto",
        "HKCU:\Software\Classes\VisualTeX.PowerPointVsto",
        "HKCU:\Software\Classes\CLSID\{F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1}",
        "HKCU:\Software\Classes\CLSID\{7E586D2D-57B0-4D14-AB24-EBA9021A5E6D}",
        "HKCU:\Software\Classes\VisualTeX.Formula.1",
        "HKCU:\Software\Classes\VisualTeX.Formula",
        "HKCU:\Software\Classes\CLSID\{8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B}",
        "HKCU:\Software\Classes\Interface\{6C672AF0-7321-4D21-B325-868CB34592C2}",
        "HKCU:\Software\Classes\Interface\{A59B7798-6F24-4CF0-B378-E951BFFAFB3A}",
        "HKCU:\Software\Classes\TypeLib\{DF66EC66-3B3A-4675-A7BE-30456A04EB96}",
        "HKCU:\Software\Classes\AppID\{3C72FF7F-B04A-4FD0-AA7D-61D110D8B3C1}",
        "HKCU:\Software\Classes\AppID\VisualTeX.FormulaOleServer.exe"
    )) {
        Remove-Item -LiteralPath $key -Recurse -Force -ErrorAction SilentlyContinue
    }

    foreach ($legacyKey in @(
        "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\VisualTeX",
        "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{69C6A866-755B-4C5A-BACB-EEA28B03C724}"
    )) {
        Remove-Item -LiteralPath $legacyKey -Recurse -Force -ErrorAction SilentlyContinue
    }
    Remove-Item -LiteralPath (Join-Path $env:LOCALAPPDATA "VisualTeX\OfficeCatalog") -Recurse -Force -ErrorAction SilentlyContinue

    # WPS Office add-in whitelist entries (the MSI removes the machine copies).
    foreach ($whitelistEntry in @(
        @("HKCU:\Software\Kingsoft\Office\WPS\AddinsWL", "VisualTeX.WordVsto"),
        @("HKCU:\Software\Kingsoft\Office\WPP\AddinsWL", "VisualTeX.PowerPointVsto")
    )) {
        Remove-ItemProperty -LiteralPath $whitelistEntry[0] -Name $whitelistEntry[1] -Force -ErrorAction SilentlyContinue
    }

    foreach ($machineKey in @(
        "HKLM:\Software\Microsoft\Office\Word\Addins\VisualTeX.WordVsto",
        "HKLM:\Software\Microsoft\Office\PowerPoint\Addins\VisualTeX.PowerPointVsto",
        "HKLM:\Software\Classes\VisualTeX.WordVsto",
        "HKLM:\Software\Classes\VisualTeX.PowerPointVsto"
    )) {
        if (Test-Path -LiteralPath $machineKey) {
            throw "Machine-wide Office registration remains after uninstall: $machineKey"
        }
    }

    $certificateScript = Join-Path $PSScriptRoot "remove_windows_office_certificate.ps1"
    if (Test-Path -LiteralPath $certificateScript -PathType Leaf) {
        $certificateLog = Join-Path $logRoot "certificate-remove-$stamp.log"
        Write-UninstallLog "Removing the VisualTeX current-user Office HTTPS certificate. Certificate log: $certificateLog"
        $certificateExitCode = Wait-DirectProcess (Resolve-PowerShellExecutable "x64") @(
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy", "Bypass",
            "-File", (Quote-ProcessArgument $certificateScript),
            "-LogPath", (Quote-ProcessArgument $certificateLog),
            "-TimeoutSeconds", "15"
        )
        Write-UninstallLog "Certificate cleanup exit code: $certificateExitCode"
        if ($certificateExitCode -ne 0) {
            $detail = ""
            if (Test-Path -LiteralPath $certificateLog -PathType Leaf) {
                $detail = [string](@(Get-Content -LiteralPath $certificateLog -ErrorAction SilentlyContinue |
                    Where-Object { $_ -match 'FAILED|error|denied|timed out|timeout' }) | Select-Object -Last 1)
            }
            throw ("Removing the VisualTeX HTTPS certificate failed (exit code $certificateExitCode)" + $(if ($detail) { ": " + $detail.Trim() } else { "." }) + " Log: $certificateLog")
        }
    }

    $modeKey = "HKCU:\Software\VisualTeX\OfficeIntegration"
    if (Test-Path -LiteralPath $modeKey) {
        New-ItemProperty $modeKey -Name "Mode" -PropertyType String -Value "auto" -Force | Out-Null
        New-ItemProperty $modeKey -Name "NativeOleEnabled" -PropertyType DWord -Value 0 -Force | Out-Null
        New-ItemProperty $modeKey -Name "FilesAndRegistryVerified" -PropertyType DWord -Value 0 -Force | Out-Null
        New-ItemProperty $modeKey -Name "OfficeRuntimeVerified" -PropertyType DWord -Value 0 -Force | Out-Null
        New-ItemProperty $modeKey -Name "WordConnected" -PropertyType DWord -Value 0 -Force | Out-Null
        New-ItemProperty $modeKey -Name "PowerPointConnected" -PropertyType DWord -Value 0 -Force | Out-Null
        foreach ($runtimeValueName in @(
            "CompanionProcessRunning",
            "CompanionPortListening",
            "CompanionHttpsHealthy",
            "CompanionCertificateMatches",
            "CompanionProtocolMatches"
        )) {
            New-ItemProperty $modeKey -Name $runtimeValueName -PropertyType DWord -Value 0 -Force | Out-Null
        }
        New-ItemProperty $modeKey -Name "LastRuntimeError" -PropertyType String -Value "Office integration is not installed." -Force | Out-Null
        Remove-ItemProperty -LiteralPath $modeKey -Name "OleManifestEnabled" -Force -ErrorAction SilentlyContinue
    }

    Write-UninstallLog "VisualTeX machine-wide Ribbon and OLE Office integration removed successfully."
    exit 0
} catch {
    Write-UninstallLog ("FAILED: " + $_.Exception.ToString())
    Write-Output ("REASON: " + (Write-FailureReason $_.Exception.Message))
    exit 1
} finally {
    if ($transcriptStarted) {
        try { Stop-Transcript | Out-Null } catch { }
    }
}
