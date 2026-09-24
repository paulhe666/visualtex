[CmdletBinding()]
param(
    [string]$LogPath,
    [switch]$Elevated,
    [switch]$ArchitectureRelaunched
)

$ErrorActionPreference = "Stop"

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

function Invoke-SelfProcess {
    param(
        [ValidateSet("x86", "x64")][string]$TargetPlatform,
        [bool]$RunAsAdministrator,
        [bool]$MarkArchitectureRelaunched
    )

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

    $startParameters = @{
        FilePath = Resolve-PowerShellExecutable $TargetPlatform
        ArgumentList = ($arguments -join " ")
        PassThru = $true
    }
    if ($RunAsAdministrator) { $startParameters.Verb = "RunAs" }

    $process = Start-Process @startParameters
    try {
        # Wait only for the direct PowerShell child. Do not wait for descendant
        # processes such as Windows Installer service helpers.
        $process.WaitForExit()
        $exitCode = $process.ExitCode
    } finally {
        $process.Dispose()
    }
    exit $exitCode
}

if (-not (Test-IsAdministrator)) {
    if ($Elevated) {
        throw "VisualTeX Office integration uninstall requires administrator privileges."
    }
    $currentPlatform = if ([Environment]::Is64BitProcess) { "x64" } else { "x86" }
    Invoke-SelfProcess $currentPlatform $true $false
}

if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    if ($ArchitectureRelaunched) {
        throw "Unable to relaunch VisualTeX Office uninstall in 64-bit PowerShell."
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
            throw "VisualTeX Office MSI uninstall failed for $productCode with exit code $exitCode."
        }
        if (Test-ProductInstalled $productCode) {
            throw "VisualTeX Office MSI $productCode is still registered after uninstall."
        }
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
            throw "VisualTeX Office certificate cleanup failed with exit code $certificateExitCode."
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
    exit 1
} finally {
    if ($transcriptStarted) {
        try { Stop-Transcript | Out-Null } catch { }
    }
}
