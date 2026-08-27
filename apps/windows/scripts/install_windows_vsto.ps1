[CmdletBinding()]
param(
    [string]$MsiPath,
    [string]$PackageDirectory,
    [string]$LogPath,
    [string]$HashManifestPath,
    [ValidateSet("auto", "x86", "x64")]
    [string]$OfficePlatform = "auto",
    [string]$VisualTeXPath,
    [string]$DiagnosticReportPath,
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

function Resolve-ElevationPath([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return $Value }
    $trimmed = $Value.Trim().Trim('"')
    if ([IO.Path]::IsPathRooted($trimmed)) { return [IO.Path]::GetFullPath($trimmed) }
    return [IO.Path]::GetFullPath((Join-Path (Get-Location).Path $trimmed))
}

function Resolve-EarlyOfficePlatform {
    if ($OfficePlatform -in @("x86", "x64")) { return $OfficePlatform }

    foreach ($view in @(
        [Microsoft.Win32.RegistryView]::Registry64,
        [Microsoft.Win32.RegistryView]::Registry32
    )) {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
            [Microsoft.Win32.RegistryHive]::LocalMachine,
            $view)
        try {
            $configuration = $baseKey.OpenSubKey("SOFTWARE\Microsoft\Office\ClickToRun\Configuration")
            if ($null -ne $configuration) {
                try {
                    $platform = [string]$configuration.GetValue("Platform", "")
                    if ($platform -in @("x86", "x64")) { return $platform }
                } finally { $configuration.Dispose() }
            }
        } finally { $baseKey.Dispose() }
    }

    foreach ($candidate in @(
        @{ Platform = "x64"; View = [Microsoft.Win32.RegistryView]::Registry64 },
        @{ Platform = "x86"; View = [Microsoft.Win32.RegistryView]::Registry32 }
    )) {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
            [Microsoft.Win32.RegistryHive]::LocalMachine,
            $candidate.View)
        try {
            $word = $baseKey.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE")
            $powerPoint = $baseKey.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE")
            try {
                if ($null -ne $word -and $null -ne $powerPoint) { return $candidate.Platform }
            } finally {
                if ($null -ne $word) { $word.Dispose() }
                if ($null -ne $powerPoint) { $powerPoint.Dispose() }
            }
        } finally { $baseKey.Dispose() }
    }

    return $(if ([Environment]::Is64BitOperatingSystem) { "x64" } else { "x86" })
}

function Resolve-PowerShellExecutable([string]$TargetPlatform) {
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
        [string]$TargetProcessPlatform,
        [string]$TargetOfficePlatform,
        [bool]$RunAsAdministrator,
        [bool]$MarkArchitectureRelaunched
    )
    $arguments = New-Object System.Collections.Generic.List[string]
    foreach ($value in @("-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", (Quote-ProcessArgument $PSCommandPath))) {
        [void]$arguments.Add($value)
    }
    if ($RunAsAdministrator) { [void]$arguments.Add("-Elevated") }
    if ($MarkArchitectureRelaunched) { [void]$arguments.Add("-ArchitectureRelaunched") }
    foreach ($entry in @(
        @{ Name = "MsiPath"; Value = (Resolve-ElevationPath $MsiPath) },
        @{ Name = "PackageDirectory"; Value = (Resolve-ElevationPath $PackageDirectory) },
        @{ Name = "LogPath"; Value = (Resolve-ElevationPath $LogPath) },
        @{ Name = "HashManifestPath"; Value = (Resolve-ElevationPath $HashManifestPath) },
        @{ Name = "OfficePlatform"; Value = $TargetOfficePlatform },
        @{ Name = "VisualTeXPath"; Value = (Resolve-ElevationPath $VisualTeXPath) },
        @{ Name = "DiagnosticReportPath"; Value = (Resolve-ElevationPath $DiagnosticReportPath) }
    )) {
        if ([string]::IsNullOrWhiteSpace([string]$entry.Value)) { continue }
        [void]$arguments.Add("-$($entry.Name)")
        [void]$arguments.Add((Quote-ProcessArgument ([string]$entry.Value)))
    }
    $startParameters = @{
        FilePath = Resolve-PowerShellExecutable $TargetProcessPlatform
        ArgumentList = ($arguments -join " ")
        PassThru = $true
    }
    if ($RunAsAdministrator) { $startParameters.Verb = "RunAs" }
    $process = Start-Process @startParameters
    try {
        # Start-Process -Wait waits for the entire descendant process tree on Windows.
        # The installation intentionally starts the long-lived VisualTeX companion, so
        # using -Wait here deadlocks the NSIS parent even after the PowerShell child exits.
        # Process.WaitForExit waits only for this direct child process.
        $process.WaitForExit()
        $exitCode = $process.ExitCode
    } finally {
        $process.Dispose()
    }
    exit $exitCode
}

$earlyOfficePlatform = Resolve-EarlyOfficePlatform
$requiresArchitectureRelaunch =
    ($earlyOfficePlatform -eq "x64" -and -not [Environment]::Is64BitProcess) -or
    ($earlyOfficePlatform -eq "x86" -and [Environment]::Is64BitProcess)

if (-not (Test-IsAdministrator)) {
    if ($Elevated) {
        throw "VisualTeX Office integration requires administrator privileges for machine-wide COM registration."
    }
    $elevationProcessPlatform = if ([Environment]::Is64BitProcess) { "x64" } else { "x86" }
    Invoke-SelfProcess $elevationProcessPlatform $earlyOfficePlatform $true $false
}

if ($requiresArchitectureRelaunch) {
    if ($ArchitectureRelaunched) {
        throw "Unable to relaunch the installer in a PowerShell process matching $earlyOfficePlatform Office."
    }
    Invoke-SelfProcess $earlyOfficePlatform $earlyOfficePlatform $false $true
}

$upgradeCode = "{A81B4BF7-0E51-45CE-A5AA-5E28F6944F42}"
$displayName = "VisualTeX Windows Office Integration"
$root = Split-Path -Parent $PSScriptRoot
$logRoot = Join-Path $env:LOCALAPPDATA "VisualTeX\office\install-logs"
New-Item -Path $logRoot -ItemType Directory -Force | Out-Null
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$bootstrapLogPath = Join-Path $logRoot "vsto-bootstrap-$stamp.log"
if ([string]::IsNullOrWhiteSpace($DiagnosticReportPath)) {
    $DiagnosticReportPath = Join-Path $logRoot "vsto-diagnostic-$stamp.json"
}
$script:diagnosticChecks = New-Object System.Collections.Generic.List[object]
$script:resolvedOfficePlatform = $null
$script:resolvedVisualTeXPath = $null
$script:resolvedMsiPath = $null
$script:staticInstallVerified = $false
$script:runtimeVerified = $false
$transcriptStarted = $false

function Add-DiagnosticCheck {
    param(
        [string]$Name,
        [bool]$Passed,
        [string]$Details,
        [string]$Category = "general"
    )
    [void]$script:diagnosticChecks.Add([pscustomobject]@{
        name = $Name
        category = $Category
        passed = $Passed
        details = $Details
    })
    if ($Passed) {
        Write-Host "[PASS] $Name - $Details"
    } else {
        Write-Warning "[FAIL] $Name - $Details"
    }
}

function Write-DiagnosticReport {
    param(
        [bool]$Succeeded,
        [bool]$FilesAndRegistryVerified,
        [bool]$OfficeRuntimeVerified,
        [string]$FailureMessage = ""
    )
    $report = [ordered]@{
        schemaVersion = 1
        generatedAt = [DateTimeOffset]::Now.ToString("o")
        succeeded = $Succeeded
        filesAndRegistryVerified = $FilesAndRegistryVerified
        officeRuntimeVerified = $OfficeRuntimeVerified
        officePlatform = $script:resolvedOfficePlatform
        visualTeXPath = $script:resolvedVisualTeXPath
        msiPath = $script:resolvedMsiPath
        msiLogPath = $LogPath
        bootstrapLogPath = $bootstrapLogPath
        failureMessage = $FailureMessage
        checks = @($script:diagnosticChecks | ForEach-Object { $_ })
    }
    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $DiagnosticReportPath -Encoding UTF8
    Write-Host "Diagnostic report: $DiagnosticReportPath"
}

function Resolve-VisualTeXExecutable {
    $attempted = New-Object System.Collections.Generic.List[string]
    $explicit = if ([string]::IsNullOrWhiteSpace($VisualTeXPath)) { $null } else { $VisualTeXPath.Trim().Trim('"') }
    if (-not [string]::IsNullOrWhiteSpace($explicit)) {
        if (Test-Path -LiteralPath $explicit -PathType Container) {
            $explicit = Join-Path $explicit "VisualTeX.exe"
        }
        [void]$attempted.Add($explicit)
        if (-not (Test-Path -LiteralPath $explicit -PathType Leaf)) {
            throw "The explicitly supplied VisualTeX executable does not exist: $explicit"
        }
        return (Resolve-Path -LiteralPath $explicit).Path
    }

    $modeKey = "HKCU:\Software\VisualTeX\OfficeIntegration"
    $registered = [string](Get-ItemProperty -LiteralPath $modeKey -Name ExecutablePath -ErrorAction SilentlyContinue).ExecutablePath
    if (-not [string]::IsNullOrWhiteSpace($registered)) {
        $registered = $registered.Trim().Trim('"')
        [void]$attempted.Add($registered)
        if (Test-Path -LiteralPath $registered -PathType Leaf) {
            return (Resolve-Path -LiteralPath $registered).Path
        }
    }

    $installedRoot = Split-Path -Parent $PSScriptRoot
    $installedCandidate = Join-Path $installedRoot "VisualTeX.exe"
    [void]$attempted.Add($installedCandidate)
    if (Test-Path -LiteralPath $installedCandidate -PathType Leaf) {
        return (Resolve-Path -LiteralPath $installedCandidate).Path
    }

    $attempts = ($attempted | ForEach-Object { "  - $_" }) -join [Environment]::NewLine
    throw @"
Unable to resolve the installed VisualTeX.exe.
Explicit -VisualTeXPath: <not supplied>
Registry HKCU\Software\VisualTeX\OfficeIntegration\ExecutablePath: $(if ([string]::IsNullOrWhiteSpace($registered)) { '<missing>' } else { $registered })
Attempted paths:
$attempts
Pass -VisualTeXPath with the exact executable path.
"@
}

function Get-RegistryView([string]$Architecture) {
    if ($Architecture -eq "x86") { return [Microsoft.Win32.RegistryView]::Registry32 }
    return [Microsoft.Win32.RegistryView]::Registry64
}

function Open-RegistryBaseKey {
    param(
        [Microsoft.Win32.RegistryHive]$Hive,
        [string]$Architecture
    )
    return [Microsoft.Win32.RegistryKey]::OpenBaseKey($Hive, (Get-RegistryView $Architecture))
}

function Get-RegistryValue {
    param(
        [Microsoft.Win32.RegistryHive]$Hive,
        [string]$SubKey,
        [string]$Name,
        [string]$Architecture,
        [switch]$Optional
    )
    $baseKey = Open-RegistryBaseKey $Hive $Architecture
    try {
        $key = $baseKey.OpenSubKey($SubKey, $false)
        if ($null -eq $key) {
            if ($Optional) { return $null }
            throw "Required registry key is missing: $Hive\$SubKey ($Architecture view)"
        }
        try {
            $valueName = if ($Name -eq "(default)") { "" } else { $Name }
            if ($valueName -notin @($key.GetValueNames())) {
                if ($Optional) { return $null }
                throw "Required registry value is missing: $Hive\$SubKey::$Name ($Architecture view)"
            }
            return $key.GetValue(
                $valueName,
                $null,
                [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
        } finally {
            $key.Dispose()
        }
    } finally {
        $baseKey.Dispose()
    }
}

function Test-RegistryKey {
    param(
        [Microsoft.Win32.RegistryHive]$Hive,
        [string]$SubKey,
        [string]$Architecture
    )
    $baseKey = Open-RegistryBaseKey $Hive $Architecture
    try {
        $key = $baseKey.OpenSubKey($SubKey, $false)
        if ($null -eq $key) { return $false }
        $key.Dispose()
        return $true
    } finally {
        $baseKey.Dispose()
    }
}

function Assert-RegistryValue {
    param(
        [Microsoft.Win32.RegistryHive]$Hive,
        [string]$SubKey,
        [string]$Name,
        $Expected,
        [string]$Architecture
    )
    $actual = Get-RegistryValue $Hive $SubKey $Name $Architecture
    if ($actual -ne $Expected) {
        throw "Registry value $Hive\$SubKey::$Name is '$actual'; expected '$Expected' ($Architecture view)."
    }
    return $actual
}

function Convert-PathToFileUri([string]$Path) {
    return ([Uri](Resolve-Path -LiteralPath $Path).Path).AbsoluteUri
}

function Test-FileUriTargetsPath([string]$UriValue, [string]$ExpectedPath) {
    if ([string]::IsNullOrWhiteSpace($UriValue)) { return $false }
    $trimmed = $UriValue.Trim().TrimEnd('|')
    if ($trimmed.EndsWith("|vstolocal", [StringComparison]::OrdinalIgnoreCase)) {
        $trimmed = $trimmed.Substring(0, $trimmed.Length - "|vstolocal".Length)
    }
    try {
        $actualPath = ([Uri]$trimmed).LocalPath
        return [string]::Equals(
            [IO.Path]::GetFullPath($actualPath),
            [IO.Path]::GetFullPath($ExpectedPath),
            [StringComparison]::OrdinalIgnoreCase)
    } catch {
        return $false
    }
}

function Assert-NoOfficeProcesses {
    $processNames = @(
        "WINWORD", "POWERPNT", "EXCEL", "OUTLOOK", "ONENOTE",
        "MSACCESS", "MSPUB", "VISIO", "MSPROJECT"
    )
    $running = @(Get-Process $processNames -ErrorAction SilentlyContinue)
    if ($running.Count -gt 0) {
        $names = $running | Sort-Object ProcessName -Unique | ForEach-Object { $_.ProcessName + ".EXE" }
        throw "Close all Microsoft Office applications before installing VisualTeX Office integration. Running: $($names -join ', '). No MSI or registry changes were made."
    }
    Add-DiagnosticCheck "Office processes closed" $true "No Word, PowerPoint or other common Office process is running." "prerequisite"
}

function Get-OfficeExecutablePath {
    param(
        [ValidateSet("Word", "PowerPoint")][string]$HostName,
        [string]$Architecture
    )
    $fileName = if ($HostName -eq "Word") { "WINWORD.EXE" } else { "POWERPNT.EXE" }
    $appPath = Get-RegistryValue `
        ([Microsoft.Win32.RegistryHive]::LocalMachine) `
        "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$fileName" `
        "(default)" `
        $Architecture `
        -Optional
    if (-not [string]::IsNullOrWhiteSpace([string]$appPath) -and
        (Test-Path -LiteralPath ([string]$appPath) -PathType Leaf)) {
        return (Resolve-Path -LiteralPath ([string]$appPath)).Path
    }

    foreach ($registryArchitecture in @("x64", "x86")) {
        $installationPath = Get-RegistryValue `
            ([Microsoft.Win32.RegistryHive]::LocalMachine) `
            "SOFTWARE\Microsoft\Office\ClickToRun\Configuration" `
            "InstallationPath" `
            $registryArchitecture `
            -Optional
        if ([string]::IsNullOrWhiteSpace([string]$installationPath)) { continue }
        foreach ($candidate in @(
            (Join-Path ([string]$installationPath) "root\Office16\$fileName"),
            (Join-Path ([string]$installationPath) "Office16\$fileName")
        )) {
            if (Test-Path -LiteralPath $candidate -PathType Leaf) {
                return (Resolve-Path -LiteralPath $candidate).Path
            }
        }
    }
    return $null
}

function Resolve-OfficePlatform([string]$RequestedPlatform) {
    $clickToRun = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue
    $detected = if ($null -ne $clickToRun -and $clickToRun.Platform -in @("x86", "x64")) {
        [string]$clickToRun.Platform
    } else {
        $null
    }

    if ($RequestedPlatform -ne "auto") {
        if ($detected -and $detected -ne $RequestedPlatform) {
            throw "Requested Office platform '$RequestedPlatform' does not match Click-to-Run Office platform '$detected'."
        }
        return $RequestedPlatform
    }
    if ($detected) { return $detected }

    foreach ($architecture in @("x64", "x86")) {
        $word = Get-OfficeExecutablePath Word $architecture
        $powerPoint = Get-OfficeExecutablePath PowerPoint $architecture
        if ($word -and $powerPoint) { return $architecture }
    }
    throw "Unable to determine whether installed Office is x86 or x64. Pass -OfficePlatform x86 or x64."
}

function Assert-OfficeApplicationsInstalled([string]$Architecture) {
    $word = Get-OfficeExecutablePath Word $Architecture
    $powerPoint = Get-OfficeExecutablePath PowerPoint $Architecture
    if (-not $word) { throw "Microsoft Word was not found for the detected $Architecture Office installation." }
    if (-not $powerPoint) { throw "Microsoft PowerPoint was not found for the detected $Architecture Office installation." }
    Add-DiagnosticCheck "Microsoft Word present" $true $word "prerequisite"
    Add-DiagnosticCheck "Microsoft PowerPoint present" $true $powerPoint "prerequisite"
}

function Assert-VstoRuntimeInstalled([string]$Architecture) {
    $subKey = "SOFTWARE\Microsoft\VSTO Runtime Setup\v4R"
    $candidateViews = @($Architecture, "x86", "x64") | Select-Object -Unique
    foreach ($registryArchitecture in $candidateViews) {
        if (-not (Test-RegistryKey ([Microsoft.Win32.RegistryHive]::LocalMachine) $subKey $registryArchitecture)) {
            continue
        }
        $install = Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) $subKey "Install" $registryArchitecture -Optional
        $clr40 = Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) $subKey "VSTORFeature_CLR40" $registryArchitecture -Optional
        $version = Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) $subKey "Version" $registryArchitecture -Optional
        $installed =
            ($null -ne $install -and [int]$install -eq 1) -or
            ($null -ne $clr40 -and [int]$clr40 -eq 1) -or
            (-not [string]::IsNullOrWhiteSpace([string]$version))
        if ($installed) {
            Add-DiagnosticCheck "VSTO Runtime" $true "HKLM\$subKey; Install=$install; VSTORFeature_CLR40=$clr40; Version=$version; registryView=$registryArchitecture; Office=$Architecture" "prerequisite"
            return
        }
    }
    throw "Microsoft Visual Studio Tools for Office Runtime is missing or incomplete at HKLM\$subKey. Checked x86 and x64 registry views for Office $Architecture."
}

function Assert-NetFramework472Installed {
    $releaseValues = @()
    foreach ($architecture in @("x64", "x86")) {
        $release = Get-RegistryValue `
            ([Microsoft.Win32.RegistryHive]::LocalMachine) `
            "SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" `
            "Release" `
            $architecture `
            -Optional
        if ($null -ne $release) { $releaseValues += [int]$release }
    }
    if ($releaseValues.Count -eq 0) {
        throw ".NET Framework 4 Full installation was not found. VisualTeX Office integration targets .NET Framework 4.7.2."
    }
    $releaseValue = ($releaseValues | Measure-Object -Maximum).Maximum
    $minimumRelease = 461808
    if ($releaseValue -lt $minimumRelease) {
        throw ".NET Framework 4.7.2 or newer is required. Detected Release=$releaseValue; expected at least $minimumRelease."
    }
    Add-DiagnosticCheck ".NET Framework 4.7.2" $true "Release=$releaseValue" "prerequisite"
}

function Resolve-MachineOfficeInstallRoot([string]$Architecture) {
    $programFilesRoot = if ($Architecture -eq "x64") {
        if (-not [string]::IsNullOrWhiteSpace($env:ProgramW6432)) {
            $env:ProgramW6432
        } elseif ([Environment]::Is64BitProcess) {
            $env:ProgramFiles
        } else {
            $null
        }
    } else {
        if (-not [string]::IsNullOrWhiteSpace(${env:ProgramFiles(x86)})) {
            ${env:ProgramFiles(x86)}
        } else {
            $env:ProgramFiles
        }
    }
    if ([string]::IsNullOrWhiteSpace($programFilesRoot)) {
        throw "Unable to resolve Program Files for $Architecture Office integration. Process64=$([Environment]::Is64BitProcess); OS64=$([Environment]::Is64BitOperatingSystem)."
    }
    return Join-Path $programFilesRoot "VisualTeX\WindowsOffice\VSTO"
}

function Remove-LegacyPerUserOfficeRegistration {
    foreach ($key in @(
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
        "HKCU:\Software\Classes\TypeLib\{DF66EC66-3B3A-4675-A7BE-30456A04EB96}",
        "HKCU:\Software\Classes\AppID\{3C72FF7F-B04A-4FD0-AA7D-61D110D8B3C1}",
        "HKCU:\Software\Classes\AppID\VisualTeX.FormulaOleServer.exe"
    )) {
        Remove-Item -LiteralPath $key -Recurse -Force -ErrorAction SilentlyContinue
    }
    Add-DiagnosticCheck "Legacy per-user COM registration removed" $true "HKCU Office Addins, managed COM and OLE class registrations were cleared before machine-wide installation." "registry"
}

function Assert-CurrentUserCertificateTrusted {
    $modeKey = "HKCU:\Software\VisualTeX\OfficeIntegration"
    $thumbprint = (Get-ItemProperty -LiteralPath $modeKey -Name CertificateThumbprint -ErrorAction SilentlyContinue).CertificateThumbprint
    $certificatePath = (Get-ItemProperty -LiteralPath $modeKey -Name CertificatePath -ErrorAction SilentlyContinue).CertificatePath
    if ([string]::IsNullOrWhiteSpace([string]$thumbprint)) {
        throw "The VisualTeX current-user HTTPS certificate has not been prepared by the installer prerequisite step."
    }
    if ([string]::IsNullOrWhiteSpace([string]$certificatePath) -or
        -not (Test-Path -LiteralPath ([string]$certificatePath) -PathType Leaf)) {
        throw "The VisualTeX certificate file recorded in the registry is missing: $certificatePath"
    }
    $certificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new([string]$certificatePath)
    if ($certificate.Thumbprint -ne ([string]$thumbprint).Replace(" ", "")) {
        throw "The VisualTeX certificate file thumbprint does not match the registry value."
    }
    $store = [Security.Cryptography.X509Certificates.X509Store]::new(
        [Security.Cryptography.X509Certificates.StoreName]::Root,
        [Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
    $store.Open([Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
    try {
        $trusted = @($store.Certificates | Where-Object {
            $_.Thumbprint -eq $certificate.Thumbprint
        }).Count -gt 0
    } finally {
        $store.Close()
    }
    if (-not $trusted) {
        throw "The VisualTeX HTTPS certificate is not trusted in the current-user Root store: $($certificate.Thumbprint)"
    }
    Add-DiagnosticCheck "Current-user HTTPS certificate" $true "$($certificate.Thumbprint); $certificatePath" "prerequisite"
}

function Assert-SharedCompanionConfiguration([string]$ExecutablePath) {
    $modeKey = "HKCU:\Software\VisualTeX\OfficeIntegration"
    $state = Get-ItemProperty -LiteralPath $modeKey -ErrorAction SilentlyContinue
    if ($null -eq $state) {
        throw "The shared VisualTeX Office registry configuration is missing: $modeKey"
    }
    $registeredExecutable = [string]$state.ExecutablePath
    if ([string]::IsNullOrWhiteSpace($registeredExecutable) -or
        -not [string]::Equals(
            [IO.Path]::GetFullPath($registeredExecutable.Trim().Trim('"')),
            [IO.Path]::GetFullPath($ExecutablePath),
            [StringComparison]::OrdinalIgnoreCase)) {
        throw "Registered ExecutablePath '$registeredExecutable' does not match the exact installer path '$ExecutablePath'."
    }
    $appDataRoot = [string]$state.AppDataRoot
    if ([string]::IsNullOrWhiteSpace($appDataRoot) -or
        -not (Test-Path -LiteralPath $appDataRoot -PathType Container)) {
        throw "The shared AppDataRoot is missing or invalid: $appDataRoot"
    }
    $certificatePath = [string]$state.CertificatePath
    if ([string]::IsNullOrWhiteSpace($certificatePath) -or
        -not (Test-Path -LiteralPath $certificatePath -PathType Leaf)) {
        throw "The shared CertificatePath is missing or invalid: $certificatePath"
    }
    if ([int]$state.CompanionPort -le 0 -or [int]$state.CompanionPort -gt 65535) {
        throw "The shared CompanionPort is missing or invalid: $($state.CompanionPort)"
    }
    if ([int]$state.ProtocolVersion -le 0) {
        throw "The shared ProtocolVersion is missing or invalid: $($state.ProtocolVersion)"
    }
    $installJson = Join-Path $appDataRoot "office\install.json"
    if (-not (Test-Path -LiteralPath $installJson -PathType Leaf)) {
        throw "VisualTeX install.json is missing: $installJson"
    }
    try {
        $installState = Get-Content -LiteralPath $installJson -Raw | ConvertFrom-Json
    } catch {
        throw "VisualTeX install.json is invalid: $installJson; $($_.Exception.Message)"
    }
    if ([int]$installState.port -ne [int]$state.CompanionPort) {
        throw "install.json port=$($installState.port) does not match CompanionPort=$($state.CompanionPort)."
    }
    if ([int]$installState.protocolVersion -ne [int]$state.ProtocolVersion) {
        throw "install.json protocolVersion=$($installState.protocolVersion) does not match ProtocolVersion=$($state.ProtocolVersion)."
    }
    if ([string]::IsNullOrWhiteSpace([string]$installState.installToken) -or
        [string]$installState.installToken.Length -ne 64) {
        throw "VisualTeX install.json contains an invalid installToken: $installJson"
    }
    Add-DiagnosticCheck "Shared companion configuration" $true "ExecutablePath=$registeredExecutable; AppDataRoot=$appDataRoot; CertificatePath=$certificatePath; CompanionPort=$($state.CompanionPort); ProtocolVersion=$($state.ProtocolVersion); InstallJson=$installJson" "prerequisite"
    return $state
}

function Assert-OfficeUiResources([string]$ExecutablePath) {
    $executableRoot = Split-Path -Parent $ExecutablePath
    $candidates = @(
        (Join-Path $executableRoot "office\dialog\index.html"),
        (Join-Path $executableRoot "resources\office\dialog\index.html")
    )
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            Add-DiagnosticCheck "Office companion UI resources" $true $candidate "prerequisite"
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }
    throw "Office companion UI resource is missing. Checked: $($candidates -join '; ')"
}

function Get-MsiArchitecture([string]$Path) {
    $installer = $null
    $summary = $null
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $summary = $installer.GetType().InvokeMember(
            "SummaryInformation",
            [Reflection.BindingFlags]::GetProperty,
            $null,
            $installer,
            @($Path, 0))
        $template = [string]$summary.GetType().InvokeMember(
            "Property",
            [Reflection.BindingFlags]::GetProperty,
            $null,
            $summary,
            @([int]7))
        if ($template -match '(^|;)(x64|Intel64)(;|$)') { return "x64" }
        if ($template -match '(^|;)Intel(;|$)') { return "x86" }
        throw "Unable to map MSI SummaryInformation Template '$template' to x86 or x64."
    } finally {
        if ($null -ne $summary) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($summary) }
        if ($null -ne $installer) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($installer) }
    }
}

function Assert-MsiArchitecture([string]$Path, [string]$Architecture) {
    $msiArchitecture = Get-MsiArchitecture $Path
    if ($msiArchitecture -ne $Architecture) {
        throw "MSI architecture '$msiArchitecture' does not match installed Office '$Architecture'."
    }
    Add-DiagnosticCheck "MSI architecture" $true "$msiArchitecture; $(Split-Path -Leaf $Path)" "package"
}

function Remove-LegacyOfficeJsState {
    foreach ($key in @(
        "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\VisualTeX",
        "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{69C6A866-755B-4C5A-BACB-EEA28B03C724}"
    )) {
        Remove-Item -LiteralPath $key -Recurse -Force -ErrorAction SilentlyContinue
    }
    $developerKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
    if (Test-Path -LiteralPath $developerKey) {
        foreach ($manifestId in @(
            "7c7d3b35-56b2-4c40-88d9-c9eb836d6021",
            "fdc8d615-7e60-4586-bff4-5a1d728f9f6c"
        )) {
            Remove-ItemProperty -LiteralPath $developerKey -Name $manifestId -Force -ErrorAction SilentlyContinue
        }
    }
    $catalog = Join-Path $env:LOCALAPPDATA "VisualTeX\OfficeCatalog"
    Remove-Item -LiteralPath $catalog -Recurse -Force -ErrorAction SilentlyContinue
    Add-DiagnosticCheck "Legacy Office.js state removed" $true "TrustedCatalogs, Developer manifest values and OfficeCatalog were removed if present." "migration"
}

function Invoke-MsiExec {
    param(
        [string[]]$Arguments,
        [string]$Operation,
        [int[]]$AllowedExitCodes
    )
    Assert-NoOfficeProcesses
    $effectiveArguments = @($Arguments) + @(
        "REBOOT=ReallySuppress",
        "MSIRESTARTMANAGERCONTROL=Disable"
    )
    $process = Start-Process msiexec.exe -ArgumentList $effectiveArguments -Wait -PassThru
    Write-Host "$Operation exit code: $($process.ExitCode)"
    if ($process.ExitCode -notin $AllowedExitCodes) {
        throw "$Operation failed with exit code $($process.ExitCode). See the verbose MSI log."
    }
    return $process.ExitCode
}

function Get-RelatedProductCodes {
    $codes = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
    $installer = $null
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $related = $installer.GetType().InvokeMember(
            "RelatedProducts",
            [Reflection.BindingFlags]::GetProperty,
            $null,
            $installer,
            @($upgradeCode))
        foreach ($code in @($related)) {
            if (-not [string]::IsNullOrWhiteSpace($code)) { [void]$codes.Add([string]$code) }
        }
    } finally {
        if ($null -ne $installer) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($installer) }
    }

    foreach ($uninstallRoot in @(
        "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKCU:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    )) {
        if (-not (Test-Path -LiteralPath $uninstallRoot)) { continue }
        foreach ($key in Get-ChildItem -LiteralPath $uninstallRoot -ErrorAction SilentlyContinue) {
            $item = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction SilentlyContinue
            if ($null -eq $item) { continue }
            if ([string]$item.DisplayName -eq $displayName -and
                $key.PSChildName -match '^\{[0-9A-Fa-f-]{36}\}$') {
                [void]$codes.Add($key.PSChildName)
            }
        }
    }
    return @($codes)
}

function Wait-ForRelatedProductCount([int]$ExpectedCount, [int]$TimeoutSeconds = 20) {
    $deadline = [DateTimeOffset]::UtcNow.AddSeconds($TimeoutSeconds)
    do {
        $codes = @(Get-RelatedProductCodes)
        if ($codes.Count -eq $ExpectedCount) { return $codes }
        Start-Sleep -Milliseconds 250
    } while ([DateTimeOffset]::UtcNow -lt $deadline)
    throw "Windows Installer product state did not settle at $ExpectedCount related product(s) within $TimeoutSeconds seconds. Last state: $($codes -join ', ')."
}

function Remove-VisualTeXOfficeResiliencyEntries {
    foreach ($hostName in @("Word", "PowerPoint")) {
        foreach ($bucket in @("StartupItems", "DisabledItems")) {
            $key = "HKCU:\Software\Microsoft\Office\16.0\$hostName\Resiliency\$bucket"
            if (-not (Test-Path -LiteralPath $key)) { continue }
            $properties = (Get-ItemProperty -LiteralPath $key).PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' }
            foreach ($property in $properties) {
                if ($property.Value -isnot [byte[]]) { continue }
                $text = [Text.Encoding]::Unicode.GetString([byte[]]$property.Value)
                if ($text -match 'VisualTeX\.(Word|PowerPoint)Vsto' -or
                    $text -match 'VisualTeX\\WindowsOffice\\VSTO') {
                    Remove-ItemProperty -LiteralPath $key -Name $property.Name -Force
                }
            }
        }
    }
}

function Assert-OfficeAddinRegistration {
    param(
        [ValidateSet("Word", "PowerPoint")][string]$HostName,
        [string]$ProgId,
        [string]$AssemblyPath,
        [string]$ExpectedDescription,
        [string]$Architecture
    )
    $subKey = "Software\Microsoft\Office\$HostName\Addins\$ProgId"
    Assert-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) $subKey "FriendlyName" "VisualTeX" $Architecture | Out-Null
    Assert-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) $subKey "Description" $ExpectedDescription $Architecture | Out-Null
    Assert-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) $subKey "LoadBehavior" 3 $Architecture | Out-Null
    $manifest = Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) $subKey "Manifest" $Architecture -Optional
    $codeBase = Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) $subKey "CodeBase" $Architecture -Optional
    $location = if (-not [string]::IsNullOrWhiteSpace([string]$manifest)) { [string]$manifest } else { [string]$codeBase }
    if ([string]::IsNullOrWhiteSpace($location)) {
        throw "$HostName add-in registration does not contain Manifest or CodeBase: HKLM\$subKey"
    }
    if (-not (Test-FileUriTargetsPath $location $AssemblyPath)) {
        throw "$HostName add-in registration location '$location' does not target '$AssemblyPath'."
    }
    Add-DiagnosticCheck "$HostName add-in registration" $true "FriendlyName, Description, LoadBehavior=3 and location=$location" "registry"
}

function Assert-ComRegistration {
    param(
        [string]$ProgId,
        [string]$Clsid,
        [string]$AssemblyPath,
        [string]$ClassName,
        [string]$Architecture
    )
    $assemblyName = [Reflection.AssemblyName]::GetAssemblyName($AssemblyPath)
    $reflectionAssembly = [Reflection.Assembly]::ReflectionOnlyLoadFrom($AssemblyPath)
    $runtimeVersion = $reflectionAssembly.ImageRuntimeVersion
    $version = $assemblyName.Version.ToString()
    Assert-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "Software\Classes\$ProgId\CLSID" "(default)" $Clsid $Architecture | Out-Null
    $classKey = "Software\Classes\CLSID\$Clsid"
    Assert-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "$classKey\InprocServer32" "(default)" "mscoree.dll" $Architecture | Out-Null
    Assert-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "$classKey\InprocServer32" "Class" $ClassName $Architecture | Out-Null
    Assert-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "$classKey\InprocServer32" "Assembly" $assemblyName.FullName $Architecture | Out-Null
    Assert-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "$classKey\InprocServer32" "RuntimeVersion" $runtimeVersion $Architecture | Out-Null
    $codeBase = [string](Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "$classKey\InprocServer32" "CodeBase" $Architecture)
    if (-not (Test-FileUriTargetsPath $codeBase $AssemblyPath)) {
        throw "$ProgId COM CodeBase '$codeBase' does not target '$AssemblyPath'."
    }
    $versionKey = "$classKey\InprocServer32\$version"
    Assert-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) $versionKey "Class" $ClassName $Architecture | Out-Null
    Assert-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) $versionKey "Assembly" $assemblyName.FullName $Architecture | Out-Null
    Assert-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) $versionKey "RuntimeVersion" $runtimeVersion $Architecture | Out-Null
    $versionCodeBase = [string](Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) $versionKey "CodeBase" $Architecture)
    if (-not (Test-FileUriTargetsPath $versionCodeBase $AssemblyPath)) {
        throw "$ProgId versioned COM CodeBase '$versionCodeBase' does not target '$AssemblyPath'."
    }
    foreach ($requiredKey in @(
        "$classKey\ProgId",
        "$classKey\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
    )) {
        if (-not (Test-RegistryKey ([Microsoft.Win32.RegistryHive]::LocalMachine) $requiredKey $Architecture)) {
            throw "Managed COM registration is incomplete: HKLM\$requiredKey ($Architecture view)"
        }
    }
    Add-DiagnosticCheck "$ProgId COM registration" $true "CLSID=$Clsid; Assembly=$($assemblyName.FullName); Runtime=$runtimeVersion; Directory=$(Split-Path -Parent $AssemblyPath)" "registry"
}

function Assert-ManagedComActivation([string]$ProgId) {
    $instance = $null
    try {
        $type = [Type]::GetTypeFromProgID($ProgId, $true)
        if ($null -eq $type) {
            throw "ProgID '$ProgId' did not resolve to a COM type."
        }
        $instance = [Activator]::CreateInstance($type)
        if ($null -eq $instance) {
            throw "ProgID '$ProgId' returned a null COM instance."
        }
        Add-DiagnosticCheck "$ProgId COM activation" $true "CoCreateInstance succeeded for CLSID=$($type.GUID)." "registry"
    } catch {
        throw "$ProgId COM activation failed. This usually indicates an invalid COM registration scope or CodeBase. $($_.Exception.Message)"
    } finally {
        if ($null -ne $instance -and [Runtime.InteropServices.Marshal]::IsComObject($instance)) {
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($instance) } catch { }
        }
    }
}

function Assert-NativeOleRegistration([string]$InstallRoot, [string]$Architecture) {
    $clsid = "{8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B}"
    $iid = "{6C672AF0-7321-4D21-B325-868CB34592C2}"
    $libid = "{DF66EC66-3B3A-4675-A7BE-30456A04EB96}"
    $server = Join-Path $InstallRoot "VisualTeX.FormulaOleServer.exe"
    if (-not (Test-Path -LiteralPath $server -PathType Leaf)) { throw "Native Formula OLE LocalServer is missing: $server" }

    Assert-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "Software\Classes\VisualTeX.Formula.1\CLSID" "(default)" $clsid $Architecture | Out-Null
    Assert-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "Software\Classes\VisualTeX.Formula\CurVer" "(default)" "VisualTeX.Formula.1" $Architecture | Out-Null
    $localServer = [string](Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "Software\Classes\CLSID\$clsid\LocalServer32" "(default)" $Architecture)
    $localServer = $localServer.Trim('"')
    if (-not [string]::Equals($localServer, $server, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Native OLE LocalServer32 is '$localServer'; expected '$server' ($Architecture view)."
    }
    Assert-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "Software\Classes\CLSID\$clsid\LocalServer32" "ServerExecutable" $server $Architecture | Out-Null
    Assert-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "Software\Classes\CLSID\$clsid\DataFormats\GetSet\0" "(default)" "14,1,64,1" $Architecture | Out-Null
    Assert-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "Software\Classes\CLSID\$clsid\DataFormats\GetSet\1" "(default)" "3,1,32,1" $Architecture | Out-Null
    Assert-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "Software\Classes\CLSID\$clsid\DataFormats\GetSet\2" "(default)" "PNG,1,1,1" $Architecture | Out-Null
    Assert-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "Software\Classes\CLSID\$clsid\ProgID" "(default)" "VisualTeX.Formula.1" $Architecture | Out-Null
    Assert-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "Software\Classes\CLSID\$clsid\VersionIndependentProgID" "(default)" "VisualTeX.Formula" $Architecture | Out-Null
    Assert-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "Software\Classes\Interface\$iid\ProxyStubClsid32" "(default)" "{00020424-0000-0000-C000-000000000046}" $Architecture | Out-Null
    Assert-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "Software\Classes\Interface\$iid\TypeLib" "(default)" $libid $Architecture | Out-Null
    $typeLibraryPlatform = if ($Architecture -eq "x86") { "win32" } else { "win64" }
    $typeLibrary = [string](Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "Software\Classes\TypeLib\$libid\1.0\0\$typeLibraryPlatform" "(default)" $Architecture)
    if (-not [string]::Equals($typeLibrary, $server, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Native OLE type library path is '$typeLibrary'; expected '$server' ($Architecture view)."
    }
    Add-DiagnosticCheck "OLE LocalServer registration" $true "$server; CLSID=$clsid; TypeLib=$libid" "ole"
    return $server
}

function Stop-VisualTeXProcessesForRepair {
    $processes = @(Get-Process visualtex -ErrorAction SilentlyContinue)
    foreach ($process in $processes) {
        $path = ""
        try { $path = [string]$process.Path } catch { }
        Write-Host "Stopping stale VisualTeX process PID=$($process.Id) Path=$path before Office integration repair."
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
    }
    if ($processes.Count -gt 0) { Start-Sleep -Milliseconds 700 }
    $remaining = @(Get-Process visualtex -ErrorAction SilentlyContinue)
    if ($remaining.Count -gt 0) {
        throw "Unable to stop all VisualTeX processes before Office integration repair. Remaining PIDs: $($remaining.Id -join ', ')"
    }
    Add-DiagnosticCheck "Stale VisualTeX processes" $true "No VisualTeX process remained before Office installation and runtime verification." "prerequisite"
}

function Assert-NativeOleServerStarts([string]$ServerPath) {
    $process = $null
    try {
        $process = Start-Process -FilePath $ServerPath -ArgumentList "-Embedding" -WindowStyle Hidden -PassThru
        Start-Sleep -Milliseconds 900
        if ($process.HasExited -and $process.ExitCode -ne 0) {
            throw "VisualTeX.FormulaOleServer.exe exited with code $($process.ExitCode) during the -Embedding health probe."
        }
        Add-DiagnosticCheck "OLE LocalServer health probe" $true "The LocalServer executable accepted -Embedding without a non-zero failure." "ole"
    } finally {
        if ($null -ne $process -and -not $process.HasExited) {
            Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
        }
    }
}

try {
    Start-Transcript -Path $bootstrapLogPath -Force | Out-Null
    $transcriptStarted = $true

    $script:resolvedVisualTeXPath = Resolve-VisualTeXExecutable
    Add-DiagnosticCheck "VisualTeX executable" $true $script:resolvedVisualTeXPath "prerequisite"
    Assert-NoOfficeProcesses
    Stop-VisualTeXProcessesForRepair
    $script:resolvedOfficePlatform = Resolve-OfficePlatform $OfficePlatform
    Write-Host "Detected Office platform: $script:resolvedOfficePlatform"
    Assert-OfficeApplicationsInstalled $script:resolvedOfficePlatform
    Assert-VstoRuntimeInstalled $script:resolvedOfficePlatform
    Assert-NetFramework472Installed
    Assert-CurrentUserCertificateTrusted
    $sharedConfiguration = Assert-SharedCompanionConfiguration $script:resolvedVisualTeXPath
    $officeUiResource = Assert-OfficeUiResources $script:resolvedVisualTeXPath

    if ([string]::IsNullOrWhiteSpace($PackageDirectory)) {
        foreach ($candidate in @(
            (Join-Path (Split-Path -Parent $PSScriptRoot) "windows-office"),
            (Join-Path $root "src-tauri\resources\windows-office")
        )) {
            if (Test-Path -LiteralPath $candidate -PathType Container) {
                $PackageDirectory = $candidate
                break
            }
        }
    }

    $packageFileName = "VisualTeX-WindowsOffice-VSTO-$script:resolvedOfficePlatform.msi"
    $manifestFileName = "VisualTeX-WindowsOffice-VSTO-$script:resolvedOfficePlatform.sha256.json"
    if ([string]::IsNullOrWhiteSpace($MsiPath) -and -not [string]::IsNullOrWhiteSpace($PackageDirectory)) {
        $MsiPath = Join-Path $PackageDirectory $packageFileName
    }
    if ([string]::IsNullOrWhiteSpace($MsiPath) -or -not (Test-Path -LiteralPath $MsiPath -PathType Leaf)) {
        throw "VisualTeX $script:resolvedOfficePlatform Office MSI was not found. Pass -PackageDirectory or the exact -MsiPath."
    }
    $MsiPath = (Resolve-Path -LiteralPath $MsiPath).Path
    $script:resolvedMsiPath = $MsiPath
    if ([string]::IsNullOrWhiteSpace($HashManifestPath)) {
        $HashManifestPath = Join-Path (Split-Path -Parent $MsiPath) $manifestFileName
    }
    if (-not (Test-Path -LiteralPath $HashManifestPath -PathType Leaf)) {
        throw "Office package SHA-256 manifest is missing: $HashManifestPath"
    }
    $HashManifestPath = (Resolve-Path -LiteralPath $HashManifestPath).Path
    $hashManifest = Get-Content -LiteralPath $HashManifestPath -Raw | ConvertFrom-Json
    if ($hashManifest.architecture -ne $script:resolvedOfficePlatform) {
        throw "Package manifest architecture '$($hashManifest.architecture)' does not match installed Office '$script:resolvedOfficePlatform'."
    }
    if ($hashManifest.package.file -ne (Split-Path -Leaf $MsiPath)) {
        throw "Package manifest expects '$($hashManifest.package.file)' but received '$(Split-Path -Leaf $MsiPath)'."
    }
    Assert-MsiArchitecture $MsiPath $script:resolvedOfficePlatform
    $actualPackageHash = (Get-FileHash -LiteralPath $MsiPath -Algorithm SHA256).Hash
    if ($actualPackageHash -ne $hashManifest.package.sha256) {
        throw "MSI SHA-256 mismatch: $actualPackageHash != $($hashManifest.package.sha256)"
    }
    Add-DiagnosticCheck "MSI SHA-256" $true "$packageFileName $actualPackageHash" "package"

    if ([string]::IsNullOrWhiteSpace($LogPath)) {
        $LogPath = Join-Path $logRoot "vsto-install-$stamp.log"
    }

    Remove-LegacyOfficeJsState
    $oldProducts = @(Get-RelatedProductCodes)
    Write-Host "Related VisualTeX MSI products before install: $($oldProducts.Count)"
    foreach ($productCode in $oldProducts) {
        $uninstallLog = Join-Path $logRoot "vsto-uninstall-$($productCode.Trim('{}'))-$stamp.log"
        [void](Invoke-MsiExec @("/x", $productCode, "/passive", "/norestart", "/L*v", ('"{0}"' -f $uninstallLog)) "MSI uninstall $productCode" @(0, 1605, 3010))
    }
    [void](Wait-ForRelatedProductCount 0)
    Remove-LegacyPerUserOfficeRegistration

    [void](Invoke-MsiExec @("/i", ('"{0}"' -f $MsiPath), "/passive", "/norestart", "/L*v", ('"{0}"' -f $LogPath)) "MSI install" @(0, 3010))
    $installedProducts = @(Wait-ForRelatedProductCount 1)
    Add-DiagnosticCheck "MSI product registration" $true ($installedProducts -join ", ") "package"

    Remove-VisualTeXOfficeResiliencyEntries
    $installRoot = Resolve-MachineOfficeInstallRoot $script:resolvedOfficePlatform
    $verifiedEntries = @($hashManifest.word, $hashManifest.powerPoint, $hashManifest.formulaOleServer)
    if ($null -ne $hashManifest.dependencies) {
        $verifiedEntries += @($hashManifest.dependencies)
    }
    foreach ($entry in $verifiedEntries) {
        $installedFile = Join-Path $installRoot $entry.file
        if (-not (Test-Path -LiteralPath $installedFile -PathType Leaf)) {
            throw "Installed Office integration file is missing: $installedFile"
        }
        $actualHash = (Get-FileHash -LiteralPath $installedFile -Algorithm SHA256).Hash
        if ($actualHash -ne $entry.sha256) {
            throw "Installed file hash mismatch for $($entry.file): $actualHash != $($entry.sha256)"
        }
        Add-DiagnosticCheck "Installed file $($entry.file)" $true $actualHash "files"
    }

    $wordAssembly = Join-Path $installRoot "VisualTeX.WordVsto.dll"
    $powerPointAssembly = Join-Path $installRoot "VisualTeX.PowerPointVsto.dll"
    Assert-OfficeAddinRegistration Word "VisualTeX.WordVsto" $wordAssembly "VisualTeX native Word formula integration" $script:resolvedOfficePlatform
    Assert-OfficeAddinRegistration PowerPoint "VisualTeX.PowerPointVsto" $powerPointAssembly "VisualTeX native PowerPoint formula integration" $script:resolvedOfficePlatform
    Assert-ComRegistration "VisualTeX.WordVsto" "{F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1}" $wordAssembly "VisualTeX.WordVsto.ThisAddIn" $script:resolvedOfficePlatform
    Assert-ComRegistration "VisualTeX.PowerPointVsto" "{7E586D2D-57B0-4D14-AB24-EBA9021A5E6D}" $powerPointAssembly "VisualTeX.PowerPointVsto.ThisAddIn" $script:resolvedOfficePlatform
    Assert-ManagedComActivation "VisualTeX.WordVsto"
    Assert-ManagedComActivation "VisualTeX.PowerPointVsto"
    $oleServer = Assert-NativeOleRegistration $installRoot $script:resolvedOfficePlatform
    Assert-NativeOleServerStarts $oleServer

    $modeKey = "HKCU:\Software\VisualTeX\OfficeIntegration"
    if (-not (Test-Path -LiteralPath $modeKey)) { New-Item -Path $modeKey -Force | Out-Null }
    New-ItemProperty -LiteralPath $modeKey -Name "ExecutablePath" -PropertyType String -Value $script:resolvedVisualTeXPath -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "Mode" -PropertyType String -Value "vsto" -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "OfficePlatform" -PropertyType String -Value $script:resolvedOfficePlatform -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "NativeOleEnabled" -PropertyType DWord -Value 1 -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "FilesAndRegistryVerified" -PropertyType DWord -Value 1 -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "OfficeRuntimeVerified" -PropertyType DWord -Value 0 -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "OfficeConnectionVerificationAttempted" -PropertyType DWord -Value 0 -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "WordConnected" -PropertyType DWord -Value 0 -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "PowerPointConnected" -PropertyType DWord -Value 0 -Force | Out-Null
    foreach ($runtimeValueName in @(
        "CompanionProcessRunning",
        "CompanionPortListening",
        "CompanionHttpsHealthy",
        "CompanionCertificateMatches",
        "CompanionProtocolMatches"
    )) {
        New-ItemProperty -LiteralPath $modeKey -Name $runtimeValueName -PropertyType DWord -Value 0 -Force | Out-Null
    }
    New-ItemProperty -LiteralPath $modeKey -Name "LastRuntimeError" -PropertyType String -Value "Companion runtime verification has not started." -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "LastDiagnosticReport" -PropertyType String -Value $DiagnosticReportPath -Force | Out-Null
    Remove-ItemProperty -LiteralPath $modeKey -Name "OleManifestEnabled" -Force -ErrorAction SilentlyContinue
    $script:staticInstallVerified = $true
    Add-DiagnosticCheck "Static installation verification" $true "Machine-wide MSI, Program Files payload, HKLM Office registrations, managed COM classes and native OLE LocalServer passed." "static-install"
    Write-Host "Static files and registry installation verified successfully."

    # Do not launch the long-lived VisualTeX companion from this elevated
    # machine-wide installer process. The non-elevated NSIS parent performs the
    # companion health check after this script returns.
    Write-DiagnosticReport $true $true $false
    Write-Host "Static Office integration installed and verified. Companion runtime verification is deferred to the non-elevated installer stage."
} catch {
    Add-DiagnosticCheck "Installation result" $false $_.Exception.Message "result"
    try {
        Write-DiagnosticReport $false $script:staticInstallVerified $script:runtimeVerified $_.Exception.Message
    } catch { }
    throw
} finally {
    if ($transcriptStarted) {
        try { Stop-Transcript | Out-Null } catch { }
    }
}
