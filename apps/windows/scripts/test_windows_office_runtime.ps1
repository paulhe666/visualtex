[CmdletBinding()]
param(
    [ValidateSet("auto", "x86", "x64")]
    [string]$OfficePlatform = "auto",
    [string]$VisualTeXPath,
    [switch]$CompanionOnly,
    [switch]$ForceCloseOffice,
    [string]$ReportPath,
    [switch]$ArchitectureRelaunched
)

$ErrorActionPreference = "Stop"

function Quote-ProcessArgument([string]$Value) {
    return '"' + $Value.Replace('"', '\"') + '"'
}

function Resolve-ForwardedPath([string]$Value) {
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

$earlyOfficePlatform = Resolve-EarlyOfficePlatform
$requiresArchitectureRelaunch =
    ($earlyOfficePlatform -eq "x64" -and -not [Environment]::Is64BitProcess) -or
    ($earlyOfficePlatform -eq "x86" -and [Environment]::Is64BitProcess)
if ($requiresArchitectureRelaunch) {
    if ($ArchitectureRelaunched) {
        throw "Unable to relaunch Office runtime verification in a PowerShell process matching $earlyOfficePlatform Office."
    }
    $arguments = New-Object System.Collections.Generic.List[string]
    foreach ($value in @(
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        (Quote-ProcessArgument $PSCommandPath),
        "-ArchitectureRelaunched",
        "-OfficePlatform",
        $earlyOfficePlatform
    )) {
        [void]$arguments.Add($value)
    }
    if (-not [string]::IsNullOrWhiteSpace($VisualTeXPath)) {
        [void]$arguments.Add("-VisualTeXPath")
        [void]$arguments.Add((Quote-ProcessArgument (Resolve-ForwardedPath $VisualTeXPath)))
    }
    if (-not [string]::IsNullOrWhiteSpace($ReportPath)) {
        [void]$arguments.Add("-ReportPath")
        [void]$arguments.Add((Quote-ProcessArgument (Resolve-ForwardedPath $ReportPath)))
    }
    if ($CompanionOnly) { [void]$arguments.Add("-CompanionOnly") }
    if ($ForceCloseOffice) { [void]$arguments.Add("-ForceCloseOffice") }
    $process = Start-Process `
        -FilePath (Resolve-PowerShellExecutable $earlyOfficePlatform) `
        -ArgumentList ($arguments -join " ") `
        -PassThru
    try {
        # Wait only for the direct architecture-matched PowerShell child. The runtime
        # probe may start the long-lived VisualTeX companion, which would make
        # Start-Process -Wait block on the descendant process tree indefinitely.
        $process.WaitForExit()
        $exitCode = $process.ExitCode
    } finally {
        $process.Dispose()
    }
    exit $exitCode
}

Add-Type -AssemblyName System.Net.Http
$logRoot = Join-Path $env:LOCALAPPDATA "VisualTeX\office\install-logs"
New-Item -Path $logRoot -ItemType Directory -Force | Out-Null
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
if ([string]::IsNullOrWhiteSpace($ReportPath)) {
    $ReportPath = Join-Path $logRoot "office-runtime-$stamp.json"
}
$script:checks = New-Object System.Collections.Generic.List[object]
$script:failures = New-Object System.Collections.Generic.List[string]
$script:resolvedOfficePlatform = $null
$officeMsiUpgradeCode = "{A81B4BF7-0E51-45CE-A5AA-5E28F6944F42}"
$officeMsiDisplayName = "VisualTeX Windows Office Integration"

function Add-Check {
    param([string]$Name, [bool]$Passed, [string]$Details)
    [void]$script:checks.Add([pscustomobject]@{
        name = $Name
        passed = $Passed
        details = $Details
    })
    if ($Passed) { Write-Host "[PASS] $Name - $Details" }
    else {
        Write-Warning "[FAIL] $Name - $Details"
        [void]$script:failures.Add("${Name}: $Details")
    }
}

function Get-RegistryView([string]$Architecture) {
    if ($Architecture -eq "x86") { return [Microsoft.Win32.RegistryView]::Registry32 }
    return [Microsoft.Win32.RegistryView]::Registry64
}

function Get-RegistryValue {
    param(
        [Microsoft.Win32.RegistryHive]$Hive,
        [string]$SubKey,
        [string]$Name,
        [string]$Architecture
    )
    $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($Hive, (Get-RegistryView $Architecture))
    try {
        $key = $baseKey.OpenSubKey($SubKey, $false)
        if ($null -eq $key) { return $null }
        try {
            $valueName = if ($Name -eq "(default)") { "" } else { $Name }
            if ($valueName -notin @($key.GetValueNames())) { return $null }
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

function Resolve-OfficePlatform([string]$Requested) {
    $configuration = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue
    $detected = if ($configuration.Platform -in @("x86", "x64")) { [string]$configuration.Platform } else { $null }
    if ($Requested -ne "auto") {
        if ($detected -and $detected -ne $Requested) {
            throw "Requested Office platform '$Requested' does not match installed Click-to-Run Office '$detected'."
        }
        return $Requested
    }
    if ($detected) { return $detected }
    foreach ($architecture in @("x64", "x86")) {
        $word = Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE" "(default)" $architecture
        $powerPoint = Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE" "(default)" $architecture
        if ($word -and $powerPoint) { return $architecture }
    }
    throw "Unable to determine the installed Office architecture."
}

function Resolve-OfficeExecutablePath {
    param(
        [ValidateSet("Word", "PowerPoint")][string]$HostName,
        [string]$Architecture
    )
    $fileName = if ($HostName -eq "Word") { "WINWORD.EXE" } else { "POWERPNT.EXE" }
    $appPath = [string](Get-RegistryValue `
        ([Microsoft.Win32.RegistryHive]::LocalMachine) `
        "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\$fileName" `
        "(default)" `
        $Architecture)
    if (-not [string]::IsNullOrWhiteSpace($appPath) -and
        (Test-Path -LiteralPath $appPath -PathType Leaf)) {
        return [IO.Path]::GetFullPath($appPath)
    }
    $configuration = Get-ItemProperty `
        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" `
        -ErrorAction SilentlyContinue
    $clientFolder = [string]$configuration.ClientFolder
    if (-not [string]::IsNullOrWhiteSpace($clientFolder)) {
        $candidate = Join-Path $clientFolder $fileName
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return [IO.Path]::GetFullPath($candidate)
        }
    }
    foreach ($root in @($env:ProgramFiles, ${env:ProgramFiles(x86)}) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) {
        $candidate = Join-Path $root "Microsoft Office\Root\Office16\$fileName"
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return [IO.Path]::GetFullPath($candidate)
        }
    }
    throw "Unable to resolve the installed $HostName executable ($fileName)."
}

function Test-CertificateTrusted {
    $modeKey = "HKCU:\Software\VisualTeX\OfficeIntegration"
    $thumbprint = (Get-ItemProperty -LiteralPath $modeKey -Name CertificateThumbprint -ErrorAction SilentlyContinue).CertificateThumbprint
    $certificatePath = (Get-ItemProperty -LiteralPath $modeKey -Name CertificatePath -ErrorAction SilentlyContinue).CertificatePath
    if ([string]::IsNullOrWhiteSpace([string]$thumbprint) -or
        [string]::IsNullOrWhiteSpace([string]$certificatePath) -or
        -not (Test-Path -LiteralPath ([string]$certificatePath) -PathType Leaf)) {
        return [pscustomobject]@{ passed = $false; details = "Certificate registry values or certificate file are missing." }
    }
    try {
        $certificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new([string]$certificatePath)
        $store = [Security.Cryptography.X509Certificates.X509Store]::new(
            [Security.Cryptography.X509Certificates.StoreName]::Root,
            [Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
        $store.Open([Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
        try {
            $trusted = @($store.Certificates | Where-Object { $_.Thumbprint -eq $certificate.Thumbprint }).Count -gt 0
        } finally {
            $store.Close()
        }
        return [pscustomobject]@{
            passed = $trusted -and $certificate.Thumbprint -eq ([string]$thumbprint).Replace(" ", "")
            details = "$($certificate.Thumbprint); $certificatePath"
        }
    } catch {
        return [pscustomobject]@{ passed = $false; details = $_.Exception.Message }
    }
}

function Get-MsiInstalledStateOnce {
    $codes = New-Object System.Collections.Generic.HashSet[string] ([StringComparer]::OrdinalIgnoreCase)
    $sources = New-Object System.Collections.Generic.List[string]

    $installer = $null
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        $related = $installer.GetType().InvokeMember(
            "RelatedProducts",
            [Reflection.BindingFlags]::GetProperty,
            $null,
            $installer,
            @($officeMsiUpgradeCode))
        foreach ($code in @($related)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$code)) {
                [void]$codes.Add([string]$code)
                if (-not $sources.Contains("WindowsInstaller.RelatedProducts")) {
                    [void]$sources.Add("WindowsInstaller.RelatedProducts")
                }
            }
        }
    } catch {
        # The uninstall registry fallback below is intentionally independent of COM.
    } finally {
        if ($null -ne $installer) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($installer)
        }
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
            if ([string]$item.DisplayName -eq $officeMsiDisplayName) {
                [void]$codes.Add([string]$key.PSChildName)
                if (-not $sources.Contains($root)) { [void]$sources.Add($root) }
            }
        }
    }

    $productCodes = @($codes)
    return [pscustomobject]@{
        installed = $productCodes.Count -gt 0
        productCode = if ($productCodes.Count -gt 0) { [string]$productCodes[0] } else { "" }
        productCodes = $productCodes
        source = $sources -join ", "
    }
}

function Test-MsiInstalled([int]$TimeoutSeconds = 10) {
    $deadline = [DateTimeOffset]::UtcNow.AddSeconds($TimeoutSeconds)
    $lastState = Get-MsiInstalledStateOnce
    while (-not $lastState.installed -and [DateTimeOffset]::UtcNow -lt $deadline) {
        Start-Sleep -Milliseconds 250
        $lastState = Get-MsiInstalledStateOnce
    }
    return $lastState
}

function Get-LoadBehavior([ValidateSet("Word", "PowerPoint")][string]$HostName, [string]$Architecture) {
    $progId = if ($HostName -eq "Word") { "VisualTeX.WordVsto" } else { "VisualTeX.PowerPointVsto" }
    $subKey = "Software\Microsoft\Office\$HostName\Addins\$progId"
    $perUser = Get-RegistryValue ([Microsoft.Win32.RegistryHive]::CurrentUser) $subKey "LoadBehavior" $Architecture
    if ($null -ne $perUser) { return $perUser }
    return Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) $subKey "LoadBehavior" $Architecture
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
        throw "Unable to resolve Program Files for $Architecture Office runtime verification."
    }
    return Join-Path $programFilesRoot "VisualTeX\WindowsOffice\VSTO"
}

function Test-RegistryKeyInView {
    param(
        [Microsoft.Win32.RegistryHive]$Hive,
        [string]$SubKey,
        [string]$Architecture
    )
    $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($Hive, (Get-RegistryView $Architecture))
    try {
        $key = $baseKey.OpenSubKey($SubKey, $false)
        if ($null -eq $key) { return $false }
        $key.Dispose()
        return $true
    } finally {
        $baseKey.Dispose()
    }
}

function Convert-FileUriToPath([string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return "" }
    try {
        $uri = [Uri]$Value
        if ($uri.IsFile) { return [IO.Path]::GetFullPath($uri.LocalPath) }
    } catch { }
    return $Value.Trim('"')
}

function Get-ManagedComRegistrationState {
    param(
        [ValidateSet("Word", "PowerPoint")][string]$HostName,
        [string]$Architecture
    )
    $progId = if ($HostName -eq "Word") { "VisualTeX.WordVsto" } else { "VisualTeX.PowerPointVsto" }
    $clsid = if ($HostName -eq "Word") {
        "{F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1}"
    } else {
        "{7E586D2D-57B0-4D14-AB24-EBA9021A5E6D}"
    }
    $className = if ($HostName -eq "Word") {
        "VisualTeX.WordVsto.ThisAddIn"
    } else {
        "VisualTeX.PowerPointVsto.ThisAddIn"
    }
    $assemblyFile = if ($HostName -eq "Word") {
        "VisualTeX.WordVsto.dll"
    } else {
        "VisualTeX.PowerPointVsto.dll"
    }
    $installRoot = Resolve-MachineOfficeInstallRoot $Architecture
    $expectedAssemblyPath = Join-Path $installRoot $assemblyFile
    $progIdClsid = [string](Get-RegistryValue `
        ([Microsoft.Win32.RegistryHive]::LocalMachine) `
        "Software\Classes\$progId\CLSID" `
        "(default)" `
        $Architecture)
    $classKey = "Software\Classes\CLSID\$clsid"
    $inproc = [string](Get-RegistryValue `
        ([Microsoft.Win32.RegistryHive]::LocalMachine) `
        "$classKey\InprocServer32" `
        "(default)" `
        $Architecture)
    $registeredClass = [string](Get-RegistryValue `
        ([Microsoft.Win32.RegistryHive]::LocalMachine) `
        "$classKey\InprocServer32" `
        "Class" `
        $Architecture)
    $assembly = [string](Get-RegistryValue `
        ([Microsoft.Win32.RegistryHive]::LocalMachine) `
        "$classKey\InprocServer32" `
        "Assembly" `
        $Architecture)
    $codeBase = [string](Get-RegistryValue `
        ([Microsoft.Win32.RegistryHive]::LocalMachine) `
        "$classKey\InprocServer32" `
        "CodeBase" `
        $Architecture)
    $codeBasePath = Convert-FileUriToPath $codeBase
    $categoryKey = "$classKey\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
    $categoryPresent = Test-RegistryKeyInView `
        ([Microsoft.Win32.RegistryHive]::LocalMachine) `
        $categoryKey `
        $Architecture
    $legacyPerUserProgId = Test-RegistryKeyInView `
        ([Microsoft.Win32.RegistryHive]::CurrentUser) `
        "Software\Classes\$progId" `
        $Architecture
    $legacyPerUserClsid = Test-RegistryKeyInView `
        ([Microsoft.Win32.RegistryHive]::CurrentUser) `
        $classKey `
        $Architecture
    $legacyPerUserAddin = Test-RegistryKeyInView `
        ([Microsoft.Win32.RegistryHive]::CurrentUser) `
        "Software\Microsoft\Office\$HostName\Addins\$progId" `
        $Architecture
    $codeBaseMatches = $false
    if (-not [string]::IsNullOrWhiteSpace($codeBasePath)) {
        try {
            $codeBaseMatches = [string]::Equals(
                [IO.Path]::GetFullPath($codeBasePath),
                [IO.Path]::GetFullPath($expectedAssemblyPath),
                [StringComparison]::OrdinalIgnoreCase)
        } catch { $codeBaseMatches = $false }
    }
    $passed =
        [string]::Equals($progIdClsid, $clsid, [StringComparison]::OrdinalIgnoreCase) -and
        [string]::Equals($inproc, "mscoree.dll", [StringComparison]::OrdinalIgnoreCase) -and
        [string]::Equals($registeredClass, $className, [StringComparison]::Ordinal) -and
        -not [string]::IsNullOrWhiteSpace($assembly) -and
        $categoryPresent -and
        (Test-Path -LiteralPath $expectedAssemblyPath -PathType Leaf) -and
        $codeBaseMatches -and
        -not $legacyPerUserProgId -and
        -not $legacyPerUserClsid -and
        -not $legacyPerUserAddin
    return [pscustomobject]@{
        host = $HostName
        architecture = $Architecture
        progId = $progId
        clsid = $clsid
        progIdClsid = $progIdClsid
        inprocServer32 = $inproc
        className = $registeredClass
        assembly = $assembly
        codeBase = $codeBase
        codeBasePath = $codeBasePath
        expectedAssemblyPath = $expectedAssemblyPath
        implementedCategoryPresent = $categoryPresent
        codeBaseMatches = $codeBaseMatches
        legacyPerUserProgId = $legacyPerUserProgId
        legacyPerUserClsid = $legacyPerUserClsid
        legacyPerUserAddin = $legacyPerUserAddin
        passed = $passed
    }
}

function Get-DisabledItems([ValidateSet("Word", "PowerPoint")][string]$HostName) {
    $results = @()
    foreach ($bucket in @("DisabledItems", "StartupItems")) {
        $key = "HKCU:\Software\Microsoft\Office\16.0\$HostName\Resiliency\$bucket"
        if (-not (Test-Path -LiteralPath $key)) { continue }
        $item = Get-ItemProperty -LiteralPath $key -ErrorAction SilentlyContinue
        foreach ($property in $item.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' }) {
            if ($property.Value -is [byte[]]) {
                $unicode = [Text.Encoding]::Unicode.GetString([byte[]]$property.Value).Trim([char]0)
                $ascii = [Text.Encoding]::ASCII.GetString([byte[]]$property.Value).Trim([char]0)
                $results += [pscustomobject]@{
                    bucket = $bucket
                    name = $property.Name
                    unicode = $unicode
                    ascii = $ascii
                    mentionsVisualTeX = ($unicode -match 'VisualTeX' -or $ascii -match 'VisualTeX')
                }
            }
        }
    }
    return @($results)
}

function Get-RecentOfficeLoadEvents {
    $start = (Get-Date).AddMinutes(-15)
    try {
        return @(Get-WinEvent -FilterHashtable @{ LogName = "Application"; StartTime = $start } -MaxEvents 150 -ErrorAction Stop |
            Where-Object {
                $_.ProviderName -match 'VSTO|\.NET Runtime|Office|Application Error' -or
                $_.Message -match 'VisualTeX|VisualTeX\.WordVsto|VisualTeX\.PowerPointVsto'
            } |
            Select-Object -First 30 TimeCreated, ProviderName, Id, LevelDisplayName, Message)
    } catch {
        return @([pscustomobject]@{
            TimeCreated = Get-Date
            ProviderName = "Get-WinEvent"
            Id = 0
            LevelDisplayName = "Warning"
            Message = $_.Exception.Message
        })
    }
}

function Test-TcpPort {
    param(
        [string]$HostName,
        [int]$Port,
        [int]$TimeoutMilliseconds = 2000
    )
    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $async = $client.BeginConnect($HostName, $Port, $null, $null)
        if (-not $async.AsyncWaitHandle.WaitOne($TimeoutMilliseconds, $false)) {
            return $false
        }
        $client.EndConnect($async)
        return $true
    } catch {
        return $false
    } finally {
        $client.Close()
    }
}

function Get-RemoteCertificateThumbprint {
    param([string]$HostName, [int]$Port)
    $client = New-Object System.Net.Sockets.TcpClient
    $stream = $null
    try {
        $client.Connect($HostName, $Port)
        $callback = [Net.Security.RemoteCertificateValidationCallback]{
            param($sender, $certificate, $chain, $sslPolicyErrors)
            return $true
        }
        $stream = New-Object System.Net.Security.SslStream($client.GetStream(), $false, $callback)
        $stream.AuthenticateAsClient("localhost")
        $remoteCertificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new($stream.RemoteCertificate)
        return $remoteCertificate.Thumbprint
    } finally {
        if ($null -ne $stream) { $stream.Dispose() }
        $client.Close()
    }
}

function Get-ExceptionChain([Exception]$Exception) {
    $items = @()
    for ($current = $Exception; $null -ne $current; $current = $current.InnerException) {
        $items += [pscustomobject]@{
            type = $current.GetType().FullName
            message = $current.Message
            hresult = ('0x{0:X8}' -f $current.HResult)
            stackTrace = $current.StackTrace
        }
    }
    return @($items)
}

function Get-ProcessIdentity {
    param(
        [int]$ProcessId,
        [int]$PathRetryMilliseconds = 3000
    )
    $deadline = [DateTimeOffset]::UtcNow.AddMilliseconds($PathRetryMilliseconds)
    $processName = ""
    $processPath = ""
    do {
        $process = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
        if ($null -eq $process) { break }
        $processName = [string]$process.ProcessName
        try { $processPath = [string]$process.Path } catch { $processPath = "" }
        if ([string]::IsNullOrWhiteSpace($processPath)) {
            try {
                $cimProcess = Get-CimInstance Win32_Process -Filter "ProcessId = $ProcessId" -ErrorAction SilentlyContinue
                $processPath = [string]$cimProcess.ExecutablePath
            } catch { $processPath = "" }
        }
        if (-not [string]::IsNullOrWhiteSpace($processPath)) { break }
        Start-Sleep -Milliseconds 100
    } while ([DateTimeOffset]::UtcNow -lt $deadline)
    return [pscustomobject]@{
        processId = $ProcessId
        processName = $processName
        processPath = $processPath
    }
}

function Get-PortOwner([int]$Port, [int]$PathRetryMilliseconds = 3000) {
    $processId = $null
    $source = "none"
    try {
        $connection = Get-NetTCPConnection -State Listen -LocalPort $Port -ErrorAction Stop |
            Select-Object -First 1
        if ($null -ne $connection) {
            $processId = [int]$connection.OwningProcess
            $source = "Get-NetTCPConnection"
        }
    } catch { }

    if ($null -eq $processId) {
        # Get-NetTCPConnection is unavailable on some machines (a broken NetTCPIP
        # module, or a PSModulePath inherited from PowerShell 7). Ask CIM directly.
        try {
            $connection = Get-CimInstance -Namespace root/StandardCimv2 -ClassName MSFT_NetTCPConnection -Filter "LocalPort = $Port AND State = 2" -ErrorAction Stop |
                Select-Object -First 1
            if ($null -ne $connection) {
                $processId = [int]$connection.OwningProcess
                $source = "MSFT_NetTCPConnection"
            }
        } catch { }
    }

    if ($null -eq $processId) {
        try {
            # Full path (PATH may lack System32), and no state text: netstat
            # localizes it. A listener has no remote endpoint (0.0.0.0:0 / [::]:0).
            $netstat = Join-Path $env:SystemRoot "System32\netstat.exe"
            foreach ($line in (& $netstat -ano -p tcp 2>$null)) {
                $columns = @($line.Trim() -split '\s+')
                if ($columns.Count -lt 5 -or $columns[0] -ne 'TCP' -or $columns[1] -notmatch ":$Port$") { continue }
                if ($columns[2] -notmatch '^(0\.0\.0\.0|\[::\]):0$' -and $line -notmatch 'LISTENING') { continue }
                $processId = [int]$columns[-1]
                $source = "netstat"
                break
            }
        } catch { }
    }

    if ($null -ne $processId) {
        $identity = Get-ProcessIdentity $processId $PathRetryMilliseconds
        return [pscustomobject]@{
            processId = $identity.processId
            processName = $identity.processName
            processPath = $identity.processPath
            source = $source
        }
    }
    return [pscustomobject]@{
        processId = $null
        processName = ""
        processPath = ""
        source = "none"
    }
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Start-CompanionAsInteractiveUser([string]$Executable) {
    if (Test-IsAdministrator) {
        throw "Companion runtime verification must run in the interactive user's non-elevated session. The machine-wide installer must return before starting VisualTeX."
    }
    return Start-Process -FilePath $Executable -ArgumentList "--office-background" -WindowStyle Hidden -PassThru
}

# CompanionPort is the standard port, which install.json names too. When it is taken
# or reserved, VisualTeX listens on another one and writes the port it bound to
# CompanionActivePort at every start.
function Get-CompanionListeningPort([int]$RegisteredPort) {
    $active = 0
    try {
        $key = Get-Item -LiteralPath "HKCU:\Software\VisualTeX\OfficeIntegration" -ErrorAction SilentlyContinue
        if ($null -ne $key) { $active = [int]$key.GetValue("CompanionActivePort", 0) }
    } catch { $active = 0 }
    if ($active -gt 0 -and $active -le 65535) { return $active }
    return $RegisteredPort
}

function Test-CompanionRuntime {
    param(
        [object]$IntegrationState,
        [string]$ExplicitVisualTeXPath
    )

    $errors = New-Object System.Collections.Generic.List[string]
    $stage = "configuration"
    $registeredExecutable = [string]$IntegrationState.ExecutablePath
    $executable = if ([string]::IsNullOrWhiteSpace($ExplicitVisualTeXPath)) {
        $registeredExecutable
    } else {
        $ExplicitVisualTeXPath.Trim().Trim('"')
    }
    if (-not [string]::IsNullOrWhiteSpace($executable) -and (Test-Path -LiteralPath $executable -PathType Leaf)) {
        $executable = (Resolve-Path -LiteralPath $executable).Path
    }
    if (-not [string]::IsNullOrWhiteSpace($registeredExecutable) -and (Test-Path -LiteralPath $registeredExecutable -PathType Leaf)) {
        $registeredExecutable = (Resolve-Path -LiteralPath $registeredExecutable).Path
    }
    $appDataRoot = [string]$IntegrationState.AppDataRoot
    $certificatePath = [string]$IntegrationState.CertificatePath
    $expectedThumbprint = ([string]$IntegrationState.CertificateThumbprint).Replace(" ", "")
    $registeredPort = [int]$IntegrationState.CompanionPort
    $port = Get-CompanionListeningPort $registeredPort
    $foreignPort = 0
    $foreignPortError = ""
    $expectedProtocol = [int]$IntegrationState.ProtocolVersion
    $installJsonPath = Join-Path $appDataRoot "office\install.json"
    $companionLogRoot = Join-Path $env:LOCALAPPDATA "VisualTeX\office\logs"
    $startupLog = Join-Path $companionLogRoot "startup.log"
    $companionLog = Join-Path $companionLogRoot "companion.log"
    $startedProcessId = $null
    $startedProcessExitCode = $null
    $processRunning = $false
    $matchingProcess = $null
    $portListening = $false
    $portOwner = [pscustomobject]@{ processId = $null; processName = ""; processPath = ""; source = "none" }
    $httpsHealthy = $false
    $protocolMatches = $false
    $certificateMatches = $false
    $localThumbprint = ""
    $remoteThumbprint = ""
    $tlsPolicyErrors = ""
    $healthRaw = ""
    $health = $null
    $exceptionChain = @()

    if ([string]::IsNullOrWhiteSpace($executable)) {
        [void]$errors.Add("ExecutablePath is missing. Supply -VisualTeXPath and repair the integration registry state.")
    } elseif (-not (Test-Path -LiteralPath $executable -PathType Leaf)) {
        [void]$errors.Add("The exact VisualTeX executable is missing: $executable")
    }
    if (-not [string]::IsNullOrWhiteSpace($ExplicitVisualTeXPath) -and
        -not [string]::IsNullOrWhiteSpace($registeredExecutable) -and
        -not [string]::Equals(
            [IO.Path]::GetFullPath($executable),
            [IO.Path]::GetFullPath($registeredExecutable),
            [StringComparison]::OrdinalIgnoreCase)) {
        [void]$errors.Add("Explicit VisualTeXPath '$executable' does not match registered ExecutablePath '$registeredExecutable'.")
    }
    if ([string]::IsNullOrWhiteSpace($appDataRoot)) {
        [void]$errors.Add("AppDataRoot is missing from HKCU\Software\VisualTeX\OfficeIntegration.")
    } elseif (-not (Test-Path -LiteralPath $appDataRoot -PathType Container)) {
        [void]$errors.Add("The registered AppDataRoot does not exist: $appDataRoot")
    }
    if ($registeredPort -le 0 -or $registeredPort -gt 65535) {
        [void]$errors.Add("CompanionPort is missing or invalid: $registeredPort")
    }
    if ($expectedProtocol -le 0) {
        [void]$errors.Add("ProtocolVersion is missing or invalid: $expectedProtocol")
    }
    if (-not (Test-Path -LiteralPath $installJsonPath -PathType Leaf)) {
        [void]$errors.Add("install.json is missing: $installJsonPath")
    } else {
        try {
            $installState = Get-Content -LiteralPath $installJsonPath -Raw | ConvertFrom-Json
            if ([string]::IsNullOrWhiteSpace([string]$installState.installToken) -or
                [string]$installState.installToken.Length -ne 64) {
                [void]$errors.Add("install.json contains an invalid installToken.")
            }
            if ([int]$installState.port -ne $registeredPort) {
                [void]$errors.Add("install.json port=$($installState.port) does not match registry CompanionPort=$registeredPort.")
            }
            if ([int]$installState.protocolVersion -ne $expectedProtocol) {
                [void]$errors.Add("install.json protocolVersion=$($installState.protocolVersion) does not match registry ProtocolVersion=$expectedProtocol.")
            }
        } catch {
            $exceptionChain = @(Get-ExceptionChain $_.Exception)
            [void]$errors.Add("install.json is invalid: $($_.Exception.Message)")
        }
    }

    if ($errors.Count -eq 0) {
        # Inspect the companion port before starting VisualTeX. A foreign
        # listener is not a failure yet: VisualTeX then binds another port and
        # names it. It is reported if VisualTeX does not, instead of a generic
        # background-process timeout or early-exit failure.
        $portListening = Test-TcpPort "127.0.0.1" $port 750
        if ($portListening) {
            $portOwner = Get-PortOwner $port
            if (-not [string]::IsNullOrWhiteSpace([string]$portOwner.processPath) -and
                -not [string]::Equals(
                    [IO.Path]::GetFullPath([string]$portOwner.processPath),
                    [IO.Path]::GetFullPath($executable),
                    [StringComparison]::OrdinalIgnoreCase)) {
                $foreignPort = $port
                $foreignPortError = "Port $port is already occupied by '$($portOwner.processPath)' PID=$($portOwner.processId) process='$($portOwner.processName)', not by '$executable'."
                $portListening = $false
            } elseif ($null -ne $portOwner.processId) {
                $matchingProcess = Get-Process -Id ([int]$portOwner.processId) -ErrorAction SilentlyContinue
            }
        }
    }

    if ($errors.Count -eq 0 -and -not $portListening) {
        $stage = "process-start"
        foreach ($process in @(Get-Process -ErrorAction SilentlyContinue)) {
            $processPath = $null
            try { $processPath = $process.Path } catch { continue }
            if (-not [string]::IsNullOrWhiteSpace([string]$processPath) -and
                [string]::Equals(
                    [IO.Path]::GetFullPath([string]$processPath),
                    [IO.Path]::GetFullPath($executable),
                    [StringComparison]::OrdinalIgnoreCase)) {
                $matchingProcess = $process
                break
            }
        }
        if ($null -eq $matchingProcess) {
            try {
                $matchingProcess = Start-CompanionAsInteractiveUser $executable
                if ($null -eq $matchingProcess) {
                    throw "The interactive user shell did not expose the started VisualTeX companion process within 8 seconds."
                }
                $startedProcessId = $matchingProcess.Id
            } catch {
                $exceptionChain = @(Get-ExceptionChain $_.Exception)
                [void]$errors.Add("Unable to start the exact VisualTeX executable '$executable': $($_.Exception.Message)")
            }
        }
    }

    if ($errors.Count -eq 0 -and -not $portListening) {
        $stage = "port-listen"
        $deadline = [DateTimeOffset]::UtcNow.AddSeconds(20)
        do {
            if ($null -ne $matchingProcess -and $matchingProcess.HasExited) {
                $startedProcessExitCode = $matchingProcess.ExitCode
                [void]$errors.Add("VisualTeX background PID=$($matchingProcess.Id) exited with code $startedProcessExitCode before the companion became ready.")
                break
            }
            # The port is known once VisualTeX has bound it; the other program's
            # port is not VisualTeX's.
            $port = Get-CompanionListeningPort $registeredPort
            $portListening = ($port -ne $foreignPort) -and (Test-TcpPort "127.0.0.1" $port 750)
            if ($portListening) { break }
            Start-Sleep -Milliseconds 250
        } while ([DateTimeOffset]::UtcNow -lt $deadline)
        if (-not $portListening -and $foreignPort -gt 0 -and $errors.Count -eq 0) {
            $stage = "port-conflict"
            [void]$errors.Add("$foreignPortError VisualTeX did not start its companion on another port within 20 seconds. Check '$startupLog' and '$companionLog'.")
        } elseif (-not $portListening) {
            [void]$errors.Add("No process listened on 127.0.0.1:$port within 20 seconds. Check '$startupLog' and '$companionLog'.")
        }
    }

    if ($portListening -and $errors.Count -eq 0) {
        $portOwner = Get-PortOwner $port
        if (-not [string]::IsNullOrWhiteSpace([string]$portOwner.processPath) -and
            -not [string]::Equals(
                [IO.Path]::GetFullPath([string]$portOwner.processPath),
                [IO.Path]::GetFullPath($executable),
                [StringComparison]::OrdinalIgnoreCase)) {
            $stage = "port-conflict"
            [void]$errors.Add("Port $port is owned by '$($portOwner.processPath)' PID=$($portOwner.processId), not by '$executable'.")
        } elseif ($null -eq $matchingProcess -and $null -ne $portOwner.processId) {
            $matchingProcess = Get-Process -Id ([int]$portOwner.processId) -ErrorAction SilentlyContinue
        }
    }
    if ($portListening -and $null -eq $matchingProcess) {
        # Windows did not name the port's owner: find the running VisualTeX by path.
        foreach ($process in @(Get-Process -Name visualtex -ErrorAction SilentlyContinue)) {
            $processPath = $null
            try { $processPath = $process.Path } catch { continue }
            if (-not [string]::IsNullOrWhiteSpace([string]$processPath) -and
                [string]::Equals(
                    [IO.Path]::GetFullPath([string]$processPath),
                    [IO.Path]::GetFullPath($executable),
                    [StringComparison]::OrdinalIgnoreCase)) {
                $matchingProcess = $process
                break
            }
        }
    }
    if ($null -ne $matchingProcess -and -not $matchingProcess.HasExited) {
        $processRunning = $true
    }
    $actualProcessId = if ($null -ne $matchingProcess) { $matchingProcess.Id } else { $portOwner.processId }

    if ($portListening -and $errors.Count -eq 0) {
        $stage = "https-handshake"
        $handler = [System.Net.Http.HttpClientHandler]::new()
        $handler.UseProxy = $false
        $handler.Proxy = $null
        $client = [System.Net.Http.HttpClient]::new($handler)
        $client.Timeout = [TimeSpan]::FromSeconds(5)
        try {
            # Read the actual server certificate through a separate TLS probe.
            # The HttpClient request itself intentionally uses the Windows trust
            # chain so an untrusted current-user Root certificate fails normally.
            $remoteThumbprint = (Get-RemoteCertificateThumbprint "127.0.0.1" $port).Replace(" ", "")
            $response = $client.GetAsync("https://127.0.0.1:$port/health").GetAwaiter().GetResult()
            $healthRaw = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            $tlsPolicyErrors = "None"
            if (-not $response.IsSuccessStatusCode) {
                [void]$errors.Add("The companion /health endpoint returned HTTP $([int]$response.StatusCode): $healthRaw")
            } else {
                $health = $healthRaw | ConvertFrom-Json
                $httpsHealthy = [bool]$health.ok
                $protocolMatches = [int]$health.protocolVersion -eq $expectedProtocol
                if (-not $httpsHealthy) {
                    [void]$errors.Add("The companion HTTPS /health endpoint returned ok=false: $healthRaw")
                }
                if (-not $protocolMatches) {
                    [void]$errors.Add("Companion protocolVersion=$($health.protocolVersion) does not match expected $expectedProtocol.")
                }
            }
        } catch {
            $exceptionChain = @(Get-ExceptionChain $_.Exception)
            if ([string]::IsNullOrWhiteSpace($tlsPolicyErrors)) { $tlsPolicyErrors = "Windows trust-chain validation failed or the TLS handshake did not complete." }
            [void]$errors.Add("HTTPS /health failed with proxy disabled: $($_.Exception.Message); TLS=$tlsPolicyErrors")
        } finally {
            $client.Dispose()
            $handler.Dispose()
        }
    }

    if ($portListening -and $errors.Count -eq 0) {
        $stage = "certificate-match"
        if ([string]::IsNullOrWhiteSpace($expectedThumbprint)) {
            [void]$errors.Add("CertificateThumbprint is missing from HKCU\Software\VisualTeX\OfficeIntegration.")
        } elseif ([string]::IsNullOrWhiteSpace($certificatePath) -or
            -not (Test-Path -LiteralPath $certificatePath -PathType Leaf)) {
            [void]$errors.Add("The registered companion certificate file is missing: $certificatePath")
        } else {
            try {
                $localCertificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new($certificatePath)
                $localThumbprint = $localCertificate.Thumbprint.Replace(" ", "")
                if ([string]::IsNullOrWhiteSpace($remoteThumbprint)) {
                    $remoteThumbprint = (Get-RemoteCertificateThumbprint "127.0.0.1" $port).Replace(" ", "")
                }
                $certificateMatches =
                    [string]::Equals($localThumbprint, $expectedThumbprint, [StringComparison]::OrdinalIgnoreCase) -and
                    [string]::Equals($remoteThumbprint, $expectedThumbprint, [StringComparison]::OrdinalIgnoreCase)
                if (-not $certificateMatches) {
                    [void]$errors.Add("Certificate mismatch: registry=$expectedThumbprint; file=$localThumbprint; server=$remoteThumbprint")
                }
            } catch {
                $exceptionChain = @(Get-ExceptionChain $_.Exception)
                [void]$errors.Add("Unable to inspect the companion certificate: $($_.Exception.Message)")
            }
        }
    }

    if ($protocolMatches -and $certificateMatches -and $errors.Count -eq 0) { $stage = "complete" }
    return [pscustomobject]@{
        stage = $stage
        executable = $executable
        registeredExecutable = $registeredExecutable
        appDataRoot = $appDataRoot
        installJsonPath = $installJsonPath
        startupLog = $startupLog
        companionLog = $companionLog
        port = $port
        expectedProtocol = $expectedProtocol
        processId = $actualProcessId
        startedProcessId = $startedProcessId
        startedProcessExitCode = $startedProcessExitCode
        processRunning = $processRunning
        portListening = $portListening
        portOwner = $portOwner
        httpsHealthy = $httpsHealthy
        certificateMatches = $certificateMatches
        protocolMatches = $protocolMatches
        certificatePath = $certificatePath
        expectedThumbprint = $expectedThumbprint
        localThumbprint = $localThumbprint
        remoteThumbprint = $remoteThumbprint
        tlsPolicyErrors = $tlsPolicyErrors
        healthRaw = $healthRaw
        health = $health
        exceptionChain = @($exceptionChain)
        errors = @($errors)
    }
}

function Test-ManagedComActivation([string]$ProgId) {
    $instance = $null
    try {
        $type = [Type]::GetTypeFromProgID($ProgId, $true)
        $instance = [Activator]::CreateInstance($type)
        if ($null -eq $instance) { throw "CoCreateInstance returned null." }
        return [pscustomobject]@{
            progId = $ProgId
            passed = $true
            clsid = [string]$type.GUID
            error = ""
            hresult = "0x00000000"
        }
    } catch {
        $errorRecord = $_
        $exception = $errorRecord.Exception
        return [pscustomobject]@{
            progId = $ProgId
            passed = $false
            clsid = ""
            error = $exception.ToString()
            hresult = ('0x{0:X8}' -f ([uint32]$exception.HResult))
        }
    } finally {
        if ($null -ne $instance -and [Runtime.InteropServices.Marshal]::IsComObject($instance)) {
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($instance) } catch { }
        }
    }
}

function Test-OleLocalServer([string]$Architecture) {
    $clsid = "{8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B}"
    $subKey = "Software\Classes\CLSID\$clsid\LocalServer32"
    $server = [string](Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) $subKey "ServerExecutable" $Architecture)
    if ([string]::IsNullOrWhiteSpace($server)) {
        $server = [string](Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) $subKey "(default)" $Architecture)
        $server = $server.Trim('"')
    }
    if ([string]::IsNullOrWhiteSpace($server) -or
        -not (Test-Path -LiteralPath $server -PathType Leaf)) {
        return [pscustomobject]@{
            healthy = $false
            server = $server
            clsid = $clsid
            embeddingProbe = $false
            error = "The registered OLE LocalServer executable is missing."
        }
    }

    $process = $null
    try {
        $process = Start-Process -FilePath $server -ArgumentList "-Embedding" -WindowStyle Hidden -PassThru
        Start-Sleep -Milliseconds 900
        if ($process.HasExited -and $process.ExitCode -ne 0) {
            return [pscustomobject]@{
                healthy = $false
                server = $server
                clsid = $clsid
                embeddingProbe = $false
                error = "The OLE LocalServer exited with code $($process.ExitCode) during the -Embedding probe."
            }
        }
        return [pscustomobject]@{
            healthy = $true
            server = $server
            clsid = $clsid
            embeddingProbe = $true
            error = ""
        }
    } catch {
        return [pscustomobject]@{
            healthy = $false
            server = $server
            clsid = $clsid
            embeddingProbe = $false
            error = $_.Exception.ToString()
        }
    } finally {
        if ($null -ne $process -and -not $process.HasExited) {
            Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
        }
    }
}

function Release-ComObject([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) } catch { }
    }
}

function Get-ComAddInItem([object]$Collection, [object]$Index) {
    if ($null -eq $Collection) { return $null }
    return $Collection.Item($Index)
}

function Test-OfficeComAddIn {
    param(
        [ValidateSet("Word", "PowerPoint")][string]$HostName,
        [string]$ProgId,
        [int]$StartupTimeoutSeconds = 15
    )
    $application = $null
    $startedProcess = $null
    $officeExecutable = ""
    $addIns = $null
    $addIn = $null
    $stage = "resolve-office-executable"
    $inventory = @()
    $connectAttempted = $false
    try {
        $officeExecutable = Resolve-OfficeExecutablePath $HostName $script:resolvedOfficePlatform
        $rotProgId = if ($HostName -eq "Word") { "Word.Application" } else { "PowerPoint.Application" }
        $stage = "start-desktop-office"
        $startedProcess = Start-Process -FilePath $officeExecutable -PassThru
        if ($null -eq $startedProcess) {
            throw "$HostName desktop process did not start."
        }

        $stage = "attach-running-office"
        $attachDeadline = [DateTimeOffset]::UtcNow.AddSeconds($StartupTimeoutSeconds)
        do {
            Start-Sleep -Milliseconds 300
            try {
                $application = [Runtime.InteropServices.Marshal]::GetActiveObject($rotProgId)
            } catch { $application = $null }
        } while ($null -eq $application -and [DateTimeOffset]::UtcNow -lt $attachDeadline)
        if ($null -eq $application) {
            throw "The normally started $HostName desktop application did not register '$rotProgId' in the Running Object Table within $StartupTimeoutSeconds seconds."
        }

        $stage = "enumerate-com-addins"
        $deadline = [DateTimeOffset]::UtcNow.AddSeconds($StartupTimeoutSeconds)
        do {
            Release-ComObject $addIn
            $addIn = $null
            Release-ComObject $addIns
            $addIns = $application.COMAddIns
            $inventory = @()
            if ($null -ne $addIns) {
                $count = 0
                try { $count = [int]$addIns.Count } catch { $count = 0 }
                for ($index = 1; $index -le $count; $index++) {
                    $candidate = $null
                    try {
                        $candidate = Get-ComAddInItem $addIns $index
                        if ($null -eq $candidate) { continue }
                        $candidateProgId = [string]$candidate.ProgId
                        $candidateDescription = [string]$candidate.Description
                        $candidateConnected = $false
                        try { $candidateConnected = [bool]$candidate.Connect } catch { }
                        $inventory += [pscustomobject]@{
                            index = $index
                            progId = $candidateProgId
                            description = $candidateDescription
                            connected = $candidateConnected
                        }
                        if ([string]::Equals($candidateProgId, $ProgId, [StringComparison]::OrdinalIgnoreCase)) {
                            $addIn = $candidate
                            $candidate = $null
                        }
                    } catch {
                        $inventory += [pscustomobject]@{
                            index = $index
                            progId = ""
                            description = ""
                            connected = $false
                            enumerationError = $_.Exception.Message
                        }
                    } finally {
                        Release-ComObject $candidate
                    }
                }
            }
            if ($null -ne $addIn) { break }
            Start-Sleep -Milliseconds 300
        } while ([DateTimeOffset]::UtcNow -lt $deadline)

        if ($null -eq $addIn) {
            $discovered = @($inventory | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.progId) } | ForEach-Object { $_.progId })
            return [pscustomobject]@{
                host = $HostName
                progId = $ProgId
                enumerated = $false
                connected = $false
                connectAttempted = $false
                description = ""
                discoveredProgIds = $discovered
                inventory = $inventory
                stage = $stage
                startupMode = "desktop-executable-rot"
                executable = $officeExecutable
                processId = if ($null -ne $startedProcess) { $startedProcess.Id } else { $null }
                error = "The normally started $HostName desktop application did not enumerate '$ProgId' within $StartupTimeoutSeconds seconds. Discovered COM add-ins: $($discovered -join ', ')"
            }
        }

        $stage = "read-connect"
        $connected = [bool]$addIn.Connect
        if (-not $connected) {
            $stage = "connect-addin"
            $connectAttempted = $true
            $addIn.Connect = $true
            $connectDeadline = [DateTimeOffset]::UtcNow.AddSeconds(5)
            do {
                Start-Sleep -Milliseconds 200
                $connected = [bool]$addIn.Connect
            } while (-not $connected -and [DateTimeOffset]::UtcNow -lt $connectDeadline)
        }
        return [pscustomobject]@{
            host = $HostName
            progId = $ProgId
            enumerated = $true
            connected = $connected
            connectAttempted = $connectAttempted
            description = [string]$addIn.Description
            discoveredProgIds = @($inventory | ForEach-Object { $_.progId })
            inventory = $inventory
            stage = "complete"
            startupMode = "desktop-executable-rot"
            executable = $officeExecutable
            processId = if ($null -ne $startedProcess) { $startedProcess.Id } else { $null }
            error = if ($connected) { "" } else { "Office enumerated '$ProgId' but COMAddIn.Connect remained false." }
        }
    } catch {
        return [pscustomobject]@{
            host = $HostName
            progId = $ProgId
            enumerated = $null -ne $addIn
            connected = $false
            connectAttempted = $connectAttempted
            description = if ($null -ne $addIn) { [string]$addIn.Description } else { "" }
            discoveredProgIds = @($inventory | ForEach-Object { $_.progId })
            inventory = $inventory
            stage = $stage
            startupMode = "desktop-executable-rot"
            executable = $officeExecutable
            processId = if ($null -ne $startedProcess) { $startedProcess.Id } else { $null }
            error = $_.Exception.ToString()
        }
    } finally {
        if ($null -ne $application) {
            try { $application.Quit() } catch { }
        }
        Release-ComObject $addIn
        Release-ComObject $addIns
        Release-ComObject $application
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        if ($null -ne $startedProcess) {
            try {
                if (-not $startedProcess.HasExited -and -not $startedProcess.WaitForExit(3000)) {
                    Stop-Process -Id $startedProcess.Id -Force -ErrorAction SilentlyContinue
                }
            } catch { }
            try { $startedProcess.Dispose() } catch { }
        }
    }
}

function Write-Report {
    param(
        [object]$WordResult,
        [object]$PowerPointResult,
        [object]$CertificateState,
        [object]$MsiState,
        [object]$OleState,
        [object]$CompanionState,
        [object]$WordComRegistration,
        [object]$PowerPointComRegistration,
        [object[]]$WordDisabledItems,
        [object[]]$PowerPointDisabledItems,
        [object[]]$Events
    )
    $report = [ordered]@{
        schemaVersion = 1
        generatedAt = [DateTimeOffset]::Now.ToString("o")
        officePlatform = $script:resolvedOfficePlatform
        succeeded = $script:failures.Count -eq 0
        checks = @($script:checks | ForEach-Object { $_ })
        word = $WordResult
        powerPoint = $PowerPointResult
        certificate = $CertificateState
        msi = $MsiState
        oleLocalServer = $OleState
        companion = $CompanionState
        wordComRegistration = $WordComRegistration
        powerPointComRegistration = $PowerPointComRegistration
        wordDisabledItems = $WordDisabledItems
        powerPointDisabledItems = $PowerPointDisabledItems
        recentLoadEvents = $Events
        failures = @($script:failures | ForEach-Object { $_ })
    }
    $report | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $ReportPath -Encoding UTF8
    Write-Host "Runtime diagnostic report: $ReportPath"
}

$wordResult = $null
$powerPointResult = $null
$certificateState = $null
$msiState = $null
$oleState = $null
$companionState = $null
$wordComRegistration = $null
$powerPointComRegistration = $null
$wordDisabledItems = @()
$powerPointDisabledItems = @()
$events = @()

try {
    $running = @(Get-Process WINWORD, POWERPNT -ErrorAction SilentlyContinue)
    if ($running.Count -gt 0) {
        if (-not $ForceCloseOffice) {
            throw "Word and PowerPoint must both be closed before runtime verification. Save your documents and close Office, or retry with the force-close option. Running: $($running.ProcessName -join ', ')"
        }
        $runningNames = @($running.ProcessName | Sort-Object -Unique)
        foreach ($process in $running) {
            Stop-Process -Id $process.Id -Force -ErrorAction Stop
        }
        $deadline = [DateTime]::UtcNow.AddSeconds(10)
        do {
            Start-Sleep -Milliseconds 200
            $stillRunning = @(Get-Process WINWORD, POWERPNT -ErrorAction SilentlyContinue)
        } while ($stillRunning.Count -gt 0 -and [DateTime]::UtcNow -lt $deadline)
        if ($stillRunning.Count -gt 0) {
            throw "Unable to force-close all Office applications before runtime verification. Still running: $($stillRunning.ProcessName -join ', ')"
        }
        Add-Check "Word and PowerPoint closed" $true "Force-closed running Office applications: $($runningNames -join ', ')."
    } else {
        Add-Check "Word and PowerPoint closed" $true "No existing WINWORD.EXE or POWERPNT.EXE process was found."
    }

    $script:resolvedOfficePlatform = Resolve-OfficePlatform $OfficePlatform
    Add-Check "Office platform" $true $script:resolvedOfficePlatform

    $modeKey = "HKCU:\Software\VisualTeX\OfficeIntegration"
    $installationState = Get-ItemProperty -LiteralPath $modeKey -ErrorAction SilentlyContinue
    $filesAndRegistryVerified = $null -ne $installationState -and
        [int]$installationState.FilesAndRegistryVerified -eq 1 -and
        [string]$installationState.Mode -eq "vsto"
    Add-Check "Files and registry installation stage" $filesAndRegistryVerified "Mode=$($installationState.Mode); FilesAndRegistryVerified=$($installationState.FilesAndRegistryVerified)"

    $certificateState = Test-CertificateTrusted
    Add-Check "Current-user HTTPS certificate" ([bool]$certificateState.passed) ([string]$certificateState.details)

    $companionState = Test-CompanionRuntime $installationState $VisualTeXPath
    $portOwnerPathMatches =
        -not [string]::IsNullOrWhiteSpace([string]$companionState.portOwner.processPath) -and
        [string]::Equals(
            [IO.Path]::GetFullPath([string]$companionState.portOwner.processPath),
            [IO.Path]::GetFullPath([string]$companionState.executable),
            [StringComparison]::OrdinalIgnoreCase)
    $portOwnerPidMatchesVerifiedProcess =
        $null -ne $companionState.processId -and
        $null -ne $companionState.portOwner.processId -and
        [int]$companionState.processId -eq [int]$companionState.portOwner.processId -and
        [bool]$companionState.processRunning
    # The server presenting VisualTeX's own certificate and protocol over HTTPS is
    # VisualTeX, even when Windows will not name the port's owner process.
    $companionProvenByCertificate =
        [bool]$companionState.portListening -and
        $null -eq $companionState.portOwner.processId -and
        [bool]$companionState.httpsHealthy -and
        [bool]$companionState.certificateMatches -and
        [bool]$companionState.protocolMatches
    $portOwnedByVisualTeX =
        [bool]$companionState.portListening -and
        ($portOwnerPathMatches -or $portOwnerPidMatchesVerifiedProcess -or $companionProvenByCertificate)
    $processVerified = [bool]$companionState.processRunning -or $companionProvenByCertificate
    Add-Check "Companion process" $processVerified "Executable=$($companionState.executable); PID=$($companionState.processId); StartedPID=$($companionState.startedProcessId); ExitCode=$($companionState.startedProcessExitCode); AppDataRoot=$($companionState.appDataRoot)$(if ($companionProvenByCertificate -and -not $companionState.processRunning) { '; process not visible to this session, verified by its HTTPS certificate' })"
    Add-Check "Companion TCP port" $portOwnedByVisualTeX "127.0.0.1:$($companionState.port); Owner=$($companionState.portOwner.processName) PID=$($companionState.portOwner.processId) Path=$($companionState.portOwner.processPath)$(if ($companionProvenByCertificate) { '; owner lookup unavailable, VisualTeX verified by its HTTPS certificate and protocol' })"
    Add-Check "Companion HTTPS health" ([bool]$companionState.httpsHealthy) $(if ($companionState.httpsHealthy) { "ok=true; Raw=$($companionState.healthRaw)" } else { "Stage=$($companionState.stage); Errors=$($companionState.errors -join '; '); TLS=$($companionState.tlsPolicyErrors); Exceptions=$($companionState.exceptionChain | ConvertTo-Json -Depth 5 -Compress)" })
    if ([string]$companionState.stage -eq "port-conflict") {
        Write-Host "[SKIP] Companion certificate and protocol - port ownership must be corrected before TLS validation."
    } else {
        Add-Check "Companion certificate" ([bool]$companionState.certificateMatches) "Registry=$($companionState.expectedThumbprint); File=$($companionState.localThumbprint); Server=$($companionState.remoteThumbprint); TLS=$($companionState.tlsPolicyErrors)"
        Add-Check "Companion protocol" ([bool]$companionState.protocolMatches) "Expected=$($companionState.expectedProtocol); Actual=$($companionState.health.protocolVersion); Response=$($companionState.healthRaw)"
    }

    $msiState = Test-MsiInstalled 10
    Add-Check "Office MSI installed" ([bool]$msiState.installed) "ProductCode=$($msiState.productCode); Source=$($msiState.source)"

    $wordLoadBehaviorBefore = Get-LoadBehavior Word $script:resolvedOfficePlatform
    $powerPointLoadBehaviorBefore = Get-LoadBehavior PowerPoint $script:resolvedOfficePlatform
    Add-Check "Word LoadBehavior before startup" ($wordLoadBehaviorBefore -eq 3) "LoadBehavior=$wordLoadBehaviorBefore"
    Add-Check "PowerPoint LoadBehavior before startup" ($powerPointLoadBehaviorBefore -eq 3) "LoadBehavior=$powerPointLoadBehaviorBefore"

    $wordComRegistration = Get-ManagedComRegistrationState Word $script:resolvedOfficePlatform
    Add-Check `
        "Word managed COM registration" `
        ([bool]$wordComRegistration.passed) `
        "Architecture=$($wordComRegistration.architecture); Scope=HKLM; CLSID=$($wordComRegistration.progIdClsid); Inproc=$($wordComRegistration.inprocServer32); Class=$($wordComRegistration.className); CodeBase=$($wordComRegistration.codeBasePath); Category=$($wordComRegistration.implementedCategoryPresent); LegacyHKCU=$($wordComRegistration.legacyPerUserProgId -or $wordComRegistration.legacyPerUserClsid -or $wordComRegistration.legacyPerUserAddin)"
    $powerPointComRegistration = Get-ManagedComRegistrationState PowerPoint $script:resolvedOfficePlatform
    Add-Check `
        "PowerPoint managed COM registration" `
        ([bool]$powerPointComRegistration.passed) `
        "Architecture=$($powerPointComRegistration.architecture); Scope=HKLM; CLSID=$($powerPointComRegistration.progIdClsid); Inproc=$($powerPointComRegistration.inprocServer32); Class=$($powerPointComRegistration.className); CodeBase=$($powerPointComRegistration.codeBasePath); Category=$($powerPointComRegistration.implementedCategoryPresent); LegacyHKCU=$($powerPointComRegistration.legacyPerUserProgId -or $powerPointComRegistration.legacyPerUserClsid -or $powerPointComRegistration.legacyPerUserAddin)"

    $wordComActivation = Test-ManagedComActivation "VisualTeX.WordVsto"
    Add-Check "Word managed COM activation" ([bool]$wordComActivation.passed) "ProgID=$($wordComActivation.progId); CLSID=$($wordComActivation.clsid); HRESULT=$($wordComActivation.hresult); Error=$($wordComActivation.error)"
    $powerPointComActivation = Test-ManagedComActivation "VisualTeX.PowerPointVsto"
    Add-Check "PowerPoint managed COM activation" ([bool]$powerPointComActivation.passed) "ProgID=$($powerPointComActivation.progId); CLSID=$($powerPointComActivation.clsid); HRESULT=$($powerPointComActivation.hresult); Error=$($powerPointComActivation.error)"

    $oleState = Test-OleLocalServer $script:resolvedOfficePlatform
    Add-Check "OLE LocalServer" ([bool]$oleState.healthy) "CLSID=$($oleState.clsid); Server=$($oleState.server); EmbeddingProbe=$($oleState.embeddingProbe); Error=$($oleState.error)"

    if ($CompanionOnly) {
        $wordConnected = $null -ne $installationState -and [int]$installationState.WordConnected -eq 1
        $powerPointConnected = $null -ne $installationState -and [int]$installationState.PowerPointConnected -eq 1
        $wordResult = [pscustomobject]@{
            host = "Word"
            progId = "VisualTeX.WordVsto"
            connected = $wordConnected
            skipped = $true
            description = "COMAddIn.Connect was not tested during companion-only installation verification."
            error = ""
        }
        $powerPointResult = [pscustomobject]@{
            host = "PowerPoint"
            progId = "VisualTeX.PowerPointVsto"
            connected = $powerPointConnected
            skipped = $true
            description = "COMAddIn.Connect was not tested during companion-only installation verification."
            error = ""
        }
        Write-Host "[SKIP] Word/PowerPoint COMAddIn.Connect - companion-only verification does not launch Office."
    } else {
        $wordResult = Test-OfficeComAddIn Word "VisualTeX.WordVsto"
        Add-Check "Word COMAddIn.Connect" ([bool]$wordResult.connected) $(if ($wordResult.connected) { "True" } else { "False; $($wordResult.error)" })
        $wordLoadBehaviorAfter = Get-LoadBehavior Word $script:resolvedOfficePlatform
        Add-Check "Word LoadBehavior after startup" ($wordLoadBehaviorAfter -eq 3) "Before=$wordLoadBehaviorBefore; After=$wordLoadBehaviorAfter"

        $powerPointResult = Test-OfficeComAddIn PowerPoint "VisualTeX.PowerPointVsto"
        Add-Check "PowerPoint COMAddIn.Connect" ([bool]$powerPointResult.connected) $(if ($powerPointResult.connected) { "True" } else { "False; $($powerPointResult.error)" })
        $powerPointLoadBehaviorAfter = Get-LoadBehavior PowerPoint $script:resolvedOfficePlatform
        Add-Check "PowerPoint LoadBehavior after startup" ($powerPointLoadBehaviorAfter -eq 3) "Before=$powerPointLoadBehaviorBefore; After=$powerPointLoadBehaviorAfter"

        if (-not $wordResult.connected) {
            $wordDisabledItems = @(Get-DisabledItems Word)
            Write-Warning "Word add-in did not connect. LoadBehavior after startup=$wordLoadBehaviorAfter; DisabledItems=$($wordDisabledItems | ConvertTo-Json -Depth 5 -Compress)"
        }
        if (-not $powerPointResult.connected) {
            $powerPointDisabledItems = @(Get-DisabledItems PowerPoint)
            Write-Warning "PowerPoint add-in did not connect. LoadBehavior after startup=$powerPointLoadBehaviorAfter; DisabledItems=$($powerPointDisabledItems | ConvertTo-Json -Depth 5 -Compress)"
        }
    }
    if ($script:failures.Count -gt 0) {
        $events = @(Get-RecentOfficeLoadEvents)
        Write-Warning "Recent Office/VSTO/.NET load events: $($events | ConvertTo-Json -Depth 5 -Compress)"
    }

    if (-not (Test-Path -LiteralPath $modeKey)) { New-Item -Path $modeKey -Force | Out-Null }
    if (-not $CompanionOnly) {
        New-ItemProperty -LiteralPath $modeKey -Name "OfficeConnectionVerificationAttempted" -PropertyType DWord -Value 1 -Force | Out-Null
        New-ItemProperty -LiteralPath $modeKey -Name "WordConnected" -PropertyType DWord -Value ([int][bool]$wordResult.connected) -Force | Out-Null
        New-ItemProperty -LiteralPath $modeKey -Name "PowerPointConnected" -PropertyType DWord -Value ([int][bool]$powerPointResult.connected) -Force | Out-Null
    }
    $companionVerified =
        [bool]$processVerified -and
        [bool]$companionState.portListening -and
        [bool]$companionState.httpsHealthy -and
        [bool]$companionState.certificateMatches -and
        [bool]$companionState.protocolMatches
    New-ItemProperty -LiteralPath $modeKey -Name "CompanionProcessRunning" -PropertyType DWord -Value ([int][bool]$processVerified) -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "CompanionPortListening" -PropertyType DWord -Value ([int][bool]$companionState.portListening) -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "CompanionHttpsHealthy" -PropertyType DWord -Value ([int][bool]$companionState.httpsHealthy) -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "CompanionCertificateMatches" -PropertyType DWord -Value ([int][bool]$companionState.certificateMatches) -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "CompanionProtocolMatches" -PropertyType DWord -Value ([int][bool]$companionState.protocolMatches) -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "OfficeRuntimeVerified" -PropertyType DWord -Value ([int]$companionVerified) -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "RuntimeVerificationPending" -PropertyType DWord -Value ([int](-not $companionVerified)) -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "LastRuntimeError" -PropertyType String -Value $(if ($script:failures.Count -eq 0) { "" } else { $script:failures -join "; " }) -Force | Out-Null
    New-ItemProperty -LiteralPath $modeKey -Name "LastRuntimeReport" -PropertyType String -Value $ReportPath -Force | Out-Null

    Write-Report $wordResult $powerPointResult $certificateState $msiState $oleState $companionState $wordComRegistration $powerPointComRegistration $wordDisabledItems $powerPointDisabledItems $events
    if ($script:failures.Count -gt 0) {
        throw "Office runtime verification failed: $($script:failures -join '; ')"
    }

    if ($CompanionOnly) {
        Write-Host "installed and runtime-verified: static Office installation and VisualTeX companion runtime passed."
        Write-Warning "Word/PowerPoint COMAddIn.Connect has not been tested because Office applications were not launched."
    } else {
        Write-Host "Native Office integration installed and verified successfully."
    }
} catch {
    if ($script:failures.Count -eq 0) {
        Add-Check "Runtime verification" $false $_.Exception.Message
    }
    if ($null -eq $wordResult) {
        $wordResult = [pscustomobject]@{ host = "Word"; progId = "VisualTeX.WordVsto"; connected = $false; description = ""; error = $_.Exception.Message }
    }
    if ($null -eq $powerPointResult) {
        $powerPointResult = [pscustomobject]@{ host = "PowerPoint"; progId = "VisualTeX.PowerPointVsto"; connected = $false; description = ""; error = $_.Exception.Message }
    }
    if ($null -eq $certificateState) { $certificateState = [pscustomobject]@{ passed = $false; details = "Not evaluated" } }
    if ($null -eq $msiState) { $msiState = [pscustomobject]@{ installed = $false; productCode = "" } }
    if ($null -eq $oleState) { $oleState = [pscustomobject]@{ healthy = $false; server = ""; clsid = "{8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B}"; embeddingProbe = $false; error = "Not evaluated" } }
    if ($null -eq $companionState) { $companionState = [pscustomobject]@{ executable = ""; appDataRoot = ""; port = 0; expectedProtocol = 0; processRunning = $false; portListening = $false; httpsHealthy = $false; certificateMatches = $false; protocolMatches = $false; certificatePath = ""; expectedThumbprint = ""; localThumbprint = ""; remoteThumbprint = ""; health = $null; errors = @("Not evaluated") } }
    if ($null -eq $wordComRegistration) { $wordComRegistration = [pscustomobject]@{ host = "Word"; architecture = $script:resolvedOfficePlatform; passed = $false; error = "Not evaluated" } }
    if ($null -eq $powerPointComRegistration) { $powerPointComRegistration = [pscustomobject]@{ host = "PowerPoint"; architecture = $script:resolvedOfficePlatform; passed = $false; error = "Not evaluated" } }
    try {
        $modeKey = "HKCU:\Software\VisualTeX\OfficeIntegration"
        if (-not (Test-Path -LiteralPath $modeKey)) { New-Item -Path $modeKey -Force | Out-Null }
        $failedRuntimeValues = @(
            "CompanionProcessRunning",
            "CompanionPortListening",
            "CompanionHttpsHealthy",
            "CompanionCertificateMatches",
            "CompanionProtocolMatches",
            "OfficeRuntimeVerified"
        )
        if (-not $CompanionOnly) {
            New-ItemProperty -LiteralPath $modeKey -Name "OfficeConnectionVerificationAttempted" -PropertyType DWord -Value 1 -Force | Out-Null
            $failedRuntimeValues += @("WordConnected", "PowerPointConnected")
        }
        foreach ($valueName in $failedRuntimeValues) {
            New-ItemProperty -LiteralPath $modeKey -Name $valueName -PropertyType DWord -Value 0 -Force | Out-Null
        }
        New-ItemProperty -LiteralPath $modeKey -Name "RuntimeVerificationPending" -PropertyType DWord -Value 1 -Force | Out-Null
        $runtimeError = if ($script:failures.Count -gt 0) { $script:failures -join "; " } else { $_.Exception.Message }
        New-ItemProperty -LiteralPath $modeKey -Name "LastRuntimeError" -PropertyType String -Value $runtimeError -Force | Out-Null
        New-ItemProperty -LiteralPath $modeKey -Name "LastRuntimeReport" -PropertyType String -Value $ReportPath -Force | Out-Null
        Write-Report $wordResult $powerPointResult $certificateState $msiState $oleState $companionState $wordComRegistration $powerPointComRegistration $wordDisabledItems $powerPointDisabledItems $events
    } catch { }
    throw
}
