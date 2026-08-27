[CmdletBinding()]
param(
    [string]$ExpectedAppVersion = "1.2.5",
    [string]$ExpectedOfficeMsiVersion = "1.0.42.0"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot

function Get-RegistryView([string]$Architecture) {
    if ($Architecture -eq "x86") { return [Microsoft.Win32.RegistryView]::Registry32 }
    return [Microsoft.Win32.RegistryView]::Registry64
}

function Test-RegistryKey {
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
    } finally { $baseKey.Dispose() }
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
            return $key.GetValue($valueName, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
        } finally { $key.Dispose() }
    } finally { $baseKey.Dispose() }
}

$integrationState = Get-ItemProperty -LiteralPath "HKCU:\Software\VisualTeX\OfficeIntegration" -ErrorAction SilentlyContinue
$officePlatform = [string]$integrationState.OfficePlatform
if ($officePlatform -notin @("x86", "x64")) {
    $clickToRun = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue
    $officePlatform = if ($clickToRun.Platform -in @("x86", "x64")) { [string]$clickToRun.Platform } else { "x64" }
}
Write-Host "Installed Office integration platform=$officePlatform"
$installRoot = Join-Path $env:LOCALAPPDATA "VisualTeX"
$programFilesRoot = $env:ProgramFiles
if ($officePlatform -eq "x86" -and -not [string]::IsNullOrWhiteSpace(${env:ProgramFiles(x86)})) {
    $programFilesRoot = ${env:ProgramFiles(x86)}
}
$vstoRoot = Join-Path $programFilesRoot "VisualTeX\WindowsOffice\VSTO"
$registeredAppPath = [string]$integrationState.ExecutablePath
$appPath = if (-not [string]::IsNullOrWhiteSpace($registeredAppPath)) {
    [IO.Path]::GetFullPath($registeredAppPath.Trim().Trim('"'))
} else {
    Join-Path $installRoot "visualtex.exe"
}
$wordDll = Join-Path $vstoRoot "VisualTeX.WordVsto.dll"
$powerPointDll = Join-Path $vstoRoot "VisualTeX.PowerPointVsto.dll"
$oleServer = Join-Path $vstoRoot "VisualTeX.FormulaOleServer.exe"
foreach ($path in @($appPath, $wordDll, $powerPointDll, $oleServer)) {
    if (-not (Test-Path -LiteralPath $path)) { throw "Installed file is missing: $path" }
}

$appVersion = (Get-Item -LiteralPath $appPath).VersionInfo.ProductVersion
Write-Host "Installed VisualTeX ProductVersion=$appVersion"
if (-not $appVersion.StartsWith($ExpectedAppVersion, [StringComparison]::Ordinal)) {
    throw "Installed VisualTeX version is $appVersion; expected $ExpectedAppVersion."
}

$expectedWordDll = Join-Path $root "src-windows\VisualTeX.WordVsto\bin\$officePlatform\Release\net472\VisualTeX.WordVsto.dll"
$expectedPowerPointDll = Join-Path $root "src-windows\VisualTeX.PowerPointVsto\bin\$officePlatform\Release\net472\VisualTeX.PowerPointVsto.dll"
$oleBuildPlatform = if ($officePlatform -eq "x86") { "Win32" } else { "x64" }
$expectedOleServer = Join-Path $root "src-windows\artifacts\formula-ole-server\$oleBuildPlatform\Release\VisualTeX.FormulaOleServer.exe"
foreach ($entry in @(
    @{ Installed = $wordDll; Expected = $expectedWordDll; Label = "Word VSTO" },
    @{ Installed = $powerPointDll; Expected = $expectedPowerPointDll; Label = "PowerPoint VSTO" },
    @{ Installed = $oleServer; Expected = $expectedOleServer; Label = "Formula OLE server" }
)) {
    $installedHash = (Get-FileHash -LiteralPath $entry.Installed -Algorithm SHA256).Hash
    $expectedHash = (Get-FileHash -LiteralPath $entry.Expected -Algorithm SHA256).Hash
    Write-Host ("{0} SHA256={1}" -f $entry.Label, $installedHash)
    if ($installedHash -ne $expectedHash) {
        throw "$($entry.Label) installed hash does not match the current release build."
    }
}

$wordLoadBehavior = Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "Software\Microsoft\Office\Word\Addins\VisualTeX.WordVsto" "LoadBehavior" $officePlatform
$powerPointLoadBehavior = Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "Software\Microsoft\Office\PowerPoint\Addins\VisualTeX.PowerPointVsto" "LoadBehavior" $officePlatform
if ([int]$wordLoadBehavior -ne 3) { throw "Word LoadBehavior is not 3 in the $officePlatform registry view." }
if ([int]$powerPointLoadBehavior -ne 3) { throw "PowerPoint LoadBehavior is not 3 in the $officePlatform registry view." }
Write-Host "Word and PowerPoint LoadBehavior=3 in the $officePlatform registry view."

foreach ($entry in @(
    @{ ProgId = "VisualTeX.WordVsto"; Clsid = "{F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1}"; AssemblyPath = $wordDll },
    @{ ProgId = "VisualTeX.PowerPointVsto"; Clsid = "{7E586D2D-57B0-4D14-AB24-EBA9021A5E6D}"; AssemblyPath = $powerPointDll }
)) {
    $registeredClsid = [string](Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) "Software\Classes\$($entry.ProgId)\CLSID" "(default)" $officePlatform)
    if (-not [string]::Equals($registeredClsid, $entry.Clsid, [StringComparison]::OrdinalIgnoreCase)) {
        throw "$($entry.ProgId) CLSID is missing or incorrect in the $officePlatform registry view."
    }
    $classKey = "Software\Classes\CLSID\$($entry.Clsid)\InprocServer32"
    $inproc = [string](Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) $classKey "(default)" $officePlatform)
    $codeBase = [string](Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) $classKey "CodeBase" $officePlatform)
    if (-not [string]::Equals($inproc, "mscoree.dll", [StringComparison]::OrdinalIgnoreCase)) {
        throw "$($entry.ProgId) InprocServer32 is not mscoree.dll in the $officePlatform registry view."
    }
    $codeBasePath = ([Uri]$codeBase).LocalPath
    if (-not [string]::Equals([IO.Path]::GetFullPath($codeBasePath), [IO.Path]::GetFullPath($entry.AssemblyPath), [StringComparison]::OrdinalIgnoreCase)) {
        throw "$($entry.ProgId) CodeBase does not point to the installed assembly in the $officePlatform registry view."
    }
}
Write-Host "Managed Word and PowerPoint COM registration verified in HKLM for $officePlatform."

foreach ($legacyPath in @(
    "Software\Microsoft\Office\Word\Addins\VisualTeX.WordVsto",
    "Software\Microsoft\Office\PowerPoint\Addins\VisualTeX.PowerPointVsto",
    "Software\Classes\VisualTeX.WordVsto",
    "Software\Classes\VisualTeX.PowerPointVsto",
    "Software\Classes\CLSID\{F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1}",
    "Software\Classes\CLSID\{7E586D2D-57B0-4D14-AB24-EBA9021A5E6D}"
)) {
    if (Test-RegistryKey ([Microsoft.Win32.RegistryHive]::CurrentUser) $legacyPath $officePlatform) {
        throw "Legacy per-user Office registration still shadows the machine-wide installation: HKCU\$legacyPath"
    }
}
Write-Host "No legacy HKCU managed COM or Office Addins registration remains."

$clsid = "{8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B}"
$localServerSubKey = "Software\Classes\CLSID\$clsid\LocalServer32"
$localServer = [string](Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) $localServerSubKey "(default)" $officePlatform)
$serverExecutable = [string](Get-RegistryValue ([Microsoft.Win32.RegistryHive]::LocalMachine) $localServerSubKey "ServerExecutable" $officePlatform)
Write-Host "OLE LocalServer32=$localServer ($officePlatform registry view)"
if (($localServer.Trim('"')) -ne $oleServer -or $serverExecutable -ne $oleServer) {
    throw "Formal OLE registration does not point to the installed $ExpectedAppVersion server."
}

$installer = New-Object -ComObject WindowsInstaller.Installer
$expectedOfficeProductCode = if ($officePlatform -eq "x86") {
    "{48ABC5AF-2963-4BE6-86E3-F03950ECD270}"
} else {
    "{8BF4D9CB-320D-4AEB-929F-7E04812795AF}"
}
$installedOfficeVersions = @()
foreach ($productCode in @($expectedOfficeProductCode)) {
    try {
        $version = $installer.ProductInfo($productCode, "VersionString")
        if ($version) {
            $installedOfficeVersions += $version
            Write-Host "Office integration $productCode VersionString=$version"
        }
    }
    catch { }
}
if ($installedOfficeVersions.Count -eq 0 -or $installedOfficeVersions -notcontains $ExpectedOfficeMsiVersion) {
    throw "Office integration MSI $ExpectedOfficeMsiVersion is not registered as installed."
}

Write-Host "Installed VisualTeX Windows release passed verification."
