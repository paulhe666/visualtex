[CmdletBinding()]
param(
    [string]$MsiPath,
    [string]$PackageDirectory,
    [string]$LogPath,
    [string]$HashManifestPath,
    [ValidateSet("auto", "x86", "x64")]
    [string]$OfficePlatform = "auto"
)

$ErrorActionPreference = "Stop"
$upgradeCode = "{A81B4BF7-0E51-45CE-A5AA-5E28F6944F42}"
$displayName = "VisualTeX Windows Office Integration"
$root = Split-Path -Parent $PSScriptRoot
$bootstrapLogRoot = Join-Path $env:LOCALAPPDATA "VisualTeX\office\install-logs"
New-Item $bootstrapLogRoot -ItemType Directory -Force | Out-Null
$bootstrapStamp = Get-Date -Format "yyyyMMdd-HHmmss"
$bootstrapLogPath = Join-Path $bootstrapLogRoot "vsto-bootstrap-$bootstrapStamp.log"
$transcriptStarted = $false
try {
    Start-Transcript -Path $bootstrapLogPath -Force | Out-Null
    $transcriptStarted = $true
} catch {
    Write-Warning "Unable to start VisualTeX bootstrap transcript: $($_.Exception.Message)"
}

function Assert-NoOfficeProcesses {
    $running = @(Get-Process WINWORD, POWERPNT -ErrorAction SilentlyContinue)
    if ($running.Count -eq 0) { return }
    $names = $running |
        Sort-Object ProcessName -Unique |
        ForEach-Object { $_.ProcessName + ".EXE" }
    throw "Close Microsoft Word and PowerPoint before installing VisualTeX Office integration. Running: $($names -join ', '). No Office files or registrations were changed."
}

function Invoke-MsiExec([string[]]$Arguments, [string]$Operation) {
    Assert-NoOfficeProcesses
    $effectiveArguments = @($Arguments) + @(
        "REBOOT=ReallySuppress",
        "MSIRESTARTMANAGERCONTROL=Disable"
    )
    $process = Start-Process msiexec.exe -ArgumentList $effectiveArguments -Wait -PassThru
    Write-Host "$Operation exit code: $($process.ExitCode)"
    if ($process.ExitCode -notin @(0, 1605, 3010)) {
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
        if (-not (Test-Path $uninstallRoot)) { continue }
        foreach ($key in Get-ChildItem $uninstallRoot -ErrorAction SilentlyContinue) {
            $item = Get-ItemProperty $key.PSPath -ErrorAction SilentlyContinue
            if ($null -eq $item) { continue }
            $displayNameProperty = $item.PSObject.Properties["DisplayName"]
            if ($null -ne $displayNameProperty -and
                [string]$displayNameProperty.Value -eq $displayName -and
                $key.PSChildName -match '^\{[0-9A-Fa-f-]{36}\}$') {
                [void]$codes.Add($key.PSChildName)
            }
        }
    }
    return @($codes)
}

function Wait-ForRelatedProductCount(
    [int]$ExpectedCount,
    [int]$TimeoutSeconds = 15
) {
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
            if (-not (Test-Path $key)) { continue }
            $properties = (Get-ItemProperty $key).PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' }
            foreach ($property in $properties) {
                if ($property.Value -isnot [byte[]]) { continue }
                $text = [Text.Encoding]::Unicode.GetString([byte[]]$property.Value)
                if ($text -match 'VisualTeX\.(Word|PowerPoint)Vsto' -or $text -match 'VisualTeX\\WindowsOffice\\VSTO') {
                    Remove-ItemProperty $key -Name $property.Name -Force
                }
            }
        }
    }
}

function Get-RegistryView([string]$Architecture) {
    if ($Architecture -eq "x86") { return [Microsoft.Win32.RegistryView]::Registry32 }
    return [Microsoft.Win32.RegistryView]::Registry64
}

function Get-HkcuRegistryValue([string]$SubKey, [string]$Name, [string]$Architecture) {
    $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        [Microsoft.Win32.RegistryHive]::CurrentUser,
        (Get-RegistryView $Architecture))
    try {
        $key = $baseKey.OpenSubKey($SubKey, $false)
        if ($null -eq $key) { throw "Required registry key is missing: HKCU\$SubKey ($Architecture view)" }
        try {
            $valueName = if ($Name -eq "(default)") { "" } else { $Name }
            if ($valueName -notin @($key.GetValueNames())) {
                throw "Required registry value is missing: HKCU\$SubKey::$Name ($Architecture view)"
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

function Test-HkcuRegistryKey([string]$SubKey, [string]$Architecture) {
    $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        [Microsoft.Win32.RegistryHive]::CurrentUser,
        (Get-RegistryView $Architecture))
    try {
        $key = $baseKey.OpenSubKey($SubKey, $false)
        if ($null -eq $key) { return $false }
        $key.Dispose()
        return $true
    } finally {
        $baseKey.Dispose()
    }
}

function Test-HklmRegistryKey([string]$SubKey, [string]$Architecture) {
    $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        [Microsoft.Win32.RegistryHive]::LocalMachine,
        (Get-RegistryView $Architecture))
    try {
        $key = $baseKey.OpenSubKey($SubKey, $false)
        if ($null -eq $key) { return $false }
        $key.Dispose()
        return $true
    } finally {
        $baseKey.Dispose()
    }
}

function Assert-HkcuRegistryValue(
    [string]$SubKey,
    [string]$Name,
    $Expected,
    [string]$Architecture
) {
    $actual = Get-HkcuRegistryValue $SubKey $Name $Architecture
    if ($actual -ne $Expected) {
        throw "Registry value HKCU\$SubKey::$Name is '$actual'; expected '$Expected' ($Architecture view)."
    }
}

function Resolve-OfficePlatform([string]$RequestedPlatform) {
    if ($RequestedPlatform -ne "auto") { return $RequestedPlatform }
    $clickToRun = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue
    if ($null -ne $clickToRun -and $clickToRun.Platform -in @("x86", "x64")) {
        return [string]$clickToRun.Platform
    }
    foreach ($architecture in @("x64", "x86")) {
        if (Test-HklmRegistryKey "SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE" $architecture) {
            return $architecture
        }
    }
    throw "Unable to determine whether installed Office is x86 or x64. Pass -OfficePlatform x86 or x64."
}

function Assert-NativeOleRegistration([string]$InstallRoot, [string]$Architecture) {
    $clsid = "{8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B}"
    $iid = "{6C672AF0-7321-4D21-B325-868CB34592C2}"
    $libid = "{DF66EC66-3B3A-4675-A7BE-30456A04EB96}"
    $server = Join-Path $InstallRoot "VisualTeX.FormulaOleServer.exe"
    if (-not (Test-Path $server)) { throw "Native Formula OLE LocalServer is missing: $server" }

    Assert-HkcuRegistryValue "Software\Classes\VisualTeX.Formula.1\CLSID" "(default)" $clsid $Architecture
    Assert-HkcuRegistryValue "Software\Classes\VisualTeX.Formula\CurVer" "(default)" "VisualTeX.Formula.1" $Architecture
    $localServer = [string](Get-HkcuRegistryValue "Software\Classes\CLSID\$clsid\LocalServer32" "(default)" $Architecture)
    $localServer = $localServer.Trim('"')
    if (-not [string]::Equals($localServer, $server, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Native OLE LocalServer32 is '$localServer'; expected '$server' ($Architecture view)."
    }
    Assert-HkcuRegistryValue "Software\Classes\CLSID\$clsid\LocalServer32" "ServerExecutable" $server $Architecture
    Assert-HkcuRegistryValue "Software\Classes\CLSID\$clsid\DataFormats\GetSet\0" "(default)" "14,1,64,1" $Architecture
    Assert-HkcuRegistryValue "Software\Classes\CLSID\$clsid\DataFormats\GetSet\1" "(default)" "3,1,32,1" $Architecture
    Assert-HkcuRegistryValue "Software\Classes\CLSID\$clsid\DataFormats\GetSet\2" "(default)" "PNG,1,1,1" $Architecture
    Assert-HkcuRegistryValue "Software\Classes\CLSID\$clsid\ProgID" "(default)" "VisualTeX.Formula.1" $Architecture
    Assert-HkcuRegistryValue "Software\Classes\CLSID\$clsid\VersionIndependentProgID" "(default)" "VisualTeX.Formula" $Architecture
    Assert-HkcuRegistryValue "Software\Classes\Interface\$iid\ProxyStubClsid32" "(default)" "{00020424-0000-0000-C000-000000000046}" $Architecture
    Assert-HkcuRegistryValue "Software\Classes\Interface\$iid\TypeLib" "(default)" $libid $Architecture
    $typeLibraryPlatform = if ($Architecture -eq "x86") { "win32" } else { "win64" }
    $typeLibrary = [string](Get-HkcuRegistryValue "Software\Classes\TypeLib\$libid\1.0\0\$typeLibraryPlatform" "(default)" $Architecture)
    if (-not [string]::Equals($typeLibrary, $server, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Native OLE type library path is '$typeLibrary'; expected '$server' ($Architecture view)."
    }
}

function Assert-ComRegistration(
    [string]$ProgId,
    [string]$Clsid,
    [string]$AssemblyFile,
    [string]$ClassName,
    [string]$Architecture
) {
    Assert-HkcuRegistryValue "Software\Classes\$ProgId\CLSID" "(default)" $Clsid $Architecture
    $classKey = "Software\Classes\CLSID\$Clsid"
    Assert-HkcuRegistryValue "$classKey\InprocServer32" "(default)" "mscoree.dll" $Architecture
    Assert-HkcuRegistryValue "$classKey\InprocServer32" "Class" $ClassName $Architecture
    $codeBase = [string](Get-HkcuRegistryValue "$classKey\InprocServer32" "CodeBase" $Architecture)
    if ($codeBase -notmatch ([regex]::Escape($AssemblyFile) + '$')) {
        throw "COM CodeBase does not target ${AssemblyFile}: $codeBase"
    }
    foreach ($requiredKey in @(
        "$classKey\InprocServer32\1.0.0.0",
        "$classKey\ProgId",
        "$classKey\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}"
    )) {
        if (-not (Test-HkcuRegistryKey $requiredKey $Architecture)) {
            throw "Managed COM registration is incomplete: HKCU\$requiredKey ($Architecture view)"
        }
    }
}

try {
Assert-NoOfficeProcesses
$resolvedOfficePlatform = Resolve-OfficePlatform $OfficePlatform
Write-Host "Detected Office platform: $resolvedOfficePlatform"
Write-Host "Bootstrap log: $bootstrapLogPath"

& (Join-Path $PSScriptRoot "uninstall_windows_ole.ps1")
& (Join-Path $PSScriptRoot "ensure_windows_office_certificate.ps1")

if ([string]::IsNullOrWhiteSpace($PackageDirectory)) {
    foreach ($candidate in @(
        (Join-Path (Split-Path -Parent $PSScriptRoot) "windows-office"),
        (Join-Path $root "src-tauri\resources\windows-office")
    )) {
        if (Test-Path $candidate) {
            $PackageDirectory = $candidate
            break
        }
    }
}

$packageFileName = "VisualTeX-WindowsOffice-VSTO-$resolvedOfficePlatform.msi"
$manifestFileName = "VisualTeX-WindowsOffice-VSTO-$resolvedOfficePlatform.sha256.json"
if ([string]::IsNullOrWhiteSpace($MsiPath) -and -not [string]::IsNullOrWhiteSpace($PackageDirectory)) {
    $MsiPath = Join-Path $PackageDirectory $packageFileName
}
if ([string]::IsNullOrWhiteSpace($MsiPath) -or -not (Test-Path $MsiPath)) {
    throw "VisualTeX $resolvedOfficePlatform VSTO MSI was not found. Pass -PackageDirectory or the exact -MsiPath."
}
$MsiPath = (Resolve-Path $MsiPath).Path
if ([string]::IsNullOrWhiteSpace($HashManifestPath)) {
    $HashManifestPath = Join-Path (Split-Path -Parent $MsiPath) $manifestFileName
}
if (-not (Test-Path $HashManifestPath)) {
    throw "VSTO SHA-256 manifest is missing: $HashManifestPath"
}
$HashManifestPath = (Resolve-Path $HashManifestPath).Path
$hashManifest = Get-Content $HashManifestPath -Raw | ConvertFrom-Json
if ($hashManifest.architecture -ne $resolvedOfficePlatform) {
    throw "Package architecture '$($hashManifest.architecture)' does not match installed Office '$resolvedOfficePlatform'."
}
if ($hashManifest.package.file -ne (Split-Path -Leaf $MsiPath)) {
    throw "Package manifest expects '$($hashManifest.package.file)' but received '$(Split-Path -Leaf $MsiPath)'."
}
$actualPackageHash = (Get-FileHash $MsiPath -Algorithm SHA256).Hash
if ($actualPackageHash -ne $hashManifest.package.sha256) {
    throw "MSI SHA-256 mismatch: $actualPackageHash != $($hashManifest.package.sha256)"
}
Write-Host "SHA-256 verified before install: $packageFileName $actualPackageHash"

$logRoot = Join-Path $env:LOCALAPPDATA "VisualTeX\office\install-logs"
New-Item $logRoot -ItemType Directory -Force | Out-Null
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
if ([string]::IsNullOrWhiteSpace($LogPath)) { $LogPath = Join-Path $logRoot "vsto-install-$stamp.log" }

$oldProducts = @(Get-RelatedProductCodes)
Write-Host "Related VisualTeX MSI products before install: $($oldProducts.Count)"
foreach ($productCode in $oldProducts) {
    $uninstallLog = Join-Path $logRoot "vsto-uninstall-$($productCode.Trim('{}'))-$stamp.log"
    [void](Invoke-MsiExec @("/x", $productCode, "/passive", "/norestart", "/L*v", ('"{0}"' -f $uninstallLog)) "MSI uninstall $productCode")
}
$remaining = @(Wait-ForRelatedProductCount 0)

[void](Invoke-MsiExec @("/i", ('"{0}"' -f $MsiPath), "/passive", "/norestart", "/L*v", ('"{0}"' -f $LogPath)) "MSI install")
$installedProducts = @(Wait-ForRelatedProductCount 1)

Remove-VisualTeXOfficeResiliencyEntries
$installRoot = Join-Path $env:LOCALAPPDATA "VisualTeX\WindowsOffice\VSTO"
$verifiedEntries = @($hashManifest.word, $hashManifest.powerPoint, $hashManifest.formulaOleServer)
if ($null -ne $hashManifest.dependencies) {
    $verifiedEntries += @($hashManifest.dependencies)
}
foreach ($entry in $verifiedEntries) {
    $installedFile = Join-Path $installRoot $entry.file
    if (-not (Test-Path $installedFile)) { throw "Installed VSTO file is missing: $installedFile" }
    $actualHash = (Get-FileHash $installedFile -Algorithm SHA256).Hash
    if ($actualHash -ne $entry.sha256) { throw "Installed DLL hash mismatch for $($entry.file): $actualHash != $($entry.sha256)" }
    Write-Host "SHA-256 verified: $($entry.file) $actualHash"
}

foreach ($key in @(
    "Software\Microsoft\Office\Word\Addins\VisualTeX.WordVsto",
    "Software\Microsoft\Office\PowerPoint\Addins\VisualTeX.PowerPointVsto"
)) {
    Assert-HkcuRegistryValue $key "LoadBehavior" 3 $resolvedOfficePlatform
}
Assert-ComRegistration "VisualTeX.WordVsto" "{F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1}" "VisualTeX.WordVsto.dll" "VisualTeX.WordVsto.ThisAddIn" $resolvedOfficePlatform
Assert-ComRegistration "VisualTeX.PowerPointVsto" "{7E586D2D-57B0-4D14-AB24-EBA9021A5E6D}" "VisualTeX.PowerPointVsto.dll" "VisualTeX.PowerPointVsto.ThisAddIn" $resolvedOfficePlatform
Assert-NativeOleRegistration $installRoot $resolvedOfficePlatform

$modeKey = "HKCU:\Software\VisualTeX\OfficeIntegration"
if (-not (Test-Path $modeKey)) { New-Item $modeKey -Force | Out-Null }
New-ItemProperty $modeKey -Name "Mode" -PropertyType String -Value "native-vsto-ole" -Force | Out-Null
New-ItemProperty $modeKey -Name "OleManifestEnabled" -PropertyType DWord -Value 0 -Force | Out-Null
New-ItemProperty $modeKey -Name "NativeOleEnabled" -PropertyType DWord -Value 1 -Force | Out-Null
Write-Host "VisualTeX $resolvedOfficePlatform native VSTO + OLE installed and chain-verified. MSI log: $LogPath"
} finally {
    if ($transcriptStarted) {
        try { Stop-Transcript | Out-Null } catch { }
    }
}
