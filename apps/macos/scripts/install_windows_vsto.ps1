[CmdletBinding()]
param(
    [string]$MsiPath,
    [string]$LogPath,
    [string]$HashManifestPath
)

$ErrorActionPreference = "Stop"
$upgradeCode = "{A81B4BF7-0E51-45CE-A5AA-5E28F6944F42}"
$displayName = "VisualTeX Windows Office Integration"
$root = Split-Path -Parent $PSScriptRoot

function Invoke-MsiExec([string[]]$Arguments, [string]$Operation) {
    $process = Start-Process msiexec.exe -ArgumentList $Arguments -Wait -PassThru
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
            if ($item.DisplayName -eq $displayName -and $key.PSChildName -match '^\{[0-9A-Fa-f-]{36}\}$') {
                [void]$codes.Add($key.PSChildName)
            }
        }
    }
    return @($codes)
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

function Assert-RegistryValue([string]$Path, [string]$Name, $Expected) {
    if (-not (Test-Path $Path)) { throw "Required registry key is missing: $Path" }
    $actual = (Get-ItemProperty $Path -Name $Name -ErrorAction Stop).$Name
    if ($actual -ne $Expected) { throw "Registry value $Path::$Name is '$actual'; expected '$Expected'." }
}

function Assert-ComRegistration([string]$ProgId, [string]$Clsid, [string]$AssemblyFile, [string]$ClassName) {
    Assert-RegistryValue "HKCU:\Software\Classes\$ProgId\CLSID" "(default)" $Clsid
    $classKey = "HKCU:\Software\Classes\CLSID\$Clsid"
    Assert-RegistryValue "$classKey\InprocServer32" "(default)" "mscoree.dll"
    Assert-RegistryValue "$classKey\InprocServer32" "Class" $ClassName
    $codeBase = (Get-ItemProperty "$classKey\InprocServer32" -Name CodeBase).CodeBase
    if ($codeBase -notmatch ([regex]::Escape($AssemblyFile) + '$')) { throw "COM CodeBase does not target ${AssemblyFile}: $codeBase" }
    foreach ($requiredKey in @("$classKey\InprocServer32\1.0.0.0", "$classKey\ProgId", "$classKey\Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}")) {
        if (-not (Test-Path $requiredKey)) { throw "Managed COM registration is incomplete: $requiredKey" }
    }
}

$office = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue
if ($null -ne $office -and $office.Platform -ne "x64") {
    throw "This VisualTeX package is x64, but installed Office reports platform '$($office.Platform)'."
}

& (Join-Path $PSScriptRoot "uninstall_windows_ole.ps1")
& (Join-Path $PSScriptRoot "ensure_windows_office_certificate.ps1")

if ([string]::IsNullOrWhiteSpace($MsiPath)) {
    $MsiPath = Get-ChildItem (Join-Path $root "src-windows") -Filter "VisualTeX-WindowsOffice-VSTO.msi" -Recurse -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName
}
if ([string]::IsNullOrWhiteSpace($MsiPath) -or -not (Test-Path $MsiPath)) {
    throw "VisualTeX VSTO MSI was not found. Pass the exact extracted MSI path with -MsiPath."
}
$MsiPath = (Resolve-Path $MsiPath).Path
if ([string]::IsNullOrWhiteSpace($HashManifestPath)) {
    $HashManifestPath = Join-Path (Split-Path -Parent $MsiPath) "VisualTeX-WindowsOffice-VSTO.sha256.json"
}
if (-not (Test-Path $HashManifestPath)) {
    throw "VSTO SHA-256 manifest is missing: $HashManifestPath"
}

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
$remaining = @(Get-RelatedProductCodes)
if ($remaining.Count -ne 0) { throw "Stale VisualTeX MSI products remain after cleanup: $($remaining -join ', ')" }

[void](Invoke-MsiExec @("/i", ('"{0}"' -f $MsiPath), "/passive", "/norestart", "/L*v", ('"{0}"' -f $LogPath)) "MSI install")
$installedProducts = @(Get-RelatedProductCodes)
if ($installedProducts.Count -ne 1) { throw "Expected exactly one VisualTeX MSI product after install; found $($installedProducts.Count)." }

Remove-VisualTeXOfficeResiliencyEntries
$installRoot = Join-Path $env:LOCALAPPDATA "VisualTeX\WindowsOffice\VSTO"
$hashManifest = Get-Content $HashManifestPath -Raw | ConvertFrom-Json
foreach ($entry in @($hashManifest.word, $hashManifest.powerPoint)) {
    $installedFile = Join-Path $installRoot $entry.file
    if (-not (Test-Path $installedFile)) { throw "Installed VSTO file is missing: $installedFile" }
    $actualHash = (Get-FileHash $installedFile -Algorithm SHA256).Hash
    if ($actualHash -ne $entry.sha256) { throw "Installed DLL hash mismatch for $($entry.file): $actualHash != $($entry.sha256)" }
    Write-Host "SHA-256 verified: $($entry.file) $actualHash"
}

foreach ($key in @(
    "HKCU:\Software\Microsoft\Office\Word\Addins\VisualTeX.WordVsto",
    "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\VisualTeX.PowerPointVsto"
)) {
    Assert-RegistryValue $key "LoadBehavior" 3
}
Assert-ComRegistration "VisualTeX.WordVsto" "{F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1}" "VisualTeX.WordVsto.dll" "VisualTeX.WordVsto.ThisAddIn"
Assert-ComRegistration "VisualTeX.PowerPointVsto" "{7E586D2D-57B0-4D14-AB24-EBA9021A5E6D}" "VisualTeX.PowerPointVsto.dll" "VisualTeX.PowerPointVsto.ThisAddIn"

$modeKey = "HKCU:\Software\VisualTeX\OfficeIntegration"
if (-not (Test-Path $modeKey)) { New-Item $modeKey -Force | Out-Null }
New-ItemProperty $modeKey -Name "Mode" -PropertyType String -Value "vsto" -Force | Out-Null
New-ItemProperty $modeKey -Name "OleManifestEnabled" -PropertyType DWord -Value 0 -Force | Out-Null
Write-Host "VisualTeX x64 VSTO installed and chain-verified. MSI log: $LogPath"
