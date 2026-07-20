[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$displayName = "VisualTeX Windows Office Integration"
$uninstallRoots = @(
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKCU:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
)
$productCodes = @()
foreach ($root in $uninstallRoots) {
    if (-not (Test-Path $root)) { continue }
    foreach ($key in Get-ChildItem $root -ErrorAction SilentlyContinue) {
        $item = Get-ItemProperty $key.PSPath -ErrorAction SilentlyContinue
        if ($item.DisplayName -ne $displayName) { continue }

        if ($key.PSChildName -match '^\{[0-9A-Fa-f-]{36}\}$') {
            $productCodes += $key.PSChildName
            continue
        }

        foreach ($candidate in @($item.QuietUninstallString, $item.UninstallString)) {
            if (-not [string]::IsNullOrWhiteSpace($candidate) -and
                $candidate -match '\{[0-9A-Fa-f-]{36}\}') {
                $productCodes += $Matches[0]
                break
            }
        }
    }
}
$productCodes = @($productCodes | Sort-Object -Unique)

if ($productCodes.Count -gt 0) {
    foreach ($productCode in $productCodes) {
        $process = Start-Process msiexec.exe -ArgumentList @(
            "/x",
            $productCode,
            "/passive",
            "/norestart"
        ) -Wait -PassThru
        if ($process.ExitCode -notin @(0, 1605, 3010)) {
            throw "VisualTeX VSTO MSI uninstall failed for $productCode with exit code $($process.ExitCode)."
        }
    }
} else {
    foreach ($key in @(
        "HKCU:\Software\Microsoft\Office\Word\Addins\VisualTeX.WordVsto",
        "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\VisualTeX.PowerPointVsto",
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
        Remove-Item $key -Recurse -Force -ErrorAction SilentlyContinue
    }
}

$modeKey = "HKCU:\Software\VisualTeX\OfficeIntegration"
if (Test-Path $modeKey) {
    New-ItemProperty $modeKey -Name "Mode" -PropertyType String -Value "auto" -Force | Out-Null
    New-ItemProperty $modeKey -Name "NativeOleEnabled" -PropertyType DWord -Value 0 -Force | Out-Null
}
Write-Host "VisualTeX native VSTO + OLE Office integration removed for the current user."
