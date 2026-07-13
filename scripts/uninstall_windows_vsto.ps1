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
        "HKCU:\Software\Classes\CLSID\{7E586D2D-57B0-4D14-AB24-EBA9021A5E6D}"
    )) {
        Remove-Item $key -Recurse -Force -ErrorAction SilentlyContinue
    }
}

$modeKey = "HKCU:\Software\VisualTeX\OfficeIntegration"
if (Test-Path $modeKey) {
    New-ItemProperty $modeKey -Name "Mode" -PropertyType String -Value "auto" -Force | Out-Null
}
Write-Host "VisualTeX VSTO Office integration removed for the current user."
