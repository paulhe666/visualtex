[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$catalog = Join-Path $env:LOCALAPPDATA "VisualTeX\OfficeCatalog"
$catalogShareName = "VisualTeXOfficeCatalog"
$catalogId = "{69C6A866-755B-4C5A-BACB-EEA28B03C724}"
$catalogKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$catalogId"
$legacyCatalogKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\VisualTeX"
$developerKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
$modeKey = "HKCU:\Software\VisualTeX\OfficeIntegration"
$manifestIds = @(
    "7c7d3b35-56b2-4c40-88d9-c9eb836d6021",
    "fdc8d615-7e60-4586-bff4-5a1d728f9f6c"
)

Remove-Item $catalogKey -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item $legacyCatalogKey -Recurse -Force -ErrorAction SilentlyContinue
foreach ($manifestId in $manifestIds) {
    Remove-ItemProperty $developerKey -Name $manifestId -Force -ErrorAction SilentlyContinue
}
if (Get-SmbShare -Name $catalogShareName -ErrorAction SilentlyContinue) {
    try {
        Remove-SmbShare -Name $catalogShareName -Force -ErrorAction Stop
    } catch {
        $script = "Remove-SmbShare -Name '$catalogShareName' -Force -ErrorAction SilentlyContinue"
        $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($script))
        $process = Start-Process powershell.exe -Verb RunAs -WindowStyle Hidden -Wait -PassThru `
            -ArgumentList "-NoProfile", "-EncodedCommand", $encoded
        if ($process.ExitCode -ne 0) {
            throw "Unable to remove the VisualTeX Office catalog share."
        }
    }
}
$customUiCacheKey = "HKCU:\Software\Microsoft\Office\16.0\Common\CustomUIValidationCache"
if (Test-Path $customUiCacheKey) {
    $properties = (Get-ItemProperty $customUiCacheKey).PSObject.Properties |
        Where-Object { $_.Name -notmatch '^PS(Path|ParentPath|ChildName|Drive|Provider)$' }
    foreach ($property in $properties) {
        if ($manifestIds | Where-Object { $property.Name.StartsWith($_, [StringComparison]::OrdinalIgnoreCase) }) {
            Remove-ItemProperty $customUiCacheKey -Name $property.Name -Force
        }
    }
}
Remove-Item $catalog -Recurse -Force -ErrorAction SilentlyContinue
$wefRoot = Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\Wef"
if (Test-Path $wefRoot) {
    $markers = @(
        $manifestIds[0],
        $manifestIds[1],
        "VisualTeX.WindowsOle",
        "127.0.0.1:43127"
    )
    foreach ($file in Get-ChildItem $wefRoot -File -Recurse -ErrorAction SilentlyContinue) {
        $relativePath = $file.FullName.Substring($wefRoot.Length).TrimStart('\')
        $isVisualTeX = $markers | Where-Object { $relativePath -match [regex]::Escape($_) } | Select-Object -First 1
        $contentScanAllowed = $relativePath -match '^(AddinInfo|AggregatedCache|AppCommands)\\' -or
            ($relativePath -match '^\{[0-9A-Fa-f-]+\}\\' -and $relativePath -notmatch '\\Omex\\')
        if (-not $isVisualTeX -and $contentScanAllowed -and $file.Length -le 10MB) {
            try {
                $bytes = [IO.File]::ReadAllBytes($file.FullName)
                $text = [Text.Encoding]::ASCII.GetString($bytes) + [Text.Encoding]::Unicode.GetString($bytes)
                $isVisualTeX = $markers | Where-Object { $text -match [regex]::Escape($_) } | Select-Object -First 1
            } catch { }
        }
        if ($isVisualTeX) { Remove-Item $file.FullName -Force -ErrorAction SilentlyContinue }
    }
}
if (Test-Path $modeKey) {
    New-ItemProperty $modeKey -Name "OleManifestEnabled" -PropertyType DWord -Value 0 -Force | Out-Null
}
Write-Host "VisualTeX Windows OLE Office integration removed for the current user."
