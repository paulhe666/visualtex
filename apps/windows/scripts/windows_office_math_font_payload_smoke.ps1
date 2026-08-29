[CmdletBinding()]
param(
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$expectedFontHash = "6075562B771F8B82F0C179E363389684F2DD09DE30038269E2628E504BD7BE0F"
$expectedLicenseHash = "2BD69AFFC3DA00715116F713F57EAB9707E96DAF3562AD0215987B15B9C16F73"

function Read-MsiFontRows([string]$MsiPath) {
    $installer = New-Object -ComObject WindowsInstaller.Installer
    $database = $null
    $view = $null
    $record = $null
    try {
        $database = $installer.GetType().InvokeMember(
            "OpenDatabase",
            "InvokeMethod",
            $null,
            $installer,
            @($MsiPath, 0))
        $view = $database.GetType().InvokeMember(
            "OpenView",
            "InvokeMethod",
            $null,
            $database,
            @('SELECT `File_`, `FontTitle` FROM `Font`'))
        $view.GetType().InvokeMember("Execute", "InvokeMethod", $null, $view, $null) | Out-Null
        $rows = @()
        while ($true) {
            $record = $view.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $view, $null)
            if ($null -eq $record) { break }
            $fileId = $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, 1)
            $fontTitle = $record.GetType().InvokeMember("StringData", "GetProperty", $null, $record, 2)
            $rows += [ordered]@{ FileId = [string]$fileId; FontTitle = [string]$fontTitle }
            [Runtime.InteropServices.Marshal]::ReleaseComObject($record) | Out-Null
            $record = $null
        }
        return $rows
    } finally {
        if ($null -ne $record) { [Runtime.InteropServices.Marshal]::ReleaseComObject($record) | Out-Null }
        if ($null -ne $view) { [Runtime.InteropServices.Marshal]::ReleaseComObject($view) | Out-Null }
        if ($null -ne $database) { [Runtime.InteropServices.Marshal]::ReleaseComObject($database) | Out-Null }
        if ($null -ne $installer) { [Runtime.InteropServices.Marshal]::ReleaseComObject($installer) | Out-Null }
    }
}

foreach ($architecture in @("x64", "x86")) {
    $msi = Join-Path $root "src-windows\VisualTeX.WindowsOffice.Installer\bin\$architecture\$Configuration\VisualTeX-WindowsOffice-VSTO-$architecture.msi"
    if (-not (Test-Path -LiteralPath $msi -PathType Leaf)) {
        throw "Office $architecture MSI is missing: $msi"
    }

    $fontRows = @(Read-MsiFontRows $msi)
    $latinModernRow = $fontRows | Where-Object {
        $_.FileId -eq "LatinModernMathSystemFontFile" -and
        $_.FontTitle -eq "Latin Modern Math"
    }
    if ($null -eq $latinModernRow) {
        throw "Office $architecture MSI does not register Latin Modern Math in the Windows Installer Font table."
    }

    $adminRoot = Join-Path $env:TEMP ("visualtex-office-font-msi-$architecture-" + [Guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $adminRoot -Force | Out-Null
    try {
        $process = Start-Process `
            -FilePath "msiexec.exe" `
            -ArgumentList @("/a", ('"' + $msi + '"'), "/qn", ("TARGETDIR=" + $adminRoot)) `
            -Wait `
            -PassThru
        if ($process.ExitCode -ne 0) {
            throw "Administrative extraction of Office $architecture MSI failed with exit code $($process.ExitCode)."
        }
        $fonts = @(Get-ChildItem -LiteralPath $adminRoot -Recurse -File -Filter "latinmodern-math.otf")
        if ($fonts.Count -lt 2) {
            throw "Office $architecture MSI administrative image contains $($fonts.Count) Latin Modern font copies; expected the system-font and VSTO-private payloads."
        }
        foreach ($font in $fonts) {
            $hash = (Get-FileHash -LiteralPath $font.FullName -Algorithm SHA256).Hash
            if (-not [string]::Equals($hash, $expectedFontHash, [StringComparison]::OrdinalIgnoreCase)) {
                throw "Office $architecture MSI contains a damaged Latin Modern Math font: $($font.FullName) hash=$hash"
            }
        }
        $license = Get-ChildItem -LiteralPath $adminRoot -Recurse -File -Filter "GUST-FONT-LICENSE.txt" | Select-Object -First 1
        if ($null -eq $license) {
            throw "Office $architecture MSI is missing the GUST Font License."
        }
        $licenseHash = (Get-FileHash -LiteralPath $license.FullName -Algorithm SHA256).Hash
        if (-not [string]::Equals($licenseHash, $expectedLicenseHash, [StringComparison]::OrdinalIgnoreCase)) {
            throw "Office $architecture MSI contains a damaged GUST Font License: $licenseHash"
        }
        Write-Host "Office $architecture Latin Modern Math MSI payload passed: Font table registration and $($fonts.Count) verified copies."
    } finally {
        Remove-Item -LiteralPath $adminRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
