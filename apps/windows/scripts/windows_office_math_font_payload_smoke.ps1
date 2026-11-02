[CmdletBinding()]
param(
    [string]$Configuration = "Release"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$expectedFontHash = "6075562B771F8B82F0C179E363389684F2DD09DE30038269E2628E504BD7BE0F"
$expectedLicenseHash = "2BD69AFFC3DA00715116F713F57EAB9707E96DAF3562AD0215987B15B9C16F73"
$expectedTrueTypeFontHash = "BB3646F9296C10C9B5CCD95BF5680297A4E35B3BC25EA5F4F7593A9C405E0C93"
# The OMML Font menu's other bundled TrueType faces: Font table file id, file name, SHA-256.
$bundledFonts = @(
    @{ FileId = "VisualTeXTeXGyreTermesMathSystemFontFile"; File = "VisualTeXTeXGyreTermesMath.ttf"; Hash = "9F58AD9C3C92D18BB439FF31535FBA7A8038624E5E5E7605E2681FFDB9C8B1EE" },
    @{ FileId = "StixTwoMathSystemFontFile"; File = "STIXTwoMath-Regular.ttf"; Hash = "49B22767985171BCE579307B21D98F4D35012FA39E316ACC4FFCAB8644DFB69B" }
)
$expectedStixLicenseHash = "0C8825913B60D858AACDB33C4CA6660A7D64B0D6464702EFBB19313F5765861A"

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
    # A TrueType font row has no title; Windows reads it from the file at install.
    $trueTypeRow = $fontRows | Where-Object { $_.FileId -eq "VisualTeXLatinModernMathSystemFontFile" }
    if ($null -eq $trueTypeRow) {
        throw "Office $architecture MSI does not register VisualTeX Latin Modern Math in the Windows Installer Font table."
    }
    foreach ($bundled in $bundledFonts) {
        if ($null -eq ($fontRows | Where-Object { $_.FileId -eq $bundled.FileId })) {
            throw "Office $architecture MSI does not register $($bundled.File) in the Windows Installer Font table."
        }
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
        $trueTypeFonts = @(Get-ChildItem -LiteralPath $adminRoot -Recurse -File -Filter "VisualTeXLatinModernMath.ttf")
        if ($trueTypeFonts.Count -lt 2) {
            throw "Office $architecture MSI administrative image contains $($trueTypeFonts.Count) VisualTeX Latin Modern Math copies; expected the system-font and VSTO-private payloads."
        }
        foreach ($font in $trueTypeFonts) {
            $hash = (Get-FileHash -LiteralPath $font.FullName -Algorithm SHA256).Hash
            if (-not [string]::Equals($hash, $expectedTrueTypeFontHash, [StringComparison]::OrdinalIgnoreCase)) {
                throw "Office $architecture MSI contains a damaged VisualTeX Latin Modern Math font: $($font.FullName) hash=$hash"
            }
        }
        foreach ($bundled in $bundledFonts) {
            $copies = @(Get-ChildItem -LiteralPath $adminRoot -Recurse -File -Filter $bundled.File)
            if ($copies.Count -lt 2) {
                throw "Office $architecture MSI administrative image contains $($copies.Count) copies of $($bundled.File); expected the system-font and VSTO-private payloads."
            }
            foreach ($copy in $copies) {
                $hash = (Get-FileHash -LiteralPath $copy.FullName -Algorithm SHA256).Hash
                if (-not [string]::Equals($hash, $bundled.Hash, [StringComparison]::OrdinalIgnoreCase)) {
                    throw "Office $architecture MSI contains a damaged $($bundled.File): $($copy.FullName) hash=$hash"
                }
            }
        }
        $stixLicense = Get-ChildItem -LiteralPath $adminRoot -Recurse -File -Filter "STIX-Two-Math-OFL.txt" | Select-Object -First 1
        if ($null -eq $stixLicense -or -not [string]::Equals((Get-FileHash -LiteralPath $stixLicense.FullName -Algorithm SHA256).Hash, $expectedStixLicenseHash, [StringComparison]::OrdinalIgnoreCase)) {
            throw "Office $architecture MSI is missing the STIX Two Math license or carries a damaged copy."
        }
        $license = Get-ChildItem -LiteralPath $adminRoot -Recurse -File -Filter "GUST-FONT-LICENSE.txt" | Select-Object -First 1
        if ($null -eq $license) {
            throw "Office $architecture MSI is missing the GUST Font License."
        }
        $licenseHash = (Get-FileHash -LiteralPath $license.FullName -Algorithm SHA256).Hash
        if (-not [string]::Equals($licenseHash, $expectedLicenseHash, [StringComparison]::OrdinalIgnoreCase)) {
            throw "Office $architecture MSI contains a damaged GUST Font License: $licenseHash"
        }
        Write-Host "Office $architecture math font MSI payload passed: Font table registration, $($fonts.Count) verified Latin Modern copies and the bundled TrueType faces."
    } finally {
        Remove-Item -LiteralPath $adminRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
