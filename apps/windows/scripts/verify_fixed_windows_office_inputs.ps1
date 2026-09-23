[CmdletBinding()]
param(
    [string]$ExpectedOfficeMsiVersion = "1.0.43.0"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$officeRoot = Join-Path $root "src-tauri\resources\windows-office"
$x64 = Join-Path $officeRoot "VisualTeX-WindowsOffice-VSTO-x64.msi"
$x86 = Join-Path $officeRoot "VisualTeX-WindowsOffice-VSTO-x86.msi"
$x64Manifest = Join-Path $officeRoot "VisualTeX-WindowsOffice-VSTO-x64.sha256.json"
$x86Manifest = Join-Path $officeRoot "VisualTeX-WindowsOffice-VSTO-x86.sha256.json"
$vsto = Join-Path $officeRoot "vstor_redist.exe"
$vstoManifest = Join-Path $officeRoot "vstor_redist.sha256.json"
$bridge = Join-Path $root "src-tauri\binaries\visualtex-windows-office-bridge-x86_64-pc-windows-msvc.exe"
$privateRuntimeManifest = Join-Path $root "src-tauri\target\private-mathtype-runtime\visualtex-runtime-manifest.json"
$fontRoot = Join-Path $root "src-windows\artifacts\office-fonts"
$font = Join-Path $fontRoot "latinmodern-math.otf"
$fontLicense = Join-Path $fontRoot "GUST-FONT-LICENSE.txt"
$fontReadme = Join-Path $fontRoot "README-Latin-Modern-Math.txt"

$expected = @{
    $x64 = "756C1AA9622234ED91BC789157ADBE97605D646DAD9173600A0A91C0732EF625"
    $x86 = "C49F6C092A8E6E1D14C8BF0FB9A72A7B8DBC3B73B79600AC16A2F7DB9234910F"
    $vsto = "CFE1A40BBE4A50022DB2164ABDB0154984E2CECB761A23CDC81CB5754F6E0A18"
    $bridge = "9131480E504B277E54857AE8D28A40A7C433E4CEC8BB6F108676D7660C114C32"
    $privateRuntimeManifest = "F634A36CF0EA361D9A13156DB8409551E50CC7AF9CD60EC9E6977E57B34766DF"
    $font = "6075562B771F8B82F0C179E363389684F2DD09DE30038269E2628E504BD7BE0F"
    $fontLicense = "2BD69AFFC3DA00715116F713F57EAB9707E96DAF3562AD0215987B15B9C16F73"
    $fontReadme = "E6FDD5FB44B656146FA988D191F58901C09A347D3F4FBAA3B45B24A6AD7EC30B"
}

function Assert-Hash([string]$Path, [string]$ExpectedHash, [string]$Label) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "$Label is missing: $Path" }
    $actual = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    if (-not [string]::Equals($actual, $ExpectedHash, [StringComparison]::OrdinalIgnoreCase)) {
        throw "$Label SHA256 mismatch. Expected=$ExpectedHash Actual=$actual Path=$Path"
    }
}

foreach ($entry in $expected.GetEnumerator()) { Assert-Hash $entry.Key $entry.Value (Split-Path -Leaf $entry.Key) }
foreach ($manifestPath in @($x64Manifest,$x86Manifest,$vstoManifest)) {
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw "Fixed-input manifest is missing: $manifestPath" }
}

$x64Data = Get-Content -LiteralPath $x64Manifest -Raw | ConvertFrom-Json
$x86Data = Get-Content -LiteralPath $x86Manifest -Raw | ConvertFrom-Json
if ([string]$x64Data.architecture -ne "x64" -or [string]$x86Data.architecture -ne "x86") { throw "Office architecture manifests are inconsistent." }
if (-not [string]::Equals([string]$x64Data.package.sha256, $expected[$x64], [StringComparison]::OrdinalIgnoreCase)) { throw "x64 MSI manifest does not pin the accepted MSI." }
if (-not [string]::Equals([string]$x86Data.package.sha256, $expected[$x86], [StringComparison]::OrdinalIgnoreCase)) { throw "x86 MSI manifest does not pin the accepted MSI." }
$expectedOfficeComponents = @{
    x64 = @{
        word = "35EDB78232272963D22FDFAFD4BF21A63DDF749F6216A0CC527CE91D446083F8"
        powerPoint = "3294285A46239A26C24E3D6587FED919BD9978FD6243BE45659768E3B4E0B249"
        formulaOleServer = "3527651149BDAF72A68BAF663012C480A4CCEE0600AC05804BED66C0A3805900"
        contracts = "C8FE901EFC50EC2A0CAAC1630D6056A39CBE9B4EEBA255ACB1CACDDAB72C0A48"
    }
    x86 = @{
        word = "3851787A9E153285C538A7C06D28958B2D3ADDA26A493761006946F5B9B490B5"
        powerPoint = "959D5F128DFFC996322C6CC4F88BC5BFE9C1668BA9CAB803164D9831D681D94B"
        formulaOleServer = "D8A60F2B792379695200E8A656C75583B160197D4983ACF10BB81D18FE1765B7"
        contracts = "56442F932CEC9D682D440451393538BF7878F37EE2E4EFD7145C7532DF434F45"
    }
}
foreach ($entry in @(@{data=$x64Data; arch="x64"}, @{data=$x86Data; arch="x86"})) {
    $manifest = $entry.data; $accepted = $expectedOfficeComponents[$entry.arch]
    if (-not [string]::Equals([string]$manifest.word.sha256, $accepted.word, [StringComparison]::OrdinalIgnoreCase)) { throw "$($entry.arch) Word VSTO DLL identity changed." }
    if (-not [string]::Equals([string]$manifest.powerPoint.sha256, $accepted.powerPoint, [StringComparison]::OrdinalIgnoreCase)) { throw "$($entry.arch) PowerPoint VSTO DLL identity changed." }
    if (-not [string]::Equals([string]$manifest.formulaOleServer.sha256, $accepted.formulaOleServer, [StringComparison]::OrdinalIgnoreCase)) { throw "$($entry.arch) Formula OLE server identity changed." }
    $contracts = @($manifest.dependencies | Where-Object { [string]$_.file -eq "VisualTeX.WindowsOffice.Contracts.dll" })
    if ($contracts.Count -ne 1 -or -not [string]::Equals([string]$contracts[0].sha256, $accepted.contracts, [StringComparison]::OrdinalIgnoreCase)) { throw "$($entry.arch) Office contracts DLL identity changed." }
    if (-not [string]::Equals([string]$manifest.officeMathFont.sha256, $expected[$font], [StringComparison]::OrdinalIgnoreCase)) { throw "Office MSI manifest does not pin the accepted Latin Modern Math font." }
    if (-not [string]::Equals([string]$manifest.officeMathFont.licenseSha256, $expected[$fontLicense], [StringComparison]::OrdinalIgnoreCase)) { throw "Office MSI manifest does not pin the accepted GUST license." }
    if (-not [string]::Equals([string]$manifest.officeMathFont.readmeSha256, $expected[$fontReadme], [StringComparison]::OrdinalIgnoreCase)) { throw "Office MSI manifest does not pin the accepted font README." }
}
$vstoData = Get-Content -LiteralPath $vstoManifest -Raw | ConvertFrom-Json
if (-not [string]::Equals([string]$vstoData.package.sha256, $expected[$vsto], [StringComparison]::OrdinalIgnoreCase)) { throw "VSTO Runtime manifest does not pin the accepted executable." }

$runtime = Get-Content -LiteralPath $privateRuntimeManifest -Raw | ConvertFrom-Json
if ([string]$runtime.product -ne "MathType" -or [string]$runtime.validatedVersion -ne "7.9.1" -or [int]$runtime.fileCount -ne 69 -or [int64]$runtime.bytes -ne 90095921) {
    throw "Private MathType runtime manifest is not the accepted 7.9.1 profile."
}
$expectedPrivateCore = @{
    "System\MT6.dll" = "D48F3BCD64B05DA6B3B3BBF0F47DCF8C2CB9F41ADC5131A4318BAEDBDE61C4C7"
    "System\64\MT6.dll" = "541B0DCF14967256A8B1ABE4E69F1B6FA707AD84B148C28492B676722031DDCD"
    "System\MathTypeLib.exe" = "85CBD59E4B35DC3CF19C77C64C9D98737B16A35A7AB1C7935F2FC4BB9ECD2BE9"
    "MathType.exe" = "4BBC50BBEF0D8827E9E35F07B159D33743DB07065A72590C727AD8711B675C94"
    "MathPage\64\MathPage.wll" = "3056F7AD528CC0ABEC12ABB4383626A048288310C9FB665841848D99CDDF7475"
    "MT7.dsc" = "793F1C60314FC744EB5D92D75B5075B90B3E491FCC244F5E81E90641A44110EC"
}
foreach ($entry in $expectedPrivateCore.GetEnumerator()) {
    $record = $runtime.core.PSObject.Properties[$entry.Key].Value
    if ($null -eq $record -or -not [string]::Equals([string]$record.sha256, $entry.Value, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Private MathType runtime core identity mismatch: $($entry.Key)"
    }
}

function Assert-Msi([string]$Path, [bool]$Expected64Bit) {
    $installer = New-Object -ComObject WindowsInstaller.Installer
    $database = $null; $propertyView = $null; $propertyRecord = $null; $scopeView = $null; $scopeRecord = $null; $componentView = $null
    try {
        $database = $installer.OpenDatabase($Path, 0)
        $propertyView = $database.OpenView("SELECT `Value` FROM `Property` WHERE `Property`='ProductVersion'")
        $propertyView.Execute(); $propertyRecord = $propertyView.Fetch()
        if ($null -eq $propertyRecord -or [string]$propertyRecord.StringData(1) -ne $ExpectedOfficeMsiVersion) { throw "Unexpected Office MSI ProductVersion: $Path" }
        $scopeView = $database.OpenView("SELECT `Value` FROM `Property` WHERE `Property`='ALLUSERS'")
        $scopeView.Execute(); $scopeRecord = $scopeView.Fetch()
        if ($null -eq $scopeRecord -or [string]$scopeRecord.StringData(1) -ne "1") { throw "Office MSI is not machine-wide: $Path" }
        $componentView = $database.OpenView('SELECT `Component`, `Attributes` FROM `Component`')
        $componentView.Execute(); $found = @{}
        while ($true) {
            $record = $componentView.Fetch(); if ($null -eq $record) { break }
            try { $found[[string]$record.StringData(1)] = [int]$record.IntegerData(2) } finally { if ([Runtime.InteropServices.Marshal]::IsComObject($record)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($record) } }
        }
        foreach ($component in @("VstoFiles","NativeOleServerFiles","NativeOleRegistration","WordComRegistration","PowerPointComRegistration","WordOfficeRegistration","PowerPointOfficeRegistration","ModeRegistration")) {
            if (-not $found.ContainsKey($component)) { throw "Office MSI missing component '$component': $Path" }
            $is64 = ([int]$found[$component] -band 256) -ne 0
            if ($is64 -ne $Expected64Bit) { throw "Office MSI component bitness mismatch: $component Path=$Path" }
        }
    } finally {
        foreach ($value in @($componentView,$scopeRecord,$scopeView,$propertyRecord,$propertyView,$database,$installer)) { if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) } }
    }
}
Assert-Msi $x64 $true
Assert-Msi $x86 $false

Write-Output ([pscustomobject]@{
    status = "PASS"
    officeMsiVersion = $ExpectedOfficeMsiVersion
    x64MsiSha256 = $expected[$x64]
    x86MsiSha256 = $expected[$x86]
    bridgeSha256 = $expected[$bridge]
    vstoRuntimeSha256 = $expected[$vsto]
    privateMathTypeVersion = [string]$runtime.validatedVersion
    privateMathTypeFiles = [int]$runtime.fileCount
    latinModernMathSha256 = $expected[$font]
    fixedInputsOnly = $true
} | ConvertTo-Json -Depth 4)
