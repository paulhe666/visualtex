[CmdletBinding()]
param(
    [switch]$Force
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$outputRoot = Join-Path $root "src-windows\artifacts\office-fonts"
$fontPath = Join-Path $outputRoot "latinmodern-math.otf"
$licensePath = Join-Path $outputRoot "GUST-FONT-LICENSE.txt"
$readmePath = Join-Path $outputRoot "README-Latin-Modern-Math.txt"
$manifestPath = Join-Path $outputRoot "latinmodern-math.manifest.json"

$expectedArchiveSha256 = "3D906317F27279AF05EB095AA4DB5E7F3F87312E69D672E8F8928B64ADCD403C"
$expectedFontSha256 = "6075562B771F8B82F0C179E363389684F2DD09DE30038269E2628E504BD7BE0F"
$expectedLicenseSha256 = "2BD69AFFC3DA00715116F713F57EAB9707E96DAF3562AD0215987B15B9C16F73"
$expectedReadmeSha256 = "E6FDD5FB44B656146FA988D191F58901C09A347D3F4FBAA3B45B24A6AD7EC30B"
$downloadUrls = @(
    "https://mirrors.cstcloud.cn/CTAN/fonts/lm-math.zip",
    "https://mirrors.ibiblio.org/pub/mirrors/CTAN/fonts/lm-math.zip"
)

function Assert-FileHash([string]$Path, [string]$Expected, [string]$Label) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "$Label is missing: $Path"
    }
    $actual = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    if (-not [string]::Equals($actual, $Expected, [StringComparison]::OrdinalIgnoreCase)) {
        throw "$Label SHA-256 mismatch: expected=$Expected actual=$actual path=$Path"
    }
}

function Test-PreparedFont {
    try {
        Assert-FileHash $fontPath $expectedFontSha256 "Latin Modern Math font"
        Assert-FileHash $licensePath $expectedLicenseSha256 "GUST Font License"
        Assert-FileHash $readmePath $expectedReadmeSha256 "Latin Modern Math README"
        return $true
    } catch {
        return $false
    }
}

if (-not $Force -and (Test-PreparedFont)) {
    Write-Host "Verified cached Latin Modern Math Office font: $fontPath"
    exit 0
}

New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null
$tempRoot = Join-Path $env:TEMP ("visualtex-office-font-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
try {
    $archivePath = Join-Path $tempRoot "lm-math.zip"
    $downloadError = $null
    foreach ($downloadUrl in $downloadUrls) {
        try {
            Write-Host "Downloading pinned Latin Modern Math package from $downloadUrl"
            Invoke-WebRequest `
                -UseBasicParsing `
                -Uri $downloadUrl `
                -OutFile $archivePath `
                -TimeoutSec 90
            $downloadError = $null
            break
        } catch {
            $downloadError = $_
            Remove-Item -LiteralPath $archivePath -Force -ErrorAction SilentlyContinue
        }
    }
    if ($null -ne $downloadError -or -not (Test-Path -LiteralPath $archivePath -PathType Leaf)) {
        throw "Unable to download the pinned Latin Modern Math package from the verified CTAN mirrors. $downloadError"
    }
    Assert-FileHash $archivePath $expectedArchiveSha256 "Latin Modern Math package"

    $expandedRoot = Join-Path $tempRoot "expanded"
    Expand-Archive -LiteralPath $archivePath -DestinationPath $expandedRoot
    $sourceFont = Join-Path $expandedRoot "lm-math\opentype\latinmodern-math.otf"
    $sourceLicense = Join-Path $expandedRoot "lm-math\doc\GUST-FONT-LICENSE.txt"
    $sourceReadme = Join-Path $expandedRoot "lm-math\README"
    Assert-FileHash $sourceFont $expectedFontSha256 "Latin Modern Math font"
    Assert-FileHash $sourceLicense $expectedLicenseSha256 "GUST Font License"
    Assert-FileHash $sourceReadme $expectedReadmeSha256 "Latin Modern Math README"

    Copy-Item -LiteralPath $sourceFont -Destination $fontPath -Force
    Copy-Item -LiteralPath $sourceLicense -Destination $licensePath -Force
    Copy-Item -LiteralPath $sourceReadme -Destination $readmePath -Force

    $manifest = [ordered]@{
        schemaVersion = 1
        package = "lm-math"
        version = "1.959"
        source = "CTAN /fonts/lm-math.zip"
        license = "GUST Font License"
        archiveSha256 = $expectedArchiveSha256
        files = @(
            [ordered]@{
                file = "latinmodern-math.otf"
                sha256 = $expectedFontSha256
            },
            [ordered]@{
                file = "GUST-FONT-LICENSE.txt"
                sha256 = $expectedLicenseSha256
            },
            [ordered]@{
                file = "README-Latin-Modern-Math.txt"
                sha256 = $expectedReadmeSha256
            }
        )
    }
    $manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifestPath -Encoding UTF8

    Assert-FileHash $fontPath $expectedFontSha256 "Prepared Latin Modern Math font"
    Assert-FileHash $licensePath $expectedLicenseSha256 "Prepared GUST Font License"
    Assert-FileHash $readmePath $expectedReadmeSha256 "Prepared Latin Modern Math README"
    Write-Host "Prepared verified Latin Modern Math Office font: $fontPath"
} finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}
