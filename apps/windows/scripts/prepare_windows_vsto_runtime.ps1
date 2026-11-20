[CmdletBinding()]
param(
    [string]$OutputPath,
    [string]$ManifestPath,
    [string]$SourcePath,
    [string]$DownloadUrl = "https://download.microsoft.com/download/5/d/2/5d24f8f8-efbb-4b63-aa33-3785e3104713/vstor_redist.exe"
)

$ErrorActionPreference = "Stop"
$expectedSha256 = "CFE1A40BBE4A50022DB2164ABDB0154984E2CECB761A23CDC81CB5754F6E0A18"
$expectedVersion = "10.0.60917.00"
$expectedCompany = "Microsoft Corporation"
$expectedProduct = "Microsoft Visual Studio Tools for Office Runtime 2010 Redistributable"

$root = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path $root "src-tauri\resources\windows-office\vstor_redist.exe"
}
if ([string]::IsNullOrWhiteSpace($ManifestPath)) {
    $ManifestPath = Join-Path (Split-Path -Parent $OutputPath) "vstor_redist.sha256.json"
}

function Assert-MicrosoftVstoRuntimePackage([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "The VSTO Runtime redistributable is missing: $Path"
    }
    $item = Get-Item -LiteralPath $Path
    $hash = (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToUpperInvariant()
    if ($hash -ne $expectedSha256) {
        throw "Unexpected VSTO Runtime SHA-256: $hash; expected $expectedSha256."
    }
    $signature = Get-AuthenticodeSignature -FilePath $Path
    if ($signature.Status -ne [System.Management.Automation.SignatureStatus]::Valid) {
        throw "The VSTO Runtime Authenticode signature is not valid: $($signature.Status); $($signature.StatusMessage)"
    }
    if ($null -eq $signature.SignerCertificate -or
        $signature.SignerCertificate.Subject -notmatch "Microsoft Corporation") {
        throw "The VSTO Runtime signer is not Microsoft Corporation: $($signature.SignerCertificate.Subject)"
    }
    $version = $item.VersionInfo
    if ([string]$version.CompanyName -ne $expectedCompany -or
        [string]$version.ProductName -ne $expectedProduct -or
        [string]$version.ProductVersion -ne $expectedVersion) {
        throw "Unexpected VSTO Runtime metadata: Company='$($version.CompanyName)', Product='$($version.ProductName)', Version='$($version.ProductVersion)'."
    }
    return [pscustomobject]@{
        fileName = $item.Name
        size = $item.Length
        sha256 = $hash
        version = [string]$version.ProductVersion
        company = [string]$version.CompanyName
        product = [string]$version.ProductName
        signerSubject = [string]$signature.SignerCertificate.Subject
        signerThumbprint = [string]$signature.SignerCertificate.Thumbprint
    }
}

$destinationDirectory = Split-Path -Parent $OutputPath
New-Item -Path $destinationDirectory -ItemType Directory -Force | Out-Null

$resolvedSource = $null
if (-not [string]::IsNullOrWhiteSpace($SourcePath)) {
    if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) {
        throw "The explicitly supplied VSTO Runtime package does not exist: $SourcePath"
    }
    $resolvedSource = (Resolve-Path -LiteralPath $SourcePath).Path
} elseif (-not [string]::IsNullOrWhiteSpace($env:VISUALTEX_VSTO_RUNTIME_PATH)) {
    if (-not (Test-Path -LiteralPath $env:VISUALTEX_VSTO_RUNTIME_PATH -PathType Leaf)) {
        throw "VISUALTEX_VSTO_RUNTIME_PATH does not point to a file: $($env:VISUALTEX_VSTO_RUNTIME_PATH)"
    }
    $resolvedSource = (Resolve-Path -LiteralPath $env:VISUALTEX_VSTO_RUNTIME_PATH).Path
} elseif (Test-Path -LiteralPath $OutputPath -PathType Leaf) {
    try {
        $metadata = Assert-MicrosoftVstoRuntimePackage $OutputPath
        $resolvedSource = (Resolve-Path -LiteralPath $OutputPath).Path
    } catch {
        Remove-Item -LiteralPath $OutputPath -Force -ErrorAction SilentlyContinue
    }
}

$tempFile = $null
try {
    if ([string]::IsNullOrWhiteSpace($resolvedSource)) {
        $tempFile = Join-Path $env:TEMP ("visualtex-vstor-redist-{0}.exe" -f ([Guid]::NewGuid().ToString("N")))
        Write-Host "Downloading the Microsoft VSTO Runtime redistributable from the pinned official URL..."
        Invoke-WebRequest -UseBasicParsing -Uri $DownloadUrl -OutFile $tempFile
        $resolvedSource = $tempFile
    }

    $metadata = Assert-MicrosoftVstoRuntimePackage $resolvedSource
    if (-not [string]::Equals(
        [IO.Path]::GetFullPath($resolvedSource),
        [IO.Path]::GetFullPath($OutputPath),
        [StringComparison]::OrdinalIgnoreCase)) {
        Copy-Item -LiteralPath $resolvedSource -Destination $OutputPath -Force
        $metadata = Assert-MicrosoftVstoRuntimePackage $OutputPath
    }

    $manifest = [ordered]@{
        schemaVersion = 1
        source = "Microsoft Download Center"
        downloadUrl = $DownloadUrl
        preparedAt = [DateTimeOffset]::Now.ToString("o")
        package = $metadata
    }
    $manifest | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $ManifestPath -Encoding UTF8
    Write-Host "Prepared bundled VSTO Runtime: $OutputPath"
    Write-Host "SHA-256: $($metadata.sha256)"
    Write-Host "Version: $($metadata.version)"
} finally {
    if (-not [string]::IsNullOrWhiteSpace($tempFile)) {
        Remove-Item -LiteralPath $tempFile -Force -ErrorAction SilentlyContinue
    }
}
