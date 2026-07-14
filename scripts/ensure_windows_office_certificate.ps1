[CmdletBinding()]
param(
    [string]$VisualTeXPath
)

$ErrorActionPreference = "Stop"

function Find-VisualTeXCertificate {
    foreach ($candidate in @(
        (Join-Path $env:APPDATA "com.visualtex.studio\office\localhost-cert.pem"),
        (Join-Path $env:LOCALAPPDATA "com.visualtex.studio\office\localhost-cert.pem"),
        (Join-Path $env:LOCALAPPDATA "VisualTeX\office\localhost-cert.pem")
    )) {
        if (Test-Path $candidate) { return (Resolve-Path $candidate).Path }
    }
    return $null
}

function Resolve-VisualTeXExecutable {
    if (-not [string]::IsNullOrWhiteSpace($VisualTeXPath) -and (Test-Path $VisualTeXPath)) {
        return (Resolve-Path $VisualTeXPath).Path
    }
    $scriptRoot = Split-Path -Parent $PSScriptRoot
    $installedRoot = Split-Path -Parent $scriptRoot
    foreach ($candidate in @(
        (Join-Path $installedRoot "VisualTeX.exe"),
        (Join-Path $env:LOCALAPPDATA "Programs\VisualTeX\VisualTeX.exe"),
        (Join-Path $env:LOCALAPPDATA "VisualTeX\visualtex.exe")
    )) {
        if (Test-Path $candidate) { return (Resolve-Path $candidate).Path }
    }
    return $null
}

$certificatePath = Find-VisualTeXCertificate
if (-not $certificatePath) {
    $visualTeX = Resolve-VisualTeXExecutable
    if (-not $visualTeX) {
        throw "VisualTeX must be started once before the current-user Office HTTPS certificate can be trusted."
    }
    Start-Process -FilePath $visualTeX -ArgumentList "--office-background" -WindowStyle Hidden | Out-Null
    $deadline = [DateTimeOffset]::UtcNow.AddSeconds(20)
    do {
        Start-Sleep -Milliseconds 200
        $certificatePath = Find-VisualTeXCertificate
    } while (-not $certificatePath -and [DateTimeOffset]::UtcNow -lt $deadline)
    if (-not $certificatePath) {
        throw "VisualTeX did not create its Office HTTPS certificate within 20 seconds."
    }
}

$certificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new($certificatePath)
$modeKey = "HKCU:\Software\VisualTeX\OfficeIntegration"
$previousThumbprint = $null
if (Test-Path $modeKey) {
    $previousThumbprint = (Get-ItemProperty $modeKey -Name CertificateThumbprint -ErrorAction SilentlyContinue).CertificateThumbprint
}
if (-not [string]::IsNullOrWhiteSpace($previousThumbprint) -and
    $previousThumbprint -ne $certificate.Thumbprint) {
    foreach ($location in @("Cert:\CurrentUser\Root", "Cert:\CurrentUser\TrustedPeople")) {
        Get-ChildItem $location -ErrorAction SilentlyContinue |
            Where-Object { $_.Thumbprint -eq $previousThumbprint } |
            Remove-Item -Force -ErrorAction SilentlyContinue
    }
}

$alreadyTrusted = @(
    Get-ChildItem Cert:\CurrentUser\Root -ErrorAction SilentlyContinue |
        Where-Object { $_.Thumbprint -eq $certificate.Thumbprint }
).Count -gt 0
if (-not $alreadyTrusted) {
    $imported = Import-Certificate -FilePath $certificatePath -CertStoreLocation Cert:\CurrentUser\Root
    if ($null -eq $imported -or $imported.Thumbprint -ne $certificate.Thumbprint) {
        throw "VisualTeX Office HTTPS certificate import did not return the expected certificate."
    }
}

# Remove an earlier test/install copy from TrustedPeople. The leaf certificate
# is trusted through the current-user Root store so Schannel and Office agree.
Get-ChildItem Cert:\CurrentUser\TrustedPeople -ErrorAction SilentlyContinue |
    Where-Object { $_.Thumbprint -eq $certificate.Thumbprint } |
    Remove-Item -Force -ErrorAction SilentlyContinue

if (-not (Test-Path $modeKey)) { New-Item $modeKey -Force | Out-Null }
New-ItemProperty $modeKey -Name "CertificateThumbprint" -PropertyType String -Value $certificate.Thumbprint -Force | Out-Null
New-ItemProperty $modeKey -Name "CertificatePath" -PropertyType String -Value $certificatePath -Force | Out-Null
Write-Host "VisualTeX Office HTTPS certificate trusted for the current user: $($certificate.Thumbprint)"
