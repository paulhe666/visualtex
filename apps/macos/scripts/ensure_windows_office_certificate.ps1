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
function Remove-CertificateFromCurrentUserStore {
    param(
        [Security.Cryptography.X509Certificates.StoreName]$StoreName,
        [string]$Thumbprint
    )
    if ([string]::IsNullOrWhiteSpace($Thumbprint)) { return }
    $store = [Security.Cryptography.X509Certificates.X509Store]::new(
        $StoreName,
        [Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser
    )
    $store.Open([Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    try {
        foreach ($existing in @($store.Certificates | Where-Object { $_.Thumbprint -eq $Thumbprint })) {
            $store.Remove($existing)
        }
    } finally {
        $store.Close()
    }
}

function Test-CertificateInCurrentUserStore {
    param(
        [Security.Cryptography.X509Certificates.StoreName]$StoreName,
        [string]$Thumbprint
    )
    $store = [Security.Cryptography.X509Certificates.X509Store]::new(
        $StoreName,
        [Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser
    )
    $store.Open([Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly)
    try {
        return @($store.Certificates | Where-Object { $_.Thumbprint -eq $Thumbprint }).Count -gt 0
    } finally {
        $store.Close()
    }
}

if (-not [string]::IsNullOrWhiteSpace($previousThumbprint) -and
    $previousThumbprint -ne $certificate.Thumbprint) {
    Remove-CertificateFromCurrentUserStore Root $previousThumbprint
    Remove-CertificateFromCurrentUserStore TrustedPeople $previousThumbprint
}

$alreadyTrusted = Test-CertificateInCurrentUserStore Root $certificate.Thumbprint
if (-not $alreadyTrusted) {
    & certutil.exe -user -f -addstore Root $certificatePath | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "certutil.exe failed to add the VisualTeX Office HTTPS certificate to the current-user Root store (exit code $LASTEXITCODE)."
    }
    if (-not (Test-CertificateInCurrentUserStore Root $certificate.Thumbprint)) {
        throw "VisualTeX Office HTTPS certificate was not added to the current-user Root store."
    }
}

# Remove an earlier test/install copy from TrustedPeople. The leaf certificate
# is trusted through the current-user Root store so Schannel and Office agree.
Remove-CertificateFromCurrentUserStore TrustedPeople $certificate.Thumbprint

if (-not (Test-Path $modeKey)) { New-Item $modeKey -Force | Out-Null }
New-ItemProperty $modeKey -Name "CertificateThumbprint" -PropertyType String -Value $certificate.Thumbprint -Force | Out-Null
New-ItemProperty $modeKey -Name "CertificatePath" -PropertyType String -Value $certificatePath -Force | Out-Null
Write-Host "VisualTeX Office HTTPS certificate trusted for the current user: $($certificate.Thumbprint)"
