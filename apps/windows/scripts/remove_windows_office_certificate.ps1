[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$modeKey = "HKCU:\Software\VisualTeX\OfficeIntegration"
$thumbprint = $null
if (Test-Path $modeKey) {
    $thumbprint = (Get-ItemProperty $modeKey -Name CertificateThumbprint -ErrorAction SilentlyContinue).CertificateThumbprint
}

if (-not [string]::IsNullOrWhiteSpace($thumbprint)) {
    $store = [Security.Cryptography.X509Certificates.X509Store]::new(
        [Security.Cryptography.X509Certificates.StoreName]::Root,
        [Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser)
    try {
        $store.Open([Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
        $matches = @($store.Certificates | Where-Object { $_.Thumbprint -eq $thumbprint })
        foreach ($certificate in $matches) { $store.Remove($certificate) }
    } finally {
        $store.Close()
    }
}

if (Test-Path $modeKey) {
    Remove-ItemProperty $modeKey -Name CertificateThumbprint -ErrorAction SilentlyContinue
    Remove-ItemProperty $modeKey -Name CertificatePath -ErrorAction SilentlyContinue
}
Write-Host "VisualTeX current-user Office HTTPS certificate trust removed."
