[CmdletBinding()]
param(
    [string]$VisualTeXPath,
    [string]$AppDataRoot,
    [ValidateRange(1, 65535)][int]$CompanionPort = 43127,
    [ValidateRange(1, 2147483647)][int]$ProtocolVersion = 1,
    [switch]$ResolveVisualTeXPathOnly
)

$ErrorActionPreference = "Stop"
$integrationKey = "HKCU:\Software\VisualTeX\OfficeIntegration"
$attemptedExecutablePaths = New-Object System.Collections.Generic.List[string]
$attemptedCertificatePaths = New-Object System.Collections.Generic.List[string]

function Add-Attempt([System.Collections.Generic.List[string]]$List, [string]$Value) {
    if ([string]::IsNullOrWhiteSpace($Value)) { return }
    if (-not $List.Contains($Value)) { [void]$List.Add($Value) }
}

function Get-IntegrationValue([string]$Name) {
    if (-not (Test-Path -LiteralPath $integrationKey)) { return $null }
    return (Get-ItemProperty -LiteralPath $integrationKey -Name $Name -ErrorAction SilentlyContinue).$Name
}

function Resolve-VisualTeXExecutable {
    $explicit = if ([string]::IsNullOrWhiteSpace($VisualTeXPath)) { $null } else { $VisualTeXPath.Trim().Trim('"') }
    if (-not [string]::IsNullOrWhiteSpace($explicit)) {
        if (Test-Path -LiteralPath $explicit -PathType Container) {
            $explicit = Join-Path $explicit "VisualTeX.exe"
        }
        Add-Attempt $attemptedExecutablePaths $explicit
        if (-not (Test-Path -LiteralPath $explicit -PathType Leaf)) {
            throw "The explicitly supplied VisualTeX executable does not exist: $explicit"
        }
        return (Resolve-Path -LiteralPath $explicit).Path
    }

    $registered = [string](Get-IntegrationValue "ExecutablePath")
    if (-not [string]::IsNullOrWhiteSpace($registered)) {
        $registered = $registered.Trim().Trim('"')
        Add-Attempt $attemptedExecutablePaths $registered
        if (Test-Path -LiteralPath $registered -PathType Leaf) {
            return (Resolve-Path -LiteralPath $registered).Path
        }
    }

    # Installed layout: <install-root>\scripts\ensure_windows_office_certificate.ps1
    # Therefore exactly one parent of $PSScriptRoot is the application root.
    $installedRoot = Split-Path -Parent $PSScriptRoot
    $installedCandidate = Join-Path $installedRoot "VisualTeX.exe"
    Add-Attempt $attemptedExecutablePaths $installedCandidate
    if (Test-Path -LiteralPath $installedCandidate -PathType Leaf) {
        return (Resolve-Path -LiteralPath $installedCandidate).Path
    }

    $attempts = if ($attemptedExecutablePaths.Count -gt 0) {
        ($attemptedExecutablePaths | ForEach-Object { "  - $_" }) -join [Environment]::NewLine
    } else {
        "  - <none>"
    }
    throw @"
Unable to resolve the installed VisualTeX.exe.
Explicit -VisualTeXPath: $(if ([string]::IsNullOrWhiteSpace($VisualTeXPath)) { '<not supplied>' } else { $VisualTeXPath })
Registry HKCU\Software\VisualTeX\OfficeIntegration\ExecutablePath: $(if ([string]::IsNullOrWhiteSpace($registered)) { '<missing>' } else { $registered })
Attempted paths:
$attempts
Pass the exact executable path, for example:
  -VisualTeXPath "D:\Softwares\visualtex\VisualTeX.exe"
"@
}

function Resolve-AppDataRoot([string]$ExecutablePath) {
    $explicitRoot = if ([string]::IsNullOrWhiteSpace($AppDataRoot)) { $null } else { $AppDataRoot.Trim().Trim('"') }
    if (-not [string]::IsNullOrWhiteSpace($explicitRoot)) {
        return [IO.Path]::GetFullPath($explicitRoot)
    }

    $registeredRoot = [string](Get-IntegrationValue "AppDataRoot")
    if (-not [string]::IsNullOrWhiteSpace($registeredRoot)) {
        return [IO.Path]::GetFullPath($registeredRoot.Trim().Trim('"'))
    }

    $registeredCertificate = [string](Get-IntegrationValue "CertificatePath")
    if (-not [string]::IsNullOrWhiteSpace($registeredCertificate)) {
        $officeRoot = Split-Path -Parent $registeredCertificate.Trim().Trim('"')
        if (-not [string]::IsNullOrWhiteSpace($officeRoot)) {
            return Split-Path -Parent $officeRoot
        }
    }

    # These are certificate/data candidates only. They are never used to locate
    # VisualTeX.exe and are recorded in diagnostics below.
    foreach ($candidate in @(
        (Join-Path $env:APPDATA "com.visualtex.studio"),
        (Join-Path $env:LOCALAPPDATA "com.visualtex.studio")
    )) {
        Add-Attempt $attemptedCertificatePaths (Join-Path $candidate "office\localhost-cert.pem")
        if (Test-Path -LiteralPath $candidate -PathType Container) { return $candidate }
    }

    # The background process will write the exact root to the registry. Return
    # the normal roaming root only as a polling candidate until that happens.
    return (Join-Path $env:APPDATA "com.visualtex.studio")
}

function Resolve-CertificatePath([string]$ResolvedAppDataRoot) {
    $registered = [string](Get-IntegrationValue "CertificatePath")
    if (-not [string]::IsNullOrWhiteSpace($registered)) {
        $registered = $registered.Trim().Trim('"')
        Add-Attempt $attemptedCertificatePaths $registered
        if (Test-Path -LiteralPath $registered -PathType Leaf) {
            return (Resolve-Path -LiteralPath $registered).Path
        }
    }

    $candidate = Join-Path $ResolvedAppDataRoot "office\localhost-cert.pem"
    Add-Attempt $attemptedCertificatePaths $candidate
    if (Test-Path -LiteralPath $candidate -PathType Leaf) {
        return (Resolve-Path -LiteralPath $candidate).Path
    }
    return $null
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

function Add-CertificateToCurrentUserStore {
    param(
        [Security.Cryptography.X509Certificates.StoreName]$StoreName,
        [Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )
    $store = [Security.Cryptography.X509Certificates.X509Store]::new(
        $StoreName,
        [Security.Cryptography.X509Certificates.StoreLocation]::CurrentUser
    )
    $store.Open([Security.Cryptography.X509Certificates.OpenFlags]::ReadWrite)
    try {
        $store.Add($Certificate)
    } finally {
        $store.Close()
    }
}

$resolvedExecutable = Resolve-VisualTeXExecutable
if ($ResolveVisualTeXPathOnly) {
    Write-Output $resolvedExecutable
    return
}

$resolvedAppDataRoot = Resolve-AppDataRoot $resolvedExecutable
$certificatePath = Resolve-CertificatePath $resolvedAppDataRoot
$bootstrapProcess = $null

try {
    # Installation/bootstrap must never start the long-lived Office companion.
    # The dedicated mode creates the certificate and install.json, writes the
    # exact executable/app-data registry configuration, validates the result,
    # creates no WebView, schedules no warmup, and exits on its own.
    $bootstrapProcess = Start-Process `
        -FilePath $resolvedExecutable `
        -ArgumentList "--office-bootstrap" `
        -WindowStyle Hidden `
        -PassThru
    if (-not $bootstrapProcess.WaitForExit(30000)) {
        Stop-Process -Id $bootstrapProcess.Id -Force -ErrorAction SilentlyContinue
        throw "VisualTeX Office bootstrap did not exit within 30 seconds and was terminated. Executable='$resolvedExecutable'; PID=$($bootstrapProcess.Id). Check %LOCALAPPDATA%\VisualTeX\logs\app-lifecycle.log and %LOCALAPPDATA%\VisualTeX\office\logs\startup.log."
    }
    if ($bootstrapProcess.ExitCode -ne 0) {
        throw "VisualTeX Office bootstrap failed. Executable='$resolvedExecutable'; PID=$($bootstrapProcess.Id); ExitCode=$($bootstrapProcess.ExitCode). Check %LOCALAPPDATA%\VisualTeX\logs\app-lifecycle.log and %LOCALAPPDATA%\VisualTeX\office\logs\startup.log."
    }
} finally {
    if ($null -ne $bootstrapProcess) {
        $bootstrapPid = $bootstrapProcess.Id
        $bootstrapProcess.Dispose()
        if (Get-Process -Id $bootstrapPid -ErrorAction SilentlyContinue) {
            Stop-Process -Id $bootstrapPid -Force -ErrorAction SilentlyContinue
            throw "VisualTeX Office bootstrap left a residual process: PID=$bootstrapPid"
        }
    }
}

$registeredRoot = [string](Get-IntegrationValue "AppDataRoot")
if (-not [string]::IsNullOrWhiteSpace($registeredRoot)) {
    $resolvedAppDataRoot = $registeredRoot.Trim().Trim('"')
}
$certificatePath = Resolve-CertificatePath $resolvedAppDataRoot
if (-not $certificatePath) {
    $certificateAttempts = ($attemptedCertificatePaths | ForEach-Object { "  - $_" }) -join [Environment]::NewLine
    throw @"
VisualTeX Office bootstrap exited successfully, but the generated HTTPS certificate could not be resolved.
Executable: $resolvedExecutable
AppDataRoot: $resolvedAppDataRoot
Certificate paths checked:
$certificateAttempts
Check %LOCALAPPDATA%\VisualTeX\logs\app-lifecycle.log and startup.log under the registered AppDataRoot\office\logs directory.
"@
}

$certificate = [Security.Cryptography.X509Certificates.X509Certificate2]::new($certificatePath)
$previousThumbprint = [string](Get-IntegrationValue "CertificateThumbprint")
if (-not [string]::IsNullOrWhiteSpace($previousThumbprint) -and
    $previousThumbprint -ne $certificate.Thumbprint) {
    Remove-CertificateFromCurrentUserStore Root $previousThumbprint
    Remove-CertificateFromCurrentUserStore TrustedPeople $previousThumbprint
}

if (-not (Test-CertificateInCurrentUserStore Root $certificate.Thumbprint)) {
    try {
        Add-CertificateToCurrentUserStore Root $certificate
    } catch {
        throw "Unable to add the VisualTeX Office HTTPS certificate to the current-user Root store: $($_.Exception.Message)"
    }
}
if (-not (Test-CertificateInCurrentUserStore Root $certificate.Thumbprint)) {
    throw "VisualTeX Office HTTPS certificate was not added to the current-user Root store: $($certificate.Thumbprint)"
}
Remove-CertificateFromCurrentUserStore TrustedPeople $certificate.Thumbprint

if (-not (Test-Path -LiteralPath $integrationKey)) { New-Item -Path $integrationKey -Force | Out-Null }
New-ItemProperty -LiteralPath $integrationKey -Name "ExecutablePath" -PropertyType String -Value $resolvedExecutable -Force | Out-Null
New-ItemProperty -LiteralPath $integrationKey -Name "AppDataRoot" -PropertyType String -Value $resolvedAppDataRoot -Force | Out-Null
New-ItemProperty -LiteralPath $integrationKey -Name "CertificatePath" -PropertyType String -Value $certificatePath -Force | Out-Null
New-ItemProperty -LiteralPath $integrationKey -Name "CertificateThumbprint" -PropertyType String -Value $certificate.Thumbprint -Force | Out-Null
New-ItemProperty -LiteralPath $integrationKey -Name "CompanionPort" -PropertyType DWord -Value $CompanionPort -Force | Out-Null
New-ItemProperty -LiteralPath $integrationKey -Name "ProtocolVersion" -PropertyType DWord -Value $ProtocolVersion -Force | Out-Null

Write-Host "VisualTeX Office HTTPS certificate trusted for the current user."
Write-Host "ExecutablePath=$resolvedExecutable"
Write-Host "AppDataRoot=$resolvedAppDataRoot"
Write-Host "CertificatePath=$certificatePath"
Write-Host "CertificateThumbprint=$($certificate.Thumbprint)"
Write-Host "CompanionPort=$CompanionPort"
Write-Host "ProtocolVersion=$ProtocolVersion"
