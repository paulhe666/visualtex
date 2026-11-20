[CmdletBinding()]
param(
    [int]$TimeoutSeconds = 10,
    [string]$LogPath
)

$ErrorActionPreference = "Stop"
$modeKey = "HKCU:\Software\VisualTeX\OfficeIntegration"
$logRoot = Join-Path $env:LOCALAPPDATA "VisualTeX\office\install-logs"
New-Item -Path $logRoot -ItemType Directory -Force | Out-Null
if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Join-Path $logRoot ("certificate-remove-{0}.log" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
}

function Write-CertificateLog([string]$Message) {
    $line = "[{0:O}] {1}" -f (Get-Date), $Message
    Write-Host $line
    Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
}

function Test-CertificatePresent([string]$Thumbprint) {
    if ([string]::IsNullOrWhiteSpace($Thumbprint)) { return $false }
    return @(
        Get-ChildItem Cert:\CurrentUser\Root -ErrorAction SilentlyContinue |
            Where-Object { $_.Thumbprint -eq $Thumbprint }
    ).Count -gt 0
}

$thumbprint = $null
if (Test-Path -LiteralPath $modeKey) {
    $thumbprint = [string](
        Get-ItemProperty -LiteralPath $modeKey -Name CertificateThumbprint -ErrorAction SilentlyContinue
    ).CertificateThumbprint
}

try {
    if ([string]::IsNullOrWhiteSpace($thumbprint)) {
        Write-CertificateLog "No VisualTeX certificate thumbprint is registered; nothing to remove."
    } else {
        $certificateRegistryPath = "HKCU:\Software\Microsoft\SystemCertificates\Root\Certificates\$thumbprint"
        $certificateRegistryNativePath = "HKCU\Software\Microsoft\SystemCertificates\Root\Certificates\$thumbprint"
        $registryPresent = Test-Path -LiteralPath $certificateRegistryPath
        $providerPresent = Test-CertificatePresent $thumbprint

        if (-not $registryPresent -and -not $providerPresent) {
            Write-CertificateLog "Certificate $thumbprint is already absent from CurrentUser\\Root."
        } else {
            $regExe = Join-Path $env:WINDIR "System32\reg.exe"
            if (-not (Test-Path -LiteralPath $regExe -PathType Leaf)) {
                throw "reg.exe was not found at $regExe"
            }

            $stdoutPath = $LogPath + ".stdout.txt"
            $stderrPath = $LogPath + ".stderr.txt"
            Write-CertificateLog "Removing certificate $thumbprint from the CurrentUser Root registry store; timeout=${TimeoutSeconds}s."
            $process = Start-Process -FilePath $regExe `
                -ArgumentList @("delete", $certificateRegistryNativePath, "/f") `
                -WindowStyle Hidden `
                -RedirectStandardOutput $stdoutPath `
                -RedirectStandardError $stderrPath `
                -PassThru
            try {
                if (-not $process.WaitForExit([Math]::Max(1, $TimeoutSeconds) * 1000)) {
                    Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
                    throw "Certificate registry cleanup exceeded the ${TimeoutSeconds}-second timeout."
                }
                $process.Refresh()
                $exitCode = $process.ExitCode
            } finally {
                $process.Dispose()
            }

            Write-CertificateLog "reg.exe exited with code $exitCode."
            Start-Sleep -Milliseconds 300
            $registryStillPresent = Test-Path -LiteralPath $certificateRegistryPath
            $providerStillPresent = Test-CertificatePresent $thumbprint
            if ($registryStillPresent -or $providerStillPresent) {
                $stdout = if (Test-Path -LiteralPath $stdoutPath) {
                    (Get-Content -LiteralPath $stdoutPath -Raw -ErrorAction SilentlyContinue).Trim()
                } else { "" }
                $stderr = if (Test-Path -LiteralPath $stderrPath) {
                    (Get-Content -LiteralPath $stderrPath -Raw -ErrorAction SilentlyContinue).Trim()
                } else { "" }
                throw "Certificate $thumbprint remains after registry cleanup. RegistryPresent=$registryStillPresent ProviderPresent=$providerStillPresent ExitCode=$exitCode stdout='$stdout' stderr='$stderr'"
            }
            Write-CertificateLog "Certificate $thumbprint was removed from CurrentUser\\Root."
        }
    }

    if (Test-Path -LiteralPath $modeKey) {
        Remove-ItemProperty -LiteralPath $modeKey -Name CertificateThumbprint -Force -ErrorAction SilentlyContinue
        Remove-ItemProperty -LiteralPath $modeKey -Name CertificatePath -Force -ErrorAction SilentlyContinue
    }
    Write-CertificateLog "VisualTeX current-user Office HTTPS certificate trust cleanup completed."
    exit 0
} catch {
    Write-CertificateLog ("FAILED: " + $_.Exception.ToString())
    exit 1
}
