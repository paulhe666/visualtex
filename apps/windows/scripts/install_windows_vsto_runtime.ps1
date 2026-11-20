[CmdletBinding()]
param(
    [string]$RuntimeInstallerPath,
    [string]$ManifestPath,
    [switch]$CheckOnly,
    [string]$LogPath
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$logRoot = Join-Path $env:LOCALAPPDATA "VisualTeX\office\install-logs"
New-Item -Path $logRoot -ItemType Directory -Force | Out-Null
if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Join-Path $logRoot ("vsto-runtime-{0}.log" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
}
if ([string]::IsNullOrWhiteSpace($RuntimeInstallerPath)) {
    $RuntimeInstallerPath = Join-Path $root "windows-office\vstor_redist.exe"
}
if ([string]::IsNullOrWhiteSpace($ManifestPath)) {
    $ManifestPath = Join-Path (Split-Path -Parent $RuntimeInstallerPath) "vstor_redist.sha256.json"
}

function Get-RegistryView([string]$Architecture) {
    if ($Architecture -eq "x86") { return [Microsoft.Win32.RegistryView]::Registry32 }
    return [Microsoft.Win32.RegistryView]::Registry64
}

function Get-VstoRuntimeState {
    $subKey = "SOFTWARE\Microsoft\VSTO Runtime Setup\v4R"
    foreach ($architecture in @("x86", "x64")) {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
            [Microsoft.Win32.RegistryHive]::LocalMachine,
            (Get-RegistryView $architecture))
        try {
            $key = $baseKey.OpenSubKey($subKey, $false)
            if ($null -eq $key) { continue }
            try {
                $install = $key.GetValue("Install", $null)
                $clr40 = $key.GetValue("VSTORFeature_CLR40", $null)
                $version = $key.GetValue("Version", $null)
                $installed =
                    ($null -ne $install -and [int]$install -eq 1) -or
                    ($null -ne $clr40 -and [int]$clr40 -eq 1) -or
                    (-not [string]::IsNullOrWhiteSpace([string]$version))
                if ($installed) {
                    return [pscustomobject]@{
                        installed = $true
                        registryView = $architecture
                        install = $install
                        clr40 = $clr40
                        version = [string]$version
                        key = "HKLM\$subKey"
                    }
                }
            } finally {
                $key.Dispose()
            }
        } finally {
            $baseKey.Dispose()
        }
    }
    return [pscustomobject]@{
        installed = $false
        registryView = "none"
        install = $null
        clr40 = $null
        version = ""
        key = "HKLM\$subKey"
    }
}

function Assert-BundledRuntimePackage {
    if (-not (Test-Path -LiteralPath $RuntimeInstallerPath -PathType Leaf)) {
        throw "The bundled Microsoft VSTO Runtime installer is missing: $RuntimeInstallerPath"
    }
    if (-not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) {
        throw "The bundled Microsoft VSTO Runtime manifest is missing: $ManifestPath"
    }
    $manifest = Get-Content -LiteralPath $ManifestPath -Raw | ConvertFrom-Json
    $expectedHash = ([string]$manifest.package.sha256).ToUpperInvariant()
    if ([string]::IsNullOrWhiteSpace($expectedHash)) {
        throw "The bundled VSTO Runtime manifest does not contain package.sha256."
    }
    $actualHash = (Get-FileHash -LiteralPath $RuntimeInstallerPath -Algorithm SHA256).Hash.ToUpperInvariant()
    if ($actualHash -ne $expectedHash) {
        throw "The bundled VSTO Runtime SHA-256 does not match its manifest: $actualHash != $expectedHash"
    }
    $signature = Get-AuthenticodeSignature -FilePath $RuntimeInstallerPath
    if ($signature.Status -ne [System.Management.Automation.SignatureStatus]::Valid -or
        $null -eq $signature.SignerCertificate -or
        $signature.SignerCertificate.Subject -notmatch "Microsoft Corporation") {
        throw "The bundled VSTO Runtime Authenticode signature is invalid or is not signed by Microsoft Corporation. Status=$($signature.Status); Signer=$($signature.SignerCertificate.Subject)"
    }
    $version = (Get-Item -LiteralPath $RuntimeInstallerPath).VersionInfo
    if ([string]$version.CompanyName -ne "Microsoft Corporation" -or
        [string]$version.ProductName -ne "Microsoft Visual Studio Tools for Office Runtime 2010 Redistributable") {
        throw "The bundled VSTO Runtime metadata is unexpected: Company='$($version.CompanyName)', Product='$($version.ProductName)'."
    }
    return [pscustomobject]@{
        sha256 = $actualHash
        version = [string]$version.ProductVersion
        signer = [string]$signature.SignerCertificate.Subject
    }
}

$transcriptStarted = $false
try {
    Start-Transcript -Path $LogPath -Force | Out-Null
    $transcriptStarted = $true
    $state = Get-VstoRuntimeState
    if ($state.installed) {
        Write-Host "VSTO Runtime is already installed: $($state.key); View=$($state.registryView); Version=$($state.version); CLR40=$($state.clr40)."
        exit 0
    }
    if ($CheckOnly) {
        Write-Host "VSTO Runtime is not installed in either x86 or x64 registry view."
        exit 2
    }

    $package = Assert-BundledRuntimePackage
    Write-Host "Installing bundled Microsoft VSTO Runtime version $($package.version)."
    Write-Host "Package SHA-256: $($package.sha256)"
    Write-Host "Package signer: $($package.signer)"

    $process = $null
    try {
        $process = Start-Process `
            -FilePath $RuntimeInstallerPath `
            -ArgumentList @("/quiet", "/norestart") `
            -Verb RunAs `
            -Wait `
            -PassThru
    } catch {
        throw "Unable to start the Microsoft VSTO Runtime installer with administrator privileges. The user may have cancelled the UAC prompt. $($_.Exception.Message)"
    }

    Write-Host "Microsoft VSTO Runtime installer exit code: $($process.ExitCode)"
    if ($process.ExitCode -notin @(0, 1641, 3010)) {
        throw "Microsoft VSTO Runtime installation failed with exit code $($process.ExitCode)."
    }

    $deadline = [DateTimeOffset]::UtcNow.AddSeconds(30)
    do {
        $state = Get-VstoRuntimeState
        if ($state.installed) { break }
        Start-Sleep -Milliseconds 250
    } while ([DateTimeOffset]::UtcNow -lt $deadline)

    if (-not $state.installed) {
        throw "The Microsoft VSTO Runtime installer returned success, but the v4R registry state was not detected within 30 seconds."
    }
    Write-Host "Microsoft VSTO Runtime installed successfully: View=$($state.registryView); Version=$($state.version); CLR40=$($state.clr40)."
    exit 0
} catch {
    Write-Error $_.Exception.Message
    exit 1
} finally {
    if ($transcriptStarted) {
        try { Stop-Transcript | Out-Null } catch { }
    }
}
