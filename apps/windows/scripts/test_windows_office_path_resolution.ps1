[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$windowsRoot = if (-not [string]::IsNullOrWhiteSpace($env:SystemRoot)) { $env:SystemRoot } else { $env:WINDIR }
if ([string]::IsNullOrWhiteSpace($windowsRoot)) {
    throw "SystemRoot/WINDIR is unavailable; cannot locate Windows PowerShell"
}
$powerShellExe = Join-Path $windowsRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
if (-not (Test-Path -LiteralPath $powerShellExe -PathType Leaf)) {
    throw "Windows PowerShell is missing: $powerShellExe"
}
$sourceScript = Join-Path $PSScriptRoot "ensure_windows_office_certificate.ps1"
if (-not (Test-Path -LiteralPath $sourceScript -PathType Leaf)) {
    throw "Certificate script is missing: $sourceScript"
}

function Unicode-Text([int[]]$CodePoints) {
    return -join @($CodePoints | ForEach-Object { [char]$_ })
}

$unicodePathLabel = Unicode-Text @(0x8DEF, 0x5F84, 0x56DE, 0x5F52)
$unicodeCustomLabel = Unicode-Text @(0x81EA, 0x5B9A, 0x4E49, 0x76EE, 0x5F55)
$unicodeInstallLabel = Unicode-Text @(0x5B89, 0x88C5, 0x6839, 0x76EE, 0x5F55)
$unicodeRegistryLabel = Unicode-Text @(0x6CE8, 0x518C, 0x8868, 0x76EE, 0x5F55)

$registryKey = "HKCU:\Software\VisualTeX\OfficeIntegration"
$registryExisted = Test-Path -LiteralPath $registryKey
$previousExecutable = $null
$previousHadExecutable = $false
if ($registryExisted) {
    $existing = Get-ItemProperty -LiteralPath $registryKey -Name ExecutablePath -ErrorAction SilentlyContinue
    if ($null -ne $existing) {
        $previousHadExecutable = $true
        $previousExecutable = [string]$existing.ExecutablePath
    }
}

$tempRoot = Join-Path $env:TEMP ("VisualTeX Office $unicodePathLabel " + [Guid]::NewGuid().ToString("N"))
$defaultRoot = $null
New-Item -Path $tempRoot -ItemType Directory -Force | Out-Null

function New-FakeInstall([string]$Root) {
    New-Item -Path $Root -ItemType Directory -Force | Out-Null
    $executable = Join-Path $Root "VisualTeX.exe"
    Set-Content -LiteralPath $executable -Value "path-resolution-smoke" -Encoding ASCII
    return (Resolve-Path -LiteralPath $executable).Path
}

function Invoke-Resolver {
    param(
        [string]$ScriptPath,
        [string]$ExpectedPath,
        [string]$ExplicitPath
    )
    $arguments = @(
        "-NoProfile",
        "-NonInteractive",
        "-ExecutionPolicy", "Bypass",
        "-File", $ScriptPath,
        "-ResolveVisualTeXPathOnly"
    )
    if (-not [string]::IsNullOrWhiteSpace($ExplicitPath)) {
        $arguments += @("-VisualTeXPath", $ExplicitPath)
    }
    $output = & $powerShellExe @arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Resolver failed with exit code $LASTEXITCODE for expected path '$ExpectedPath'. Output: $($output -join [Environment]::NewLine)"
    }
    $actual = [string]($output | Select-Object -Last 1)
    if (-not [string]::Equals(
        [IO.Path]::GetFullPath($actual.Trim()),
        [IO.Path]::GetFullPath($ExpectedPath),
        [StringComparison]::OrdinalIgnoreCase)) {
        throw "Resolved '$actual'; expected '$ExpectedPath'."
    }
    Write-Host "[PASS] $ExpectedPath"
}

try {
    if (-not (Test-Path -LiteralPath $registryKey)) {
        New-Item -Path $registryKey -Force | Out-Null
    }

    $defaultRoot = Join-Path $env:LOCALAPPDATA ("Programs\VisualTeX Path Smoke " + [Guid]::NewGuid().ToString("N"))
    $defaultExe = New-FakeInstall $defaultRoot
    Invoke-Resolver $sourceScript $defaultExe $defaultExe

    $customRoot = Join-Path $tempRoot ("C drive spaces $unicodeCustomLabel")
    $customExe = New-FakeInstall $customRoot
    Invoke-Resolver $sourceScript $customExe $customExe
    Invoke-Resolver $sourceScript $customExe $customRoot

    $layoutRoot = Join-Path $tempRoot $unicodeInstallLabel
    $layoutExe = New-FakeInstall $layoutRoot
    $layoutScripts = Join-Path $layoutRoot "scripts"
    New-Item -Path $layoutScripts -ItemType Directory -Force | Out-Null
    $copiedScript = Join-Path $layoutScripts "ensure_windows_office_certificate.ps1"
    Copy-Item -LiteralPath $sourceScript -Destination $copiedScript -Force
    Remove-ItemProperty -LiteralPath $registryKey -Name ExecutablePath -ErrorAction SilentlyContinue
    Invoke-Resolver $copiedScript $layoutExe ""

    $registryRoot = Join-Path $tempRoot $unicodeRegistryLabel
    $registryExe = New-FakeInstall $registryRoot
    New-ItemProperty -LiteralPath $registryKey -Name ExecutablePath -PropertyType String -Value $registryExe -Force | Out-Null
    $unrelatedRoot = Join-Path $tempRoot "unrelated\scripts"
    New-Item -Path $unrelatedRoot -ItemType Directory -Force | Out-Null
    $unrelatedScript = Join-Path $unrelatedRoot "ensure_windows_office_certificate.ps1"
    Copy-Item -LiteralPath $sourceScript -Destination $unrelatedScript -Force
    Invoke-Resolver $unrelatedScript $registryExe ""

    if (Test-Path -LiteralPath "D:\") {
        $dRoot = "D:\VisualTeX Path Smoke $unicodeCustomLabel $([Guid]::NewGuid().ToString('N'))"
        try {
            $dExe = New-FakeInstall $dRoot
            Invoke-Resolver $sourceScript $dExe $dExe
        } finally {
            Remove-Item -LiteralPath $dRoot -Recurse -Force -ErrorAction SilentlyContinue
        }
    } else {
        Write-Host "[SKIP] D drive is not available on this machine."
    }

    Write-Host "VisualTeX executable path resolution regression passed."
} finally {
    if (-not (Test-Path -LiteralPath $registryKey)) {
        New-Item -Path $registryKey -Force | Out-Null
    }
    if ($previousHadExecutable) {
        New-ItemProperty -LiteralPath $registryKey -Name ExecutablePath -PropertyType String -Value $previousExecutable -Force | Out-Null
    } else {
        Remove-ItemProperty -LiteralPath $registryKey -Name ExecutablePath -ErrorAction SilentlyContinue
    }
    if (-not $registryExisted) {
        $remaining = Get-ItemProperty -LiteralPath $registryKey -ErrorAction SilentlyContinue
        $userValues = @($remaining.PSObject.Properties | Where-Object { $_.Name -notmatch '^PS' })
        if ($userValues.Count -eq 0) {
            Remove-Item -LiteralPath $registryKey -Force -ErrorAction SilentlyContinue
        }
    }
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
    if (-not [string]::IsNullOrWhiteSpace($defaultRoot)) {
        Remove-Item -LiteralPath $defaultRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

