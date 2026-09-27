[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InstallRoot,
    [ValidateSet('Check', 'Stop', 'RepairRegistration')][string]$Mode = 'Check'
)

# Invoked from the installer's temporary plugin directory, not from the old
# installation. Check never changes processes, files or registry values.
# Exit codes: 0 = clear; 10 = private runtime present; 11 = stop incomplete;
# 20 = inspection failed. Never treat a failed inspection as a running runtime.
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)

function Get-PrivateRuntimeProcesses {
    $result = @()
    # CIM exposes both x86 and x64 executable paths even from NSIS's x86
    # PowerShell. Get-Process.Path alone cannot reliably do this cross-bitness.
    $processes = @(Get-CimInstance -ClassName Win32_Process -Filter "Name='MathType.exe' OR Name='MathTypeLib.exe'" -OperationTimeoutSec 5 -ErrorAction Stop)
    foreach ($process in $processes) {
        if ([string]::IsNullOrWhiteSpace([string]$process.ExecutablePath)) {
            throw "Cannot inspect executable path for $($process.Name), PID $($process.ProcessId); no process was selected by name alone."
        }
        $path = [IO.Path]::GetFullPath([string]$process.ExecutablePath)
        if ($path.StartsWith($script:runtimePrefix, [StringComparison]::OrdinalIgnoreCase)) {
            $result += $process
        }
    }
    return $result
}

function Get-NormalizedLocalServerPath([object]$value) {
    if ($null -eq $value -or [string]::IsNullOrWhiteSpace([string]$value)) { return $null }
    $source = [Environment]::ExpandEnvironmentVariables(([string]$value).Trim())
    if ($source.StartsWith('"')) {
        $closingQuote = $source.IndexOf('"', 1)
        if ($closingQuote -gt 1) { return $source.Substring(1, $closingQuote - 1) }
        return $source.Trim('"')
    }
    $exeEnd = $source.IndexOf('.exe', [StringComparison]::OrdinalIgnoreCase)
    if ($exeEnd -ge 0) { return $source.Substring(0, $exeEnd + 4).Trim('"') }
    return ($source -split '\s+', 2)[0].Trim('"')
}

function Repair-LegacyPrivateMathTypeRegistration {
    $ownedPrefixes = @($script:runtimePrefix)
    foreach ($legacyRoot in @(
        (Join-Path $env:LOCALAPPDATA 'VisualTeX\mathtype-runtime'),
        (Join-Path $env:APPDATA 'VisualTeX\mathtype-runtime'),
        (Join-Path $env:USERPROFILE 'AppData\VisualTeX\mathtype-runtime')
    )) {
        if ([string]::IsNullOrWhiteSpace($legacyRoot)) { continue }
        $ownedPrefixes += [IO.Path]::GetFullPath($legacyRoot).TrimEnd('\', '/') + '\'
    }
    $ownedPrefixes = @($ownedPrefixes | Select-Object -Unique)
    $keyPath = 'Software\Classes\CLSID\{0002CE03-0000-0000-C000-000000000046}\LocalServer32'
    $parentPath = 'Software\Classes\CLSID\{0002CE03-0000-0000-C000-000000000046}'
    $repaired = 0
    $views = if ([Environment]::Is64BitOperatingSystem) {
        @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)
    } else {
        @([Microsoft.Win32.RegistryView]::Registry32)
    }
    foreach ($view in $views) {
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
            [Microsoft.Win32.RegistryHive]::CurrentUser,
            $view)
        try {
            $key = $baseKey.OpenSubKey($keyPath, $true)
            if ($null -eq $key) { continue }
            try {
                $registeredValue = $key.GetValue(
                    '',
                    $null,
                    [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                $serverPath = Get-NormalizedLocalServerPath $registeredValue
                if ([string]::IsNullOrWhiteSpace($serverPath)) { continue }
                try { $fullServerPath = [IO.Path]::GetFullPath($serverPath) }
                catch { continue }
                $isOwned = $false
                foreach ($prefix in $ownedPrefixes) {
                    if ($fullServerPath.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
                        $isOwned = $true
                        break
                    }
                }
                if (-not $isOwned) { continue }
                $key.DeleteValue('', $false)
                $repaired++
            } finally {
                $key.Dispose()
            }

            $emptyKey = $baseKey.OpenSubKey($keyPath, $false)
            if ($null -ne $emptyKey) {
                try {
                    if ($emptyKey.SubKeyCount -eq 0 -and $emptyKey.ValueCount -eq 0) {
                        $emptyKey.Dispose()
                        $emptyKey = $null
                        $baseKey.DeleteSubKey($keyPath, $false)
                    }
                } finally {
                    if ($null -ne $emptyKey) { $emptyKey.Dispose() }
                }
            }
            $emptyParent = $baseKey.OpenSubKey($parentPath, $false)
            if ($null -ne $emptyParent) {
                try {
                    if ($emptyParent.SubKeyCount -eq 0 -and $emptyParent.ValueCount -eq 0) {
                        $emptyParent.Dispose()
                        $emptyParent = $null
                        $baseKey.DeleteSubKey($parentPath, $false)
                    }
                } finally {
                    if ($null -ne $emptyParent) { $emptyParent.Dispose() }
                }
            }
        } finally {
            $baseKey.Dispose()
        }
    }
    Write-Output ('PRIVATE_RUNTIME_REGISTRATION_REPAIRED count=' + $repaired)
}

try {
    if ([string]::IsNullOrWhiteSpace($InstallRoot) -or -not [IO.Path]::IsPathRooted($InstallRoot)) {
        throw 'An absolute VisualTeX installation directory is required.'
    }
    $installation = [IO.Path]::GetFullPath($InstallRoot.TrimEnd(' ', '\', '/'))
    $script:runtimePrefix = [IO.Path]::GetFullPath((Join-Path $installation 'mathtype-runtime')).TrimEnd('\', '/') + '\'
    if ($Mode -eq 'RepairRegistration') {
        Repair-LegacyPrivateMathTypeRegistration
        exit 0
    }
    $targets = @(Get-PrivateRuntimeProcesses)
    if ($targets.Count -eq 0) {
        Write-Output 'PRIVATE_RUNTIME_CLEAR'
        exit 0
    }
    if ($Mode -eq 'Check') {
        Write-Output ('PRIVATE_RUNTIME_PRESENT PIDs=' + (($targets | ForEach-Object { $_.ProcessId }) -join ','))
        exit 10
    }

    # This branch is reached only after the installer's explicit consent prompt.
    # Never terminate a user's separately installed MathType. Re-resolve each PID
    # and path immediately before stopping it, and never use taskkill /IM /T.
    foreach ($target in $targets) {
        $current = Get-CimInstance -ClassName Win32_Process -Filter ('ProcessId=' + $target.ProcessId) -OperationTimeoutSec 5 -ErrorAction Stop
        if ($null -eq $current) { continue }
        if ($current.CreationDate -ne $target.CreationDate) { continue }
        if ([string]::IsNullOrWhiteSpace([string]$current.ExecutablePath)) {
            throw "Cannot revalidate runtime PID $($target.ProcessId)."
        }
        $path = [IO.Path]::GetFullPath([string]$current.ExecutablePath)
        if (-not $path.StartsWith($script:runtimePrefix, [StringComparison]::OrdinalIgnoreCase)) { continue }
        try {
            $owned = Get-Process -Id ([int]$target.ProcessId) -ErrorAction Stop
            try {
                Stop-Process -InputObject $owned -Force -ErrorAction Stop
                [void]$owned.WaitForExit(2000)
            } finally { $owned.Dispose() }
        } catch {
            if (Get-Process -Id ([int]$target.ProcessId) -ErrorAction SilentlyContinue) { throw }
        }
    }
    # A fresh enumeration also catches an immediately relaunched helper.
    Start-Sleep -Milliseconds 250
    $remaining = @(Get-PrivateRuntimeProcesses)
    if ($remaining.Count -ne 0) {
        Write-Output ('PRIVATE_RUNTIME_REMAINS PIDs=' + (($remaining | ForEach-Object { $_.ProcessId }) -join ','))
        exit 11
    }
    Write-Output ('PRIVATE_RUNTIME_STOPPED count=' + $targets.Count)
    exit 0
} catch {
    Write-Output ('PRIVATE_RUNTIME_CHECK_FAILED: ' + $_.Exception.Message)
    exit 20
}
