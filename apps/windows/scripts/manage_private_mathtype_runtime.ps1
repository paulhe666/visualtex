[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InstallRoot,
    [ValidateSet('Check', 'Stop')][string]$Mode = 'Check'
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

try {
    if ([string]::IsNullOrWhiteSpace($InstallRoot) -or -not [IO.Path]::IsPathRooted($InstallRoot)) {
        throw 'An absolute VisualTeX installation directory is required.'
    }
    $installation = [IO.Path]::GetFullPath($InstallRoot.TrimEnd(' ', '\', '/'))
    $script:runtimePrefix = [IO.Path]::GetFullPath((Join-Path $installation 'mathtype-runtime')).TrimEnd('\', '/') + '\'
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
