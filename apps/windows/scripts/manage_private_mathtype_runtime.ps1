[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InstallRoot,
    [ValidateSet('Check', 'Stop', 'RepairRegistration')][string]$Mode = 'Check',
    [switch]$Elevated
)

# Current VisualTeX native rendering is owned by one SCM service. Installation
# maintenance therefore manages that exact service identity instead of scanning
# MathType.exe/MathTypeLib.exe processes. Legacy registration repair remains
# path-scoped to known VisualTeX roots and never inspects or terminates a process.
#
# Exit codes:
#   0  = clear/stopped/repaired
#   10 = owned renderer service is running (Check only)
#   11 = owned renderer service did not stop
#   20 = ownership/inspection/elevation failed
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)

$serviceName = 'VisualTeX.NativeRender'
$expectedAccount = 'NT SERVICE\VisualTeX.NativeRender'
$programFiles64 = [Environment]::GetEnvironmentVariable('ProgramW6432')
if ([string]::IsNullOrWhiteSpace($programFiles64)) {
    $programFiles64 = [Environment]::GetFolderPath('ProgramFiles')
}
$expectedServicePath = [IO.Path]::GetFullPath(
    (Join-Path $programFiles64 'VisualTeX\NativeRender\VisualTeX.NativeRenderService.exe'))

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Quote-ProcessArgument([string]$Value) {
    return '"' + $Value.Replace('"', '\"') + '"'
}

function Invoke-ElevatedSelf {
    $arguments = New-Object System.Collections.Generic.List[string]
    foreach ($value in @(
        '-NoProfile',
        '-NonInteractive',
        '-WindowStyle',
        'Hidden',
        '-ExecutionPolicy',
        'Bypass',
        '-File',
        (Quote-ProcessArgument $PSCommandPath),
        '-InstallRoot',
        (Quote-ProcessArgument $InstallRoot),
        '-Mode',
        'Stop',
        '-Elevated'
    )) {
        [void]$arguments.Add($value)
    }

    $powershell = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $process = Start-Process -FilePath $powershell -ArgumentList ($arguments -join ' ') -Verb RunAs -PassThru -WindowStyle Hidden
    try {
        $process.WaitForExit()
        exit $process.ExitCode
    } finally {
        $process.Dispose()
    }
}

function Get-NormalizedExecutablePath([object]$Value) {
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    $source = [Environment]::ExpandEnvironmentVariables(([string]$Value).Trim())
    if ($source.StartsWith('"')) {
        $closingQuote = $source.IndexOf('"', 1)
        if ($closingQuote -le 1) { return $null }
        return [IO.Path]::GetFullPath($source.Substring(1, $closingQuote - 1))
    }
    $exeEnd = $source.IndexOf('.exe', [StringComparison]::OrdinalIgnoreCase)
    if ($exeEnd -lt 0) { return $null }
    return [IO.Path]::GetFullPath($source.Substring(0, $exeEnd + 4).Trim('"'))
}

function Get-OwnedNativeRenderService {
    $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
        [Microsoft.Win32.RegistryHive]::LocalMachine,
        [Microsoft.Win32.RegistryView]::Registry64)
    try {
        $key = $baseKey.OpenSubKey('SYSTEM\CurrentControlSet\Services\VisualTeX.NativeRender', $false)
        if ($null -eq $key) { return $null }
        try {
            $imagePath = Get-NormalizedExecutablePath $key.GetValue(
                'ImagePath',
                $null,
                [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
            $account = [string]$key.GetValue('ObjectName', '')
            $expectedPath = [string]::Equals(
                $imagePath,
                $expectedServicePath,
                [StringComparison]::OrdinalIgnoreCase)
            $expectedIdentity = [string]::Equals(
                $account,
                $expectedAccount,
                [StringComparison]::OrdinalIgnoreCase)
            if ([string]::IsNullOrWhiteSpace($imagePath) -or -not $expectedPath -or -not $expectedIdentity) {
                throw "A different service occupies '$serviceName'; refusing to control it."
            }
        } finally {
            $key.Dispose()
        }
    } finally {
        $baseKey.Dispose()
    }

    $service = Get-Service -Name $serviceName -ErrorAction Stop
    return $service
}

function Stop-OwnedNativeRenderService {
    $service = Get-OwnedNativeRenderService
    if ($null -eq $service) {
        Write-Output 'NATIVE_RENDER_SERVICE_CLEAR'
        return 0
    }
    try {
        if ($service.Status -eq [ServiceProcess.ServiceControllerStatus]::Stopped) {
            Write-Output 'NATIVE_RENDER_SERVICE_STOPPED already=true'
            return 0
        }

        if (-not (Test-IsAdministrator)) {
            if ($Elevated) {
                throw 'VisualTeX native-render service maintenance requires administrator privileges.'
            }
            Invoke-ElevatedSelf
        }

        Stop-Service -Name $serviceName -Force -ErrorAction Stop
        $service.Dispose()
        $service = Get-Service -Name $serviceName -ErrorAction Stop
        try {
            $service.WaitForStatus(
                [ServiceProcess.ServiceControllerStatus]::Stopped,
                [TimeSpan]::FromSeconds(15))
            $service.Refresh()
            if ($service.Status -ne [ServiceProcess.ServiceControllerStatus]::Stopped) {
                Write-Output 'NATIVE_RENDER_SERVICE_REMAINS_RUNNING'
                return 11
            }
        } finally {
            $service.Dispose()
        }
        Write-Output 'NATIVE_RENDER_SERVICE_STOPPED'
        return 0
    } finally {
        if ($null -ne $service) {
            try { $service.Dispose() } catch { }
        }
    }
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
    $script:runtimePrefix = [IO.Path]::GetFullPath(
        (Join-Path $installation 'mathtype-runtime')).TrimEnd('\', '/') + '\'

    if ($Mode -eq 'RepairRegistration') {
        Repair-LegacyPrivateMathTypeRegistration
        exit 0
    }

    if ($Mode -eq 'Stop') {
        exit (Stop-OwnedNativeRenderService)
    }

    $service = Get-OwnedNativeRenderService
    if ($null -eq $service) {
        Write-Output 'NATIVE_RENDER_SERVICE_CLEAR'
        exit 0
    }
    try {
        if ($service.Status -eq [ServiceProcess.ServiceControllerStatus]::Stopped) {
            Write-Output 'NATIVE_RENDER_SERVICE_CLEAR'
            exit 0
        }
        Write-Output ('NATIVE_RENDER_SERVICE_PRESENT status=' + $service.Status)
        exit 10
    } finally {
        $service.Dispose()
    }
} catch {
    Write-Output ('PRIVATE_RUNTIME_CHECK_FAILED: ' + $_.Exception.Message)
    exit 20
}
