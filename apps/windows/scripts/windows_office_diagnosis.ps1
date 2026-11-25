# Why a VisualTeX Office installation step failed, in words a user can act on. The Office install and
# uninstall scripts dot-source this file; the NSIS installer also runs it on its own (-DiagnosisMode)
# before it copies files and when an earlier VisualTeX's uninstaller fails. ASCII only (Windows
# PowerShell 5.1 reads it as ANSI).
[CmdletBinding()]
param(
    [ValidateSet('', 'InstallFolder', 'PreviousUninstall', 'Wps')][string]$DiagnosisMode = '',
    [string]$DiagnosisPath = '',
    # The result, for the installer: written in the system code page, which NSIS reads.
    [string]$DiagnosisResultPath = '',
    # With -DiagnosisMode Wps (Settings > Office > Check WPS): also ask WPS itself and search its settings.
    [switch]$AskWps
)

# The processes that have one of the files open or loaded (Windows' Restart Manager): Id, Name (the
# executable) and Title (the name Windows shows for it).
function Get-ProcessesUsingFiles([string[]]$Paths) {
    $files = @($Paths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path -LiteralPath $_ -PathType Leaf) })
    if ($files.Count -eq 0) { return @() }
    try {
    if (-not ('VisualTeX.RestartManager' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
namespace VisualTeX {
public static class RestartManager {
    [StructLayout(LayoutKind.Sequential)]
    struct RM_UNIQUE_PROCESS { public int dwProcessId; public System.Runtime.InteropServices.ComTypes.FILETIME ProcessStartTime; }
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    struct RM_PROCESS_INFO {
        public RM_UNIQUE_PROCESS Process;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)] public string strAppName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)] public string strServiceShortName;
        public int ApplicationType;
        public uint AppStatus;
        public uint TSSessionId;
        [MarshalAs(UnmanagedType.Bool)] public bool bRestartable;
    }
    [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)] static extern int RmStartSession(out uint session, int flags, StringBuilder key);
    [DllImport("rstrtmgr.dll")] static extern int RmEndSession(uint session);
    [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)] static extern int RmRegisterResources(uint session, uint fileCount, string[] files, uint applicationCount, IntPtr applications, uint serviceCount, string[] services);
    [DllImport("rstrtmgr.dll")] static extern int RmGetList(uint session, out uint needed, ref uint count, [In, Out] RM_PROCESS_INFO[] processes, ref uint rebootReasons);

    // "id|title" for every process using one of the files; empty when Windows cannot tell.
    public static string[] ProcessesUsing(string[] files) {
        uint session;
        if (RmStartSession(out session, 0, new StringBuilder(64)) != 0) return new string[0];
        try {
            if (RmRegisterResources(session, (uint)files.Length, files, 0, IntPtr.Zero, 0, null) != 0) return new string[0];
            for (int attempt = 0; attempt < 3; attempt++) {
                uint needed, count = 0, reasons = 0;
                int result = RmGetList(session, out needed, ref count, null, ref reasons);
                if (result == 0) return new string[0];
                if (result != 234) return new string[0];
                RM_PROCESS_INFO[] processes = new RM_PROCESS_INFO[needed];
                count = needed;
                result = RmGetList(session, out needed, ref count, processes, ref reasons);
                if (result == 234) continue;
                if (result != 0) return new string[0];
                List<string> list = new List<string>();
                for (int i = 0; i < count; i++) list.Add(processes[i].Process.dwProcessId + "|" + processes[i].strAppName);
                return list.ToArray();
            }
            return new string[0];
        } finally { RmEndSession(session); }
    }
}
}
'@
    }
    $entries = @([VisualTeX.RestartManager]::ProcessesUsing([string[]]$files))
    } catch {
        # No compiler or no Restart Manager: the caller reports the file without the programs.
        return @()
    }
    $result = @()
    foreach ($entry in $entries) {
        $id, $title = ([string]$entry).Split('|', 2)
        $name = ''
        try { $name = (Get-Process -Id ([int]$id) -ErrorAction Stop).ProcessName + '.exe' } catch { }
        $result += [pscustomobject]@{ Id = [int]$id; Name = $name; Title = $title }
    }
    return @($result | Sort-Object Id -Unique)
}

# "WPS Office (wps.exe, PID 1234), WINWORD.exe (PID 5678)"
function Format-ProcessList($Processes) {
    return (@($Processes | ForEach-Object {
        $exe = if ($_.Name) { [string]$_.Name } else { 'process' }
        $title = [string]$_.Title
        if ($title -and $title -ine $exe -and $title -ine [IO.Path]::GetFileNameWithoutExtension($exe)) {
            "$title ($exe, PID $($_.Id))"
        } else {
            "$exe (PID $($_.Id))"
        }
    }) -join ', ')
}

function Get-PathOwner([string]$Path) {
    try { return [string](Get-Acl -LiteralPath $Path -ErrorAction Stop).Owner } catch { return '' }
}

function Get-BaseExceptionOf($ErrorRecordOrException) {
    $exception = if ($ErrorRecordOrException -is [Management.Automation.ErrorRecord]) { $ErrorRecordOrException.Exception } else { $ErrorRecordOrException }
    if ($null -eq $exception) { return $null }
    return $exception.GetBaseException()
}

# Why this account cannot replace the file, or $null when it can. A read-only file is made writable on
# the way (the NSIS installer cannot overwrite one). Processes in $IgnoreProcessIds (VisualTeX itself,
# which the installer closes next) do not count as users of the file.
function Get-FileWriteBlock([string]$Path, [int[]]$IgnoreProcessIds = @()) {
    $item = Get-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
    if ($null -eq $item -or $item.PSIsContainer) { return $null }
    if ($item.Attributes -band [IO.FileAttributes]::ReadOnly) {
        try { $item.Attributes = $item.Attributes -band (-bnot [IO.FileAttributes]::ReadOnly) } catch { }
    }
    try {
        $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Write, [IO.FileShare]::Read)
        $stream.Dispose()
        return $null
    } catch {
        $base = Get-BaseExceptionOf $_
        if ($base -is [UnauthorizedAccessException]) {
            $readOnly = $false
            try { $readOnly = [bool]((Get-Item -LiteralPath $Path -Force).Attributes -band [IO.FileAttributes]::ReadOnly) } catch { }
            return [pscustomobject]@{ Kind = $(if ($readOnly) { 'ReadOnly' } else { 'AccessDenied' }); Path = $Path; Detail = (Get-PathOwner $Path) }
        }
        $all = @(Get-ProcessesUsingFiles @($Path))
        $users = @($all | Where-Object { $IgnoreProcessIds -notcontains $_.Id })
        if ($all.Count -gt 0 -and $users.Count -eq 0) { return $null }
        $sharing = ($base -is [IO.IOException]) -and (($base.HResult -band 0xFFFF) -in @(32, 33))
        if ($sharing -or $users.Count -gt 0) {
            return [pscustomobject]@{ Kind = 'InUse'; Path = $Path; Detail = (Format-ProcessList $users) }
        }
        return [pscustomobject]@{ Kind = 'Other'; Path = $Path; Detail = $base.Message }
    }
}

# The first reason the installer could not write its files into the installation folder: the folder
# itself, then every file directly in it and in its first-level folders (the ones a VisualTeX update
# replaces), or $null. VisualTeX's own processes running from the folder are closed by the installer
# right after this check, so files only they use are not reported.
function Get-InstallFolderProblem([string]$Root) {
    $root = [IO.Path]::GetFullPath($Root.Trim().Trim('"')).TrimEnd('\')
    try { [void][IO.Directory]::CreateDirectory($root) } catch {
        $parent = Split-Path -Parent $root
        return [pscustomobject]@{ Kind = 'FolderDenied'; Path = $root; Detail = (Get-PathOwner $parent) }
    }
    $probe = Join-Path $root ('visualtex-write-test-' + [Guid]::NewGuid().ToString('N') + '.tmp')
    try {
        [IO.File]::WriteAllText($probe, '')
        Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
    } catch {
        return [pscustomobject]@{ Kind = 'FolderDenied'; Path = $root; Detail = (Get-PathOwner $root) }
    }
    $own = @()
    foreach ($name in @('visualtex', 'visualtex-mcp')) {
        foreach ($process in @(Get-Process -Name $name -ErrorAction SilentlyContinue)) {
            $path = ''
            try { $path = [string]$process.Path } catch { }
            if ($path -and $path.StartsWith($root + '\', [StringComparison]::OrdinalIgnoreCase)) { $own += [int]$process.Id }
        }
    }
    $files = @(Get-ChildItem -LiteralPath $root -File -Force -ErrorAction SilentlyContinue)
    foreach ($folder in @(Get-ChildItem -LiteralPath $root -Directory -Force -ErrorAction SilentlyContinue)) {
        # updates\ holds the in-app update's installer, which is the program running this check: nothing
        # there is replaced (it was taken for a file in use, and the update stopped at Retry/Cancel)
        if ($folder.Name -ieq 'updates') { continue }
        $files += @(Get-ChildItem -LiteralPath $folder.FullName -File -Force -ErrorAction SilentlyContinue)
    }
    foreach ($file in $files) {
        $block = Get-FileWriteBlock $file.FullName $own
        if ($null -ne $block) { return $block }
    }
    return $null
}

# Windows Installer's exit code and the error it would have shown, from its verbose log.
function Get-MsiFailureCause([int]$ExitCode, [string]$LogPath) {
    $meaning = @{
        1601 = 'the Windows Installer service could not be accessed'
        1602 = 'the installation was cancelled'
        1603 = 'a fatal error occurred during installation'
        1605 = 'the product is not installed'
        1612 = "the product's installation package is no longer on this computer"
        1618 = 'another installation is in progress; wait for it to finish'
        1619 = 'the installation package could not be opened'
        1620 = 'the installation package could not be opened'
        1622 = 'the installation log could not be written'
        1625 = 'a system policy prohibits this installation'
        1638 = 'another version of this product is already installed'
        1641 = 'Windows is restarting to complete the installation'
        3010 = 'a restart is required to complete the installation'
    }
    $text = "Windows Installer returned $ExitCode"
    if ($meaning.ContainsKey($ExitCode)) { $text += " ($($meaning[$ExitCode]))" }
    $detail = Get-MsiLogError $LogPath
    if ($detail) { $text += ": $detail" }
    return $text
}

# "Error 1920. Service ... failed to start ..." or the action that failed; empty when the log says neither.
function Get-MsiLogError([string]$LogPath) {
    if ([string]::IsNullOrWhiteSpace($LogPath)) { return '' }
    $path = $LogPath.Trim().Trim('"')
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return '' }
    $lines = @(Get-Content -LiteralPath $path -ErrorAction SilentlyContinue)
    foreach ($line in $lines) {
        if ($line -match '(?:^|-- )Error (\d{4})\.\s*(.+)$') { return ("Error {0}. {1}" -f $Matches[1], $Matches[2].Trim()) }
    }
    foreach ($line in $lines) {
        if ($line -match 'Action ended \d{1,2}:\d{2}:\d{2}: (\S+?)\. Return value 3\.') { return "the step $($Matches[1]) failed" }
    }
    return ''
}

# Windows Security's record of removing the file; empty when it has none (or cannot be asked).
function Get-QuarantineCause([string]$Path) {
    $name = [IO.Path]::GetFileName($Path)
    if ([string]::IsNullOrWhiteSpace($name)) { return '' }
    try {
        foreach ($detection in @(Get-MpThreatDetection -ErrorAction Stop)) {
            foreach ($resource in @($detection.Resources)) {
                if ([string]$resource -notlike "*\$name") { continue }
                $threat = ''
                try { $threat = [string](Get-MpThreat -ThreatID $detection.ThreatID -ErrorAction Stop).ThreatName } catch { }
                return "Windows Security removed $name on $($detection.InitialDetectionTime)" + $(if ($threat) { " as $threat" } else { '' })
            }
        }
    } catch { }
    return ''
}

# The VisualTeX.NativeRender service's state when it gets in the way, in words; empty when it does not.
function Get-NativeRenderServiceTrouble {
    $previousPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { $query = @(& sc.exe qc VisualTeX.NativeRender 2>$null) -join ' ' } finally { $ErrorActionPreference = $previousPreference }
    if ($query -match '\b1072\b') {
        return 'Windows has marked the VisualTeX.NativeRender service for deletion, but a program still holds it open (often the Services window or the Services tab of Task Manager). Close them or restart Windows, then try again.'
    }
    $service = Get-CimInstance Win32_Service -Filter "Name='VisualTeX.NativeRender'" -ErrorAction SilentlyContinue
    if ($null -eq $service) { return '' }
    if ($service.State -in @('Stop Pending', 'Start Pending')) {
        return "The VisualTeX.NativeRender service is stuck in '$($service.State)' (process $($service.ProcessId)). Restart Windows, then try again."
    }
    $exe = ([string]$service.PathName).Trim().Trim('"')
    if ($exe -and -not (Test-Path -LiteralPath $exe -PathType Leaf)) {
        $removed = Get-QuarantineCause $exe
        return "The VisualTeX.NativeRender service points to $exe, which no longer exists" + $(if ($removed) { " ($removed)" } else { '' }) + '.'
    }
    return ''
}

# Windows Installer still lists the renderer's package (VisualTeX Native Rendering).
function Test-NativeRenderPackageRegistered {
    $installer = New-Object -ComObject WindowsInstaller.Installer
    try {
        $related = @($installer.GetType().InvokeMember('RelatedProducts', [Reflection.BindingFlags]::GetProperty, $null, $installer, @('{D51CF2F8-0CF2-4A79-BA25-80E7A77FD548}')))
        foreach ($code in $related) {
            try { if ([string]$installer.ProductInfo([string]$code, 'ProductName') -ceq 'VisualTeX Native Rendering') { return $true } } catch { }
        }
        return $false
    } catch {
        return $false
    } finally {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($installer)
    }
}

# The files of the installed Office add-ins (for Microsoft Office and for a WPS Office of either
# bitness), which Word, PowerPoint and WPS load.
function Get-OfficeAddinFiles {
    $roots = @()
    foreach ($programFiles in @($env:ProgramW6432, ${env:ProgramFiles(x86)}, $env:ProgramFiles)) {
        if ([string]::IsNullOrWhiteSpace($programFiles)) { continue }
        foreach ($leaf in @('VSTO', 'VSTO-WPS')) { $roots += (Join-Path $programFiles "VisualTeX\WindowsOffice\$leaf") }
    }
    $files = @()
    foreach ($root in @($roots | Select-Object -Unique)) {
        if (Test-Path -LiteralPath $root -PathType Container) {
            $files += @(Get-ChildItem -LiteralPath $root -File -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName })
        }
    }
    return @($files)
}

# The programs that have VisualTeX's Office add-ins loaded right now: an installation cannot replace the
# add-ins under them (Windows would swap the files only at the next restart).
function Get-OfficeAddinUsers {
    return @(Get-ProcessesUsingFiles (Get-OfficeAddinFiles))
}

# Why the newest Office uninstall of a VisualTeX installation failed, from its own log (also the logs
# of releases before 1.3.1, which said only "uninstall failed"); empty when the newest one did not fail
# or is more than an hour old.
function Get-PreviousUninstallFailure {
    $root = Join-Path $env:LOCALAPPDATA 'VisualTeX\office\install-logs'
    $log = Get-ChildItem -LiteralPath $root -Filter 'vsto-uninstall-bootstrap-*.log' -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike '*.transcript.txt' } |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($null -eq $log -or $log.LastWriteTime -lt (Get-Date).AddHours(-1)) { return '' }
    $text = [IO.File]::ReadAllText($log.FullName)
    $failure = [regex]::Match($text, 'FAILED: (?:[\w.]+Exception: )?([^\r\n]+)')
    if (-not $failure.Success) { return '' }
    $reason = $failure.Groups[1].Value.Trim()
    if ($reason -match 'Unmanaged service found|refusing to adopt') {
        if (Test-NativeRenderPackageRegistered) {
            return "The VisualTeX.NativeRender service is installed, but VisualTeX's record of it (HKLM\SOFTWARE\VisualTeX\NativeRender) was deleted, for example by a forced uninstall, and that version refuses to touch the service. This version installs it again."
        }
        return 'The VisualTeX.NativeRender service of an earlier installation was left without its package, for example by a forced uninstall, and that version refuses to remove it. This version removes it.'
    }
    $msi = [regex]::Match($reason, 'MSI uninstall failed for (\{[0-9A-Fa-f-]{36}\}) with exit code (-?\d+)')
    if ($msi.Success) {
        $msiLog = [regex]::Match($text, 'Uninstalling Office MSI ' + [regex]::Escape($msi.Groups[1].Value) + '\. MSI log: ([^\r\n]+)').Groups[1].Value
        return 'Removing the Office add-in package failed: ' + (Get-MsiFailureCause ([int]$msi.Groups[2].Value) $msiLog)
    }
    if ($reason -match 'certificate cleanup failed') {
        $certificateLog = [regex]::Match($text, 'Certificate log: ([^\r\n]+)').Groups[1].Value.Trim()
        if ($certificateLog -and (Test-Path -LiteralPath $certificateLog -PathType Leaf)) {
            $last = @(Get-Content -LiteralPath $certificateLog -ErrorAction SilentlyContinue | Where-Object { $_ -match 'FAILED|error|denied|timed out' } | Select-Object -Last 1)
            if ($last.Count -gt 0) { return "Removing the VisualTeX HTTPS certificate failed: $($last[0].Trim())" }
        }
    }
    return $reason
}

# When the Windows registry last changed the key, or $null.
function Get-RegistryKeyWriteTime([Microsoft.Win32.RegistryHive]$Hive, [Microsoft.Win32.RegistryView]$View, [string]$SubKey) {
    try {
        if (-not ('VisualTeX.RegistryTime' -as [type])) {
            Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
namespace VisualTeX {
public static class RegistryTime {
    [DllImport("advapi32.dll", CharSet = CharSet.Unicode)]
    static extern int RegQueryInfoKeyW(SafeRegistryHandle key, IntPtr className, IntPtr classLength, IntPtr reserved, IntPtr subKeys, IntPtr maxSubKey, IntPtr maxClass, IntPtr values, IntPtr maxValueName, IntPtr maxValue, IntPtr securityDescriptor, out long lastWriteTime);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    static extern uint GetLongPathNameW(string shortPath, System.Text.StringBuilder longPath, uint length);
    public static string LongPath(string path) {
        System.Text.StringBuilder buffer = new System.Text.StringBuilder(1024);
        uint length = GetLongPathNameW(path, buffer, (uint)buffer.Capacity);
        return length > 0 && length < buffer.Capacity ? buffer.ToString() : path;
    }
    public static long LastWrite(SafeRegistryHandle key) {
        long time;
        return RegQueryInfoKeyW(key, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, out time) == 0 ? time : 0;
    }
}
}
'@
        }
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($Hive, $View)
        try {
            $key = $base.OpenSubKey($SubKey, $false)
            if ($null -eq $key) { return $null }
            try {
                $time = [VisualTeX.RegistryTime]::LastWrite($key.Handle)
                if ($time -le 0) { return $null }
                return [DateTime]::FromFileTime($time)
            } finally { $key.Dispose() }
        } finally { $base.Dispose() }
    } catch { return $null }
}

function Get-RegistryValueIn([Microsoft.Win32.RegistryHive]$Hive, [Microsoft.Win32.RegistryView]$View, [string]$SubKey, [string]$Name) {
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($Hive, $View)
    try {
        $key = $base.OpenSubKey($SubKey, $false)
        if ($null -eq $key) { return $null }
        try { return $key.GetValue($Name, $null) } finally { $key.Dispose() }
    } finally { $base.Dispose() }
}

# PowerShell of the given bitness, as WPS of that bitness loads the add-ins.
function Get-PowerShellOfBitness([string]$Architecture) {
    $windows = if ($env:WINDIR) { $env:WINDIR } else { 'C:\Windows' }
    if (-not [Environment]::Is64BitOperatingSystem) { return Join-Path $windows 'System32\WindowsPowerShell\v1.0\powershell.exe' }
    if ($Architecture -eq 'x86') { return Join-Path $windows 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe' }
    if (-not [Environment]::Is64BitProcess) { return Join-Path $windows 'Sysnative\WindowsPowerShell\v1.0\powershell.exe' }
    return Join-Path $windows 'System32\WindowsPowerShell\v1.0\powershell.exe'
}

# Creates the add-in the way WPS does (a COM object of the add-in's ProgID, in a process of WPS's bitness):
# "ok", or the error WPS would meet.
function Test-AddinActivation([string]$Architecture, [string]$ProgId) {
    # The add-ins are .NET classes: in PowerShell the object can be the managed one, not a COM wrapper.
    $command = "try { `$o = New-Object -ComObject $ProgId } catch { `$e = `$_.Exception.GetBaseException(); ('0x{0:X8} {1}' -f `$e.HResult, `$e.Message); exit 0 }; if ([Runtime.InteropServices.Marshal]::IsComObject(`$o)) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject(`$o) }; 'ok'"
    $output = Join-Path ([IO.Path]::GetTempPath()) ('visualtex-activation-' + [Guid]::NewGuid().ToString('N') + '.txt')
    try {
        $process = Start-Process -FilePath (Get-PowerShellOfBitness $Architecture) -ArgumentList @('-NoProfile', '-NonInteractive', '-Command', $command) -RedirectStandardOutput $output -WindowStyle Hidden -PassThru
        if (-not $process.WaitForExit(20000)) {
            try { $process.Kill() } catch { }
            return 'no answer within 20 seconds'
        }
        $text = (Get-Content -LiteralPath $output -Raw -ErrorAction SilentlyContinue)
        if ([string]::IsNullOrWhiteSpace($text)) { return "PowerShell exited with $($process.ExitCode)" }
        return ($text -replace '\s+', ' ').Trim()
    } catch {
        return $_.Exception.Message
    } finally {
        Remove-Item -LiteralPath $output -Force -ErrorAction SilentlyContinue
    }
}

# Why each WPS Office does (not) show VisualTeX, per installation and host, for Settings > Office:
# the registration in WPS's own registry view (missing, turned off, an incomplete COM class, a missing
# file, a DLL of the other bitness), the add-in whitelist, WPS processes started before the add-ins were
# registered (they read the list only at start), and whether the add-in can be created in a process of
# WPS's bitness.
function Get-WpsDiagnosis {
    if (-not (Get-Command Get-WpsInstallations -ErrorAction SilentlyContinue)) {
        $companion = Join-Path $PSScriptRoot 'windows_wps_companion.ps1'
        if (Test-Path -LiteralPath $companion -PathType Leaf) { . $companion }
    }
    $installations = @()
    foreach ($wps in @(Get-WpsInstallations)) {
        $view = if ($wps.Architecture -eq 'x86') { [Microsoft.Win32.RegistryView]::Registry32 } else { [Microsoft.Win32.RegistryView]::Registry64 }
        $hosts = @()
        $registeredAt = $null
        foreach ($entry in @(
            @{ Host = 'writer'; Office = 'Word'; ProgId = 'VisualTeX.WordVsto'; Clsid = '{F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1}'; List = 'WPS' },
            @{ Host = 'presentation'; Office = 'PowerPoint'; ProgId = 'VisualTeX.PowerPointVsto'; Clsid = '{7E586D2D-57B0-4D14-AB24-EBA9021A5E6D}'; List = 'WPP' }
        )) {
            $addinKey = "Software\Microsoft\Office\$($entry.Office)\Addins\$($entry.ProgId)"
            $serverKey = "Software\Classes\CLSID\$($entry.Clsid)\InprocServer32"
            $machineLoad = Get-RegistryValueIn LocalMachine $view $addinKey 'LoadBehavior'
            $userLoad = Get-RegistryValueIn CurrentUser $view $addinKey 'LoadBehavior'
            $server = [string](Get-RegistryValueIn LocalMachine $view $serverKey '')
            $codeBase = [string](Get-RegistryValueIn LocalMachine $view $serverKey 'CodeBase')
            $file = ''
            if ($codeBase) {
                try { $file = ([Uri]$codeBase).LocalPath } catch { $file = $codeBase -replace '^file:///', '' -replace '/', '\' }
            }
            $state = 'ok'
            $detail = [ordered]@{}
            if ($null -eq $machineLoad -and $null -eq $userLoad) {
                $state = 'missing'
            } elseif ($null -ne $userLoad -and [int]$userLoad -ne 3) {
                $state = 'disabled'; $detail.loadBehavior = [int]$userLoad; $detail.scope = 'user'
            } elseif ($null -eq $userLoad -and [int]$machineLoad -ne 3) {
                $state = 'disabled'; $detail.loadBehavior = [int]$machineLoad; $detail.scope = 'machine'
            } elseif ($server.Trim() -ine 'mscoree.dll' -or -not $file) {
                $state = 'server'
            } elseif (-not (Test-Path -LiteralPath $file -PathType Leaf)) {
                $state = 'fileMissing'; $detail.file = $file; $detail.quarantine = (Get-QuarantineCause $file)
            } else {
                $fileArchitecture = if (Get-Command Get-PeArchitecture -ErrorAction SilentlyContinue) { Get-PeArchitecture $file } else { $null }
                if ($fileArchitecture -and $fileArchitecture -ne $wps.Architecture) { $state = 'wrongBitness'; $detail.file = $file; $detail.fileArchitecture = $fileArchitecture }
            }
            $listKey = "Software\Kingsoft\Office\$($entry.List)\AddinsWL"
            $whitelisted = ($null -ne (Get-RegistryValueIn CurrentUser $view $listKey $entry.ProgId)) -or ($null -ne (Get-RegistryValueIn LocalMachine $view $listKey $entry.ProgId))
            $time = Get-RegistryKeyWriteTime LocalMachine $view $serverKey
            if ($null -ne $time -and ($null -eq $registeredAt -or $time -gt $registeredAt)) { $registeredAt = $time }
            $activation = if ($state -eq 'ok') { Test-AddinActivation $wps.Architecture $entry.ProgId } else { '' }
            $hosts += [ordered]@{ host = $entry.Host; state = $state; detail = $detail; whitelisted = $whitelisted; activation = $activation }
        }
        # WPS reads its add-in list when it starts; one started before the registration has not seen it.
        $folder = Split-Path -Parent $wps.Path
        if ('VisualTeX.RegistryTime' -as [type]) { $folder = [VisualTeX.RegistryTime]::LongPath($folder) }
        $stale = @()
        if ($null -ne $registeredAt) {
            foreach ($process in @(Get-Process -Name wps, wpp, et -ErrorAction SilentlyContinue)) {
                $path = ''
                $started = $null
                try { $path = [string]$process.Path; $started = $process.StartTime } catch { }
                if (-not $path -or $null -eq $started) { continue }
                if ('VisualTeX.RegistryTime' -as [type]) { $path = [VisualTeX.RegistryTime]::LongPath($path) }
                if (-not ([IO.Path]::GetDirectoryName($path)).Equals($folder, [StringComparison]::OrdinalIgnoreCase)) { continue }
                if ($started -lt $registeredAt) { $stale += [ordered]@{ id = $process.Id; name = $process.ProcessName + '.exe'; started = $started.ToString('yyyy-MM-dd HH:mm') } }
            }
        }
        $installations += [ordered]@{
            path = $wps.Path
            architecture = $wps.Architecture
            version = $wps.Version
            registeredAt = $(if ($null -ne $registeredAt) { $registeredAt.ToString('yyyy-MM-dd HH:mm') } else { '' })
            hosts = $hosts
            staleProcesses = $stale
        }
    }
    # WPS's own lists of add-ins it stopped loading (turned off for slowing its start, failed to load, ...),
    # by ProgID or CLSID: a reinstall does not clear them (Repair does).
    $records = @()
    $pattern = 'VisualTeX|F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1|7E586D2D-57B0-4D14-AB24-EBA9021A5E6D'
    $lists = @()
    foreach ($app in @("WPS", "WPP")) { foreach ($list in @("AddinsBL", "AddinsCL", "AddinsLF")) { $lists += "Software\Kingsoft\Office\$app\$list" } }
    foreach ($app in @("wps", "wpp")) { $lists += "Software\Kingsoft\Office\6.0\$app\DiagDisabledAddins" }
    foreach ($list in $lists) {
        $key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey($list)
        if ($null -eq $key) { continue }
        try {
            foreach ($name in $key.GetValueNames()) {
                $value = $key.GetValue($name)
                $text = if ($value -is [byte[]]) { [Text.Encoding]::Unicode.GetString($value) } else { [string]$value }
                if ($name -match $pattern -or $text -match $pattern) { $records += ($list -replace '^Software\\Kingsoft\\Office\\', '') + ": " + $name }
            }
        } finally { $key.Dispose() }
    }
    return [ordered]@{ installations = $installations; wpsRecords = $records }
}

# The user's own on/off setting of an add-in (HKCU LoadBehavior, which WPS's COM add-ins dialog writes), to put
# back after asking WPS to connect it: the check must not switch an add-in on that the user switched off.
function Get-UserAddinSetting([string]$Office, [string]$ProgId) {
    $path = "Software\Microsoft\Office\$Office\Addins\$ProgId"
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryView]::Registry64)
    try {
        $key = $base.OpenSubKey($path, $false)
        if ($null -eq $key) { return @{ Path = $path; Exists = $false; LoadBehavior = $null } }
        try { return @{ Path = $path; Exists = $true; LoadBehavior = $key.GetValue('LoadBehavior', $null) } } finally { $key.Dispose() }
    } finally { $base.Dispose() }
}

function Restore-UserAddinSetting($Setting) {
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryView]::Registry64)
    try {
        if (-not $Setting.Exists) { $base.DeleteSubKeyTree($Setting.Path, $false); return }
        $key = $base.CreateSubKey($Setting.Path)
        try {
            if ($null -eq $Setting.LoadBehavior) { $key.DeleteValue('LoadBehavior', $false) }
            else { $key.SetValue('LoadBehavior', [int]$Setting.LoadBehavior, [Microsoft.Win32.RegistryValueKind]::DWord) }
        } finally { $key.Dispose() }
    } finally { $base.Dispose() }
}

# What WPS itself says about VisualTeX's add-ins: a WPS started for the question (hidden; WPS windows the user
# has open are separate instances and stay untouched) lists the add-in or not, has connected it or not, and when
# it has not, what it answers when asked to connect it. Asking changes nothing: the user's on/off setting is put
# back afterwards. Each question runs in its own PowerShell with a time limit, so a WPS dialog cannot hang the
# diagnosis; hidden WPS processes the question leaves behind are closed.
function Get-WpsAnswers($Installations) {
    $installations = @($Installations)
    if ($installations.Count -eq 0) { return @() }
    $architecture = if (@($installations | Where-Object { $_.architecture -ne 'x86' }).Count -gt 0) { 'x64' } else { 'x86' }
    $query = @'
param([string]$ApplicationId, [string]$ProgId, [string]$ResultPath)
$result = [ordered]@{ listed = $false; connected = $false; askedConnected = $null; answer = ''; error = '' }
$app = $null
try {
    $app = New-Object -ComObject $ApplicationId
    $item = $null
    try { $item = $app.COMAddIns.Item($ProgId) } catch { }
    if ($null -ne $item) {
        $result.listed = $true
        $result.connected = [bool]$item.Connect
        if (-not $result.connected) {
            try { $item.Connect = $true; $result.askedConnected = [bool]$item.Connect }
            catch { $result.askedConnected = $false; $result.answer = $_.Exception.Message }
        }
    }
} catch {
    $result.error = $_.Exception.Message
} finally {
    if ($null -ne $app) { try { $app.Quit() } catch { } }
}
[IO.File]::WriteAllText($ResultPath, (ConvertTo-Json $result -Compress), (New-Object Text.UTF8Encoding($false)))
'@
    $queryFile = Join-Path ([IO.Path]::GetTempPath()) ('visualtex-wps-question-' + [Guid]::NewGuid().ToString('N') + '.ps1')
    [IO.File]::WriteAllText($queryFile, $query, [Text.Encoding]::ASCII)
    $answers = @()
    try {
        foreach ($entry in @(
            @{ Host = 'writer'; App = 'KWPS.Application'; Office = 'Word'; ProgId = 'VisualTeX.WordVsto' },
            @{ Host = 'presentation'; App = 'KWpp.Application'; Office = 'PowerPoint'; ProgId = 'VisualTeX.PowerPointVsto' }
        )) {
            $answer = [ordered]@{ host = $entry.Host; listed = $false; connected = $false; askedConnected = $null; answer = ''; error = '' }
            $setting = Get-UserAddinSetting $entry.Office $entry.ProgId
            $existing = @(Get-Process -Name wps, wpp, et -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
            $resultPath = $queryFile + '.' + $entry.Host + '.json'
            try {
                $child = Start-Process -FilePath (Get-PowerShellOfBitness $architecture) -ArgumentList @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $queryFile + '"'), $entry.App, $entry.ProgId, ('"' + $resultPath + '"')) -WindowStyle Hidden -PassThru
                if (-not $child.WaitForExit(90000)) {
                    try { $child.Kill() } catch { }
                    $answer.error = 'WPS did not answer within 90 seconds'
                } elseif (Test-Path -LiteralPath $resultPath -PathType Leaf) {
                    $parsed = Get-Content -LiteralPath $resultPath -Raw -Encoding UTF8 | ConvertFrom-Json
                    foreach ($name in 'listed', 'connected', 'askedConnected', 'answer', 'error') { $answer[$name] = $parsed.$name }
                } else {
                    $answer.error = "PowerShell exited with $($child.ExitCode)"
                }
            } catch {
                $answer.error = $_.Exception.Message
            } finally {
                Remove-Item -LiteralPath $resultPath -Force -ErrorAction SilentlyContinue
                Start-Sleep -Milliseconds 800
                # only windowless WPS processes started during the question: never a WPS the user opened
                foreach ($left in @(Get-Process -Name wps, wpp, et -ErrorAction SilentlyContinue | Where-Object { $existing -notcontains $_.Id -and $_.MainWindowHandle -eq [IntPtr]::Zero })) {
                    try { Stop-Process -Id $left.Id -Force -ErrorAction Stop } catch { }
                }
                $now = Get-UserAddinSetting $entry.Office $entry.ProgId
                if ($now.Exists -ne $setting.Exists -or "$($now.LoadBehavior)" -ne "$($setting.LoadBehavior)") {
                    try { Restore-UserAddinSetting $setting } catch { $answer.error = 'WPS changed the add-in setting and it could not be put back: ' + $_.Exception.Message }
                }
            }
            $answers += $answer
        }
    } finally {
        Remove-Item -LiteralPath $queryFile -Force -ErrorAction SilentlyContinue
    }
    return $answers
}

# Places in WPS's own settings, other than its add-in whitelist and the stop-loading lists above, that name
# VisualTeX's add-ins by ProgID or CLSID: registry values under HKCU\Software\Kingsoft and WPS's settings files
# (WPS keeps its settings database open, so files are read shared). Caches, cloud files and fonts are skipped.
function Get-WpsSettingsMentions {
    $pattern = 'F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1|7E586D2D-57B0-4D14-AB24-EBA9021A5E6D|VisualTeX\.WordVsto|VisualTeX\.PowerPointVsto'
    $covered = '\\(AddinsWL|AddinsBL|AddinsCL|AddinsLF|DiagDisabledAddins)$'
    $mentions = @()
    $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryView]::Registry64)
    try {
        $stack = New-Object System.Collections.Stack
        $stack.Push('Software\Kingsoft')
        $visited = 0
        while ($stack.Count -gt 0 -and $visited -lt 60000) {
            $path = [string]$stack.Pop(); $visited++
            if ($path -match 'RecentFiles|openfilelist|hookdns|PathMemory') { continue }
            $key = $null
            try { $key = $base.OpenSubKey($path, $false) } catch { }
            if ($null -eq $key) { continue }
            try {
                if ($path -notmatch $covered) {
                    foreach ($name in $key.GetValueNames()) {
                        $value = $key.GetValue($name)
                        $text = if ($value -is [byte[]]) { [Text.Encoding]::Unicode.GetString($value) + ' ' + [Text.Encoding]::ASCII.GetString($value) } else { [string]$value }
                        if ($name -match $pattern -or $text -match $pattern) { $mentions += ('HKCU\' + $path + ': ' + $(if ($name) { $name } else { '(default)' })) }
                    }
                }
                foreach ($sub in $key.GetSubKeyNames()) { $stack.Push($path + '\' + $sub) }
            } finally { $key.Dispose() }
        }
    } finally { $base.Dispose() }
    $skip = '^(WPS Office\\[0-9.]+|office6\\cache|WPS Cloud Files|personalofflinecache|hiddenfonts|wpsphoto\+|PDF|log|.*\\addons\\pool|.*\\log)$'
    foreach ($root in @((Join-Path $env:APPDATA 'kingsoft'), (Join-Path $env:LOCALAPPDATA 'Kingsoft'))) {
        if (-not (Test-Path -LiteralPath $root -PathType Container)) { continue }
        $folders = New-Object System.Collections.Stack
        $folders.Push($root)
        while ($folders.Count -gt 0) {
            $folder = [string]$folders.Pop()
            try {
                foreach ($sub in [IO.Directory]::GetDirectories($folder)) {
                    $relative = $sub.Substring($root.Length).TrimStart('\')
                    if ($relative -notmatch $skip -and -not ([IO.File]::GetAttributes($sub) -band [IO.FileAttributes]::ReparsePoint)) { $folders.Push($sub) }
                }
                foreach ($path in [IO.Directory]::GetFiles($folder)) {
                    $file = New-Object IO.FileInfo($path)
                    $wanted = $file.Length -gt 0 -and (($file.Length -lt 32MB -and $file.Extension -match '^\.(dat|db|data|json|ini|xml|cfg|conf|kcfg|txt|mdb)$') -or $file.Length -lt 1MB)
                    if (-not $wanted) { continue }
                    $bytes = $null
                    try {
                        $stream = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
                        try { $bytes = New-Object byte[] $stream.Length; [void]$stream.Read($bytes, 0, $bytes.Length) } finally { $stream.Dispose() }
                    } catch { continue }
                    $text = [Text.Encoding]::ASCII.GetString($bytes) + ' ' + [Text.Encoding]::Unicode.GetString($bytes)
                    if ($text -match $pattern) {
                        $what = @([regex]::Matches($text, $pattern + '|user_remove|user_enable') | ForEach-Object { $_.Value } | Select-Object -Unique) -join ', '
                        $shown = if ($env:USERPROFILE) { $path.Replace($env:USERPROFILE, '%USERPROFILE%') } else { $path }
                        $mentions += ('{0} ({1:yyyy-MM-dd HH:mm}): {2}' -f $shown, $file.LastWriteTime, $what)
                    }
                }
            } catch { }
        }
    }
    return $mentions
}

function Write-DiagnosisResult([string[]]$Lines) {
    if ([string]::IsNullOrWhiteSpace($DiagnosisResultPath)) {
        foreach ($line in $Lines) { Write-Output $line }
        return
    }
    [IO.File]::WriteAllText($DiagnosisResultPath, (($Lines | ForEach-Object { ($_ -replace '[\r\n]+', ' ').Trim() }) -join "`r`n"), [Text.Encoding]::Default)
}

if ($DiagnosisMode -eq 'InstallFolder') {
    # Exit 10 with three lines (kind, path, owner or programs) when the folder cannot take the files.
    $ErrorActionPreference = 'Stop'
    try {
        $problem = Get-InstallFolderProblem $DiagnosisPath
    } catch {
        $problem = $null
    }
    if ($null -eq $problem) { Write-DiagnosisResult @(); exit 0 }
    Write-DiagnosisResult @($problem.Kind, $problem.Path, [string]$problem.Detail)
    exit 10
}
if ($DiagnosisMode -eq 'Wps') {
    # JSON on stdout for VisualTeX (Settings > Office, the WPS repair prompt). -AskWps (Settings > Office > Check
    # WPS) adds what WPS itself answers and the other places in WPS's settings that name the add-ins.
    $diagnosis = [ordered]@{ installations = @() }
    try { $diagnosis = Get-WpsDiagnosis } catch { $diagnosis.error = $_.Exception.Message }
    if ($AskWps) {
        try { $diagnosis.wpsAnswers = @(Get-WpsAnswers $diagnosis.installations) } catch { $diagnosis.wpsAnswersError = $_.Exception.Message }
        try { $diagnosis.wpsSettings = @(Get-WpsSettingsMentions) } catch { $diagnosis.wpsSettings = @() }
        $diagnosis.checkedAt = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    }
    [Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
    Write-Output (ConvertTo-Json $diagnosis -Depth 6 -Compress)
    exit 0
}
if ($DiagnosisMode -eq 'PreviousUninstall') {
    $reason = ''
    try { $reason = Get-PreviousUninstallFailure } catch { }
    Write-DiagnosisResult @($reason)
    exit 0
}
