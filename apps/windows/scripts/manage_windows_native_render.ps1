[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][ValidateSet('Install','Verify','Uninstall','EnsureFonts')][string]$Action,
    [string]$PackageDirectory='', [string]$LogPath=''
)
$ErrorActionPreference='Stop'
$serviceName='VisualTeX.NativeRender'
$expectedAccount='NT SERVICE\VisualTeX.NativeRender'
$expectedSid='S-1-5-80-3568014781-357338252-412558032-1798669891-843860098'
$programFiles64=[Environment]::GetEnvironmentVariable('ProgramW6432')
if (!$programFiles64) { $programFiles64=[Environment]::GetFolderPath('ProgramFiles') }
$installRoot=[IO.Path]::GetFullPath((Join-Path $programFiles64 'VisualTeX\NativeRender'))
$expectedExe=Join-Path $installRoot 'VisualTeX.NativeRenderService.exe'
if ($Action -ne 'Verify') {
    $identity=[Security.Principal.WindowsIdentity]::GetCurrent()
    $principal=New-Object Security.Principal.WindowsPrincipal($identity)
    if (!$principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw 'Native-render installation requires elevation from the Office installer.' }
}
function Get-OwnedService {
    $service=Get-CimInstance Win32_Service -Filter "Name='VisualTeX.NativeRender'"
    if ($service -and ($service.PathName.Trim('"') -ine $expectedExe -or $service.StartName -ine $expectedAccount)) {
        throw 'A different service occupies the native-render name; refusing to replace or stop it.'
    }
    return $service
}
function Get-OwnedProduct {
    $hive=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,[Microsoft.Win32.RegistryView]::Registry64)
    try {
        $key=$hive.OpenSubKey('Software\VisualTeX\NativeRender')
        if (!$key) { return $null }
        try {
            if ([IO.Path]::GetFullPath([string]$key.GetValue('InstallPath')).TrimEnd('\') -ine $installRoot) { throw 'Unexpected native-render installation root.' }
            $code=[string]$key.GetValue('ProductCode')
            if ($code -notmatch '^\{[0-9A-Fa-f-]{36}\}$') { throw 'Invalid native-render product identity.' }
            $installer=New-Object -ComObject WindowsInstaller.Installer
            try {
                if ($installer.ProductInfo($code,'ProductName') -cne 'VisualTeX Native Rendering' -or $installer.ProductInfo($code,'Publisher') -cne 'VisualTeX') { throw 'Native-render MSI ownership mismatch.' }
            } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($installer) }
            return $code
        } finally { $key.Dispose() }
    } finally { $hive.Dispose() }
}
function Invoke-OwnedMsi([string[]]$Arguments) {
    $process=Start-Process -FilePath "$env:windir\System32\msiexec.exe" -ArgumentList $Arguments -WindowStyle Hidden -PassThru
    try { $process.WaitForExit(); $code=$process.ExitCode } finally { $process.Dispose() }
    if ($code -notin @(0,3010)) { throw "Native-render MSI failed with exit code $code. Log: $LogPath" }
    if ($code -eq 3010) { Write-Warning 'Windows reports a restart is required; no restart was requested.' }
}
function Assert-Payload($Manifest) {
    $seen=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach($file in $Manifest.files) {
        $relative=[string]$file.path
        if (!$relative -or [IO.Path]::IsPathRooted($relative) -or $relative -match '(^|[\\/])\.\.([\\/]|$)' -or !$seen.Add($relative)) { throw 'Invalid native-render manifest path.' }
        $path=[IO.Path]::GetFullPath((Join-Path $installRoot $relative))
        if (!$path.StartsWith($installRoot+'\',[StringComparison]::OrdinalIgnoreCase) -or !(Test-Path -LiteralPath $path -PathType Leaf)) { throw "Native-render file missing: $relative" }
        if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -ine [string]$file.sha256) { throw "Native-render file mismatch: $relative" }
    }
    if (!$seen.Contains('VisualTeX.NativeRenderService.exe') -or !$seen.Contains('VisualTeX.NativeRenderWorker.exe') -or !$seen.Contains('runtime\MathType.exe')) { throw 'Incomplete native-render manifest.' }
}
# MathType draws many symbols from its own fonts (MT Extra, the Euclid fonts).
# Office installs a 1992 MT Extra (46 glyphs, version 3.1) under the same name
# "MT Extra (TrueType)", so an install-if-missing rule leaves Office's copy in
# charge and symbols such as a triangle, bowtie or parallel bars show as boxes.
# Make Windows resolve every MathType font to a copy that has all of MathType's
# glyphs. Files are only added, never overwritten or deleted, and nothing is
# removed on uninstall. A registration of an older copy (one missing glyphs that
# MathType's copy has) is pointed at MathType's copy instead; MathType's MT Extra
# keeps every code and width of Office's copy, so nothing that used it moves.
function Get-FontApi {
    if (-not ('VisualTeX.MathTypeFontApi' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
namespace VisualTeX {
public static class MathTypeFontApi {
    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] public static extern int AddFontResourceW(string file);
    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] public static extern bool RemoveFontResourceW(string file);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern IntPtr SendMessageTimeoutW(IntPtr window, uint message, IntPtr wParam, IntPtr lParam, uint flags, uint timeout, out IntPtr result);
    [DllImport("gdi32.dll")] static extern IntPtr CreateCompatibleDC(IntPtr dc);
    [DllImport("gdi32.dll")] static extern bool DeleteDC(IntPtr dc);
    [DllImport("gdi32.dll")] static extern IntPtr SelectObject(IntPtr dc, IntPtr value);
    [DllImport("gdi32.dll")] static extern bool DeleteObject(IntPtr value);
    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] static extern IntPtr CreateFontW(int height, int width, int escapement, int orientation, int weight, uint italic, uint underline, uint strikeOut, uint charSet, uint outPrecision, uint clipPrecision, uint quality, uint pitchAndFamily, string face);
    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] static extern int GetTextFaceW(IntPtr dc, int count, [Out] char[] face);
    [DllImport("gdi32.dll", CharSet = CharSet.Unicode)] static extern uint GetGlyphIndicesW(IntPtr dc, string text, int count, [Out] ushort[] indices, uint flags);
    public static void BroadcastFontChange() {
        IntPtr ignored;
        SendMessageTimeoutW(new IntPtr(0xFFFF), 0x001D, IntPtr.Zero, IntPtr.Zero, 2, 5000, out ignored);
    }
    // Codes this session cannot draw with the font Windows picks for the
    // family and style (as MathType's pictures request it); null when Windows
    // picks a different family altogether.
    public static int[] MissingCodes(string family, int weight, bool italic, int[] codes) {
        IntPtr dc = CreateCompatibleDC(IntPtr.Zero);
        IntPtr font = CreateFontW(-1000, 0, 0, 0, weight, italic ? 1u : 0u, 0, 0, 1, 7, 0, 0, 0, family);
        IntPtr old = SelectObject(dc, font);
        try {
            char[] face = new char[64];
            int length = GetTextFaceW(dc, face.Length, face);
            string actual = length > 0 ? new string(face, 0, length - 1).TrimEnd('\0') : "";
            if (!string.Equals(actual, family, StringComparison.OrdinalIgnoreCase)) return null;
            List<int> missing = new List<int>();
            ushort[] index = new ushort[1];
            foreach (int code in codes) {
                if (GetGlyphIndicesW(dc, ((char)code).ToString(), 1, index, 1) == 0xFFFFFFFF || index[0] == 0xFFFF) missing.Add(code);
            }
            return missing.ToArray();
        } finally { SelectObject(dc, old); DeleteObject(font); DeleteDC(dc); }
    }
}
}
'@
    }
}
function Read-FontInfo([string]$Path) {
    if (!(Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    try {
        $face=New-Object System.Windows.Media.GlyphTypeface (New-Object Uri ([IO.Path]::GetFullPath($Path)))
        $family=$face.Win32FamilyNames[[Globalization.CultureInfo]::GetCultureInfo('en-us')]
        if (!$family) { $family=@($face.Win32FamilyNames.Values)[0] }
        $style=$face.Win32FaceNames[[Globalization.CultureInfo]::GetCultureInfo('en-us')]
        if (!$style) { $style=@($face.Win32FaceNames.Values)[0] }
        $codes=New-Object 'Collections.Generic.HashSet[int]'
        foreach($code in $face.CharacterToGlyphMap.Keys) { [void]$codes.Add([int]$code) }
        return [pscustomobject]@{
            Path=[IO.Path]::GetFullPath($Path); Family=[string]$family; Style=[string]$style
            Weight=$face.Weight.ToOpenTypeWeight(); Italic=($face.Style -ne [System.Windows.FontStyles]::Normal)
            Codes=$codes; Hash=(Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
        }
    } catch { return $null }
}
function Get-StemKey([string]$Path) {
    $stem=[IO.Path]::GetFileNameWithoutExtension($Path).Replace(' ','').ToLowerInvariant()
    return ($stem -replace '_(\d+|visualtex(_\d+)?)$','')
}
function Ensure-MathTypeFonts {
    Add-Type -AssemblyName PresentationCore
    Get-FontApi
    $sourceRoot=Join-Path $installRoot 'runtime\Fonts\TrueType'
    if (!(Test-Path -LiteralPath $sourceRoot -PathType Container)) { throw 'The native renderer does not carry MathType fonts.' }
    $fontsRoot=Join-Path $env:windir 'Fonts'
    $keyPath='HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts'
    $bundled=@(Get-ChildItem -LiteralPath $sourceRoot -Filter *.ttf | ForEach-Object { Read-FontInfo $_.FullName })
    if ($bundled.Count -eq 0 -or @($bundled | Where-Object { $null -eq $_ }).Count) { throw 'A bundled MathType font cannot be read.' }
    $stems=@($bundled | ForEach-Object { Get-StemKey $_.Path })
    $families=@($bundled | ForEach-Object { $_.Family } | Sort-Object -Unique)
    # Registered entries that can be one of MathType's fonts (by name or file).
    $entries=@()
    $values=(Get-ItemProperty -LiteralPath $keyPath).PSObject.Properties | Where-Object { $_.Name -notlike 'PS*' }
    foreach($value in $values) {
        $data=[string]$value.Value
        if (!$data) { continue }
        $path=if ([IO.Path]::IsPathRooted($data)) { $data } else { Join-Path $fontsRoot $data }
        $related=$false
        foreach($family in $families) { if ($value.Name.IndexOf($family,[StringComparison]::OrdinalIgnoreCase) -ge 0) { $related=$true } }
        if (!$related -and ($stems -notcontains (Get-StemKey $path))) { continue }
        $entries+=[pscustomobject]@{ Name=$value.Name; Data=$data; Info=(Read-FontInfo $path); Path=$path }
    }
    $changed=$false
    $report=@()
    foreach($font in $bundled) {
        $same=@($entries | Where-Object { $_.Info -and $_.Info.Family -ieq $font.Family -and $_.Info.Style -ieq $font.Style })
        $older=@($same | Where-Object { -not $_.Info.Codes.IsSupersetOf($font.Codes) })
        $complete=@($same | Where-Object { $_.Info.Codes.IsSupersetOf($font.Codes) })
        if ($complete.Count -gt 0 -and $older.Count -eq 0) { continue }
        $target=$null
        if ($complete.Count -gt 0) {
            $target=$complete[0]
        } else {
            # Reuse an identical copy this script installed before; otherwise add
            # a new file under a name nothing else uses.
            $stem=[IO.Path]::GetFileNameWithoutExtension($font.Path)
            $fileName=$null
            for ($n=1; $n -lt 100 -and !$fileName; $n++) {
                $candidate=if ($n -eq 1) { "${stem}_visualtex.ttf" } else { "${stem}_visualtex_$n.ttf" }
                $candidatePath=Join-Path $fontsRoot $candidate
                if (!(Test-Path -LiteralPath $candidatePath)) { Copy-Item -LiteralPath $font.Path -Destination $candidatePath; $fileName=$candidate }
                elseif ((Get-FileHash -LiteralPath $candidatePath -Algorithm SHA256).Hash -ieq $font.Hash) { $fileName=$candidate }
            }
            if (!$fileName) { throw "No free file name for $($font.Family) $($font.Style)." }
            $valueName=if ($font.Style -ieq 'Regular') { "$($font.Family) (TrueType)" } else { "$($font.Family) $($font.Style) (TrueType)" }
            Set-ItemProperty -LiteralPath $keyPath -Name $valueName -Value $fileName
            $target=[pscustomobject]@{ Name=$valueName; Data=$fileName; Path=(Join-Path $fontsRoot $fileName) }
            [void][VisualTeX.MathTypeFontApi]::AddFontResourceW($target.Path)
            $report+="$($font.Family) $($font.Style): registered $fileName as '$valueName'"
        }
        foreach($entry in $older) {
            # The older file stays on disk; only its registration and this
            # session's copy go, so Windows stops picking it first.
            if ($entry.Name -ne $target.Name) { Remove-ItemProperty -LiteralPath $keyPath -Name $entry.Name -ErrorAction SilentlyContinue }
            for ($i=0; $i -lt 16 -and [VisualTeX.MathTypeFontApi]::RemoveFontResourceW($entry.Path); $i++) { }
            $report+="$($font.Family) $($font.Style): replaced older $($entry.Data) ($($entry.Info.Codes.Count) codes)"
        }
        $changed=$true
    }
    if ($changed) { [VisualTeX.MathTypeFontApi]::BroadcastFontChange() }
    # What Word in this session now gets for every MathType font. MathType's
    # pictures draw symbol fonts with 8-bit codes, which Windows maps to
    # U+F020-U+F0FF; other cmap entries of those fonts are never drawn.
    $problems=@()
    foreach($font in $bundled) {
        $codes=[int[]]@($font.Codes | Where-Object { $_ -ge 0xF020 -and $_ -le 0xF0FF })
        if ($codes.Count -eq 0) { $codes=[int[]]@($font.Codes | Where-Object { $_ -le 0xFFFF }) }
        $missing=[VisualTeX.MathTypeFontApi]::MissingCodes($font.Family,$font.Weight,[bool]$font.Italic,$codes)
        if ($null -eq $missing -or $missing.Count) {
            # Registered but not loaded in this session yet: load it and look again.
            $registered=@((Get-ItemProperty -LiteralPath $keyPath).PSObject.Properties | Where-Object {
                $_.Name -notlike 'PS*' -and [string]$_.Value -and (Get-StemKey ([string]$_.Value)) -eq (Get-StemKey $font.Path) })
            foreach($value in $registered) {
                $path=if ([IO.Path]::IsPathRooted([string]$value.Value)) { [string]$value.Value } else { Join-Path $fontsRoot ([string]$value.Value) }
                $info=Read-FontInfo $path
                if ($info -and $info.Family -ieq $font.Family -and $info.Style -ieq $font.Style -and $info.Codes.IsSupersetOf($font.Codes)) {
                    [void][VisualTeX.MathTypeFontApi]::AddFontResourceW($path)
                }
            }
            [VisualTeX.MathTypeFontApi]::BroadcastFontChange()
            $missing=[VisualTeX.MathTypeFontApi]::MissingCodes($font.Family,$font.Weight,[bool]$font.Italic,$codes)
        }
        # MathType's default styles (and VisualTeX's equations) draw from MT Extra
        # and the Euclid fonts; Fences and the Tiger fonts are optional.
        $required=($font.Family -ieq 'MT Extra' -or $font.Family -like 'Euclid*')
        $problem=if ($null -eq $missing) { "$($font.Family) $($font.Style) is not available" } elseif ($missing.Count) { "$($font.Family) $($font.Style) lacks $($missing.Count) glyphs" } else { $null }
        if ($problem -and $required) { $problems+=$problem } elseif ($problem) { Write-Warning "MathType fonts: optional $problem" }
    }
    foreach($line in $report) { Write-Output "MathType fonts: $line" }
    if ($problems.Count) { throw ('MathType fonts are still incomplete: '+($problems -join '; ')) }
    Write-Output "MathType fonts verified: $($bundled.Count) fonts."
}

$service=Get-OwnedService
$product=Get-OwnedProduct
$manifest=$null
if ($Action -eq 'Install' -or ($Action -eq 'Verify' -and $PackageDirectory)) {
    $msi=Join-Path $PackageDirectory 'VisualTeX-NativeRender-x64.msi'
    $manifestPath=Join-Path $PackageDirectory 'VisualTeX-NativeRender-x64.sha256.json'
    $manifest=Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($manifest.schema -ne 1 -or $manifest.service -cne $serviceName -or $manifest.package.name -cne 'VisualTeX-NativeRender-x64.msi') { throw 'Unexpected native-render package manifest.' }
    if ((Get-FileHash -LiteralPath $msi -Algorithm SHA256).Hash -ine $manifest.package.sha256) { throw 'Native-render MSI hash mismatch.' }
}
if ($Action -ne 'Verify' -and !$LogPath) {
    $logRoot=Join-Path $env:LOCALAPPDATA 'VisualTeX\office\install-logs'
    [void][IO.Directory]::CreateDirectory($logRoot)
    $LogPath=Join-Path $logRoot ('native-render-'+$Action.ToLowerInvariant()+'-'+(Get-Date -Format 'yyyyMMdd-HHmmss')+'.log')
}
if ($Action -eq 'Install') {
    if ($service -and !$product) { throw 'The existing service is not managed by this MSI; refusing to adopt it.' }
    Invoke-OwnedMsi @('/i',('"'+[IO.Path]::GetFullPath($msi)+'"'),'/qn','/norestart','REBOOT=ReallySuppress','/L*v',('"'+$LogPath+'"'))
    $product=Get-OwnedProduct
    $service=Get-OwnedService
    if (!$product -or !$service) {
        throw 'The native-render MSI completed without an owned service registration.'
    }
    if ($service.State -ne 'Running') {
        # A same-version MSI maintenance install does not necessarily rerun the
        # ServiceControl Start=install action after the installer deliberately
        # stopped the old service. Ownership has already been re-proved above,
        # so start exactly this service through SCM and wait for it to settle.
        Start-Service -Name $serviceName -ErrorAction Stop
        $controller=Get-Service -Name $serviceName -ErrorAction Stop
        try {
            $controller.WaitForStatus(
                [ServiceProcess.ServiceControllerStatus]::Running,
                [TimeSpan]::FromSeconds(15))
        } finally {
            $controller.Dispose()
        }
        $service=Get-OwnedService
    }
    # A font problem must not undo a working renderer install; the app reports
    # it and its Repair action runs this step again.
    try { Ensure-MathTypeFonts } catch { Write-Warning $_.Exception.Message }
}
if ($Action -eq 'EnsureFonts') {
    if (!$product) { throw 'The native renderer is not installed.' }
    Ensure-MathTypeFonts
    return
}
if ($Action -in @('Install','Verify')) {
    if (!$product -or !$service -or $service.State -ne 'Running') { throw 'The private native-render service is not installed and running.' }
    $sid=(New-Object Security.Principal.NTAccount($expectedAccount)).Translate([Security.Principal.SecurityIdentifier]).Value
    if ($sid -cne $expectedSid) { throw 'The service virtual-account identity does not match the protected data ACL.' }
    if ($manifest) { Assert-Payload $manifest }
    Write-Output "Native renderer verified: $serviceName; private account; $expectedExe"
    return
}
if ($Action -eq 'Uninstall') {
    if (!$product -and $service) { throw 'Unmanaged service found; refusing to remove it.' }
    if ($product) { Invoke-OwnedMsi @('/x',$product,'/qn','/norestart','REBOOT=ReallySuppress','/L*v',('"'+$LogPath+'"')) }
    if (Get-CimInstance Win32_Service -Filter "Name='VisualTeX.NativeRender'") { throw 'Native-render service remains after uninstall.' }
    # Windows owns profile/hive lifetime. Do not force-unload an active hive or
    # delete shared MathType fonts, public COM registrations, or user settings.
    $profile=Get-CimInstance Win32_UserProfile -Filter "SID='$expectedSid'"
    if ($profile) {
        $expectedProfile=Join-Path $env:windir 'ServiceProfiles\VisualTeX.NativeRender'
        if ($profile.Special -or [IO.Path]::GetFullPath($profile.LocalPath) -ine $expectedProfile) { throw 'Unexpected private-service profile target.' }
        if ($profile.Loaded) { Write-Warning "The private service profile is still loaded by Windows: $expectedProfile" }
        else { Remove-CimInstance -InputObject $profile }
    }
    Write-Output 'Native-render MSI removed; standalone MathType was not modified.'
}
