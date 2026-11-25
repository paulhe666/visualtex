# What an uninstall leaves after its regular steps (the Office integration's own uninstall, the input
# method's, the program files): VisualTeX's caches, logs and certificate files, the WebView folder, startup
# entries, what WPS and Office remember about the add-ins, MathType's WPS placeholder (MathType's own template
# goes back), program-folder remains, and Windows Installer records of a package that could not be removed.
# Kept unless asked: settings and formula documents (%APPDATA%\com.visualtex.studio and %APPDATA%\VisualTeX;
# the Office preferences move there from the registry) and the OCR environment. Never touched: the documents
# in Documents\VisualTeX. Runs as the user; asks for administrator rights once, and only when something under
# Program Files, the Windows folder or HKLM is left. The NSIS uninstaller shows the output. ASCII only.
param(
    [string]$InstallDir = '',
    [switch]$RemoveSettings,
    [switch]$RemoveOcr,
    [switch]$Elevated,
    [string]$OutputPath = ''
)
$ErrorActionPreference = 'Continue'
# The NSIS uninstaller starts 32-bit PowerShell, which sees the 32-bit registry: run as the system's own
# 64-bit PowerShell instead.
if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
    $native = Join-Path $env:WINDIR 'Sysnative\WindowsPowerShell\v1.0\powershell.exe'
    if (Test-Path -LiteralPath $native) {
        $forward = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
        if ($InstallDir) { $forward += @('-InstallDir', $InstallDir) }
        if ($RemoveSettings) { $forward += '-RemoveSettings' }
        if ($RemoveOcr) { $forward += '-RemoveOcr' }
        & $native @forward
        exit $LASTEXITCODE
    }
}
if (-not $InstallDir) { $InstallDir = Join-Path $env:LOCALAPPDATA 'VisualTeX' }
$settingsRoot = Join-Path $env:APPDATA 'com.visualtex.studio'
$sharedRoot = Join-Path $env:APPDATA 'VisualTeX'
$progIds = @('VisualTeX.WordVsto', 'VisualTeX.PowerPointVsto', 'VisualTeX.Formula', 'VisualTeX.Formula.1')
$clsids = @('{F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1}', '{7E586D2D-57B0-4D14-AB24-EBA9021A5E6D}', '{8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B}')
$interfaces = @('{6C672AF0-7321-4D21-B325-868CB34592C2}', '{A59B7798-6F24-4CF0-B378-E951BFFAFB3A}')
$typeLibs = @('{DF66EC66-3B3A-4675-A7BE-30456A04EB96}')
$appIds = @('{3C72FF7F-B04A-4FD0-AA7D-61D110D8B3C1}', 'VisualTeX.FormulaOleServer.exe')
$upgradeCodes = @('{A81B4BF7-0E51-45CE-A5AA-5E28F6944F42}', '{D51CF2F8-0CF2-4A79-BA25-80E7A77FD548}')
$namePattern = 'VisualTeX\.(WordVsto|PowerPointVsto|Formula)|F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1|7E586D2D-57B0-4D14-AB24-EBA9021A5E6D|8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B'
# MathType's WPS startup template is replaced by this placeholder (office/wps_mathtype_template.rs)
$placeholderSha256 = '1105415A9B8EEB51E3DDFD0A0636D8464103EAFD58AA4E8FE1C0C7BDE6C0817D'
$views = @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)

$lines = New-Object System.Collections.Generic.List[string]
function Say([string]$Text) { [void]$lines.Add($Text); Write-Output $Text }
function Hive([string]$Name) { if ($Name -eq 'HKLM') { [Microsoft.Win32.RegistryHive]::LocalMachine } else { [Microsoft.Win32.RegistryHive]::CurrentUser } }
function Test-KeyIn([string]$HiveName, [string]$Path) {
    foreach ($view in $views) {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey((Hive $HiveName), $view)
        try { $key = $base.OpenSubKey($Path); if ($key) { $key.Dispose(); return $true } } finally { $base.Dispose() }
    }
    return $false
}
function Remove-KeyIn([string]$HiveName, [string]$Path) {
    foreach ($view in $views) {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey((Hive $HiveName), $view)
        try {
            $key = $base.OpenSubKey($Path)
            if ($null -eq $key) { continue }
            $key.Dispose()
            try { $base.DeleteSubKeyTree($Path, $false); Say ("removed {0}\{1} [{2}]" -f $HiveName, $Path, $view) }
            catch { Say ("could not remove {0}\{1}: {2}" -f $HiveName, $Path, $_.Exception.Message) }
        } finally { $base.Dispose() }
    }
}
function Remove-MatchingValues([string]$HiveName, [string]$Path, [string]$Pattern) {
    foreach ($view in $views) {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey((Hive $HiveName), $view)
        try {
            $key = $null
            try { $key = $base.OpenSubKey($Path, $true) } catch { }
            if ($null -eq $key) { continue }
            try {
                foreach ($name in $key.GetValueNames()) {
                    $value = $key.GetValue($name)
                    $text = if ($value -is [byte[]]) { [Text.Encoding]::Unicode.GetString($value) + ' ' + [Text.Encoding]::ASCII.GetString($value) } else { [string]$value }
                    if ($name -match $Pattern -or $text -match $Pattern) {
                        try { $key.DeleteValue($name, $false); Say ("removed {0}\{1} value {2}" -f $HiveName, $Path, $name) } catch { }
                    }
                }
            } finally { $key.Dispose() }
        } finally { $base.Dispose() }
    }
}
function Remove-PathNow([string]$Path) {
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return }
    Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object { try { $_.Attributes = 'Normal' } catch { } }
    Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue
    if ((Test-Path -LiteralPath $Path) -and $Elevated) {
        # files owned by the renderer's service account: an administrator takes them over first
        & takeown.exe /f $Path /r /d y 2>&1 | Out-Null
        & icacls.exe $Path /grant '*S-1-5-32-544:F' /t /c /q 2>&1 | Out-Null
        Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path -LiteralPath $Path) { Say ('could not remove (in use or denied): ' + $Path) } else { Say ('removed ' + $Path) }
}
# VisualTeX's own OCR environment: a folder with VisualTeX's marker file, never any other folder
function Test-OcrRoot([string]$Path) {
    $marker = Join-Path $Path '.visualtex-ocr-root.json'
    if (-not (Test-Path -LiteralPath $marker -PathType Leaf)) { return $false }
    try { $m = Get-Content -LiteralPath $marker -Raw | ConvertFrom-Json; return ($m.product -eq 'VisualTeX' -and $m.kind -eq 'ocr-runtime') } catch { return $false }
}
function Test-HasPython([string]$Path) {
    foreach ($p in 'python\python.exe', 'venv\Scripts\python.exe') { if (Test-Path -LiteralPath (Join-Path $Path $p) -PathType Leaf) { return $true } }
    return $false
}
# Pack-Guid: Windows Installer's registry form of a product code
function Pack-Guid([string]$Guid) {
    $g = $Guid.Trim('{}').Replace('-', '').ToUpperInvariant()
    $reverse = { param($s) -join ($s.ToCharArray()[($s.Length - 1)..0]) }
    $out = (& $reverse $g.Substring(0, 8)) + (& $reverse $g.Substring(8, 4)) + (& $reverse $g.Substring(12, 4))
    for ($i = 16; $i -lt 32; $i += 2) { $out += $g.Substring($i + 1, 1) + $g.Substring($i, 1) }
    return $out
}
function Get-OwnProducts {
    $found = @()
    try {
        $installer = New-Object -ComObject WindowsInstaller.Installer
        foreach ($code in $upgradeCodes) { foreach ($p in @($installer.GetType().InvokeMember('RelatedProducts', 'GetProperty', $null, $installer, @($code)))) { if ($p) { $found += [string]$p } } }
    } catch { }
    return @($found | ForEach-Object { $_.ToUpperInvariant() } | Select-Object -Unique)
}

if ($Elevated) {
    # ---------- the parts that need administrator rights ----------
    if (Get-CimInstance Win32_Service -Filter "Name='VisualTeX.NativeRender'" -ErrorAction SilentlyContinue) {
        & sc.exe stop VisualTeX.NativeRender | Out-Null
        Start-Sleep -Seconds 2
        & sc.exe delete VisualTeX.NativeRender | Out-Null
        Say 'removed the service VisualTeX.NativeRender'
    }
    foreach ($code in @(Get-OwnProducts)) {
        $p = Start-Process msiexec.exe -ArgumentList @('/x', $code, '/qn', '/norestart', 'REBOOT=ReallySuppress') -Wait -PassThru
        Say ("Windows Installer package {0}: msiexec /x exit code {1}" -f $code, $p.ExitCode)
        if (@(Get-OwnProducts) -contains $code) {
            # a package Windows Installer cannot remove (its cached copy is gone) loses its registration
            $packed = Pack-Guid $code
            foreach ($path in @("SOFTWARE\Classes\Installer\Products\$packed", "SOFTWARE\Classes\Installer\Features\$packed",
                                "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\$packed",
                                "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\$code")) { Remove-KeyIn 'HKLM' $path }
            foreach ($upgrade in $upgradeCodes) { Remove-MatchingValues 'HKLM' ('SOFTWARE\Classes\Installer\UpgradeCodes\' + (Pack-Guid $upgrade)) $packed }
        }
    }
    foreach ($id in $progIds) { Remove-KeyIn 'HKLM' ('Software\Classes\' + $id) }
    foreach ($id in $clsids) { Remove-KeyIn 'HKLM' ('Software\Classes\CLSID\' + $id) }
    foreach ($id in $interfaces) { Remove-KeyIn 'HKLM' ('Software\Classes\Interface\' + $id) }
    foreach ($id in $typeLibs) { Remove-KeyIn 'HKLM' ('Software\Classes\TypeLib\' + $id) }
    foreach ($id in $appIds) { Remove-KeyIn 'HKLM' ('Software\Classes\AppID\' + $id) }
    foreach ($office in 'Word', 'PowerPoint') { foreach ($id in 'VisualTeX.WordVsto', 'VisualTeX.PowerPointVsto') { Remove-KeyIn 'HKLM' ("Software\Microsoft\Office\$office\Addins\$id") } }
    foreach ($app in 'WPS', 'WPP', 'ET') { foreach ($list in 'AddinsWL', 'AddinsBL', 'AddinsCL', 'AddinsLF') { Remove-MatchingValues 'HKLM' ("Software\Kingsoft\Office\$app\$list") $namePattern } }
    Remove-KeyIn 'HKLM' 'SOFTWARE\VisualTeX'
    foreach ($root in @($env:ProgramW6432, ${env:ProgramFiles(x86)}) | Where-Object { $_ } | Select-Object -Unique) {
        Remove-PathNow (Join-Path $root 'VisualTeX')
        Remove-PathNow (Join-Path $root 'VisualTeX Input')
    }
    $serviceProfile = Join-Path $env:windir 'ServiceProfiles\VisualTeX.NativeRender'
    if (Test-Path -LiteralPath $serviceProfile) { Remove-PathNow $serviceProfile }
    # the MathType fonts the renderer added, if its own uninstall did not run
    $renderScript = Join-Path $PSScriptRoot 'manage_windows_native_render.ps1'
    if (Test-Path -LiteralPath $renderScript) {
        try { foreach ($line in @(& $renderScript -Action RemoveFonts 2>&1)) { Say ([string]$line) } } catch { Say ('MathType fonts: ' + $_.Exception.Message) }
    }
    if ($OutputPath) { [IO.File]::WriteAllLines($OutputPath, [string[]]$lines) }
    exit 0
}

# ---------- as the user ----------
# The Office preferences in the registry (language, shortcuts, formula formats, the Office and WPS
# switches) move into the settings folder; VisualTeX puts them back when it is installed again.
$officeKey = 'HKCU:\Software\VisualTeX\OfficeIntegration'
if (-not $RemoveSettings -and (Test-Path -LiteralPath $officeKey)) {
    $preferences = [ordered]@{}
    foreach ($p in (Get-ItemProperty -LiteralPath $officeKey).PSObject.Properties) {
        if ($p.Name -like 'PS*') { continue }
        if ($p.Name -match '(Shortcut|Format)$' -or $p.Name -in @('Language', 'MicrosoftOffice', 'Wps')) { $preferences[$p.Name] = $p.Value }
    }
    if ($preferences.Count -gt 0) {
        $target = Join-Path $settingsRoot 'office\office-preferences.json'
        New-Item -ItemType Directory -Force -Path (Split-Path -Parent $target) | Out-Null
        [IO.File]::WriteAllText($target, (ConvertTo-Json $preferences), (New-Object Text.UTF8Encoding($false)))
        Say ('kept the Office preferences in ' + $target)
    }
}
Remove-KeyIn 'HKCU' 'Software\VisualTeX'
foreach ($office in 'Word', 'PowerPoint') {
    foreach ($id in 'VisualTeX.WordVsto', 'VisualTeX.PowerPointVsto') { Remove-KeyIn 'HKCU' ("Software\Microsoft\Office\$office\Addins\$id") }
    foreach ($version in '16.0', '15.0') { foreach ($bucket in 'DisabledItems', 'StartupItems', 'DisabledAddins') { Remove-MatchingValues 'HKCU' ("Software\Microsoft\Office\$version\$office\Resiliency\$bucket") $namePattern } }
}
foreach ($id in $progIds) { Remove-KeyIn 'HKCU' ('Software\Classes\' + $id) }
foreach ($id in $clsids) { Remove-KeyIn 'HKCU' ('Software\Classes\CLSID\' + $id) }
foreach ($id in $typeLibs) { Remove-KeyIn 'HKCU' ('Software\Classes\TypeLib\' + $id) }
Remove-KeyIn 'HKCU' 'Software\Classes\VisualTeX.Markdown'
Remove-KeyIn 'HKCU' 'Software\Classes\Applications\visualtex.exe'
foreach ($ext in '.md', '.markdown') { Remove-MatchingValues 'HKCU' ("Software\Classes\$ext\OpenWithProgids") 'VisualTeX' }
Remove-MatchingValues 'HKCU' 'Software\Microsoft\Windows\CurrentVersion\Run' 'VisualTeX|visualtex'
foreach ($app in 'WPS', 'WPP', 'ET') { foreach ($list in 'AddinsWL', 'AddinsBL', 'AddinsCL', 'AddinsLF') { Remove-MatchingValues 'HKCU' ("Software\Kingsoft\Office\$app\$list") $namePattern } }
foreach ($app in 'wps', 'wpp', 'et', 'pdf') { Remove-MatchingValues 'HKCU' ("Software\Kingsoft\Office\6.0\$app\DiagDisabledAddins") $namePattern }
# the HTTPS certificate, if the Office uninstall could not remove it
foreach ($store in 'Root', 'My') {
    foreach ($cert in @(Get-ChildItem -Path ("Cert:\CurrentUser\" + $store) -ErrorAction SilentlyContinue | Where-Object { $_.Subject -match 'VisualTeX' -or $_.FriendlyName -match 'VisualTeX' -or $_.Issuer -match 'VisualTeX' })) {
        Remove-KeyIn 'HKCU' ("Software\Microsoft\SystemCertificates\$store\Certificates\" + $cert.Thumbprint)
    }
}
# MathType's WPS template comes back in place of VisualTeX's placeholder
$startupTemplate = Join-Path $env:APPDATA 'kingsoft\wps\startup\MathType Commands 2016.dotm'
if ((Test-Path -LiteralPath $startupTemplate -PathType Leaf) -and (Get-FileHash -LiteralPath $startupTemplate -Algorithm SHA256).Hash -eq $placeholderSha256) {
    Remove-Item -LiteralPath $startupTemplate -Force -ErrorAction SilentlyContinue
    $kept = Join-Path $env:LOCALAPPDATA 'VisualTeX\office\mathtype-wps\MathType Commands 2016.dotm'
    if (Test-Path -LiteralPath $kept -PathType Leaf) { Move-Item -LiteralPath $kept -Destination $startupTemplate -Force -ErrorAction SilentlyContinue; Say 'MathType''s WPS template is back in WPS''s startup folder' }
    else { Say 'removed VisualTeX''s placeholder for MathType''s WPS template' }
}

# Caches, logs and files a new installation writes again. The WebView folder holds the app's settings
# besides its caches: it goes whole only when the settings folder has their copy (settings-snapshot.json,
# written by VisualTeX 1.3.1 and later while it runs); without one only the caches go.
$webView = Join-Path $env:LOCALAPPDATA 'com.visualtex.studio'
if ($RemoveSettings -or (Test-Path -LiteralPath (Join-Path $settingsRoot 'settings-snapshot.json'))) {
    Remove-PathNow $webView
} elseif (Test-Path -LiteralPath $webView) {
    $profileRoot = Join-Path $webView 'EBWebView'
    foreach ($relative in 'Default\Cache', 'Default\Code Cache', 'Default\GPUCache', 'Default\DawnGraphiteCache', 'Default\DawnWebGPUCache', 'Default\Service Worker\CacheStorage', 'Default\Service Worker\ScriptCache', 'GrShaderCache', 'GraphiteDawnCache', 'ShaderCache', 'Crashpad') {
        Remove-PathNow (Join-Path $profileRoot $relative)
    }
    Say ('kept the settings in ' + $webView + ' (no copy of them in the settings folder yet)')
}
foreach ($relative in 'anywhere.log', 'office\logs', 'office\formulas', 'office\localhost-cert.pem', 'office\localhost-key.pem', 'office\certificate.json', 'office\install.json') {
    Remove-PathNow (Join-Path $settingsRoot $relative)
}
Get-ChildItem -LiteralPath $env:TEMP -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '^visualtex' } | ForEach-Object { Remove-PathNow $_.FullName }

# The OCR environment: where the configuration points, the default and older places, VisualTeX's own
# folders only (the marker file); kept unless asked.
$ocrRoots = New-Object System.Collections.Generic.List[string]
$storageFile = Join-Path $sharedRoot 'ocr-storage.json'
if (Test-Path -LiteralPath $storageFile) { try { $configured = (Get-Content -LiteralPath $storageFile -Raw | ConvertFrom-Json).runtimeRoot; if ($configured) { [void]$ocrRoots.Add([string]$configured) } } catch { } }
foreach ($candidate in (Join-Path $env:LOCALAPPDATA 'VisualTeXData\ocr-runtime'), (Join-Path $settingsRoot 'ocr-runtime'), (Join-Path $env:APPDATA 'VisualTeX-OCR'), (Join-Path $env:LOCALAPPDATA 'VisualTeX-OCR'), (Join-Path $env:USERPROFILE 'VisualTeX-OCR'), (Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'VisualTeX-OCR')) { [void]$ocrRoots.Add($candidate) }
$keptOcr = @()
foreach ($root in @($ocrRoots | Select-Object -Unique)) {
    if (-not (Test-Path -LiteralPath $root) -or -not (Test-OcrRoot $root)) { continue }
    if ($RemoveOcr) { Remove-PathNow $root }
    elseif (Test-HasPython $root) { $keptOcr += $root }
    else { Remove-PathNow $root }   # what moving the storage away left (marker, logs, empty caches)
}
if ($RemoveOcr) { Remove-PathNow $storageFile }
# the default OCR folder's parent once no file is left in it (with or without the box: empty folders are no
# environment; they stayed after every uninstall that kept the OCR environment)
$dataRoot = Join-Path $env:LOCALAPPDATA 'VisualTeXData'
if ((Test-Path -LiteralPath $dataRoot) -and -not (Get-ChildItem -LiteralPath $dataRoot -Recurse -File -Force -ErrorAction SilentlyContinue)) { Remove-PathNow $dataRoot }

# Settings and formula documents, only when asked; an OCR environment kept above stays where it is
if ($RemoveSettings) {
    foreach ($folder in $settingsRoot, $sharedRoot) {
        if (-not (Test-Path -LiteralPath $folder)) { continue }
        foreach ($item in @(Get-ChildItem -LiteralPath $folder -Force)) {
            if ($keptOcr -contains $item.FullName -or ($item.Name -eq 'ocr-storage.json' -and $keptOcr.Count -gt 0)) { continue }
            Remove-PathNow $item.FullName
        }
        if (-not (Get-ChildItem -LiteralPath $folder -Force)) { Remove-PathNow $folder }
    }
}

# What only an administrator can remove: one prompt, only if something is there
$needAdmin = (@(Get-OwnProducts).Count -gt 0) -or [bool](Get-CimInstance Win32_Service -Filter "Name='VisualTeX.NativeRender'" -ErrorAction SilentlyContinue)
foreach ($root in @($env:ProgramW6432, ${env:ProgramFiles(x86)}) | Where-Object { $_ }) { foreach ($leaf in 'VisualTeX', 'VisualTeX Input') { if (Test-Path -LiteralPath (Join-Path $root $leaf)) { $needAdmin = $true } } }
if (Test-Path -LiteralPath (Join-Path $env:windir 'ServiceProfiles\VisualTeX.NativeRender')) { $needAdmin = $true }
if (@(Get-ChildItem -LiteralPath (Join-Path $env:windir 'Fonts') -File -ErrorAction SilentlyContinue | Where-Object { $_.BaseName -match '_visualtex(_\d+)?$' }).Count -gt 0) { $needAdmin = $true }
foreach ($id in $clsids) { if (Test-KeyIn 'HKLM' ('Software\Classes\CLSID\' + $id)) { $needAdmin = $true } }
if (Test-KeyIn 'HKLM' 'SOFTWARE\VisualTeX') { $needAdmin = $true }
if ($needAdmin) {
    $output = Join-Path $env:TEMP ('vtx-uninstall-' + [Guid]::NewGuid().ToString('N') + '.txt')
    $arguments = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $PSCommandPath + '"'), '-Elevated', '-InstallDir', ('"' + $InstallDir + '"'), '-OutputPath', ('"' + $output + '"'))
    try {
        $child = Start-Process -FilePath (Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe') -ArgumentList $arguments -Verb RunAs -WindowStyle Hidden -Wait -PassThru
        if (Test-Path -LiteralPath $output) { foreach ($line in [IO.File]::ReadAllLines($output)) { Say $line }; Remove-Item -LiteralPath $output -Force -ErrorAction SilentlyContinue }
        Say ('administrator cleanup finished (exit code ' + $child.ExitCode + ')')
    } catch {
        Say ('administrator cleanup was not allowed: ' + $_.Exception.Message + ' - Program Files\VisualTeX and the HKLM entries stay until it runs')
    }
}

if ($keptOcr.Count -gt 0) { Say ('kept the OCR environment: ' + ($keptOcr -join '; ')) }
if (-not $RemoveSettings) { Say ('kept settings and formula documents: ' + $settingsRoot + '; ' + $sharedRoot) }
exit 0
