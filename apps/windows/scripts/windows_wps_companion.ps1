# WPS Office loads the same COM add-ins as Microsoft Office, from the Office registration in
# WPS's own registry view. The machine-wide MSI registers the add-ins for the bitness of the
# installed Microsoft Office, so a WPS of the other bitness (32-bit WPS beside 64-bit Office,
# the usual case for older and free WPS editions) never sees them: they are not even listed
# under WPS's COM add-ins. This file adds a second registration for that bitness, from the
# other VisualTeX VSTO package, and removes it again. Dot-sourced by install_windows_vsto.ps1
# and uninstall_windows_vsto.ps1; ASCII only (Windows PowerShell 5.1 reads it as ANSI).

$script:WpsCompanionClsids = [ordered]@{
    Word = "{F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1}"
    PowerPoint = "{7E586D2D-57B0-4D14-AB24-EBA9021A5E6D}"
}
$script:WpsCompanionOleClsid = "{8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B}"
$script:WpsCompanionOleAppId = "{3C72FF7F-B04A-4FD0-AA7D-61D110D8B3C1}"
$script:WpsCompanionOleInterfaces = [ordered]@{
    "{6C672AF0-7321-4D21-B325-868CB34592C2}" = "IVisualTeXFormulaObject"
    "{A59B7798-6F24-4CF0-B378-E951BFFAFB3A}" = "IVisualTeXFormulaMetadata"
}
$script:WpsCompanionTypeLib = "{DF66EC66-3B3A-4675-A7BE-30456A04EB96}"
$script:WpsCompanionRecordKey = "Software\VisualTeX\OfficeIntegration"

function Get-WpsCompanionView([string]$Architecture) {
    if ($Architecture -eq "x86") { return [Microsoft.Win32.RegistryView]::Registry32 }
    return [Microsoft.Win32.RegistryView]::Registry64
}

function Get-PeArchitecture([string]$Path) {
    try {
        $stream = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
        try {
            $reader = New-Object IO.BinaryReader($stream)
            if ($reader.ReadUInt16() -ne 0x5A4D) { return $null }
            $stream.Position = 0x3C
            $stream.Position = $reader.ReadInt32()
            if ($reader.ReadUInt32() -ne 0x4550) { return $null }
            $machine = $reader.ReadUInt16()
            if ($machine -eq 0x14C) { return "x86" }
            if ($machine -eq 0x8664) { return "x64" }
            return $null
        } finally { $stream.Dispose() }
    } catch {
        return $null
    }
}

# Every WPS Office installation on this PC: per-user (LocalAppData) and per-machine, from the
# install roots WPS records and from the usual folders. Path, version and bitness of wps.exe.
function Get-WpsInstallations {
    $candidates = New-Object System.Collections.Generic.List[string]
    foreach ($hive in @([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryHive]::LocalMachine)) {
        foreach ($view in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
            $baseKey = $null
            try {
                $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
                $key = $baseKey.OpenSubKey("Software\Kingsoft\Office\6.0\Common", $false)
                if ($null -eq $key) { continue }
                try {
                    $root = [string]$key.GetValue("InstallRoot", "")
                    if (-not [string]::IsNullOrWhiteSpace($root)) {
                        [void]$candidates.Add((Join-Path $root.Trim().Trim('"') "office6\wps.exe"))
                    }
                } finally { $key.Dispose() }
            } catch {
            } finally {
                if ($null -ne $baseKey) { $baseKey.Dispose() }
            }
        }
    }
    foreach ($parent in @(
        (Join-Path $env:LOCALAPPDATA "Kingsoft\WPS Office"),
        $(if ($env:ProgramW6432) { Join-Path $env:ProgramW6432 "Kingsoft\WPS Office" }),
        $(if (${env:ProgramFiles(x86)}) { Join-Path ${env:ProgramFiles(x86)} "Kingsoft\WPS Office" })
    )) {
        if ([string]::IsNullOrWhiteSpace($parent) -or -not (Test-Path -LiteralPath $parent -PathType Container)) { continue }
        foreach ($folder in @(Get-ChildItem -LiteralPath $parent -Directory -ErrorAction SilentlyContinue)) {
            [void]$candidates.Add((Join-Path $folder.FullName "office6\wps.exe"))
        }
    }

    $seen = New-Object System.Collections.Generic.HashSet[string]([StringComparer]::OrdinalIgnoreCase)
    $installations = New-Object System.Collections.Generic.List[object]
    foreach ($candidate in $candidates) {
        if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) { continue }
        $path = [IO.Path]::GetFullPath($candidate)
        if (-not $seen.Add($path)) { continue }
        $architecture = Get-PeArchitecture $path
        if ($null -eq $architecture) { continue }
        $version = ""
        try { $version = ([string](Get-Item -LiteralPath $path).VersionInfo.ProductVersion) -replace '\s*,\s*', '.' } catch { }
        [void]$installations.Add([pscustomobject]@{
            Path = $path
            Architecture = $architecture
            Version = $version
        })
    }
    return $installations.ToArray()
}

function Set-WpsCompanionValue($Key, [string]$Name, $Value, [Microsoft.Win32.RegistryValueKind]$Kind = [Microsoft.Win32.RegistryValueKind]::String) {
    $Key.SetValue($Name, $Value, $Kind)
}

function Open-WpsCompanionKey($BaseKey, [string]$SubKey) {
    $key = $BaseKey.CreateSubKey($SubKey)
    if ($null -eq $key) { throw "Unable to create registry key $SubKey." }
    return $key
}

function Write-WpsCompanionRegistry([string]$Architecture, [string]$Root) {
    $view = Get-WpsCompanionView $Architecture
    $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view)
    try {
        foreach ($hostName in @("Word", "PowerPoint")) {
            $progId = "VisualTeX.${hostName}Vsto"
            $className = "$progId.ThisAddIn"
            $assemblyPath = Join-Path $Root "$progId.dll"
            $assemblyName = [Reflection.AssemblyName]::GetAssemblyName($assemblyPath)
            # The MSI's form: mscoree does not decode an escaped URI (%20 -> file not found).
            $codeBase = "file:///$assemblyPath"
            $clsid = $script:WpsCompanionClsids[$hostName]

            $classKey = Open-WpsCompanionKey $baseKey "Software\Classes\CLSID\$clsid"
            try {
                Set-WpsCompanionValue $classKey "" "VisualTeX $hostName VSTO"
                foreach ($subKey in @("InprocServer32", "InprocServer32\$($assemblyName.Version)")) {
                    $server = Open-WpsCompanionKey $classKey $subKey
                    try {
                        if ($subKey -eq "InprocServer32") {
                            Set-WpsCompanionValue $server "" "mscoree.dll"
                            Set-WpsCompanionValue $server "ThreadingModel" "Both"
                        }
                        Set-WpsCompanionValue $server "Class" $className
                        Set-WpsCompanionValue $server "Assembly" $assemblyName.FullName
                        Set-WpsCompanionValue $server "RuntimeVersion" "v4.0.30319"
                        Set-WpsCompanionValue $server "CodeBase" $codeBase
                    } finally { $server.Dispose() }
                }
                $progIdKey = Open-WpsCompanionKey $classKey "ProgId"
                try { Set-WpsCompanionValue $progIdKey "" $progId } finally { $progIdKey.Dispose() }
                (Open-WpsCompanionKey $classKey "Implemented Categories\{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}").Dispose()
            } finally { $classKey.Dispose() }

            # The ProgID key is shared by both registry views; the MSI already wrote it.
            $addin = Open-WpsCompanionKey $baseKey "Software\Microsoft\Office\$hostName\Addins\$progId"
            try {
                Set-WpsCompanionValue $addin "FriendlyName" "VisualTeX"
                Set-WpsCompanionValue $addin "Description" "VisualTeX native $hostName formula integration"
                Set-WpsCompanionValue $addin "LoadBehavior" 3 ([Microsoft.Win32.RegistryValueKind]::DWord)
                Set-WpsCompanionValue $addin "CommandLineSafe" 0 ([Microsoft.Win32.RegistryValueKind]::DWord)
                Set-WpsCompanionValue $addin "CodeBase" $codeBase
            } finally { $addin.Dispose() }

            $listName = if ($hostName -eq "Word") { "WPS" } else { "WPP" }
            $whitelist = Open-WpsCompanionKey $baseKey "Software\Kingsoft\Office\$listName\AddinsWL"
            try { Set-WpsCompanionValue $whitelist $progId "" } finally { $whitelist.Dispose() }
        }

        # The formula OLE class for this view, served by the package's own LocalServer.
        $server = Join-Path $Root "VisualTeX.FormulaOleServer.exe"
        $oleKey = Open-WpsCompanionKey $baseKey "Software\Classes\CLSID\$script:WpsCompanionOleClsid"
        try {
            Set-WpsCompanionValue $oleKey "" "VisualTeX Formula"
            Set-WpsCompanionValue $oleKey "AppID" $script:WpsCompanionOleAppId
            foreach ($entry in @(
                @("AuxUserType\2", "Formula"),
                @("AuxUserType\3", "VisualTeX"),
                @("DataFormats\GetSet\0", "14,1,64,1"),
                @("DataFormats\GetSet\1", "3,1,32,1"),
                @("DataFormats\GetSet\2", "PNG,1,1,1"),
                @("InprocHandler32", "Ole32.dll"),
                @("ProgID", "VisualTeX.Formula.1"),
                @("TypeLib", $script:WpsCompanionTypeLib),
                @("Version", "1.0"),
                @("VersionIndependentProgID", "VisualTeX.Formula"),
                @("DefaultIcon", ('"{0}",0' -f $server)),
                @("Verb\0", "&Edit,0,2"),
                @("Verb\1", "&Open,0,2"),
                @("MiscStatus", "0"),
                @("MiscStatus\1", "147473")
            )) {
                $sub = Open-WpsCompanionKey $oleKey $entry[0]
                try { Set-WpsCompanionValue $sub "" $entry[1] } finally { $sub.Dispose() }
            }
            (Open-WpsCompanionKey $oleKey "Insertable").Dispose()
            $localServer = Open-WpsCompanionKey $oleKey "LocalServer32"
            try {
                Set-WpsCompanionValue $localServer "" ('"{0}"' -f $server)
                Set-WpsCompanionValue $localServer "ServerExecutable" $server
            } finally { $localServer.Dispose() }
        } finally { $oleKey.Dispose() }
        foreach ($iid in $script:WpsCompanionOleInterfaces.Keys) {
            $interface = Open-WpsCompanionKey $baseKey "Software\Classes\Interface\$iid"
            try {
                Set-WpsCompanionValue $interface "" $script:WpsCompanionOleInterfaces[$iid]
                foreach ($entry in @(
                    @("ProxyStubClsid32", "{00020424-0000-0000-C000-000000000046}"),
                    @("Version", "1.0")
                )) {
                    $sub = Open-WpsCompanionKey $interface $entry[0]
                    try { Set-WpsCompanionValue $sub "" $entry[1] } finally { $sub.Dispose() }
                }
                $typeLib = Open-WpsCompanionKey $interface "TypeLib"
                try {
                    Set-WpsCompanionValue $typeLib "" $script:WpsCompanionTypeLib
                    Set-WpsCompanionValue $typeLib "Version" "1.0"
                } finally { $typeLib.Dispose() }
            } finally { $interface.Dispose() }
        }
        # TypeLib is shared by both views: add only this platform's entry beside the MSI's.
        $platformKey = if ($Architecture -eq "x86") { "win32" } else { "win64" }
        $typeLibPath = Open-WpsCompanionKey $baseKey "Software\Classes\TypeLib\$script:WpsCompanionTypeLib\1.0\0\$platformKey"
        try { Set-WpsCompanionValue $typeLibPath "" $server } finally { $typeLibPath.Dispose() }
    } finally {
        $baseKey.Dispose()
    }
}

function Remove-WpsCompanionRegistry([string]$Architecture) {
    $view = Get-WpsCompanionView $Architecture
    $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view)
    try {
        foreach ($subKey in @(
            "Software\Classes\CLSID\$($script:WpsCompanionClsids.Word)",
            "Software\Classes\CLSID\$($script:WpsCompanionClsids.PowerPoint)",
            "Software\Microsoft\Office\Word\Addins\VisualTeX.WordVsto",
            "Software\Microsoft\Office\PowerPoint\Addins\VisualTeX.PowerPointVsto",
            "Software\Classes\CLSID\$script:WpsCompanionOleClsid"
        ) + @($script:WpsCompanionOleInterfaces.Keys | ForEach-Object { "Software\Classes\Interface\$_" })) {
            try { $baseKey.DeleteSubKeyTree($subKey, $false) } catch { }
        }
        $platformKey = if ($Architecture -eq "x86") { "win32" } else { "win64" }
        try { $baseKey.DeleteSubKeyTree("Software\Classes\TypeLib\$script:WpsCompanionTypeLib\1.0\0\$platformKey", $false) } catch { }
        foreach ($entry in @(@("WPS", "VisualTeX.WordVsto"), @("WPP", "VisualTeX.PowerPointVsto"))) {
            $whitelist = $baseKey.OpenSubKey("Software\Kingsoft\Office\$($entry[0])\AddinsWL", $true)
            if ($null -eq $whitelist) { continue }
            try { $whitelist.DeleteValue($entry[1], $false) } finally { $whitelist.Dispose() }
        }
    } finally {
        $baseKey.Dispose()
    }
}

# The record of the extra registration (HKLM, 64-bit view): platform and folder.
function Get-WpsCompanionRecord {
    $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
    try {
        $key = $baseKey.OpenSubKey($script:WpsCompanionRecordKey, $false)
        if ($null -eq $key) { return $null }
        try {
            $platform = [string]$key.GetValue("WpsCompanionPlatform", "")
            $root = [string]$key.GetValue("WpsCompanionRoot", "")
            if ($platform -notin @("x86", "x64")) { return $null }
            return [pscustomobject]@{ Architecture = $platform; Root = $root }
        } finally { $key.Dispose() }
    } finally { $baseKey.Dispose() }
}

function Set-WpsCompanionRecord([string]$Architecture, [string]$Root) {
    $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
    try {
        $key = $baseKey.CreateSubKey($script:WpsCompanionRecordKey)
        try {
            if ([string]::IsNullOrWhiteSpace($Architecture)) {
                $key.DeleteValue("WpsCompanionPlatform", $false)
                $key.DeleteValue("WpsCompanionRoot", $false)
            } else {
                $key.SetValue("WpsCompanionPlatform", $Architecture)
                $key.SetValue("WpsCompanionRoot", $Root)
            }
        } finally { $key.Dispose() }
    } finally { $baseKey.Dispose() }
}

function Resolve-WpsCompanionRoot([string]$Architecture) {
    $programFiles = if ($Architecture -eq "x86") {
        if (${env:ProgramFiles(x86)}) { ${env:ProgramFiles(x86)} } else { $env:ProgramFiles }
    } else {
        if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }
    }
    return Join-Path $programFiles "VisualTeX\WindowsOffice\VSTO-WPS"
}

# $PrimaryArchitecture: the bitness the MSI registered, whose keys must stay (empty when the
# MSI is being removed as well).
function Remove-WpsCompanionRegistration([string]$PrimaryArchitecture = "") {
    $record = Get-WpsCompanionRecord
    if ($null -eq $record) { return $false }
    if ($record.Architecture -ne $PrimaryArchitecture) { Remove-WpsCompanionRegistry $record.Architecture }
    $expectedRoot = Resolve-WpsCompanionRoot $record.Architecture
    if (-not [string]::IsNullOrWhiteSpace($record.Root) -and
        [string]::Equals([IO.Path]::GetFullPath($record.Root), [IO.Path]::GetFullPath($expectedRoot), [StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $expectedRoot -Recurse -Force -ErrorAction SilentlyContinue
        if (Test-Path -LiteralPath $expectedRoot) {
            Write-Warning "The WPS add-in folder is still in use and will remain until WPS is closed: $expectedRoot"
        }
    }
    Set-WpsCompanionRecord "" ""
    return $true
}

function Test-WpsCompanionFile([string]$Path, [string]$Sha256) {
    return (Test-Path -LiteralPath $Path -PathType Leaf) -and
        ((Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash -ieq $Sha256)
}

# Installs the add-ins of package $Architecture for the WPS of that bitness: the package's
# files go to Program Files\VisualTeX\WindowsOffice\VSTO-WPS and the add-ins, their COM
# classes, WPS's whitelist and the formula OLE class are registered in that registry view.
function Install-WpsCompanionRegistration {
    param(
        [Parameter(Mandatory = $true)][ValidateSet("x86", "x64")][string]$Architecture,
        [Parameter(Mandatory = $true)][string]$PackageDirectory,
        [Parameter(Mandatory = $true)][string]$PrimaryArchitecture,
        [Parameter(Mandatory = $true)][string]$LogDirectory
    )
    $msi = Join-Path $PackageDirectory "VisualTeX-WindowsOffice-VSTO-$Architecture.msi"
    $manifestPath = Join-Path $PackageDirectory "VisualTeX-WindowsOffice-VSTO-$Architecture.sha256.json"
    if (-not (Test-Path -LiteralPath $msi -PathType Leaf)) { throw "The $Architecture Office package for WPS is missing: $msi" }
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw "The $Architecture Office package manifest is missing: $manifestPath" }
    $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
    if ($manifest.architecture -ne $Architecture) { throw "Package manifest architecture '$($manifest.architecture)' is not '$Architecture'." }
    if ((Get-FileHash -LiteralPath $msi -Algorithm SHA256).Hash -ine $manifest.package.sha256) { throw "MSI SHA-256 mismatch: $msi" }

    $entries = @($manifest.word, $manifest.powerPoint, $manifest.formulaOleServer)
    if ($null -ne $manifest.dependencies) { $entries += @($manifest.dependencies) }
    $target = Resolve-WpsCompanionRoot $Architecture
    $current = $true
    foreach ($entry in $entries) {
        if (-not (Test-WpsCompanionFile (Join-Path $target $entry.file) $entry.sha256)) { $current = $false; break }
    }

    if (-not $current) {
        # An administrative image only unpacks the package; nothing is registered.
        $extract = Join-Path ([IO.Path]::GetTempPath()) ("visualtex-wps-" + [Guid]::NewGuid().ToString("N"))
        $log = Join-Path $LogDirectory ("vsto-wps-$Architecture-" + (Get-Date -Format "yyyyMMdd-HHmmss") + ".log")
        try {
            $process = Start-Process msiexec.exe -ArgumentList @("/a", ('"{0}"' -f $msi), "/qn", ('TARGETDIR="{0}"' -f $extract), "/L*v", ('"{0}"' -f $log)) -Wait -PassThru -WindowStyle Hidden
            if ($process.ExitCode -ne 0) { throw "Unpacking the $Architecture Office package failed with exit code $($process.ExitCode). Log: $log" }
            $wordDll = Get-ChildItem -LiteralPath $extract -Recurse -Filter "VisualTeX.WordVsto.dll" -File | Select-Object -First 1
            if ($null -eq $wordDll) { throw "The $Architecture Office package did not contain VisualTeX.WordVsto.dll." }
            $source = $wordDll.DirectoryName
            foreach ($entry in $entries) {
                if (-not (Test-WpsCompanionFile (Join-Path $source $entry.file) $entry.sha256)) {
                    throw "Unpacked file hash mismatch: $($entry.file)"
                }
            }
            New-Item -Path $target -ItemType Directory -Force | Out-Null
            try {
                Copy-Item -Path (Join-Path $source "*") -Destination $target -Recurse -Force
            } catch {
                throw "Unable to update the WPS add-in files in $target. Close WPS Office and try again. $($_.Exception.Message)"
            }
        } finally {
            Remove-Item -LiteralPath $extract -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    foreach ($entry in $entries) {
        if (-not (Test-WpsCompanionFile (Join-Path $target $entry.file) $entry.sha256)) {
            throw "Installed WPS add-in file is missing or different: $(Join-Path $target $entry.file)"
        }
    }
    Write-WpsCompanionRegistry $Architecture $target
    Set-WpsCompanionRecord $Architecture $target
    return $target
}

# Reads the registration back in its view: add-in keys, COM classes, whitelist, OLE class.
function Assert-WpsCompanionRegistration([string]$Architecture, [string]$Root) {
    $view = Get-WpsCompanionView $Architecture
    $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view)
    try {
        foreach ($hostName in @("Word", "PowerPoint")) {
            $progId = "VisualTeX.${hostName}Vsto"
            $assemblyPath = Join-Path $Root "$progId.dll"
            $assemblyName = [Reflection.AssemblyName]::GetAssemblyName($assemblyPath)
            $expected = @(
                @("Software\Microsoft\Office\$hostName\Addins\$progId", "LoadBehavior", 3),
                @("Software\Microsoft\Office\$hostName\Addins\$progId", "CodeBase", "file:///$assemblyPath"),
                @("Software\Classes\$progId\CLSID", "", $script:WpsCompanionClsids[$hostName]),
                @("Software\Classes\CLSID\$($script:WpsCompanionClsids[$hostName])\InprocServer32", "", "mscoree.dll"),
                @("Software\Classes\CLSID\$($script:WpsCompanionClsids[$hostName])\InprocServer32", "Class", "$progId.ThisAddIn"),
                @("Software\Classes\CLSID\$($script:WpsCompanionClsids[$hostName])\InprocServer32", "Assembly", $assemblyName.FullName),
                @("Software\Classes\CLSID\$($script:WpsCompanionClsids[$hostName])\InprocServer32", "CodeBase", "file:///$assemblyPath")
            )
            foreach ($check in $expected) {
                $key = $baseKey.OpenSubKey($check[0], $false)
                if ($null -eq $key) { throw "WPS add-in registration is missing: HKLM\$($check[0]) ($Architecture view)" }
                try {
                    $actual = $key.GetValue($check[1], $null)
                    if ([string]$actual -ne [string]$check[2]) {
                        throw "WPS add-in registration HKLM\$($check[0])::$($check[1]) is '$actual'; expected '$($check[2])' ($Architecture view)."
                    }
                } finally { $key.Dispose() }
            }
        }
        $server = Join-Path $Root "VisualTeX.FormulaOleServer.exe"
        $localServer = $baseKey.OpenSubKey("Software\Classes\CLSID\$script:WpsCompanionOleClsid\LocalServer32", $false)
        if ($null -eq $localServer) { throw "The formula OLE class is not registered in the $Architecture view." }
        try {
            if (([string]$localServer.GetValue("", "")).Trim('"') -ine $server) { throw "The $Architecture formula OLE class does not point to $server." }
        } finally { $localServer.Dispose() }
        if ((Get-PeArchitecture (Join-Path $Root "VisualTeX.WordVsto.dll")) -ne $Architecture) {
            throw "The WPS add-in in $Root is not built for $Architecture."
        }
    } finally {
        $baseKey.Dispose()
    }
}

# Starts the add-in's COM class in a PowerShell of the given bitness, as WPS would.
function Test-WpsCompanionActivation([string]$Architecture, [string]$ProgId) {
    $windowsRoot = if ($env:WINDIR) { $env:WINDIR } else { "C:\Windows" }
    $powershell = if ($Architecture -eq "x86" -and [Environment]::Is64BitOperatingSystem) {
        Join-Path $windowsRoot "SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    } elseif (-not [Environment]::Is64BitProcess -and [Environment]::Is64BitOperatingSystem) {
        Join-Path $windowsRoot "Sysnative\WindowsPowerShell\v1.0\powershell.exe"
    } else {
        Join-Path $windowsRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
    }
    $command = "`$ErrorActionPreference='Stop'; `$t=[Type]::GetTypeFromProgID('$ProgId', `$true); `$o=[Activator]::CreateInstance(`$t); if (`$null -eq `$o) { exit 3 }; exit 0"
    $output = & $powershell -NoProfile -NonInteractive -ExecutionPolicy Bypass -Command $command 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "$ProgId did not start in a $Architecture process (exit $LASTEXITCODE). $($output -join ' ')"
    }
}
