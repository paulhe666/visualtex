<#
.SYNOPSIS
    Register or unregister the VisualTeX WPS Writer COM add-in.

.DESCRIPTION
    WPS loads IDTExtensibility2 COM add-ins from the regular Office Addins key
    plus its own whitelist:
        HKCU\Software\kingsoft\office\WPS\AddinsWL\<ProgID> = ""
    This script writes the COM class registration, the Office Addins entry and
    the WPS whitelist for both the 32-bit and 64-bit registry views. It writes
    per-user (HKCU) registrations by default; pass -PerMachine for HKLM (needs
    an elevated shell).

    The whole build output directory, not just the add-in DLL, must live next to
    the registered assembly: VisualTeX.WindowsOffice.Contracts.dll and the
    System.Text.Json dependency chain are resolved from the add-in directory.

.PARAMETER AssemblyPath
    Full path to VisualTeX.WpsWriterVsto.dll.

.PARAMETER Unregister
    Remove the registration instead of writing it.

.PARAMETER PerMachine
    Register under HKLM instead of HKCU.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File install_wps_addin.ps1 `
        -AssemblyPath "C:\path\to\VisualTeX.WpsWriterVsto\bin\x64\Release\net472\VisualTeX.WpsWriterVsto.dll"
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$AssemblyPath,

    [switch]$Unregister,

    [switch]$PerMachine
)

$ErrorActionPreference = 'Stop'

$progId = 'VisualTeX.WpsWriterVsto'
$clsid = '{5D6E7F80-91A2-4B3C-8D4E-5F60718293A4}'
$className = 'VisualTeX.WpsWriterVsto.ThisAddIn'
$addInCategory = '{62C8FE65-4EBB-45E7-B440-6E39B2CDBF29}'

if (-not (Test-Path -LiteralPath $AssemblyPath)) {
    throw "Assembly not found: $AssemblyPath"
}
$AssemblyPath = (Resolve-Path -LiteralPath $AssemblyPath).Path
$assemblyDirectory = Split-Path -Parent $AssemblyPath
$assemblyName = [System.Reflection.AssemblyName]::GetAssemblyName($AssemblyPath)
$codeBase = ([System.Uri]$AssemblyPath).AbsoluteUri

$contractsPath = Join-Path $assemblyDirectory 'VisualTeX.WindowsOffice.Contracts.dll'
if (-not (Test-Path -LiteralPath $contractsPath)) {
    Write-Warning "VisualTeX.WindowsOffice.Contracts.dll is not next to the add-in. Copy the full build output directory."
}

$hive = if ($PerMachine) { [Microsoft.Win32.RegistryHive]::LocalMachine } else { [Microsoft.Win32.RegistryHive]::CurrentUser }
$views = @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)

function Set-RegValue {
    param(
        [Microsoft.Win32.RegistryKey]$Root,
        [string]$SubKey,
        [string]$Name,
        $Value
    )
    $key = $Root.CreateSubKey($SubKey, $true)
    try {
        if ([string]::IsNullOrEmpty($Name)) {
            $key.SetValue('', [string]$Value)
        }
        else {
            $key.SetValue($Name, $Value)
        }
    }
    finally {
        $key.Close()
    }
}

function Remove-RegTree {
    param(
        [Microsoft.Win32.RegistryKey]$Root,
        [string]$SubKey
    )
    try { $Root.DeleteSubKeyTree($SubKey, $false) } catch { }
}

foreach ($view in $views) {
    $root = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
    try {
        if ($Unregister) {
            Remove-RegTree $root "Software\Classes\$progId"
            Remove-RegTree $root "Software\Classes\CLSID\$clsid"
            Remove-RegTree $root "Software\Microsoft\Office\Word\Addins\$progId"
            Remove-RegTree $root "Software\kingsoft\office\WPS\AddinsWL\$progId"
            Remove-RegTree $root "Software\Kingsoft\Office\WPS\AddinsWL\$progId"
            Write-Host "Unregistered $progId from $($view.ToString())."
            continue
        }

        Set-RegValue $root "Software\Classes\$progId" '' $progId
        Set-RegValue $root "Software\Classes\$progId\CLSID" '' $clsid
        Set-RegValue $root "Software\Classes\CLSID\$clsid" '' $progId
        Set-RegValue $root "Software\Classes\CLSID\$clsid\InprocServer32" '' 'mscoree.dll'
        Set-RegValue $root "Software\Classes\CLSID\$clsid\InprocServer32" 'ThreadingModel' 'Both'
        Set-RegValue $root "Software\Classes\CLSID\$clsid\InprocServer32" 'Class' $className
        Set-RegValue $root "Software\Classes\CLSID\$clsid\InprocServer32" 'Assembly' $assemblyName.FullName
        Set-RegValue $root "Software\Classes\CLSID\$clsid\InprocServer32" 'RuntimeVersion' 'v4.0.30319'
        Set-RegValue $root "Software\Classes\CLSID\$clsid\InprocServer32" 'CodeBase' $codeBase
        Set-RegValue $root "Software\Classes\CLSID\$clsid\Implemented Categories\$addInCategory" '' ''

        $addInKey = "Software\Microsoft\Office\Word\Addins\$progId"
        Set-RegValue $root $addInKey 'FriendlyName' 'VisualTeX WPS Writer'
        Set-RegValue $root $addInKey 'Description' 'VisualTeX MathType formula add-in for WPS Writer'
        Set-RegValue $root $addInKey 'LoadBehavior' 3
        Set-RegValue $root $addInKey 'CommandLineSafe' 0

        Set-RegValue $root "Software\kingsoft\office\WPS\AddinsWL\$progId" '' ''
        Set-RegValue $root "Software\Kingsoft\Office\WPS\AddinsWL\$progId" '' ''

        Write-Host "Registered $progId in $($view.ToString())."
    }
    finally {
        $root.Close()
    }
}

if ($Unregister) {
    Write-Host 'Restart WPS Writer to finish unloading the add-in.'
}
else {
    Write-Host "CodeBase: $codeBase"
    Write-Host 'Start WPS Writer and check the VisualTeX tab.'
}
