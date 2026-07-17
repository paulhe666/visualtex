[CmdletBinding()]
param(
    [string]$ExpectedAppVersion = "1.2.0",
    [string]$ExpectedOfficeMsiVersion = "1.0.33.0"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$installRoot = Join-Path $env:LOCALAPPDATA "VisualTeX"
$vstoRoot = Join-Path $installRoot "WindowsOffice\VSTO"
$appPath = Join-Path $installRoot "visualtex.exe"
$wordDll = Join-Path $vstoRoot "VisualTeX.WordVsto.dll"
$powerPointDll = Join-Path $vstoRoot "VisualTeX.PowerPointVsto.dll"
$oleServer = Join-Path $vstoRoot "VisualTeX.FormulaOleServer.exe"
foreach ($path in @($appPath, $wordDll, $powerPointDll, $oleServer)) {
    if (-not (Test-Path -LiteralPath $path)) { throw "Installed file is missing: $path" }
}

$appVersion = (Get-Item -LiteralPath $appPath).VersionInfo.ProductVersion
Write-Host "Installed VisualTeX ProductVersion=$appVersion"
if (-not $appVersion.StartsWith($ExpectedAppVersion, [StringComparison]::Ordinal)) {
    throw "Installed VisualTeX version is $appVersion; expected $ExpectedAppVersion."
}

$expectedWordDll = Join-Path $root "src-windows\VisualTeX.WordVsto\bin\x64\Release\net48\VisualTeX.WordVsto.dll"
$expectedPowerPointDll = Join-Path $root "src-windows\VisualTeX.PowerPointVsto\bin\x64\Release\net48\VisualTeX.PowerPointVsto.dll"
$expectedOleServer = Join-Path $root "src-windows\artifacts\formula-ole-server\x64\Release\VisualTeX.FormulaOleServer.exe"
foreach ($entry in @(
    @{ Installed = $wordDll; Expected = $expectedWordDll; Label = "Word VSTO" },
    @{ Installed = $powerPointDll; Expected = $expectedPowerPointDll; Label = "PowerPoint VSTO" },
    @{ Installed = $oleServer; Expected = $expectedOleServer; Label = "Formula OLE server" }
)) {
    $installedHash = (Get-FileHash -LiteralPath $entry.Installed -Algorithm SHA256).Hash
    $expectedHash = (Get-FileHash -LiteralPath $entry.Expected -Algorithm SHA256).Hash
    Write-Host ("{0} SHA256={1}" -f $entry.Label, $installedHash)
    if ($installedHash -ne $expectedHash) {
        throw "$($entry.Label) installed hash does not match the current release build."
    }
}

$wordAddIn = Get-ItemProperty -LiteralPath "HKCU:\Software\Microsoft\Office\Word\Addins\VisualTeX.WordVsto"
$powerPointAddIn = Get-ItemProperty -LiteralPath "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\VisualTeX.PowerPointVsto"
if ([int]$wordAddIn.LoadBehavior -ne 3) { throw "Word LoadBehavior is not 3." }
if ([int]$powerPointAddIn.LoadBehavior -ne 3) { throw "PowerPoint LoadBehavior is not 3." }
Write-Host "Word and PowerPoint LoadBehavior=3."

$clsid = "{8FF7F5AA-0D60-48D5-ADBD-65A64B4C827B}"
$localServerKey = "Registry::HKEY_CURRENT_USER\Software\Classes\CLSID\$clsid\LocalServer32"
$localServer = (Get-Item -LiteralPath $localServerKey).GetValue("")
$serverExecutable = (Get-ItemProperty -LiteralPath $localServerKey).ServerExecutable
Write-Host "OLE LocalServer32=$localServer"
if (($localServer.Trim('"')) -ne $oleServer -or $serverExecutable -ne $oleServer) {
    throw "Formal OLE registration does not point to the installed $ExpectedAppVersion server."
}

$installer = New-Object -ComObject WindowsInstaller.Installer
$officeProductCodes = @(
    "{B4E2A791-6C35-4F8D-9A20-7E1C5B3D8642}",
    "{5F9C2D18-A743-4B6E-8D01-C2E7A5943B60}"
)
$installedOfficeVersions = @()
foreach ($productCode in $officeProductCodes) {
    try {
        $version = $installer.ProductInfo($productCode, "VersionString")
        if ($version) {
            $installedOfficeVersions += $version
            Write-Host "Office integration $productCode VersionString=$version"
        }
    }
    catch { }
}
if ($installedOfficeVersions.Count -eq 0 -or $installedOfficeVersions -notcontains $ExpectedOfficeMsiVersion) {
    throw "Office integration MSI $ExpectedOfficeMsiVersion is not registered as installed."
}

Write-Host "Installed VisualTeX Windows release passed verification."
