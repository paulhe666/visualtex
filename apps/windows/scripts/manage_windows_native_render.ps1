[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][ValidateSet('Install','Verify','Uninstall')][string]$Action,
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
