param([string]$Label='stage01')
$ErrorActionPreference='Stop'
if($env:VISUALTEX_VSTO_ACCEPTANCE -eq '1'){throw 'Manual-service acceptance mode must not be enabled.'}
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$source=Join-Path $root 'apps\windows\src-windows\VisualTeX.WordVsto\bin\x64\Release\net472'
$hash=(Get-FileHash -LiteralPath (Join-Path $source 'VisualTeX.WordVsto.dll') -Algorithm SHA256).Hash
$installBase=Join-Path $env:LOCALAPPDATA 'VisualTeX\office\remediation-builds'
$destination=Join-Path $installBase $hash
$keyPath='Software\Classes\CLSID\{F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1}'
$userBase=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,[Microsoft.Win32.RegistryView]::Registry64)
$machineBase=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,[Microsoft.Win32.RegistryView]::Registry64)
function Copy-Registration($from,$to){
 foreach($name in $from.GetValueNames()){$to.SetValue($name,$from.GetValue($name,$null,[Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames),$from.GetValueKind($name))}
 foreach($name in $from.GetSubKeyNames()){
  $childFrom=$from.OpenSubKey($name);$childTo=$to.CreateSubKey($name)
  try{Copy-Registration $childFrom $childTo}finally{$childFrom.Dispose();$childTo.Dispose()}
 }
}
$existing=$userBase.OpenSubKey($keyPath)
if($existing){
 try{
  $existingServer=$existing.OpenSubKey('InprocServer32')
  try{$oldCodeBase=[string]$existingServer.GetValue('CodeBase')}finally{$existingServer.Dispose()}
  if(-not ([Uri]$oldCodeBase).LocalPath.StartsWith($installBase+'\',[StringComparison]::OrdinalIgnoreCase)) {throw 'An unrelated per-user Word COM registration exists; refusing to replace it.'}
 }finally{$existing.Dispose()}
}
[void][IO.Directory]::CreateDirectory($destination)
$files=@()
foreach($file in Get-ChildItem -LiteralPath $source -File){
 $target=Join-Path $destination $file.Name
 $expected=(Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
 if(Test-Path -LiteralPath $target){if((Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -ne $expected){throw "Immutable stage file differs: $target"}}
 else{Copy-Item -LiteralPath $file.FullName -Destination $target}
 $files+=@{file=$file.Name;sha256=$expected}
}
$machine=$machineBase.OpenSubKey($keyPath)
if(!$machine){throw 'The installed VisualTeX Word COM registration is missing.'}
$user=$userBase.CreateSubKey($keyPath)
try{
 Copy-Registration $machine $user
 $codeBase=([Uri](Join-Path $destination 'VisualTeX.WordVsto.dll')).AbsoluteUri
 foreach($sub in @('InprocServer32','InprocServer32\1.0.0.0')){
  $server=$user.OpenSubKey($sub,$true)
  try{$server.SetValue('CodeBase',$codeBase,[Microsoft.Win32.RegistryValueKind]::String)}finally{$server.Dispose()}
 }
}finally{$machine.Dispose();$user.Dispose();$userBase.Dispose();$machineBase.Dispose()}
$record=@{label=$Label;time=[DateTime]::UtcNow.ToString('o');source=$source;directory=$destination;wordSha256=$hash;registryKey=$keyPath;previousUserCodeBase=$oldCodeBase;existingWordProcesses=@(Get-Process WINWORD | Select-Object Id,StartTime,Path);files=$files;manualServiceAcceptance=$false}
$record | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath (Join-Path $PSScriptRoot "evidence\$Label-installed.json") -Encoding UTF8
Write-Output ("STAGE_INSTALLED|"+$destination)
