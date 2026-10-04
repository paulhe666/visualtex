param([ValidateSet('physics','math-exam','math-exam-a3')][string]$Document='physics',[Parameter(Mandatory=$true)][string]$Label)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$root=Join-Path $PSScriptRoot 'evidence\real-mathtype-20260926'
$source=Join-Path $repo 'apps\windows\src-windows\VisualTeX.WordVsto\bin\x64\Debug\net472'
$hash=(Get-FileHash (Join-Path $source 'VisualTeX.WordVsto.dll')).Hash
$stage=Join-Path $root ('candidates\'+$hash)
$keyPath='Software\Classes\CLSID\{F1B68342-F9C6-4E7D-A9C6-A2F64C3558A1}'
$user=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser,[Microsoft.Win32.RegistryView]::Registry64)
$machine=[Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine,[Microsoft.Win32.RegistryView]::Registry64)
$existing=$user.OpenSubKey($keyPath)
if($existing){$existing.Dispose();throw 'A per-user COM override already exists; refusing to replace it.'}
if(Test-Path (Join-Path $root ($Label+'-process.json'))){throw 'Candidate Word already started for this label'}
[void][IO.Directory]::CreateDirectory($stage)
foreach($file in Get-ChildItem $source -File){
 $target=Join-Path $stage $file.Name
 if(Test-Path $target){if((Get-FileHash $target).Hash -ne (Get-FileHash $file.FullName).Hash){throw ('Immutable candidate dependency differs: '+$target)}}
 else{[IO.File]::Copy($file.FullName,$target,$false)}
}
function Copy-Registration($from,$to){
 foreach($n in $from.GetValueNames()){$to.SetValue($n,$from.GetValue($n,$null,[Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames),$from.GetValueKind($n))}
 foreach($n in $from.GetSubKeyNames()){$f=$from.OpenSubKey($n);$t=$to.CreateSubKey($n);try{Copy-Registration $f $t}finally{$f.Dispose();$t.Dispose()}}
}
$base=$machine.OpenSubKey($keyPath)
if(!$base){throw 'Installed machine Word add-in class is missing'}
$created=$false
$codebase=([Uri](Join-Path $stage 'VisualTeX.WordVsto.dll')).AbsoluteUri
try{
 $new=$user.CreateSubKey($keyPath);$created=$true
 try{
  Copy-Registration $base $new
  foreach($sub in @('InprocServer32','InprocServer32\1.0.0.0')){$s=$new.OpenSubKey($sub,$true);try{$s.SetValue('CodeBase',$codebase,[Microsoft.Win32.RegistryValueKind]::String)}finally{$s.Dispose()}}
 }finally{$new.Dispose()}
 & (Join-Path $PSScriptRoot 'start-real-doc-stage.ps1') -Document $Document -Label $Label
 $log=Join-Path $root ($Label+'-hook.log')
 $deadline=[DateTime]::UtcNow.AddSeconds(15)
 do {
  if((Test-Path $log) -and ([IO.File]::ReadAllText($log).Contains('assembly='+ (Join-Path $stage 'VisualTeX.WordVsto.dll')))){break}
  if([DateTime]::UtcNow -ge $deadline){throw 'Candidate assembly not confirmed; preserve recorded PID and inspect.'}
  Start-Sleep -Milliseconds 300
 }while($true)
 [pscustomobject]@{stage=$stage;wordSha256=$hash;candidateVerified=$true;originalWordProcessesUntouched=$true}|ConvertTo-Json
}finally{
 if($created){
  $s=$user.OpenSubKey($keyPath+'\InprocServer32');$current=if($s){$s.GetValue('CodeBase')}else{$null};if($s){$s.Dispose()}
  if($current -ne $codebase){throw 'COM override changed concurrently; refusing to delete unknown registration'}
  $user.DeleteSubKeyTree($keyPath,$false)
  [IO.File]::WriteAllText((Join-Path $root ($Label+'-registration-restored.json')),(@{restored=$true;originalUserKeyExisted=$false;stage=$stage;hash=$hash;time=[DateTimeOffset]::Now.ToString('o')}|ConvertTo-Json),[Text.UTF8Encoding]::new($false))
  Write-Output 'TEMPORARY COM OVERRIDE REMOVED; installed registration unchanged.'
 }
 $base.Dispose();$user.Dispose();$machine.Dispose()
}
