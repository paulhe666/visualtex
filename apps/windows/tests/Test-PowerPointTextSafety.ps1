param(
    [Parameter(Mandatory=$true)][string]$CandidateDirectory,
    [Parameter(Mandatory=$true)][string]$ArtifactDirectory,
    [ValidateSet('powerpoint-undo-grouping','powerpoint-text-safety','powerpoint-basic-functions','powerpoint-native-styles','all','word-exact-mathtype-omml-core-clone','word-mathtype-manual-coexistence')][string]$Mode='powerpoint-text-safety',
    [int]$ExpectedExitCode=0
)
$ErrorActionPreference='Stop'
$candidate=(Resolve-Path -LiteralPath $CandidateDirectory).Path
$artifact=[IO.Path]::GetFullPath($ArtifactDirectory)
if(Test-Path -LiteralPath $artifact){throw 'A fresh evidence directory is required.'}
if(@(Get-Process POWERPNT -ErrorAction SilentlyContinue).Count -ne 0){throw 'Existing PowerPoint is not owned; no test started.'}
function Read-Environment {
    $keys=@(foreach($appName in @('Word','PowerPoint')){
        foreach($hive in @('HKCU','HKLM')){
            $path=$hive+':\Software\Microsoft\Office\'+$appName+'\Addins\VisualTeX.'+$appName+'Vsto'
            $exists=Test-Path -LiteralPath $path
            $values=if($exists){Get-ItemProperty -LiteralPath $path | Select-Object LoadBehavior,Manifest,FriendlyName,Description}else{$null}
            [pscustomobject]@{Path=$path;Exists=$exists;Values=$values}
        }
    })
    $files=@(foreach($name in @('Word','PowerPoint')){
        $path=Join-Path $env:ProgramFiles ('VisualTeX\WindowsOffice\VSTO\VisualTeX.'+$name+'Vsto.dll')
        [pscustomobject]@{Path=$path;Hash=$(if(Test-Path -LiteralPath $path){(Get-FileHash -LiteralPath $path).Hash}else{$null})}
    })
    $processes=@(Get-Process WINWORD,POWERPNT -ErrorAction SilentlyContinue | ForEach-Object {
        [pscustomobject]@{Id=$_.Id;Name=$_.ProcessName;StartUtc=$_.StartTime.ToUniversalTime().ToString('o')}
    } | Sort-Object Id)
    [pscustomobject]@{User=[Security.Principal.WindowsIdentity]::GetCurrent().Name;Keys=$keys;Files=$files;Processes=$processes}
}
$before=Read-Environment
[void][IO.Directory]::CreateDirectory($artifact)
$before | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $artifact 'environment-before.json') -Encoding UTF8
& (Join-Path $candidate 'VisualTeX.VstoFlowAcceptance.exe') $artifact ('--mode='+$Mode)
$resultCode=$LASTEXITCODE
# Quit is asynchronous; wait only for observed process identity equality. Never
# kill, attach to, close or save any pre-existing or unknown Office process.
$deadline=[DateTime]::UtcNow.AddSeconds(8)
do {
    $after=Read-Environment
    $processesEqual=($before.Processes | ConvertTo-Json -Compress) -ceq ($after.Processes | ConvertTo-Json -Compress)
    if($processesEqual){break}
    Start-Sleep -Milliseconds 250
} while([DateTime]::UtcNow -lt $deadline)
$after | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $artifact 'environment-after.json') -Encoding UTF8
$summary=[ordered]@{
    ExitCode=$resultCode
    ExpectedExitCode=$ExpectedExitCode
    Mode=$Mode
    Candidate=$candidate
    WordDllSha256=(Get-FileHash -LiteralPath (Join-Path $candidate 'VisualTeX.WordVsto.dll')).Hash
    PowerPointDllSha256=(Get-FileHash -LiteralPath (Join-Path $candidate 'VisualTeX.PowerPointVsto.dll')).Hash
    KeysUnchanged=(($before.Keys | ConvertTo-Json -Depth 6 -Compress) -ceq ($after.Keys | ConvertTo-Json -Depth 6 -Compress))
    InstalledFilesUnchanged=(($before.Files | ConvertTo-Json -Compress) -ceq ($after.Files | ConvertTo-Json -Compress))
    ProcessesUnchanged=$processesEqual
}
$summary | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $artifact 'environment-checks.json') -Encoding UTF8
$summary | ConvertTo-Json
if($resultCode -ne $ExpectedExitCode -or !$summary.KeysUnchanged -or !$summary.InstalledFilesUnchanged -or !$processesEqual){exit 1}
exit 0
