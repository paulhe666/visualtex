param([string]$Label='stage04ai',[int]$ExpectedProcessId=76424)
$ErrorActionPreference='Stop'
if($Label -notmatch '^[a-zA-Z0-9-]+$'){throw 'Invalid label.'}
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$source=Join-Path $root 'apps\windows\src-tauri\target\release\visualtex.exe'
$installed=Join-Path $env:LOCALAPPDATA 'VisualTeX\visualtex.exe'
$registered=(Get-ItemProperty 'HKCU:\Software\VisualTeX\OfficeIntegration').ExecutablePath
if($registered -ne $installed){throw 'Current companion registration differs.'}
$current=Get-Process -Id $ExpectedProcessId
if($current.ProcessName -ne 'visualtex' -or $current.Path -ne $installed){throw 'Expected installed idle companion process differs.'}
$visible=& (Join-Path $PSScriptRoot 'word-ui.ps1') -Action windows | ConvertFrom-Json
if(@($visible | Where-Object {$_.pid -eq $ExpectedProcessId -and $_.title -ne 'com.visualtex.studio-siw'}).Count){throw 'A VisualTeX editor/main window is open; do not restart it.'}
$latestAction=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'evidence\ui-actions.ndjson') -Encoding UTF8 -Tail 40 | ForEach-Object {ConvertFrom-Json $_} | Where-Object {$_.action -eq 'observed-editor-release'} | Select-Object -Last 1
if(!$latestAction -or $latestAction.status -notin @('completed','cancelled')){throw 'The latest real editor release was not terminal.'}
$latestSession=Get-Content -LiteralPath (Join-Path $env:APPDATA ('com.visualtex.studio\office\sessions\'+$latestAction.id+'\session.json')) -Raw | ConvertFrom-Json
if($latestSession.status -notin @('completed','cancelled')){throw 'Real session state is not terminal.'}
$backup=Join-Path $PSScriptRoot ('evidence\'+$Label+'-main-before.exe')
if(Test-Path -LiteralPath $backup){throw 'A main candidate backup already exists; inspect before retrying.'}
$beforeHash=(Get-FileHash -LiteralPath $installed -Algorithm SHA256).Hash
$sourceHash=(Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$wordsBefore=@(Get-Process WINWORD | Select-Object Id,StartTime)
$oleBefore=@(Get-Process VisualTeX.FormulaOleServer | Select-Object Id,StartTime,Path)
Copy-Item -LiteralPath $installed -Destination $backup
# Only the checked idle VisualTeX desktop companion is restarted. No Word,
# PowerPoint, MathType or FormulaOleServer process is stopped.
Stop-Process -Id $ExpectedProcessId
Wait-Process -Id $ExpectedProcessId -Timeout 10 -ErrorAction SilentlyContinue
if(Get-Process -Id $ExpectedProcessId -ErrorAction SilentlyContinue){throw 'The old companion did not exit.'}
Copy-Item -LiteralPath $source -Destination $installed -Force
if((Get-FileHash -LiteralPath $installed -Algorithm SHA256).Hash -ne $sourceHash){throw 'Installed main hash mismatch.'}
$started=Start-Process -FilePath $installed -ArgumentList '--office-background' -WindowStyle Hidden -PassThru
$record=@{time=[DateTime]::UtcNow.ToString('o');label=$Label;previousPid=$ExpectedProcessId;pid=$started.Id;installed=$installed;beforeSha256=$beforeHash;afterSha256=$sourceHash;backup=$backup;lastCompletedEditor=$latestAction.id;wordProcessesBefore=$wordsBefore;oleProcessesBefore=$oleBefore;nsisInstalled=$false}
$record | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $PSScriptRoot ('evidence\'+$Label+'-main-installed.json')) -Encoding UTF8
Write-Output ('MAIN_CANDIDATE_INSTALLED|'+$started.Id+'|'+$sourceHash)
