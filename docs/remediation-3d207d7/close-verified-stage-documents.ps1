param([string]$Label='stage04ag-close-verified',[int[]]$OnlyProcessIds=@())
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$evidence=Join-Path $PSScriptRoot 'evidence'
$creation=Get-Content (Join-Path $evidence 'stage04ag-owned-document-creation-records.json') -Raw -Encoding UTF8 | ConvertFrom-Json
# Preserve original process, current investigation, and unresolved historical
# document-loss/bookmark-recovery specimens. Never terminate a process forcibly.
$preserve=@(27664,60152,55728,100440,101980,82668,81452)
$log=Join-Path $evidence ($Label+'.ndjson')
function Record($item){$item.time=[DateTime]::UtcNow.ToString('o');[IO.File]::AppendAllText($log,($item|ConvertTo-Json -Compress -Depth 6)+[Environment]::NewLine,[Text.UTF8Encoding]::new($false));$item|ConvertTo-Json -Compress -Depth 6}
foreach($file in Get-ChildItem -LiteralPath $evidence -Filter '*-word-process.json'){
 $stage=Get-Content $file.FullName -Raw | ConvertFrom-Json
 $stagePid=[int]$stage.pid
 if($OnlyProcessIds.Count -gt 0 -and $OnlyProcessIds -notcontains $stagePid){continue}
 if($preserve -contains $stagePid){continue}
 $process=Get-Process -Id $stagePid -ErrorAction SilentlyContinue
 if(!$process -or $process.ProcessName -ne 'WINWORD'){continue}
 $recordedStart=[DateTimeOffset]::Parse($stage.started).UtcDateTime
 if([Math]::Abs(($process.StartTime.ToUniversalTime()-$recordedStart).TotalSeconds) -gt 10){Record @{action='preserved';pid=$stagePid;reason='process-start-mismatch'};continue}
 $knownNumbers=@($creation | Where-Object {$_.pid -eq $stagePid -and $_.stdout -match '^NEW_DOCUMENT\|[^|]*?([0-9]+)\|'} | ForEach-Object {if($_.stdout -match '^NEW_DOCUMENT\|[^|]*?([0-9]+)\|'){$Matches[1]}})
 $word=$null
 try{
  $word=Connect-RunningWord $stagePid
  for($index=$word.Documents.Count;$index -ge 1;$index--){
   $document=$word.Documents.Item($index)
   try{
    $name=[string]$document.Name;$full=[string]$document.FullName
    $number=if($name -match '^[^0-9]+([0-9]+)$'){$Matches[1]}else{''}
    $ownedUnsaved=$full -eq $name -and $knownNumbers -contains $number
    $ownedEmpty=$full -eq $name -and $document.Saved -and $document.Content.End -eq 1
    $ownedReopened=$false
    if($stagePid -eq 88240){
     $reopenedRecord=Get-Content (Join-Path $evidence 'stage04l-identical-save-reopen.json') -Raw | ConvertFrom-Json
     $ownedReopened=$reopenedRecord.pid -eq $stagePid -and $full -eq $reopenedRecord.path
    }
    if(!$ownedUnsaved -and !$ownedEmpty -and !$ownedReopened){Record @{action='preserved';pid=$stagePid;document=$name;reason='not-recorded-test-or-empty-starter'};continue}
    Record @{action='closing-without-save';pid=$stagePid;stage=$stage.stage;document=$name;end=$document.Content.End;maths=$document.OMaths.Count;oles=$document.InlineShapes.Count;fields=$document.Fields.Count;emptyStarter=$ownedEmpty}
    $document.Close(0)
    Record @{action='closed-without-save';pid=$stagePid;document=$name}
   }finally{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($document)}
  }
  if($word.Documents.Count -eq 0){$word.Quit(0);Record @{action='quit-empty-owned-word';pid=$stagePid}}
 }catch{Record @{action='preserved-after-error';pid=$stagePid;error=$_.Exception.ToString()}}
 finally{if($word){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)}}
}
