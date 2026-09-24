param([string]$Label='stage01',[int]$RestartEmptyProcessId=0,[switch]$TraceRecoveryXml,[string]$InstalledStage='')
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'word-connection.ps1')
if($RestartEmptyProcessId){
 $record=Get-Content -LiteralPath (Join-Path $PSScriptRoot "evidence\$Label-word-process.json") -Raw | ConvertFrom-Json
 if($RestartEmptyProcessId -ne $record.pid -or $RestartEmptyProcessId -eq 27664){throw 'Only this stage-owned process can be restarted.'}
 $word=Connect-RunningWord $RestartEmptyProcessId
 if($word.Documents.Count -ne 1 -or $word.ActiveDocument.Content.End -ne 1 -or !$word.ActiveDocument.Saved){throw 'Stage document is not an unchanged empty document; refusing to close it.'}
 $word.Quit(0)
 [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
if($env:VISUALTEX_VSTO_ACCEPTANCE -eq '1'){throw 'Manual-service acceptance mode is prohibited.'}
$env:VISUALTEX_WORD_HOOK_TRACE_PATH=Join-Path $PSScriptRoot "evidence\$Label-word-hook.log"
$env:VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG=Join-Path $PSScriptRoot "evidence\$Label-word-bulk.log"
$env:VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG=Join-Path $PSScriptRoot "evidence\$Label-word-redraw.log"
$env:VISUALTEX_NUMBERED_PERF_TRACE='1'
$env:VISUALTEX_NUMBERED_PERF_TRACE_PATH=Join-Path $PSScriptRoot "evidence\$Label-numbering-perf.log"
$env:VISUALTEX_VSTO_TRACE_FORMAT_PERF='1'
if($TraceRecoveryXml){$env:VISUALTEX_VSTO_TRACE_RECOVERY_XML='1'}else{$env:VISUALTEX_VSTO_TRACE_RECOVERY_XML=$null}
$stageWord=Start-Process -FilePath 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE' -ArgumentList '/x' -PassThru
if(!$InstalledStage){$InstalledStage=$Label}
$record=@{pid=$stageWord.Id;stage=$Label;started=[DateTime]::UtcNow.ToString('o');source="$InstalledStage-installed.json";protectedWordPid=27664}
$record|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $PSScriptRoot "evidence\$Label-word-process.json") -Encoding UTF8
$record|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $PSScriptRoot 'evidence\target-word.json') -Encoding UTF8
$deadline=[DateTime]::UtcNow.AddSeconds(25)
$ready=$false
do {
 $stageWord.Refresh()
 if($stageWord.HasExited){throw 'The new Word process exited before creating its window.'}
 $readyWord=$null
 try {
  $readyWord=Connect-RunningWord $stageWord.Id
  $ready=$readyWord.Documents.Count -gt 0
 } catch {
  if($_.Exception.Message -notlike 'No native Word document window*' -and $_.Exception.HResult -ne -2147418111){throw}
 } finally {if($readyWord){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($readyWord)}}
 if($ready){break}
 Start-Sleep -Milliseconds 400
} while([DateTime]::UtcNow -lt $deadline)
if(!$ready){throw 'The recorded Word process has not exposed its native document; inspect without starting another.'}
$stageWord.Id
