param([ValidateSet('physics','math-exam','math-exam-a3')][string]$Document='physics',[Parameter(Mandatory=$true)][string]$Label)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$root=Join-Path $PSScriptRoot 'evidence\real-mathtype-20260926'
if($Label -notmatch '^[a-z0-9-]+$'){throw 'Invalid evidence label'}
$source=Join-Path $root ($Document+'-disk-original.docx')
$copy=Join-Path $root ($Label+'.docx')
$recordPath=Join-Path $root ($Label+'-process.json')
if((Test-Path $copy) -or (Test-Path $recordPath)){throw 'Evidence already exists; do not start duplicate Word'}
[IO.File]::Copy($source,$copy,$false)
$env:VISUALTEX_VSTO_ACCEPTANCE=$null
$env:VISUALTEX_VSTO_REDRAW_ACCEPTANCE_LOG=Join-Path $root ($Label+'-redraw.log')
$env:VISUALTEX_WORD_HOOK_TRACE_PATH=Join-Path $root ($Label+'-hook.log')
$env:VISUALTEX_VSTO_BULK_ACCEPTANCE_LOG=Join-Path $root ($Label+'-bulk.log')
$p=Start-Process 'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE' -ArgumentList @('/x',('"'+$copy+'"')) -PassThru
$record=[ordered]@{pid=$p.Id;label=$Label;document=([IO.Path]::GetFileName($copy));path=$copy;source=$source;started=[DateTimeOffset]::Now.ToString('o')}
[IO.File]::WriteAllText($recordPath,($record|ConvertTo-Json),[Text.UTF8Encoding]::new($false))
$deadline=[DateTime]::UtcNow.AddSeconds(20)
do{
 $app=$null;$doc=$null
 try{
  $app=Connect-RunningWord $p.Id
  $doc=$app.ActiveDocument
  if($doc.FullName -ne $copy){throw 'Unexpected owned document'}
  $record.hwnd=$doc.ActiveWindow.Hwnd
  $record.compatibility=$doc.CompatibilityMode
  $record.paragraphs=$doc.Paragraphs.Count
  $record.shapes=$doc.InlineShapes.Count
  $record.maths=$doc.OMaths.Count
  [IO.File]::WriteAllText($recordPath,($record|ConvertTo-Json),[Text.UTF8Encoding]::new($false))
  $record|ConvertTo-Json
  return
 }catch {if([DateTime]::UtcNow -ge $deadline){throw};Start-Sleep -Milliseconds 500}
 finally{if($doc){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc)};if($app){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)}}
}while([DateTime]::UtcNow -lt $deadline)
throw 'Owned Word not ready; inspect recorded PID rather than relaunch.'
