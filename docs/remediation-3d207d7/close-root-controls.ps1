$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$evidence=Join-Path $PSScriptRoot 'evidence'
# Only the two completed radical A/B controls. No process-wide close or Quit.
$controls=@(
 @{stage='stage04al';record='stage04al-cases-roundtrip-omml-__8.json'},
 @{stage='stage04ah';record='stage04ak-cases-ab-__20.json'}
)
$rows=@()
foreach($control in $controls){
 $processRecord=Get-Content (Join-Path $evidence ($control.stage+'-word-process.json')) -Raw -Encoding UTF8|ConvertFrom-Json
 $expected=Get-Content (Join-Path $evidence $control.record) -Raw -Encoding UTF8|ConvertFrom-Json
 $process=Get-Process -Id $processRecord.pid
 if($processRecord.pid -eq 27664 -or $process.ProcessName -ne 'WINWORD' -or [Math]::Abs(($process.StartTime.ToUniversalTime()-[DateTimeOffset]::Parse($processRecord.started).UtcDateTime).TotalSeconds) -gt 10){throw 'Recorded test process identity changed.'}
 $word=Connect-RunningWord $processRecord.pid
 try{
  $document=$word.Documents.Item($expected.name)
  try{
   if($document.Name -ne $expected.name -or $document.FullName -ne $expected.fullName -or $document.Content.End -ne $expected.end -or $document.OMaths.Count -ne $expected.maths.Count -or $document.InlineShapes.Count -ne $expected.shapes.Count -or $document.Fields.Count -ne $expected.fields.Count -or $document.Tables.Count -ne $expected.tables.Count -or $document.Bookmarks.Count -ne $expected.bookmarks.Count){throw 'Completed control differs from its recorded evidence; preserving it.'}
   foreach($entry in $expected.bookmarks){
    if(!$document.Bookmarks.Exists($entry.name)){throw 'A recorded control bookmark changed; preserving document.'}
    $range=$document.Bookmarks.Item($entry.name).Range
    if($range.Start -ne $entry.start -or $range.End -ne $entry.end){throw 'A recorded control bookmark moved; preserving document.'}
   }
   $document.Close(0)
   $rows+=@{action='closed-without-save';pid=$processRecord.pid;document=$expected.name;source=$control.record;time=[DateTime]::UtcNow.ToString('o')}
  }finally{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($document)}
 }finally{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)}
}
$rows|ConvertTo-Json -Depth 4|Set-Content (Join-Path $evidence 'stage04ao-closed-root-controls.json') -Encoding UTF8
$rows|ConvertTo-Json -Depth 4
