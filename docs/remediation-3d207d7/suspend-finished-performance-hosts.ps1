param([string[]]$Stages=@('perf-baseline','perf-profile','perf-a','perf-candidate-a','perf-b'))
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$rows=@()
foreach($stage in $Stages){
 $file=Join-Path $PSScriptRoot "evidence\$stage-word-process.json"
 if(!(Test-Path -LiteralPath $file)){continue}
 $r=Get-Content -LiteralPath $file -Raw -Encoding UTF8|ConvertFrom-Json
 if($r.pid -in @(27664,107976) -or !$r.started){throw 'Original Word process is protected.'}
 $process=Get-Process -Id $r.pid -ErrorAction SilentlyContinue
 if(!$process){continue}
 if($process.ProcessName -ne 'WINWORD' -or [Math]::Abs(($process.StartTime.ToUniversalTime()-[DateTimeOffset]::Parse($r.started).UtcDateTime).TotalSeconds)-gt 10){throw 'Recorded test process identity changed.'}
 $word=Connect-RunningWord $r.pid
 try {
  $before=@();for($i=1;$i -le $word.Documents.Count;$i++){
   $d=$word.Documents.Item($i)
   try{$before+=@{name=$d.Name;end=$d.Content.End;maths=$d.OMaths.Count;oles=$d.InlineShapes.Count;saved=$d.Saved}}
   finally{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($d)}
  }
  $active=@(Get-ChildItem (Join-Path $env:APPDATA 'com.visualtex.studio\office\sessions') -Directory | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 100 | ForEach-Object {
   $p=Join-Path $_.FullName 'session.json'
   if(Test-Path -LiteralPath $p){try{$s=Get-Content $p -Raw -Encoding UTF8|ConvertFrom-Json;if($s.status -in @('editing','committing') -and $s.sourceDocumentId -in $before.name){$s.id}}catch{}}
  })
  if($active.Count){Write-Output "PRESERVED_BUSY_HOST|$stage|$($r.pid)|$($active -join ',')";continue}
  $addin=$word.COMAddIns.Item('VisualTeX.WordVsto')
  try{if($addin.Connect){$addin.Connect=$false}}finally{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($addin)}
  foreach($expected in $before){
   $d=$word.Documents.Item($expected.name)
   try{if($d.Content.End -ne $expected.end -or $d.OMaths.Count -ne $expected.maths -or $d.InlineShapes.Count -ne $expected.oles -or $d.Saved -ne $expected.saved){throw 'Test document changed while detaching its idle add-in.'}}
   finally{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($d)}
  }
  $rows+=@{stage=$stage;pid=$r.pid;documents=$before;action='detached idle test add-in; no Word document closed, saved or edited'}
  Write-Output "DETACHED_TEST_ADDIN|$stage|$($r.pid)"
 }finally{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)}
}
$path=Join-Path $PSScriptRoot ('evidence\idle-addin-detach-'+[DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss')+'.json')
[IO.File]::WriteAllText($path,($rows|ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($false))
