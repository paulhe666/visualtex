param([string[]]$Stages=@('perf-baseline','perf-profile','perf-a','perf-candidate-a','perf-b'))
$ErrorActionPreference='Stop';[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$stamp=[DateTime]::UtcNow.ToString('yyyyMMdd-HHmmss');$records=@()
foreach($stage in $Stages){
 $file=Join-Path $PSScriptRoot "evidence\$stage-word-process.json"
 if(!(Test-Path $file)){continue}
 $r=Get-Content $file -Raw -Encoding UTF8|ConvertFrom-Json
 if($r.pid -in @(27664,107976) -or !$r.started){throw 'Original Word process is protected.'}
 $p=Get-Process -Id $r.pid -ErrorAction SilentlyContinue;if(!$p){continue}
 if($p.ProcessName -ne 'WINWORD' -or [Math]::Abs(($p.StartTime.ToUniversalTime()-[DateTimeOffset]::Parse($r.started).UtcDateTime).TotalSeconds)-gt 10){throw 'Recorded test process identity changed.'}
 $word=Connect-RunningWord $r.pid
 try {
  $docs=@();for($i=1;$i -le $word.Documents.Count;$i++){
   $d=$word.Documents.Item($i)
   try {
    $untitledPattern='^(Document|'+[char]0x6587+[char]0x6863+')\d+$'
    if($d.Path -or $d.Name -notmatch $untitledPattern){throw 'An unowned or saved document is present; preserving this Word process.'}
    $name=$d.Name;$xml=$d.Content.WordOpenXML
    $evidence=Join-Path $PSScriptRoot ("evidence\$stamp-preserved-$stage-"+($name -replace '[^0-9a-zA-Z]','_')+'.xml')
    [IO.File]::WriteAllText($evidence,$xml,[Text.UTF8Encoding]::new($false))
    $docs+=@{name=$name;end=$d.Content.End;maths=$d.OMaths.Count;oles=$d.InlineShapes.Count;xml=$evidence}
   }finally{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($d)}
  }
  foreach($before in $docs){
   $d=$word.Documents.Item($before.name)
   try{if($d.Content.End -ne $before.end -or $d.OMaths.Count -ne $before.maths -or $d.InlineShapes.Count -ne $before.oles){throw 'Owned test document changed before cleanup.'};$d.Close(0)}
   finally{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($d)}
  }
  try{if($word.Documents.Count -eq 0){$word.Quit(0)}}catch{if(Get-Process -Id $r.pid -ErrorAction SilentlyContinue){throw}}
  $records+=@{stage=$stage;pid=$r.pid;documents=$docs;action='closed only recorded disposable UI-test documents after preserving XML'}
  Write-Output "CLOSED_TEST_HOST|$stage|$($r.pid)"
 }finally{try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)}catch{}}
}
[IO.File]::WriteAllText((Join-Path $PSScriptRoot "evidence\$stamp-finished-test-cleanup.json"),($records|ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($false))
