param([int]$WordProcessId,[string]$DocumentName,[string]$Label='field-navigation')
$ErrorActionPreference='Stop';[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
function Release-Com($o){if($null -ne $o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($o)}}
$word=Connect-RunningWord $WordProcessId;$doc=$word.ActiveDocument
if($doc.Name -ne $DocumentName){throw 'Explicit document is not active.'}
$selectionBefore=@($word.Selection.Start,$word.Selection.End);$rows=@()
foreach($name in @('REF','GOTOBUTTON','MACROBUTTON','SEQ','DATE')){
 $cursor=0
 for($i=0;$i -lt 3;$i++){
  $scope=$doc.Range($cursor,$cursor);$found=$null;$fields=$null
  try{
   $watch=[Diagnostics.Stopwatch]::StartNew();$found=$scope.GoTo(7,1,1,$name);$elapsed=$watch.Elapsed.TotalMilliseconds
   $probe=$doc.Range($found.Start,[Math]::Min($doc.Content.End,$found.Start+1))
   $fields=$probe.Fields;$items=@()
   for($j=1;$j -le $fields.Count;$j++){$f=$fields.Item($j);$code=$f.Code;try{$items+=@{type=$f.Type;start=$code.Start;end=$code.End;code=$code.Text}}finally{Release-Com $code;Release-Com $f}}
   $rows+=@{name=$name;from=$cursor;start=$found.Start;end=$found.End;text=$probe.Text;fields=$items;elapsedMs=$elapsed}
   if($found.Start -lt $cursor){break};$cursor=if($items.Count){$f2=$fields.Item(1);$result=$f2.Result;try{$result.End+1}finally{Release-Com $result;Release-Com $f2}}else{$found.End+1}
   Release-Com $probe;$probe=$null
  }catch{$rows+=@{name=$name;from=$cursor;error=$_.Exception.Message;hresult=$_.Exception.HResult};break}
  finally{Release-Com $fields;Release-Com $found;Release-Com $scope}
 }
}
$after=@($word.Selection.Start,$word.Selection.End)
if(($selectionBefore -join ':') -ne ($after -join ':')){throw 'Read-only range navigation moved the actual selection.'}
$report=@{document=$doc.Name;rows=$rows;selectionUnchanged=$true}
[IO.File]::WriteAllText((Join-Path $PSScriptRoot "evidence\$Label.json"),($report|ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($false))
$report|ConvertTo-Json -Depth 6
Release-Com $doc;Release-Com $word
