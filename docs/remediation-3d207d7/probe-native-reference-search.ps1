param([int]$WordProcessId,[string]$DocumentName,[string]$Label='reference-search')
$ErrorActionPreference='Stop';[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
function Release-Com($o){if($null -ne $o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($o)}}
$word=Connect-RunningWord $WordProcessId;$doc=$word.ActiveDocument
if($doc.Name -ne $DocumentName){throw 'Explicit document is not active.'}
$targets=@();$fields=$doc.Fields
for($i=1;$i -le $fields.Count;$i++){
 $f=$fields.Item($i);$c=$null
 try{if($f.Type -notin @(50,88)){continue};$c=$f.Code;$targets+=@{type=$f.Type;start=$c.Start;end=$c.End;text=$c.Text}}
 finally{Release-Com $c;Release-Com $f}
}
Release-Com $fields
$report=@{document=$doc.Name;fields=$targets;results=@()}
$token=($targets|Select-Object -First 1).text
$match=[regex]::Match($token,'(?i)\bREF\s+(\S+)')
if(!$match.Success){throw 'A real external REF is required.'}
foreach($text in @(('^d REF '+$match.Groups[1].Value), '^d GOTOBUTTON', '^d', $match.Groups[1].Value)){
 $scope=$doc.Content.Duplicate;$find=$scope.Find
 try{
  $find.ClearFormatting();$find.Text=$text;$find.Forward=$true;$find.Wrap=0;$find.Format=$false;$find.MatchWildcards=$false;$find.MatchCase=$false;$find.MatchWholeWord=$false
  $watch=[Diagnostics.Stopwatch]::StartNew();$found=$find.Execute();$elapsed=$watch.Elapsed.TotalMilliseconds
  $report.results+=@{search=$text;found=$found;start=$scope.Start;end=$scope.End;text=([string]$scope.Text).Substring(0,[Math]::Min(250,([string]$scope.Text).Length));elapsedMs=$elapsed}
 }finally{Release-Com $find;Release-Com $scope}
}
[IO.File]::WriteAllText((Join-Path $PSScriptRoot "evidence\$Label.json"),($report|ConvertTo-Json -Depth 5),[Text.UTF8Encoding]::new($false))
$report|ConvertTo-Json -Depth 5
Release-Com $doc;Release-Com $word
