param([string]$Label='reference-rectangles')
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$target=Get-Content (Join-Path $PSScriptRoot 'evidence\target-word.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$word=Connect-RunningWord $target.pid
$doc=$word.ActiveDocument
if($doc.Name -ne $target.document -or $target.pid -eq 27664){throw 'Not the owned test document.'}
$field=@($doc.Fields | Where-Object {[int]$_.Type -eq 50})[0]
if(!$field){throw 'No actual GOTOBUTTON.'}
$nested=$field.Code.Fields.Item(1)
$ranges=@{outerResult=$field.Result;outerCode=$field.Code;complete=$doc.Range($field.Code.Start-1,$field.Result.End+1);nestedResult=$nested.Result}
$rows=@()
foreach($entry in $ranges.GetEnumerator()) {
 $range=$entry.Value;[int]$left=0;[int]$top=0;[int]$width=0;[int]$height=0
 $doc.ActiveWindow.GetPoint([ref]$left,[ref]$top,[ref]$width,[ref]$height,$range)
 $rows+=@{range=$entry.Key;start=$range.Start;end=$range.End;text=$range.Text;left=$left;top=$top;width=$width;height=$height}
}
$rows|ConvertTo-Json -Depth 4|Set-Content (Join-Path $PSScriptRoot "evidence\$Label.json") -Encoding UTF8
$rows|ConvertTo-Json -Depth 4
