param([string]$Label='number-paragraph')
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$target=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'evidence\target-word.json') -Raw -Encoding UTF8|ConvertFrom-Json
$word=Connect-RunningWord $target.pid
$doc=$word.ActiveDocument
if($target.pid -eq 27664 -or $doc.Name -ne $target.document){throw 'Only the recorded owned test document may be inspected.'}
function ReadRange($range){
 $f=$range.Font
 $text=[string]$range.Text
 return @{start=$range.Start;end=$range.End;text=$text;codes=@($text.ToCharArray()|ForEach-Object{[int]$_});ascii=$f.NameAscii;other=$f.NameOther;name=$f.Name;size=$f.Size;style=$range.Style.NameLocal}
}
$rows=@()
foreach($field in $doc.Fields){
 if($field.Type -ne 12 -or $field.Code.Text -notmatch '\bVisualTeXEquation\b'){continue}
 $range=$field.Result;$p=$range.Paragraphs.Item(1).Range
 $characters=@();for($position=$p.End-3;$position -le $p.End;$position++){
  if($position -ge $doc.Content.Start -and $position -lt $doc.Content.End){$characters+=ReadRange ($doc.Range($position,$position+1))}
 }
 $rows+=@{code=$field.Code.Text;result=(ReadRange $range);paragraph=(ReadRange $p);endingCharacters=$characters}
}
$rows|ConvertTo-Json -Depth 7|Set-Content -LiteralPath (Join-Path $PSScriptRoot "evidence\$Label.json") -Encoding UTF8
$rows|ForEach-Object {[pscustomobject]@{result=$_.result.text;ascii=$_.result.ascii;paragraphEnd=$_.paragraph.end;ending=($_.endingCharacters|ConvertTo-Json -Depth 4 -Compress)}}|ConvertTo-Json -Depth 3
