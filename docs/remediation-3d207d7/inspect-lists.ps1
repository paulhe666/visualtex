param([string]$Label='list-evidence')
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$target=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'evidence\target-word.json') -Raw | ConvertFrom-Json
$word=Connect-RunningWord $target.pid
$document=$word.ActiveDocument
if($document.Name -ne $target.document){throw 'The active document is not the recorded test document.'}
$rows=@()
foreach($paragraph in $document.Paragraphs){
 $range=$paragraph.Range
 if($range.Information(12)){continue}
 $list=$range.ListFormat
 $format=$range.ParagraphFormat
 $rows+=@{start=$range.Start;end=$range.End;text=$range.Text;style=$range.Style.NameLocal;listType=[int]$list.ListType;listString=$list.ListString;listValue=$list.ListValue;level=$list.ListLevelNumber;leftIndent=$format.LeftIndent;firstLineIndent=$format.FirstLineIndent;rightIndent=$format.RightIndent;font=$range.Font.NameAscii;size=$range.Font.Size}
}
$result=@{pid=$target.pid;document=$document.Name;readOnly=$true;paragraphs=$rows}
$result | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath (Join-Path $PSScriptRoot "evidence\$Label.json") -Encoding UTF8
$rows | Select-Object -First 6 | ConvertTo-Json -Depth 5
