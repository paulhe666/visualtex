param([string]$Label='reference-format',[int]$WordProcessId=0)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'word-connection.ps1')
if(!$WordProcessId){$WordProcessId=(Get-Content (Join-Path $PSScriptRoot 'evidence\target-word.json') -Raw | ConvertFrom-Json).pid}
$word=Connect-RunningWord $WordProcessId
$document=$word.ActiveDocument
function FormatOf($range) {
 $font=$range.Font
 $style=$range.Style
 $item=[ordered]@{start=$range.Start;end=$range.End;text=$range.Text;style=$style.NameLocal}
 foreach($property in @('Name','NameAscii','NameOther','NameFarEast','NameBi','Size','SizeBi','Bold','BoldBi','Italic','ItalicBi','Underline','Color','Hidden','Position','Spacing','Scaling','Subscript','Superscript')){$item[$property]=$font.$property}
 return $item
}
$paragraphs=@()
foreach($p in $document.Paragraphs){
 $r=$p.Range
 if($r.Information(12)){continue}
 $paragraphs+=@{range=(FormatOf $r);first=(FormatOf ($document.Range($r.Start,[Math]::Min($r.Start+1,$r.End))));mark=(FormatOf ($document.Range($r.End-1,$r.End)))}
}
$fields=@()
foreach($f in $document.Fields){
 if([int]$f.Type -notin @(3,12,50,51)){continue}
 $r=$f.Result;$c=$f.Code
 $fields+=@{type=[int]$f.Type;code=(FormatOf $c);result=(FormatOf $r);codeFirst=(FormatOf ($document.Range($c.Start,[Math]::Min($c.Start+2,$c.End))));preceding=(FormatOf ($document.Range([Math]::Max(0,$c.Start-2),[Math]::Max(0,$c.Start-1))))}
}
$bookmarks=@()
foreach($b in $document.Bookmarks){if($b.Name -match '^(VTEq|ZEqnNum)'){$bookmarks+=@{name=$b.Name;format=(FormatOf $b.Range)}}}
$record=@{document=$document.Name;pid=$WordProcessId;selection=(FormatOf $word.Selection.Range);paragraphs=$paragraphs;fields=$fields;bookmarks=$bookmarks}
$json=$record | ConvertTo-Json -Depth 12
[IO.File]::WriteAllText((Join-Path $PSScriptRoot "evidence\$Label.json"),$json,[Text.UTF8Encoding]::new($false))
$json
