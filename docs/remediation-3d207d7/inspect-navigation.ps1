param([string]$Label='navigation-state')
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$target=Get-Content (Join-Path $PSScriptRoot 'evidence\target-word.json') -Raw|ConvertFrom-Json
$word=Connect-RunningWord $target.pid
$document=$word.ActiveDocument
$selection=$word.Selection
$range=$selection.Range
$fields=@()
foreach($field in $document.Fields){
 if([int]$field.Type -eq 51 -and $field.Code.Text -match '^\s*MACROBUTTON MTPlaceRef\b'){
  $fields+=@{type=[int]$field.Type;codeStart=$field.Code.Start;codeEnd=$field.Code.End;resultStart=$field.Result.Start;resultEnd=$field.Result.End;code=$field.Code.Text;rangeText=$document.Range($field.Code.Start-1,$field.Result.End+1).Text}
 }
}
$selectionFields=@()
foreach($field in $selection.Fields){$selectionFields+=@{type=[int]$field.Type;codeStart=$field.Code.Start;codeEnd=$field.Code.End;resultEnd=$field.Result.End;code=$field.Code.Text}}
$bookmarkStates=@()
foreach($bookmark in $document.Bookmarks){$bookmarkStates+=@{name=$bookmark.Name;start=$bookmark.Range.Start;end=$bookmark.Range.End;text=$bookmark.Range.Text}}
$record=@{pid=$target.pid;document=$document.Name;end=$document.Content.End;selection=@{type=[int]$selection.Type;start=$selection.Start;end=$selection.End;text=$selection.Text;rangeStart=$range.Start;rangeEnd=$range.End;rangeText=$range.Text;fields=$selectionFields};numberHosts=$fields;bookmarks=$bookmarkStates}
$record|ConvertTo-Json -Depth 8|Set-Content (Join-Path $PSScriptRoot "evidence\$Label.json") -Encoding UTF8
$range.WordOpenXML|Set-Content (Join-Path $PSScriptRoot "evidence\$Label-selection.xml") -Encoding UTF8
$record|ConvertTo-Json -Depth 8
