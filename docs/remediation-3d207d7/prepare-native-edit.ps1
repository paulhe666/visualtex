param([int]$Index=1,[string]$From='',[string]$To='x',[string]$Label='native-user-edit')
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$target=Get-Content (Join-Path $PSScriptRoot 'evidence\target-word.json') -Raw | ConvertFrom-Json
if(!$target.pid -or $target.pid -eq 27664){throw 'Only a stage-owned test process may receive input.'}
$word=Connect-RunningWord $target.pid
$document=$word.ActiveDocument
if(!$From){$From=[char]::ConvertFromUtf32(0x1D44E)}
$math=$document.OMaths.Item($Index)
$range=$math.Range
$matches=@()
for($position=$range.Start;$position -le $range.End-$From.Length;$position++){
 $probe=$document.Range($position,$position+$From.Length)
 try{if($probe.Text -ceq $From){$matches+=@{start=$probe.Start;end=$probe.End}}}
 finally{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($probe)}
}
if($matches.Count -ne 1){throw "Expected one exact user-edit character range, found $($matches.Count)."}
$before=@{maths=$document.OMaths.Count;shapes=$document.InlineShapes.Count;fields=$document.Fields.Count;text=$range.Text}
$document.Activate()
$selection=$document.Range($matches[0].start,$matches[0].end)
$selection.Select()
$word.Selection.TypeText($To)
$record=@{action='Select exact native equation text and Selection.TypeText';document=$document.Name;pid=$target.pid;from=$From;to=$To;selection=$matches[0];before=$before;after=@{maths=$document.OMaths.Count;shapes=$document.InlineShapes.Count;fields=$document.Fields.Count;text=$document.OMaths.Item($Index).Range.Text}}
$record | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $PSScriptRoot "evidence\$Label.json") -Encoding UTF8
$record | ConvertTo-Json -Depth 6
