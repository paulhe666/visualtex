param([int]$WordProcessId=40948,[string]$ExpectedDocument='math-before.docx',[ValidateSet('Inventory','CenterPositionOnly')][string]$Action='Inventory')
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$root=Join-Path $PSScriptRoot 'evidence\real-mathtype-20260926'
$a=Connect-RunningWord $WordProcessId;$d=$a.ActiveDocument
if($d.Name -ne $ExpectedDocument -or !$d.FullName.StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'Only owned evidence copies allowed'}
function Inventory($doc){
 $items=@()
 for($i=1;$i -le $doc.Paragraphs.Count;$i++){
  $p=$doc.Paragraphs.Item($i);$r=$p.Range;$fm=$p.Format
  $shapes=@()
  foreach($s in $r.InlineShapes){
   if([int]$s.Type -ne 1){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($s);continue}
   $sr=$s.Range;$fields=$sr.Fields;$f=$null;$fr=$null
   if($fields.Count -eq 1){$f=$fields.Item(1);$fr=$f.Result}
   $shapes+=@{start=$sr.Start;end=$sr.End;width=$s.Width;height=$s.Height;position=if($fr){$fr.Font.Position}else{$sr.Font.Position};font=if($fr){$fr.Font.Name}else{$sr.Font.Name};size=if($fr){$fr.Font.Size}else{$sr.Font.Size};progId=$s.OLEFormat.ProgID}
   foreach($v in @($fr,$f,$fields,$sr,$s)){if($v){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($v)}}
  }
  $items+=@{ordinal=$i;text=$r.Text;start=$r.Start;end=$r.End;alignment=[int]$fm.BaseLineAlignment;lineRule=[int]$fm.LineSpacingRule;line=$fm.LineSpacing;shapes=$shapes}
  foreach($v in @($fm,$r,$p)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($v)}
 }
 return $items
}
$tag=$ExpectedDocument.Replace('.docx','')
$items=Inventory $d
[IO.File]::WriteAllText((Join-Path $root ($tag+'-alignment-inventory.json')),($items|ConvertTo-Json -Depth 8),[Text.UTF8Encoding]::new($false))
if($Action -eq 'Inventory'){
 $d.Save()
 Write-Output ('Saved owned evidence only: '+$d.FullName)
}else{
 $record=$a.UndoRecord;$record.StartCustomRecord('Diagnostic center-alignment coordinate probe')
 try{
  foreach($row in $items){
   if($row.alignment -ne 1){continue}
   foreach($shape in $row.shapes){
    if($shape.progId -notmatch 'Equation'){continue}
    $range=$d.Range($shape.start,$shape.end);$fields=$range.Fields;$field=$fields.Item(1);$result=$field.Result
    $result.Font.Position=0
    foreach($v in @($result,$field,$fields,$range)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($v)}
   }
  }
 }finally{$record.EndCustomRecord();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($record)}
 try{
  & (Join-Path $PSScriptRoot 'inspect-real-doc-page.ps1') -WordProcessId $WordProcessId -ExpectedDocument $ExpectedDocument -Needle ([string][char]0x82E5+[char]0x5173+[char]0x4E8E) -Label 'math-center-position-only' -Zoom 130
 }finally{
  if(!$d.Undo(1)){throw 'Diagnostic undo failed'}
  Write-Output 'Diagnostic position edits undone; original faulty evidence retained.'
 }
}
foreach($v in @($d,$a)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($v)}
