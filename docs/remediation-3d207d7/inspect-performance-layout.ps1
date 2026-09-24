param([Parameter(Mandatory=$true)][int]$WordProcessId,[string]$DocumentName,[string]$Label='layout')
$ErrorActionPreference='Stop';[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
function Release-Com($v){if($v -and [Runtime.InteropServices.Marshal]::IsComObject($v)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($v)}}
$word=Connect-RunningWord $WordProcessId;$d=$word.ActiveDocument
if($DocumentName -and $d.Name -ne $DocumentName){throw 'Expected active document is unavailable; no activation or mutation was attempted.'}
$items=@();$tables=@();$maths=$d.OMaths;$shapes=$d.InlineShapes;$ts=$d.Tables
try {
 for($i=1;$i -le $maths.Count;$i++){
  $m=$maths.Item($i);$r=$m.Range;$f=$r.ParagraphFormat;$font=$r.Font
  try {$items+=@{index=$i;start=$r.Start;end=$r.End;text=$r.Text;type=[int]$m.Type;inTable=[bool]$r.Information(12);font=$font.Name;size=$font.Size;position=$font.Position;lineRule=[int]$f.LineSpacingRule;lineSpacing=$f.LineSpacing;grid=$f.DisableLineHeightGrid;spaceBefore=$f.SpaceBefore;spaceAfter=$f.SpaceAfter;alignment=[int]$f.Alignment}}
  finally{Release-Com $font;Release-Com $f;Release-Com $r;Release-Com $m}
 }
 for($i=1;$i -le $ts.Count;$i++){
  $t=$ts.Item($i);$rs=$t.Rows;$r=$t.Range
  try {for($j=1;$j -le $rs.Count;$j++){$row=$rs.Item($j);try{$tables+=@{index=$i;row=$j;columns=$t.Columns.Count;start=$r.Start;end=$r.End;height=$row.Height;rule=[int]$row.HeightRule;topPadding=$t.TopPadding;bottomPadding=$t.BottomPadding}}finally{Release-Com $row}}}
  finally{Release-Com $r;Release-Com $rs;Release-Com $t}
 }
 $s=$word.Selection;$content=$d.Content;$formats=@()
 try {$report=@{pid=$WordProcessId;document=$d.Name;omml=$maths.Count;inlineShapes=$shapes.Count;selection=@{start=$s.Start;end=$s.End;text=$s.Text};mathFont=$d.OMathFontName;maths=$items;tables=$tables;documentEnd=$content.End}}
 finally{Release-Com $content;Release-Com $s}
 $json=$report|ConvertTo-Json -Depth 8
 [IO.File]::WriteAllText((Join-Path $PSScriptRoot "evidence\$Label-layout.json"),$json,[Text.UTF8Encoding]::new($false));$json
}finally{Release-Com $ts;Release-Com $shapes;Release-Com $maths;Release-Com $d;Release-Com $word}
