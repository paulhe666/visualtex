param([int]$WordProcessId=0,[string]$Mode = 'inventory', [string]$DocumentName = [string]::Empty, [string]$DocumentNumber = '', [string]$Label = 'original', [switch]$OriginalOnly)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding
$root = Join-Path $PSScriptRoot 'evidence'
[void][IO.Directory]::CreateDirectory($root)
function Esc([string]$s) { return $s.Replace("`r", '<CR>').Replace("`n", '<LF>').Replace("`t", '<TAB>').Replace([string][char]7, '<CELL>').Replace([string][char]11, '<BR>') }
function Release($o) { if ($null -ne $o -and [Runtime.InteropServices.Marshal]::IsComObject($o)) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} } }
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$word = Connect-RunningWord $WordProcessId
$reports = @()
for ($i=1; $i -le $word.Documents.Count; $i++) {
  $d = $word.Documents.Item($i)
  if ($DocumentName -and $d.Name -ne $DocumentName) { Release $d; continue }
  if ($DocumentNumber -and $d.Name -notmatch ('^[^0-9]*'+[regex]::Escape($DocumentNumber)+'$')) { Release $d; continue }
  if ($OriginalOnly -and $d.Name -notmatch '^[^0-9]*(1|22|26|28|29|33|35|36)$') { Release $d; continue }
  $name = $d.Name; $key = $Label+'-'+($name -replace '[^0-9a-zA-Z_-]', '_')
  $xml = $d.Content.WordOpenXML
  [IO.File]::WriteAllText((Join-Path $root ($key+'.xml')), $xml, [Text.UTF8Encoding]::new($false))
  $r = [ordered]@{name=$name; fullName=$d.FullName; saved=$d.Saved; end=$d.Content.End; paragraphs=$d.Paragraphs.Count; tables=@(); maths=@(); shapes=@(); controls=@(); bookmarks=@(); fields=@(); paras=@(); variables=@()}
  for ($j=1; $j -le $d.Tables.Count; $j++) {
    $t=$d.Tables.Item($j)
    $tr=[ordered]@{index=$j; start=$t.Range.Start; end=$t.Range.End; rows=$t.Rows.Count; cols=$t.Columns.Count; nesting=$t.NestingLevel; text=(Esc $t.Range.Text); cells=@()}
    for ($k=1;$k -le $t.Range.Cells.Count;$k++) { $c=$t.Range.Cells.Item($k); $tr.cells+=@{row=$c.RowIndex;col=$c.ColumnIndex;start=$c.Range.Start;end=$c.Range.End;math=$c.Range.OMaths.Count;text=(Esc $c.Range.Text)}; Release $c }
    $r.tables += $tr; Release $t
  }
  for ($j=1;$j -le $d.OMaths.Count;$j++) {
    $m=$d.OMaths.Item($j);$q=$m.Range
    $r.maths+=@{index=$j;start=$q.Start;end=$q.End;type=[int]$m.Type;font=$q.Font.Name;size=$q.Font.Size;text=(Esc $q.Text);tables=$q.Tables.Count;controlTags=@($q.ContentControls|ForEach-Object {$_.Tag})}; Release $q;Release $m
  }
  for($j=1;$j -le $d.InlineShapes.Count;$j++) {
    $s=$d.InlineShapes.Item($j);$prog='';try {$prog=$s.OLEFormat.ProgID}catch{}
    $r.shapes+=@{index=$j;start=$s.Range.Start;end=$s.Range.End;type=[int]$s.Type;prog=$prog;width=$s.Width;height=$s.Height;alternative=$s.AlternativeText};Release $s
  }
  for($j=1;$j -le $d.ContentControls.Count;$j++) { $c=$d.ContentControls.Item($j);$r.controls+=@{index=$j;id=$c.ID;tag=$c.Tag;title=$c.Title;start=$c.Range.Start;end=$c.Range.End;type=[int]$c.Type;text=(Esc $c.Range.Text)};Release $c }
  for($j=1;$j -le $d.Bookmarks.Count;$j++) { $b=$d.Bookmarks.Item($j);$r.bookmarks+=@{name=$b.Name;start=$b.Range.Start;end=$b.Range.End;text=(Esc $b.Range.Text)};Release $b }
  for($j=1;$j -le $d.Fields.Count;$j++) { $f=$d.Fields.Item($j);$r.fields+=@{index=$j;type=[int]$f.Type;start=$f.Code.Start;end=$f.Result.End;code=(Esc $f.Code.Text);result=(Esc $f.Result.Text);nested=$f.Code.Fields.Count};Release $f }
  for($j=1;$j -le $d.Paragraphs.Count;$j++) { $p=$d.Paragraphs.Item($j);$r.paras+=@{index=$j;start=$p.Range.Start;end=$p.Range.End;text=(Esc $p.Range.Text);font=$p.Range.Font.Name;size=$p.Range.Font.Size;hidden=$p.Range.Font.Hidden;alignment=[int]$p.Alignment;lineRule=[int]$p.LineSpacingRule;line=$p.LineSpacing;before=$p.SpaceBefore;after=$p.SpaceAfter;inTable=$p.Range.Information(12)};Release $p }
  for($j=1;$j -le $d.Variables.Count;$j++) { $v=$d.Variables.Item($j);$r.variables+=@{name=$v.Name;value=$v.Value};Release $v }
  [xml]$parsed=$xml
  $ns=[Xml.XmlNamespaceManager]::new($parsed.NameTable)
  $ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
  $ns.AddNamespace('m','http://schemas.openxmlformats.org/officeDocument/2006/math')
  $r.xmlFonts=@($parsed.SelectNodes('//m:oMath//w:rFonts',$ns)|ForEach-Object {$_.OuterXml}|Sort-Object -Unique)
  $r.xmlMathCount=$parsed.SelectNodes('//m:oMath',$ns).Count
  $r.customXml=@($parsed.SelectNodes('//*[local-name()="part" and contains(@*[local-name()="name"],"customXml/item") and not(contains(@*[local-name()="name"],"Props"))]')|ForEach-Object {$_.InnerXml})
  [IO.File]::WriteAllText((Join-Path $root ($key+'.json')), ($r|ConvertTo-Json -Depth 12), [Text.UTF8Encoding]::new($false))
  $reports += [ordered]@{name=$name;key=$key;saved=$r.saved;end=$r.end;paragraphs=$r.paragraphs;tables=$r.tables;maths=$r.maths;shapes=$r.shapes;controls=$r.controls;fields=$r.fields;bookmarks=$r.bookmarks;paras=$r.paras;xmlFonts=$r.xmlFonts;xmlMathCount=$r.xmlMathCount}
  Release $d
}
$installed=@()
foreach($p in (Get-Process WINWORD,visualtex -ErrorAction SilentlyContinue)) { $installed+=@{name=$p.ProcessName;id=$p.Id;path=$p.Path;start=$p.StartTime.ToString('o')} }
$summary=@{time=[DateTime]::Now.ToString('o');active=$word.ActiveDocument.Name;statusBar=[string]$word.StatusBar;version=$word.Version;build=$word.Build;processes=$installed;documents=$reports}
$dest=Join-Path $root ($Label+'-inventory.json')
[IO.File]::WriteAllText($dest, ($summary|ConvertTo-Json -Depth 15), [Text.UTF8Encoding]::new($false))
foreach($a in $reports) { Write-Output ('DOCUMENT|'+$a.name+'|OMML='+$a.maths.Count+'|OLE='+$a.shapes.Count+'|tables='+$a.tables.Count+'|fields='+$a.fields.Count+'|end='+$a.end);foreach($t in $a.tables) {Write-Output ('TABLE|'+$t.start+':'+$t.end+'|'+$t.rows+'x'+$t.cols)};foreach($m in $a.maths) {Write-Output ('MATH|'+$m.index+'|'+$m.start+':'+$m.end+'|'+$m.font+'|'+$m.size)} }
Write-Output ('INVENTORY|'+$dest)
Release $word
