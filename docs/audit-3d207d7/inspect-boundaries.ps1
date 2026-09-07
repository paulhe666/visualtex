param([string]$Label='boundaries')
$ErrorActionPreference='Stop';[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$w=[Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application');$d=$w.ActiveDocument
$out=Join-Path $PSScriptRoot 'evidence'
[xml]$xml=$d.Content.WordOpenXML;$ns=[Xml.XmlNamespaceManager]::new($xml.NameTable);$ns.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main');$ns.AddNamespace('m','http://schemas.openxmlformats.org/officeDocument/2006/math')
$r=[ordered]@{document=$d.Name;mathFont=$d.OMathFontName;anchors=@();separators=@()}
foreach($b in $xml.SelectNodes('//w:bookmarkStart[starts-with(@w:name,"VTOMML_")]',$ns)){$name=$b.GetAttribute('name',$ns.LookupNamespace('w'));$anc=@();$a=$b.ParentNode;while($a -and $anc.Count -lt 6){$anc+=$a.Name;$a=$a.ParentNode};$r.anchors+=@{name=$name;parents=$anc;insideMath=($null -ne $b.SelectSingleNode('ancestor::m:oMath',$ns));parentXml=$b.ParentNode.OuterXml.Substring(0,[Math]::Min(2400,$b.ParentNode.OuterXml.Length))}}
for($i=1;$i -lt $d.Tables.Count;$i++){$a=$d.Tables.Item($i);$b=$d.Tables.Item($i+1);$s=$d.Range($a.Range.End,$b.Range.Start);$r.separators+=@{start=$s.Start;end=$s.End;text=$s.Text;codes=@($s.Text.ToCharArray()|ForEach-Object{[int]$_});tables=$s.Tables.Count;shapes=$s.InlineShapes.Count;maths=$s.OMaths.Count;fields=$s.Fields.Count;bookmarks=@($s.Bookmarks|ForEach-Object{@{name=$_.Name;start=$_.Range.Start;end=$_.Range.End}});frames=$s.Frames.Count;paragraphs=$s.Paragraphs.Count;previous=@($a.Range.Start,$a.Range.End);next=@($b.Range.Start,$b.Range.End)}}
[IO.File]::WriteAllText((Join-Path $out ($Label+'.json')),($r|ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false));$r|ConvertTo-Json -Depth 10
