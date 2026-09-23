$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
function Rel($o){if($null-ne$o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($o)}catch{}}}
$xml=[IO.File]::ReadAllText('C:/Users/pojian_liao/Desktop/devspace/visualtex1.2.3-reference-3d207d7/docs/remediation-3d207d7/evidence/numbered-mt-failure-beforefix-latex_100_inline_100_display_docx.xml',[Text.Encoding]::UTF8)
$w=$null;$d=$null;$s=$null
try{
 $w=New-Object -ComObject Word.Application;$w.Visible=$false;$w.DisplayAlerts=0
 $d=$w.Documents.Add();$d.Content.InsertXML($xml)
 Write-Output ('CLONE|compat='+$d.CompatibilityMode+'|shapes='+$d.InlineShapes.Count+'|fields='+$d.Fields.Count)
 foreach($idx in @(1,2,4,116)){
  if($idx -gt $d.InlineShapes.Count){continue}
  $s=$d.InlineShapes.Item($idx);$r=$s.Range;$width=[single]$s.Width;$height=[single]$s.Height
  $p=$r.Paragraphs.Item(1);$place=$false
  foreach($f in @($p.Range.Fields)){try{if(([string]$f.Code.Text)-match 'MTPlaceRef'){$place=$true}}catch{};Rel $f}
  Rel $p
  $steps=@()
  foreach($op in @('unlock','width','height','lock')){
   try{
    if($op-eq'unlock'){$s.LockAspectRatio=0}
    elseif($op-eq'width'){$s.Width=$width}
    elseif($op-eq'height'){$s.Height=$height}
    else{$s.LockAspectRatio=-1}
    $steps+=($op+':ok')
   }catch{$steps+=($op+':FAIL:0x'+('{0:X8}' -f ($_.Exception.HResult -band 0xffffffff))+':'+$_.Exception.Message)}
  }
  Write-Output ('SHAPE|'+$idx+'|range='+$r.Start+':'+$r.End+'|numbered='+$place+'|'+($steps -join '|'))
  Rel $r;Rel $s;$s=$null
 }
}finally{if($d){try{$d.Close(0)}catch{}};if($w){try{$w.Quit()}catch{}};Rel $s;Rel $d;Rel $w;[GC]::Collect();[GC]::WaitForPendingFinalizers()}
