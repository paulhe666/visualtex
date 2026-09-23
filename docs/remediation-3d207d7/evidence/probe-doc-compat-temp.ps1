$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
function Rel($o){if($null-ne$o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($o)}catch{}}}
$w=$null;$d=$null
try{
 $w=New-Object -ComObject Word.Application;$w.Visible=$false;$w.DisplayAlerts=0
 $d=$w.Documents.Open('C:/Users/pojian_liao/Downloads/latex_100_inline_100_display.docx',$false,$true)
 Write-Output ('DISK|compat='+$d.CompatibilityMode+'|saved='+$d.Saved+'|shapes='+$d.InlineShapes.Count+'|fields='+$d.Fields.Count+'|paras='+$d.Paragraphs.Count)
}finally{if($d){try{$d.Close(0)}catch{}};if($w){try{$w.Quit()}catch{}};Rel $d;Rel $w;[GC]::Collect();[GC]::WaitForPendingFinalizers()}
