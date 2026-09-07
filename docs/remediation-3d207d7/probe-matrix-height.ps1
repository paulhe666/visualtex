param([int]$WordProcessId,[string]$DocumentName,[int]$Index=12,[string]$Label='matrix-height')
$ErrorActionPreference='Stop';[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
Add-Type -AssemblyName System.Drawing,System.Windows.Forms
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class HeightUi{[DllImport("user32.dll")]public static extern bool SetProcessDPIAware();[DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);}'
[void][HeightUi]::SetProcessDPIAware()
$target=Get-Content (Join-Path $PSScriptRoot 'evidence\target-word.json') -Raw|ConvertFrom-Json
if($target.pid -ne $WordProcessId -or $WordProcessId -in @(107976,27664)){throw 'Only the owned test process is allowed.'}
$w=Connect-RunningWord $WordProcessId;$d=$w.ActiveDocument
if($d.Name -ne $DocumentName -or $d.OMaths.Count -lt 20){throw 'Needs the specified 20-formula test document.'}
$m=$d.OMaths.Item($Index);$r=$m.Range;$t=$r.Tables.Item(1);$row=$t.Rows.Item(1);$window=$w.ActiveWindow
$format=$r.ParagraphFormat
$before=@{height=$row.Height;rule=$row.HeightRule;top=$t.TopPadding;bottom=$t.BottomPadding;mathText=$r.Text;count=$d.OMaths.Count;end=$d.Content.End;lineRule=$format.LineSpacingRule;lineSpacing=$format.LineSpacing}
function Shot($suffix){
 $window.ScrollIntoView($r,$true);[void][HeightUi]::SetForegroundWindow([IntPtr]$window.Hwnd);Start-Sleep -Milliseconds 400
 $x=0;$y=0;$width=0;$height=0;$window.GetPoint([ref]$x,[ref]$y,[ref]$width,[ref]$height,$r)
 $rect=[Drawing.Rectangle]::new([Math]::Max(0,$x-30),[Math]::Max(0,$y-60),[Math]::Max(600,$width+60),[Math]::Max(240,$height+120))
 $bmp=[Drawing.Bitmap]::new($rect.Width,$rect.Height);$g=[Drawing.Graphics]::FromImage($bmp)
 try{$g.CopyFromScreen($rect.Location,[Drawing.Point]::Empty,$rect.Size);$bmp.Save((Join-Path $PSScriptRoot "evidence\$Label-$suffix.png"),[Drawing.Imaging.ImageFormat]::Png)}finally{$g.Dispose();$bmp.Dispose()}
 Write-Output "$suffix|rowHeight=$($row.Height)|mathPixels=$height|top=$($t.TopPadding)|bottom=$($t.BottomPadding)"
}
try{
 Shot 'before'
 $row.HeightRule=1;$row.Height=60;Shot 'row60'
 $row.Height=$before.height;$t.TopPadding=3;$t.BottomPadding=3;Shot 'padding3'
 $t.TopPadding=$before.top;$t.BottomPadding=$before.bottom
 $format.LineSpacingRule=3;$format.LineSpacing=36;Shot 'line-atleast36'
 $format.LineSpacing=48;Shot 'line-atleast48'
 $format.LineSpacingRule=0
 foreach($after in @(1,2,3,6)){$format.SpaceAfter=$after;Shot ('paragraph-after'+$after)}
 $format.SpaceAfter=0
}finally{
 $format.LineSpacingRule=$before.lineRule;$format.LineSpacing=$before.lineSpacing
 $row.HeightRule=$before.rule;$row.Height=$before.height;$t.TopPadding=$before.top;$t.BottomPadding=$before.bottom
 if($d.OMaths.Count -ne $before.count -or $d.Content.End -ne $before.end -or $r.Text -ne $before.mathText){throw 'Test content changed unexpectedly.'}
 [IO.File]::WriteAllText((Join-Path $PSScriptRoot "evidence\$Label.json"),($before|ConvertTo-Json),[Text.UTF8Encoding]::new($false))
 foreach($o in @($format,$row,$t,$r,$m,$window,$d,$w)){if($null -ne $o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($o)}}
}
