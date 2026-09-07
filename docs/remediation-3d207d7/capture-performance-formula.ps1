param([int]$WordProcessId,[string]$DocumentName,[int]$Index=12,[string]$Label='formula-visual')
$ErrorActionPreference='Stop';[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
Add-Type -AssemblyName System.Drawing,System.Windows.Forms
Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public static class CaptureUi{[DllImport("user32.dll")]public static extern bool SetProcessDPIAware();[DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);}'
[void][CaptureUi]::SetProcessDPIAware()
$record=Get-Content (Join-Path $PSScriptRoot 'evidence\target-word.json') -Raw|ConvertFrom-Json
if($record.pid -ne $WordProcessId -or $WordProcessId -in @(107976,27664)){throw 'Requires the owned test Word.'}
$w=Connect-RunningWord $WordProcessId;$d=$w.ActiveDocument
if($d.Name -ne $DocumentName -or $d.OMaths.Count+$d.InlineShapes.Count -lt 20){throw 'Requires the specified document with at least 20 formulas.'}
$m=$d.OMaths.Item($Index);$r=$m.Range;$window=$w.ActiveWindow;$f=$r.ParagraphFormat;$font=$r.Font
try{
 [void][CaptureUi]::SetForegroundWindow([IntPtr]$window.Hwnd);$window.ScrollIntoView($r,$true);Start-Sleep -Milliseconds 400
 $x=0;$y=0;$width=0;$height=0;$window.GetPoint([ref]$x,[ref]$y,[ref]$width,[ref]$height,$r)
 $bounds=[System.Windows.Forms.Screen]::PrimaryScreen.Bounds
 $rect=[Drawing.Rectangle]::Intersect($bounds,[Drawing.Rectangle]::new([Math]::Max(0,$x-25),[Math]::Max(0,$y-45),[Math]::Max(600,$width+50),[Math]::Max(210,$height+90)))
 $bmp=[Drawing.Bitmap]::new($rect.Width,$rect.Height);$g=[Drawing.Graphics]::FromImage($bmp)
 try{$g.CopyFromScreen($rect.Location,[Drawing.Point]::Empty,$rect.Size);$bmp.Save((Join-Path $PSScriptRoot "evidence\$Label.png"),[Drawing.Imaging.ImageFormat]::Png)}finally{$g.Dispose();$bmp.Dispose()}
 $report=@{pid=$WordProcessId;document=$DocumentName;index=$Index;text=$r.Text;size=$font.Size;type=[int]$m.Type;maths=$d.OMaths.Count;oles=$d.InlineShapes.Count;tables=$d.Tables.Count;fields=$d.Fields.Count;spaceAfter=$f.SpaceAfter;spaceBefore=$f.SpaceBefore;lineRule=[int]$f.LineSpacingRule;pixelHeight=$height}
 $json=$report|ConvertTo-Json;[IO.File]::WriteAllText((Join-Path $PSScriptRoot "evidence\$Label.json"),$json,[Text.UTF8Encoding]::new($false));$json
}finally{foreach($o in @($font,$f,$window,$r,$m,$d,$w)){if($null -ne $o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($o)}}}
