param([int]$WordProcessId,[string]$DocumentName,[ValidateSet('math','ole','prose')][string]$Kind='math',[int]$Index=12,[int]$Position=2,[string]$Label='font-state')
$ErrorActionPreference='Stop';[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
Add-Type -AssemblyName UIAutomationClient,UIAutomationTypes
Add-Type -TypeDefinition @'
using System;using System.Runtime.InteropServices;
public static class FontStateClick {
[StructLayout(LayoutKind.Sequential)] public struct Point { public int x,y; }
[DllImport("user32.dll")]public static extern bool SetProcessDPIAware();
[DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);
[DllImport("user32.dll")]public static extern bool ShowWindow(IntPtr h,int c);
[DllImport("user32.dll")]public static extern IntPtr GetForegroundWindow();
[DllImport("user32.dll")]public static extern bool SetCursorPos(int x,int y);
[DllImport("user32.dll")]public static extern void mouse_event(uint f,uint x,uint y,uint d,UIntPtr e);
[DllImport("user32.dll")]static extern IntPtr WindowFromPoint(Point point);
[DllImport("user32.dll")]static extern uint GetWindowThreadProcessId(IntPtr h,out uint pid);
public static uint ProcessAt(int x,int y){uint pid;GetWindowThreadProcessId(WindowFromPoint(new Point{x=x,y=y}),out pid);return pid;}
}
'@
[void][FontStateClick]::SetProcessDPIAware()
$w=Connect-RunningWord $WordProcessId;$d=$w.ActiveDocument
if($d.Name -ne $DocumentName -or $WordProcessId -in @(107976,27664)){throw 'Not the owned performance document.'}
$win=$w.ActiveWindow;$h=[IntPtr]$win.Hwnd;[void][FontStateClick]::ShowWindow($h,3);[void][FontStateClick]::SetForegroundWindow($h)
$math=$null;$shape=$null
if($Kind -eq 'math'){$math=$d.OMaths.Item($Index);$range=$math.Range}elseif($Kind -eq 'ole'){$shape=$d.InlineShapes.Item($Index);$range=$shape.Range}else{$range=$d.Range($Position,$Position+1)}
try{
 $win.ScrollIntoView($range,$true);Start-Sleep -Milliseconds 400
 $x=0;$y=0;$width=0;$height=0;$win.GetPoint([ref]$x,[ref]$y,[ref]$width,[ref]$height,$range)
 $px=$x+[int]($width/2);$py=$y+[int]($height/2)
 if($width -le 0 -or $height -le 0 -or [FontStateClick]::ProcessAt($px,$py) -ne $WordProcessId){throw 'Target point is not visible in owned Word.'}
 $before=@{start=$w.Selection.Start;end=$w.Selection.End}
 [void][FontStateClick]::SetCursorPos($px,$py);[FontStateClick]::mouse_event(2,0,0,0,[UIntPtr]::Zero);[FontStateClick]::mouse_event(4,0,0,0,[UIntPtr]::Zero)
 Start-Sleep -Milliseconds 800
 if([FontStateClick]::GetForegroundWindow() -ne $h){throw 'Word lost foreground; do not accept this sample.'}
 $after=@{start=$w.Selection.Start;end=$w.Selection.End}
 $root=[Windows.Automation.AutomationElement]::FromHandle($h)
 $condition=[Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::NameProperty,'字号')
 $controls=$root.FindAll([Windows.Automation.TreeScope]::Descendants,$condition)
 $state=@();foreach($control in $controls){$state+=@{type=$control.Current.ControlType.ProgrammaticName;enabled=$control.Current.IsEnabled;offscreen=$control.Current.IsOffscreen}}
 $report=@{pid=$WordProcessId;document=$DocumentName;kind=$Kind;index=$Index;before=$before;after=$after;target=@{start=$range.Start;end=$range.End};controls=$state;method='physical mouse click, then real Ribbon UIA read'}
 $json=$report|ConvertTo-Json -Depth 5;[IO.File]::WriteAllText((Join-Path $PSScriptRoot "evidence\$Label.json"),$json,[Text.UTF8Encoding]::new($false));$json
}finally{foreach($o in @($range,$math,$shape,$win,$d,$w)){if($null -ne $o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($o)}}}
