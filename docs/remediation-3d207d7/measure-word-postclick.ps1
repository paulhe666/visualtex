param([int]$WordProcessId,[string]$DocumentName,[string]$Label='postclick',[int]$Count=20,[int]$ObserveMilliseconds=5000,[ValidateSet('prefix','after-inline')][string]$Mode='prefix')
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
Add-Type -TypeDefinition @'
using System;using System.Collections.Generic;using System.Diagnostics;using System.Runtime.InteropServices;using System.Threading;
public static class PostClickProbe {
 [StructLayout(LayoutKind.Sequential)] public struct Point { public int x,y; }
 [StructLayout(LayoutKind.Sequential)] public struct Cursor { public int size,flags;public IntPtr handle;public Point point; }
 [DllImport("user32.dll")] static extern bool GetCursorInfo(ref Cursor cursor);
 [DllImport("user32.dll")] static extern IntPtr LoadCursor(IntPtr instance,IntPtr name);
 [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
 [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hwnd);
 [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hwnd,int command);
 [DllImport("user32.dll")] static extern IntPtr WindowFromPoint(Point point);
 public static uint ProcessAtPoint(int x,int y){uint pid;GetWindowThreadProcessId(WindowFromPoint(new Point{x=x,y=y}),out pid);return pid;}
 [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
 [DllImport("user32.dll")] public static extern bool SetCursorPos(int x,int y);
 [DllImport("user32.dll")] static extern void mouse_event(uint flags,uint x,uint y,uint data,UIntPtr extra);
 [DllImport("user32.dll",SetLastError=true)] static extern IntPtr SendMessageTimeout(IntPtr hwnd,uint msg,UIntPtr wp,IntPtr lp,uint flags,uint timeout,out UIntPtr result);
 [StructLayout(LayoutKind.Sequential)] public struct LastInput { public uint size,time; }
 [DllImport("user32.dll")] static extern bool GetLastInputInfo(ref LastInput input);
 public class Sample {public double ms,pingMs;public long cursor;public bool busy,responding;public uint foregroundPid,inputTick;public int x,y;}
 public static Sample[] ClickAndObserve(IntPtr hwnd,int x,int y,int duration) {
  var wait=LoadCursor(IntPtr.Zero,new IntPtr(32514));var starting=LoadCursor(IntPtr.Zero,new IntPtr(32650));
  SetCursorPos(x,y);var watch=Stopwatch.StartNew();mouse_event(2,0,0,0,UIntPtr.Zero);mouse_event(4,0,0,0,UIntPtr.Zero);
  var rows=new List<Sample>();
  while(watch.ElapsedMilliseconds<duration) {
   var cursor=new Cursor();cursor.size=Marshal.SizeOf(typeof(Cursor));if(!GetCursorInfo(ref cursor))throw new System.ComponentModel.Win32Exception();
   uint pid;GetWindowThreadProcessId(GetForegroundWindow(),out pid);
   var ping=Stopwatch.StartNew();UIntPtr result;bool responds=SendMessageTimeout(hwnd,0,UIntPtr.Zero,IntPtr.Zero,2,20,out result)!=IntPtr.Zero;
   var input=new LastInput();input.size=(uint)Marshal.SizeOf(typeof(LastInput));GetLastInputInfo(ref input);
   rows.Add(new Sample{ms=watch.Elapsed.TotalMilliseconds,pingMs=ping.Elapsed.TotalMilliseconds,cursor=cursor.handle.ToInt64(),busy=cursor.handle==wait||cursor.handle==starting,responding=responds,foregroundPid=pid,inputTick=input.time,x=cursor.point.x,y=cursor.point.y});
   Thread.Sleep(10);
  }
  return rows.ToArray();
 }
}
'@
function Release-Com($v){if($v -and [Runtime.InteropServices.Marshal]::IsComObject($v)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($v)}}
[void][PostClickProbe]::SetProcessDPIAware()
$word=Connect-RunningWord $WordProcessId
$doc=$word.ActiveDocument
if($doc.Name -ne $DocumentName){throw 'Recorded test document is not active; no mouse input was sent.'}
$total=$doc.OMaths.Count+$doc.InlineShapes.Count
if($total -lt 20){throw 'Performance sampling requires at least 20 real formulas.'}
$window=$word.ActiveWindow;$hwnd=[IntPtr]$window.Hwnd
[void][PostClickProbe]::ShowWindow($hwnd,3);[void][PostClickProbe]::SetForegroundWindow($hwnd);Start-Sleep -Milliseconds 500
if([PostClickProbe]::GetForegroundWindow() -ne $hwnd){throw 'Test Word is not foreground.'}
$targets=@();$paragraphs=$doc.Paragraphs
for($i=1;$i -le $paragraphs.Count;$i++){
 $p=$paragraphs.Item($i);$r=$p.Range
 try {
  if($r.Tables.Count -gt 0 -or ([string]$r.Text).Trim().Length -lt 12){continue}
  $anchor=$r.Start
  if($Mode -eq 'after-inline'){
   $localMath=$r.OMaths;$localShapes=$r.InlineShapes
   try{
    if($localMath.Count -eq 1){$formula=$localMath.Item(1);$formulaRange=$formula.Range;try{$anchor=$formulaRange.End}finally{Release-Com $formulaRange;Release-Com $formula}}
    elseif($localShapes.Count -eq 1){$formula=$localShapes.Item(1);$formulaRange=$formula.Range;try{$anchor=$formulaRange.End}finally{Release-Com $formulaRange;Release-Com $formula}}
    else{continue}
   }finally{Release-Com $localMath;Release-Com $localShapes}
  }
  foreach($offset in @(2,5)){
   $position=$anchor+$offset;if($position+1 -ge $r.End){continue};$probe=$doc.Range($position,$position+1)
   try {if($probe.OMaths.Count -gt 0 -or $probe.InlineShapes.Count -gt 0 -or $probe.Fields.Count -gt 0){continue};$targets+=@{start=$position;paragraph=$i}}finally{Release-Com $probe}
  }
 }finally{Release-Com $r;Release-Com $p}
}
if($targets.Count -lt $Count){throw 'Not enough verified prose targets.'}
$results=@();$process=[Diagnostics.Process]::GetProcessById($WordProcessId)
try {
 for($i=0;$i -lt $Count;$i++){
  $target=$targets[$i];$range=$doc.Range($target.start,$target.start+1)
  try {
   $window.ScrollIntoView($range,$true);Start-Sleep -Milliseconds 250
   $left=0;$top=0;$width=0;$height=0;$window.GetPoint([ref]$left,[ref]$top,[ref]$width,[ref]$height,$range)
   if($width -le 0 -or $height -le 0){throw 'No visible text rectangle.'}
   $x=$left+[int]($width/2);$y=$top+[int]($height/2)
   $hit=$window.RangeFromPoint($x,$y);$expectedStart=$hit.Start;$expectedEnd=$hit.End;Release-Com $hit
   if($expectedStart -lt $target.start-1 -or $expectedStart -gt $target.start+1){throw 'Physical target does not correspond to the requested text.'}
   if([PostClickProbe]::GetForegroundWindow() -ne $hwnd){throw 'Foreground changed; refusing to click another window.'}
   if([PostClickProbe]::ProcessAtPoint($x,$y) -ne $WordProcessId){throw 'Another window covers the intended document point; no click was sent.'}
   $process.Refresh();$cpu=$process.TotalProcessorTime.TotalMilliseconds
   # No COM calls while sampling: observe the actual global cursor, foreground and native Word message responsiveness.
   $samples=[PostClickProbe]::ClickAndObserve($hwnd,$x,$y,$ObserveMilliseconds)
   $process.Refresh();$cpu=$process.TotalProcessorTime.TotalMilliseconds-$cpu
   $selection=$word.Selection;$selectionStart=$selection.Start;$selectionEnd=$selection.End;Release-Com $selection
   $changes=0;$busyEntries=0;$busyMs=0.0;$lastBusyMs=0.0;$foregroundChanges=0
   for($j=0;$j -lt $samples.Count;$j++){
    $s=$samples[$j];if($j -gt 0){$prev=$samples[$j-1];if($s.cursor -ne $prev.cursor){$changes++};if($s.foregroundPid -ne $prev.foregroundPid){$foregroundChanges++};if($prev.busy){$busyMs+=$s.ms-$prev.ms}}
    if($s.busy){$lastBusyMs=$s.ms;if($j -eq 0 -or !$samples[$j-1].busy){$busyEntries++}}
   }
   $record=@{index=$i+1;target=$target;expectedStart=$expectedStart;expectedEnd=$expectedEnd;selectionStart=$selectionStart;selectionEnd=$selectionEnd;cursorChanges=$changes;busyEntries=$busyEntries;busyMilliseconds=[Math]::Round($busyMs,2);lastBusyMilliseconds=[Math]::Round($lastBusyMs,2);foregroundChanges=$foregroundChanges;unresponsiveSamples=@($samples|Where-Object {!$_.responding}).Count;cpuMs=$cpu;samples=$samples}
   $results+=$record
   Write-Output ("POSTCLICK|index=$($i+1)|busyEntries=$busyEntries|busyMs=$([Math]::Round($busyMs,2))|cursorChanges=$changes|foregroundChanges=$foregroundChanges|cpuMs=$cpu")
   if($selectionStart -lt $expectedStart-1 -or $selectionStart -gt $expectedEnd+1){throw 'The click did not settle at the requested prose location.'}
  }finally{Release-Com $range}
 }
}finally{
 $report=@{label=$Label;pid=$WordProcessId;document=$DocumentName;formulaCount=$total;mode=$Mode;requestedClicks=$Count;completedClicks=$results.Count;observeMilliseconds=$ObserveMilliseconds;rows=$results;method='Physical mouse input followed by Win32-only sampling; no COM polling during observation.'}
 [IO.File]::WriteAllText((Join-Path $PSScriptRoot "evidence\$Label-postclick.json"),($report|ConvertTo-Json -Depth 8),[Text.UTF8Encoding]::new($false))
 Release-Com $paragraphs;Release-Com $window;Release-Com $doc;Release-Com $word
}
