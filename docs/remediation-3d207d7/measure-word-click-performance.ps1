param(
 [int]$WordProcessId,
 [string]$DocumentName,
 [string]$Label='click-perf',
 [int]$Count=20
)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
if(!('VisualTeXClickPerf' -as [type])){
 Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class VisualTeXClickPerf {
 [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
 [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hwnd);
 [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr hwnd);
 [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hwnd, int command);
 [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
 [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
 [DllImport("user32.dll")] public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extra);
 public const uint LEFTDOWN=0x0002;
 public const uint LEFTUP=0x0004;
 public static void Click(int x,int y){SetCursorPos(x,y);mouse_event(LEFTDOWN,0,0,0,UIntPtr.Zero);mouse_event(LEFTUP,0,0,0,UIntPtr.Zero);}
}
'@
}
[void][VisualTeXClickPerf]::SetProcessDPIAware()
function Release-Com($value){if($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($value)}}
$word=$null;$document=$null;$window=$null;$content=$null;$paragraphs=$null
try {
 $word=Connect-RunningWord $WordProcessId
 for($i=1;$i -le $word.Documents.Count;$i++){ $candidate=$word.Documents.Item($i);if($candidate.Name -eq $DocumentName){$document=$candidate;break}else{Release-Com $candidate}}
 if($null -eq $document){throw "Document not found: $DocumentName"}
 $document.Activate();Start-Sleep -Milliseconds 300
 $window=$word.ActiveWindow;$content=$document.Content;$paragraphs=$content.Paragraphs
 $targets=New-Object System.Collections.Generic.List[object]
 for($i=1;$i -le $paragraphs.Count;$i++){
  $paragraph=$null;$range=$null;$tables=$null
  try {
   $paragraph=$paragraphs.Item($i);$range=$paragraph.Range.Duplicate
   $tables=$range.Tables;if($tables.Count -gt 0){continue}
   $text=[string]$range.Text
   if([string]::IsNullOrWhiteSpace($text) -or $text.Length -lt 12){continue}
   # Pick two ordinary-text character positions near the start of each prose paragraph.
   foreach($offset in @(2,6)){
    $start=[Math]::Min($range.End-2,$range.Start+$offset)
    if($start -le $range.Start -or $start -ge $range.End-1){continue}
    $targets.Add([pscustomobject]@{paragraph=$i;start=$start;end=$start+1;matchStart=$range.Start;matchEnd=[Math]::Min($range.End-1,$range.Start+12);text=$text.Substring(0,[Math]::Min(32,$text.Length)).Replace("`r",'').Replace("`n",'')})
    if($targets.Count -ge $Count){break}
   }
   if($targets.Count -ge $Count){break}
  } finally {Release-Com $tables;Release-Com $range;Release-Com $paragraph}
 }
 if($targets.Count -lt $Count){throw "Only $($targets.Count) prose click targets were available; expected $Count."}
 $rows=@();$process=[Diagnostics.Process]::GetProcessById($WordProcessId)
 for($index=0;$index -lt $Count;$index++){
  $target=$targets[$index];$probe=$null;$selection=$null
  try {
   $probe=$document.Range([int]$target.start,[int]$target.end)
   $window.ScrollIntoView($probe,$true);Start-Sleep -Milliseconds 80
   $left=0;$top=0;$width=0;$height=0
   $window.GetPoint([ref]$left,[ref]$top,[ref]$width,[ref]$height,$probe)
   if($width -le 0 -or $height -le 0){throw "Invalid target rectangle at $($target.start): ${left},${top},${width},${height}"}
   $wordHwnd=[IntPtr]$window.Hwnd
   [void][VisualTeXClickPerf]::ShowWindow($wordHwnd,9);[void][VisualTeXClickPerf]::BringWindowToTop($wordHwnd);[void][VisualTeXClickPerf]::SetForegroundWindow($wordHwnd)
   $foregroundDeadline=[DateTime]::UtcNow.AddSeconds(2);$foregroundPid=0
   do {$foreground=[VisualTeXClickPerf]::GetForegroundWindow();[uint32]$owner=0;[void][VisualTeXClickPerf]::GetWindowThreadProcessId($foreground,[ref]$owner);$foregroundPid=[int]$owner;if($foregroundPid -eq $WordProcessId){break};Start-Sleep -Milliseconds 25} while([DateTime]::UtcNow -lt $foregroundDeadline)
   if($foregroundPid -ne $WordProcessId){throw "The test Word process did not obtain foreground; actual pid=$foregroundPid."}
   Start-Sleep -Milliseconds 50
   $process.Refresh();$cpuBefore=$process.TotalProcessorTime.TotalMilliseconds
   $watch=[Diagnostics.Stopwatch]::StartNew();[VisualTeXClickPerf]::Click($left+[Math]::Max(1,[int]($width/2)),$top+[Math]::Max(1,[int]($height/2)))
   $deadline=[DateTime]::UtcNow.AddSeconds(8);$matched=$false;$selectionStart=-1
   do {
    Release-Com $selection;$selection=$word.Selection;$selectionStart=$selection.Start
    if($selectionStart -ge $target.matchStart -and $selectionStart -le $target.matchEnd){$matched=$true;break}
    Start-Sleep -Milliseconds 10
   } while([DateTime]::UtcNow -lt $deadline)
   $watch.Stop();$process.Refresh();$cpuAfter=$process.TotalProcessorTime.TotalMilliseconds
   $rows += [pscustomobject]@{index=$index+1;paragraph=$target.paragraph;targetStart=$target.start;matchStart=$target.matchStart;matchEnd=$target.matchEnd;selectionStart=$selectionStart;matched=$matched;elapsedMs=$watch.Elapsed.TotalMilliseconds;cpuMs=$cpuAfter-$cpuBefore;left=$left;top=$top;width=$width;height=$height;text=$target.text}
   if(!$matched){throw "Physical click $($index+1) did not land on the expected prose range: target=$($target.start), match=$($target.matchStart):$($target.matchEnd), selection=$selectionStart, rect=${left},${top},${width},${height}."}
   Start-Sleep -Milliseconds 120
  } finally {Release-Com $selection;Release-Com $probe}
 }
 $elapsed=@($rows|ForEach-Object {$_.elapsedMs}|Sort-Object)
 $avg=($elapsed|Measure-Object -Average).Average;$max=($elapsed|Measure-Object -Maximum).Maximum;$p95=$elapsed[[Math]::Min($elapsed.Count-1,[Math]::Floor(($elapsed.Count-1)*0.95))]
 $report=[pscustomobject]@{label=$Label;pid=$WordProcessId;document=$DocumentName;formulaCount=$document.OMaths.Count;tableCount=$document.Tables.Count;clicks=$rows.Count;averageMs=[Math]::Round($avg,2);p95Ms=[Math]::Round($p95,2);maxMs=[Math]::Round($max,2);rows=$rows}
 $out=Join-Path $PSScriptRoot 'evidence';[void][IO.Directory]::CreateDirectory($out);$path=Join-Path $out ($Label+'-click-performance.json')
 [IO.File]::WriteAllText($path,($report|ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($false))
 Write-Output ("CLICK_PERF|document=$DocumentName|formulas=$($report.formulaCount)|tables=$($report.tableCount)|clicks=$($report.clicks)|avgMs=$($report.averageMs)|p95Ms=$($report.p95Ms)|maxMs=$($report.maxMs)")
 Write-Output ("EVIDENCE|"+$path)
} finally {Release-Com $paragraphs;Release-Com $content;Release-Com $window;Release-Com $document;Release-Com $word}
