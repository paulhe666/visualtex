param([string]$Stage,[string]$DocumentName,[string]$CloseEmptyDocument='', [switch]$DismissDocumentChanged)
$ErrorActionPreference='Stop';[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$record=Get-Content (Join-Path $PSScriptRoot "evidence\$Stage-word-process.json") -Raw -Encoding UTF8|ConvertFrom-Json
$process=Get-Process -Id $record.pid
if($record.pid -eq 27664 -or $record.pid -eq 107976 -or $process.ProcessName -ne 'WINWORD' -or [Math]::Abs(($process.StartTime.ToUniversalTime()-[DateTimeOffset]::Parse($record.started).UtcDateTime).TotalSeconds) -gt 10){throw 'Test process identity changed.'}
if($DismissDocumentChanged){
 Add-Type -TypeDefinition @'
using System;using System.Text;using System.Runtime.InteropServices;
public static class PerfDialog {
 public delegate bool EnumProc(IntPtr hwnd,IntPtr data);
 [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc callback,IntPtr data);
 [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent,EnumProc callback,IntPtr data);
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hwnd,StringBuilder text,int length);
 [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hwnd,uint message,UIntPtr w,IntPtr l);
 [DllImport("user32.dll")] public static extern int GetDlgCtrlID(IntPtr hwnd);
}
'@
 $script:dialog=[IntPtr]::Zero
 [void][PerfDialog]::EnumWindows({param($h,$l) [uint32]$p=0;[void][PerfDialog]::GetWindowThreadProcessId($h,[ref]$p);$s=[Text.StringBuilder]::new(256);[void][PerfDialog]::GetWindowText($h,$s,256);if($p -eq $record.pid -and $s.ToString() -eq 'VisualTeX 批量导入'){$script:dialog=$h};return $true},[IntPtr]::Zero)
 if($script:dialog -eq [IntPtr]::Zero){throw 'The expected import-error dialog is missing.'}
 $script:knownError=$false;$script:ok=[IntPtr]::Zero
 [void][PerfDialog]::EnumChildWindows($script:dialog,{param($h,$l) $s=[Text.StringBuilder]::new(1024);[void][PerfDialog]::GetWindowText($h,$s,1024);if($s.ToString() -eq 'The active Word document changed while the VisualTeX editor was open.'){$script:knownError=$true};if([PerfDialog]::GetDlgCtrlID($h) -eq 2 -and $s.ToString() -eq '确定'){$script:ok=$h};return $true},[IntPtr]::Zero)
 if(!$script:knownError -or $script:ok -eq [IntPtr]::Zero){throw 'Dialog contents differ; no button was clicked.'}
 [void][PerfDialog]::PostMessage($script:ok,0xF5,[UIntPtr]::Zero,[IntPtr]::Zero);Start-Sleep -Milliseconds 500
}
$word=Connect-RunningWord $record.pid
try {
 if(!$DocumentName){
  if($word.Documents.Count -ne 1){throw 'Automatic seed adoption requires one unchanged empty test document.'}
  $seed=$word.ActiveDocument
  if(!$seed.Saved -or $seed.Path -or $seed.Content.End -ne 1){throw 'The test seed contains content; preserving it.'}
  $DocumentName=$seed.Name
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($seed)
 }
 $doc=$word.Documents.Item($DocumentName)
 if($doc.Name -ne $DocumentName){throw 'Target document missing.'}
 if($CloseEmptyDocument){
  if($CloseEmptyDocument -eq $DocumentName){throw 'Cannot close the target.'}
  $empty=$word.Documents.Item($CloseEmptyDocument)
  if(!$empty.Saved -or $empty.Path -or $empty.Content.End -ne 1 -or $empty.OMaths.Count -ne 0 -or $empty.InlineShapes.Count -ne 0 -or $empty.Fields.Count -ne 0){throw 'The test seed is no longer an unchanged empty document; preserving it.'}
  $empty.Close(0);[void][Runtime.InteropServices.Marshal]::ReleaseComObject($empty)
 }
 $doc.Activate()
 $record|Add-Member -NotePropertyName document -NotePropertyValue $DocumentName -Force
 $record|ConvertTo-Json|Set-Content (Join-Path $PSScriptRoot 'evidence\target-word.json') -Encoding UTF8
 Write-Output ("PERF_TARGET|pid=$($record.pid)|document=$DocumentName|omml=$($doc.OMaths.Count)|ole=$($doc.InlineShapes.Count)|end=$($doc.Content.End)")
 [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc)
}finally{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)}
