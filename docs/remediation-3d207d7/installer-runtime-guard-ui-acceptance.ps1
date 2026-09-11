[CmdletBinding()]
param([string]$InstallerPath, [string]$Label = 'runtime-guard-ui-v3')
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$root = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
if (-not $InstallerPath) { $InstallerPath = Join-Path $root 'apps/windows/src-tauri/target/release/bundle/nsis/VisualTeX_1.2.6_x64-setup.exe' }
$InstallerPath = [IO.Path]::GetFullPath($InstallerPath)
$reportPath = Join-Path $PSScriptRoot ('evidence/' + $Label + '.json')
if (Test-Path -LiteralPath $reportPath) { throw 'Evidence already exists; select a new label.' }
Add-Type @'
using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
public static class RuntimeInstallerUI {
 public class Item { public long Handle; public int Id; public string Text; public string Class; public bool Enabled; }
 delegate bool EnumProc(IntPtr h, IntPtr p);
 [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc c, IntPtr p);
 [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h, EnumProc c, IntPtr p);
 [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint p);
 [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
 [DllImport("user32.dll")] static extern bool IsWindowEnabled(IntPtr h);
 [DllImport("user32.dll")] static extern int GetDlgCtrlID(IntPtr h);
 [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder b, int n);
 [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder b, int n);
 [DllImport("user32.dll")] static extern bool PostMessage(IntPtr h, uint m, IntPtr w, IntPtr l);
 [DllImport("user32.dll")] static extern IntPtr GetAncestor(IntPtr h, uint flags);
 [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("user32.dll")] static extern bool SetProcessDpiAwarenessContext(IntPtr context);
 [StructLayout(LayoutKind.Sequential)] struct Rect { public int Left,Top,Right,Bottom; }
 [StructLayout(LayoutKind.Sequential)] struct Point { public int X,Y; }
 [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h,out Rect r);
 [DllImport("user32.dll")] static extern IntPtr WindowFromPoint(Point p);
 [DllImport("user32.dll")] static extern bool SetCursorPos(int x,int y);
 [DllImport("user32.dll")] static extern void mouse_event(uint f,uint x,uint y,uint data,UIntPtr extra);
 public static List<Item> Read(int pid) {
  var items = new List<Item>();
  Action<IntPtr> add = h => { uint owner; GetWindowThreadProcessId(h, out owner); if(owner!=pid || !IsWindowVisible(h))return;
   var text=new StringBuilder(8192);var cls=new StringBuilder(256);GetWindowText(h,text,text.Capacity);GetClassName(h,cls,cls.Capacity);
   items.Add(new Item { Handle=h.ToInt64(),Id=GetDlgCtrlID(h),Text=text.ToString(),Class=cls.ToString(),Enabled=IsWindowEnabled(h) }); };
  EnumWindows((h,p)=> { uint owner;GetWindowThreadProcessId(h,out owner);if(owner==pid && IsWindowVisible(h)) {add(h);EnumChildWindows(h,(c,q)=>{add(c);return true;},IntPtr.Zero);}return true;},IntPtr.Zero);
  return items;
 }
 public static void Click(int pid, long handle) {
  var h=new IntPtr(handle);uint owner;GetWindowThreadProcessId(h,out owner);
  if(owner!=pid || !IsWindowVisible(h) || !IsWindowEnabled(h))throw new Exception("Refusing an unowned or disabled UI target");
  SetForegroundWindow(GetAncestor(h,2));
  System.Threading.Thread.Sleep(120);
  // Deliver the native Next-button notification to its dialog. Other buttons
  // use BM_CLICK so radio/check state changes are handled by Windows itself.
  if(GetDlgCtrlID(h)==1) {
   if(!PostMessage(GetAncestor(h,2),0x0111,new IntPtr(1),h))throw new Exception("Next notification failed");
  } else if(!PostMessage(h,0x00F5,IntPtr.Zero,IntPtr.Zero))throw new Exception("BM_CLICK failed");
 }
}
'@
function Assert-True([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
function Wait-For([scriptblock]$Condition, [string]$Description, [int]$Seconds=18) {
 $until=[DateTime]::UtcNow.AddSeconds($Seconds)
 do { if (& $Condition) { return }; Start-Sleep -Milliseconds 120 } while([DateTime]::UtcNow -lt $until)
 throw ('Timed out: ' + $Description + '; UI=' + (([RuntimeInstallerUI]::Read($script:setup.Id) | ForEach-Object { $_.Text }) -join ' | '))
}
function Has-Text([string]$Text) { return @([RuntimeInstallerUI]::Read($script:setup.Id) | Where-Object { $_.Text.Contains($Text) }).Count -gt 0 }
function Click-Id([int]$Id) {
 $buttons=@([RuntimeInstallerUI]::Read($script:setup.Id) | Where-Object { $_.Id -eq $Id -and $_.Class -eq 'Button' -and $_.Enabled })
 if ($buttons.Count -ne 1) { throw "Expected one enabled button ID $Id; found $($buttons.Count)." }
 [RuntimeInstallerUI]::Click($script:setup.Id,$buttons[0].Handle)
 Start-Sleep -Milliseconds 150
}
function Click-Text([string]$Text) {
 $buttons=@([RuntimeInstallerUI]::Read($script:setup.Id) | Where-Object { $_.Text.Contains($Text) -and $_.Class -eq 'Button' -and $_.Enabled })
 if ($buttons.Count -ne 1) { throw "Expected unique button '$Text'." }
 [RuntimeInstallerUI]::Click($script:setup.Id,$buttons[0].Handle)
 Start-Sleep -Milliseconds 150
}
function Start-Probe([string]$Exe) {
 $si=[Diagnostics.ProcessStartInfo]::new();$si.FileName=$Exe;$si.Arguments='127.0.0.1 -t';$si.UseShellExecute=$false;$si.CreateNoWindow=$true
 $p=[Diagnostics.Process]::Start($si);$script:ownedProbes.Add($p);return $p
}
function Is-Running($Process) { $Process.Refresh();return -not $Process.HasExited }
function Registry-Snapshot {
 $result=@()
 foreach($hive in @([Microsoft.Win32.RegistryHive]::CurrentUser,[Microsoft.Win32.RegistryHive]::LocalMachine)) {
  foreach($view in @([Microsoft.Win32.RegistryView]::Registry32,[Microsoft.Win32.RegistryView]::Registry64)) {
   $base=[Microsoft.Win32.RegistryKey]::OpenBaseKey($hive,$view)
   try { foreach($keyPath in @('Software\Classes\CLSID\{0002CE03-0000-0000-C000-000000000046}\LocalServer32','Software\Classes\Equation.DSMT4\CLSID','Software\Microsoft\Windows\CurrentVersion\App Paths\MathType.exe')) {
    $key=$base.OpenSubKey($keyPath);try { $result+=('{0}|{1}|{2}|{3}' -f $hive,$view,$keyPath,$(if($key){$key.GetValue('')}else{'<absent>'})) }finally{if($key){$key.Dispose()}}
   }}finally{$base.Dispose()}
  }
 }
 return ($result -join "`n")
}
$work=[IO.Path]::GetFullPath((Join-Path $env:TEMP ('VisualTeX Runtime UI ' + [guid]::NewGuid().ToString('N'))))
$install=Join-Path $work "app o'brien"
$private=Join-Path $install 'mathtype-runtime'
$outside=Join-Path $install 'mathtype-runtime-external'
[void][IO.Directory]::CreateDirectory($private);[void][IO.Directory]::CreateDirectory($outside)
$ping32=Join-Path $env:WINDIR 'SysWOW64/ping.exe';$ping64=Join-Path $env:WINDIR 'System32/ping.exe'
foreach($folder in @($private,$outside)) {
 Copy-Item -LiteralPath $ping32 -Destination (Join-Path $folder 'MathType.exe')
 Copy-Item -LiteralPath $ping64 -Destination (Join-Path $folder 'MathTypeLib.exe')
}
$script:ownedProbes=[Collections.Generic.List[Diagnostics.Process]]::new()
$script:setup=$null
$report=[ordered]@{label=$Label;installer=$InstallerPath;sha256=(Get-FileHash -LiteralPath $InstallerPath).Hash;testRoot=$install;started=[DateTimeOffset]::Now.ToString('o');passed=$false;checks=@()}
$beforeRegistry=Registry-Snapshot
$beforeWord=@(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object {$_.Id})
$prompt='Close the VisualTeX-owned runtime and continue?'
try {
 $inside32=Start-Probe (Join-Path $private 'MathType.exe');$inside64=Start-Probe (Join-Path $private 'MathTypeLib.exe')
 $outside32=Start-Probe (Join-Path $outside 'MathType.exe');$outside64=Start-Probe (Join-Path $outside 'MathTypeLib.exe')
 $si=[Diagnostics.ProcessStartInfo]::new();$si.FileName=$InstallerPath;$si.UseShellExecute=$false
 # /D consumes the remainder of the command line. No trailing space or quotes.
 $si.Arguments='/VISUALTEXACCEPTANCE /VISUALTEXOFFICE=skip /VISUALTEXOCR=none /D='+$install
 $script:setup=[Diagnostics.Process]::Start($si)
 $report.setupPid=$script:setup.Id
 Wait-For { Has-Text 'Welcome to VisualTeX Setup' } 'welcome page'
 Click-Id 1
 Wait-For { Has-Text 'VisualTeX only' } 'Office options'
 Click-Text 'VisualTeX only';Click-Id 1
 Wait-For { Has-Text $prompt } 'runtime closure consent'
 Assert-True ((Is-Running $inside32) -and (Is-Running $inside64)) 'Runtime was stopped before consent.'
 $report.checks+= 'Consent displayed; private x86/x64 probes still running.'
 Click-Id 7
 Wait-For { (Has-Text 'VisualTeX only') -and -not (Has-Text $prompt) } 'No returns to Office options'
 Assert-True ((Is-Running $inside32) -and (Is-Running $inside64) -and (Is-Running $outside32) -and (Is-Running $outside64)) 'No affected a probe.'
 $report.checks+= 'No: all four probes survived; no payload extracted.'
 Click-Id 1;Wait-For { Has-Text $prompt } 'second consent prompt';Click-Id 6
 Wait-For { Has-Text 'Choose Install Location' } 'continue after Yes'
 Assert-True ((-not (Is-Running $inside32)) -and (-not (Is-Running $inside64))) 'Yes did not stop both private probes.'
 Assert-True ((Is-Running $outside32) -and (Is-Running $outside64)) 'A same-name external probe was stopped.'
 $report.checks+= 'Yes: both private probes stopped; both outside-prefix probes survived.'
 Click-Id 3;Wait-For { Has-Text 'VisualTeX only' } 'back to Office options';Click-Id 1
 Wait-For { Has-Text 'Choose Install Location' } 'no false positive with only external MathType names'
 Assert-True (-not (Has-Text $prompt)) 'External processes caused a private-runtime prompt.'
 $report.checks+= 'Only external same-name processes: no false-positive prompt.'
 # Recreate the private helpers after the early check. The pre-extraction guard
 # must ask again. Choose No there, so no product files or registry are written.
 $late32=Start-Probe (Join-Path $private 'MathType.exe');$late64=Start-Probe (Join-Path $private 'MathTypeLib.exe')
 Click-Id 1;Wait-For { Has-Text 'Install offline OCR resources' } 'OCR option page'
 # The custom-page controls are enumerated before nsDialogs finishes entering
 # its modal loop. Let that loop settle before posting the Next button click.
 Start-Sleep -Milliseconds 700
 Click-Id 1
 Start-Sleep -Milliseconds 700
 if ((Has-Text 'Install offline OCR resources') -and -not (Has-Text $prompt)) { Click-Id 1 }
 Wait-For { Has-Text $prompt } 'late-runtime preinstall recheck'
 Assert-True ((Is-Running $late32) -and (Is-Running $late64)) 'Late helpers were stopped without consent.'
 Click-Id 7
 Start-Sleep -Milliseconds 500
 Assert-True (-not (Test-Path -LiteralPath (Join-Path $install 'visualtex.exe'))) 'Declining preinstall guard still extracted the app.'
 Assert-True ((Is-Running $late32) -and (Is-Running $late64) -and (Is-Running $outside32) -and (Is-Running $outside64)) 'Declining late closure affected processes.'
 $report.checks+= 'Restart after early check: preinstall asks again; No leaves processes and payload untouched.'
 Assert-True ((Registry-Snapshot) -ceq $beforeRegistry) 'MathType registration changed.'
 Assert-True ((@((Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object {$_.Id}) | Sort-Object) -join ',') -eq (($beforeWord | Sort-Object) -join ',')) 'Word process list changed.'
 $report.checks+= 'MathType COM/App Paths registry and Word process list unchanged.'
 $report.passed=$true
} catch {
 $report.error=$_.Exception.ToString()
 if($script:setup -and -not $script:setup.HasExited){$report.ui=@([RuntimeInstallerUI]::Read($script:setup.Id))}
} finally {
 # Only this script's known test processes are closed. Files remain as evidence.
 if($script:setup){try{if(-not $script:setup.HasExited){$script:setup.Kill();[void]$script:setup.WaitForExit(2000)}}catch{}}
 foreach($p in $script:ownedProbes){try{if(-not $p.HasExited){$p.Kill();[void]$p.WaitForExit(1000)}}catch{}finally{$p.Dispose()}}
 $report.finished=[DateTimeOffset]::Now.ToString('o')
 [IO.File]::WriteAllText($reportPath,($report | ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($false))
 $report | ConvertTo-Json -Depth 6
}
if(-not $report.passed){exit 1}
