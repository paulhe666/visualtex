param([string]$ExpectedTitle='', [ValidateSet('list','find','close')][string]$Action='list', [string]$OwnerManifest='')
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=New-Object System.Text.UTF8Encoding($false)
Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class OwnedWindowInventory {
 public delegate bool EnumProc(IntPtr hwnd,IntPtr param);
 [DllImport("user32.dll")]public static extern bool EnumWindows(EnumProc callback,IntPtr param);
 [DllImport("user32.dll")]public static extern bool IsWindowVisible(IntPtr hwnd);
 [DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint pid);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]public static extern int GetWindowText(IntPtr hwnd,StringBuilder title,int count);
 [DllImport("user32.dll")]public static extern bool PostMessage(IntPtr hwnd,uint msg,IntPtr wp,IntPtr lp);
 public static long[] Windows(){var list=new List<long>();EnumWindows((h,p)=>{if(IsWindowVisible(h))list.Add(h.ToInt64());return true;},IntPtr.Zero);return list.ToArray();}
 public static string Title(IntPtr hwnd){var s=new StringBuilder(1024);GetWindowText(hwnd,s,s.Capacity);return s.ToString();}
}
'@
if($Action -eq 'close'){
 $root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../target/owned-ime-review'))+[IO.Path]::DirectorySeparatorChar
 $path=[IO.Path]::GetFullPath($OwnerManifest)
 if(!$path.StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'Window ownership manifest must be test-owned'}
 $owner=Get-Content -LiteralPath $path -Raw|ConvertFrom-Json
 if($owner.schema -ne 'visualtex-test-window-v1' -or $owner.title -notmatch '^VisualTeX-IME-[a-z0-9-]{8,80}$' -or $owner.previousWindows -contains $owner.hwnd){throw 'Invalid window ownership record'}
 $window=[IntPtr][long]$owner.hwnd
 [uint32]$pidValue=0;[void][OwnedWindowInventory]::GetWindowThreadProcessId($window,[ref]$pidValue)
 if($pidValue -ne $owner.pid -or ![OwnedWindowInventory]::Title($window).StartsWith($owner.title,[StringComparison]::Ordinal)){throw 'Window identity changed; it will not be closed'}
 # WM_CLOSE addresses only this newly created top-level window. Never kill the
 # browser process: it also owns the user's pre-existing windows and tabs.
 if(![OwnedWindowInventory]::PostMessage($window,0x10,[IntPtr]::Zero,[IntPtr]::Zero)){throw 'Owned window close request failed'}
 @{closedWindow=$owner.hwnd;processKilled=$false}|ConvertTo-Json;exit
}
$found=@()
foreach($hwnd in [OwnedWindowInventory]::Windows()){
 if($Action -eq 'list'){$found+=$hwnd;continue}
 if($ExpectedTitle -notmatch '^VisualTeX-IME-[a-z0-9-]{8,80}$'){throw 'Exact owned title prefix required'}
 $title=[OwnedWindowInventory]::Title([IntPtr]$hwnd)
 if(!$title.StartsWith($ExpectedTitle,[StringComparison]::Ordinal)){continue}
 [uint32]$pidValue=0;[void][OwnedWindowInventory]::GetWindowThreadProcessId([IntPtr]$hwnd,[ref]$pidValue)
 $process=Get-CimInstance Win32_Process -Filter "ProcessId=$pidValue"
 if($process.Name -notin @('chrome.exe','msedge.exe')){throw 'Owned title is not in a browser window'}
 $found+=@{hwnd=$hwnd;pid=$pidValue;title=$title;executable=$process.ExecutablePath}
}
ConvertTo-Json -InputObject @($found) -Depth 5 -Compress
