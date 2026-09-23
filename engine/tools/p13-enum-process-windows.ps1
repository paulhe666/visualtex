param([Parameter(Mandatory=$true)][int]$ProcessId)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
Add-Type @'
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class P13WindowEnum {
  public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr hWnd, StringBuilder text, int count);
}
'@
$rows=New-Object System.Collections.Generic.List[object]
$callback=[P13WindowEnum+EnumWindowsProc]{
  param([IntPtr]$h,[IntPtr]$l)
  [uint32]$ownerPid=0
  [void][P13WindowEnum]::GetWindowThreadProcessId($h,[ref]$ownerPid)
  if($ownerPid -eq [uint32]$ProcessId){
    $title=[Text.StringBuilder]::new(1024);[void][P13WindowEnum]::GetWindowText($h,$title,$title.Capacity)
    $class=[Text.StringBuilder]::new(512);[void][P13WindowEnum]::GetClassName($h,$class,$class.Capacity)
    $rows.Add([pscustomobject]@{handle=$h.ToInt64();visible=[P13WindowEnum]::IsWindowVisible($h);class=$class.ToString();title=$title.ToString()})
  }
  return $true
}
[void][P13WindowEnum]::EnumWindows($callback,[IntPtr]::Zero)
[pscustomobject]@{pid=$ProcessId;count=$rows.Count;windows=$rows}|ConvertTo-Json -Depth 4 -Compress
