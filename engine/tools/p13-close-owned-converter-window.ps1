param(
    [Parameter(Mandatory=$true)][int]$ProcessId,
    [Parameter(Mandatory=$true)][long]$WindowHandle
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)

$process = Get-CimInstance Win32_Process -Filter ("ProcessId=" + $ProcessId)
if (-not $process) { throw "Owned companion PID is not running: $ProcessId" }
$expectedRoot = 'C:\Users\pojian_liao\Desktop\devspace\visualtex1.2.3-reference-3d207d7\apps\windows\src-tauri\target\debug\'
if (-not $process.ExecutablePath.StartsWith($expectedRoot, [StringComparison]::OrdinalIgnoreCase) -or
    -not $process.ExecutablePath.EndsWith('visualtex.exe', [StringComparison]::OrdinalIgnoreCase) -or
    $process.CommandLine -notmatch '--office-background') {
    throw "Refusing to operate on a non-owned companion process."
}

Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;
public static class P13OwnedWindow {
  [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hWnd);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int maxCount);
  [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
}
'@

$hwnd = [IntPtr]$WindowHandle
if (-not [P13OwnedWindow]::IsWindow($hwnd)) { throw "Converter HWND is no longer valid." }
[uint32]$owner = 0
[void][P13OwnedWindow]::GetWindowThreadProcessId($hwnd, [ref]$owner)
if ([int]$owner -ne $ProcessId) { throw "Converter HWND belongs to another process." }
$text = [Text.StringBuilder]::new(512)
[void][P13OwnedWindow]::GetWindowText($hwnd, $text, $text.Capacity)
$title = $text.ToString()
if (-not $title.StartsWith('VisualTeX', [StringComparison]::Ordinal) -or
    $title.IndexOf('Office', [StringComparison]::Ordinal) -lt 0) {
    throw "Refusing to close unexpected window title: $title"
}
if (-not [P13OwnedWindow]::PostMessage($hwnd, 0x0010, [IntPtr]::Zero, [IntPtr]::Zero)) {
    throw "WM_CLOSE could not be posted."
}
Start-Sleep -Milliseconds 800
[pscustomobject]@{
    pid = $ProcessId
    hwnd = $WindowHandle
    title = $title
    stillExists = [P13OwnedWindow]::IsWindow($hwnd)
}|ConvertTo-Json -Compress
