param([Parameter(Mandatory=$true)][int]$WordProcessId,[Parameter(Mandatory=$true)][string]$ExpectedDocument,[Parameter(Mandatory=$true)][string]$Label)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
Add-Type -AssemblyName System.Drawing,System.Windows.Forms
Add-Type @'
using System;using System.Runtime.InteropServices;
public static class RealDocScreen {
 [StructLayout(LayoutKind.Sequential)] public struct Rect{public int Left,Top,Right,Bottom;}
 [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h,out Rect r);
 [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h,IntPtr after,int x,int y,int w,int ht,uint flags);
 [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h,int command);
 [DllImport("user32.dll")] public static extern bool SetProcessDPIAware();
}
'@
[void][RealDocScreen]::SetProcessDPIAware()
$a=Connect-RunningWord $WordProcessId
$d=$a.ActiveDocument
$root=Join-Path $PSScriptRoot 'evidence\real-mathtype-20260926'
if($d.Name -ne $ExpectedDocument -or !$d.FullName.StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'Only an owned evidence document may be positioned for screenshots'}
$h=[IntPtr]$d.ActiveWindow.Hwnd
$old=[RealDocScreen+Rect]::new();[void][RealDocScreen]::GetWindowRect($h,[ref]$old)
try {
 [void][RealDocScreen]::ShowWindow($h,9)
 [void][RealDocScreen]::SetWindowPos($h,[IntPtr](-1),40,40,1850,1800,0x0010)
 Start-Sleep -Milliseconds 700
 $r=[RealDocScreen+Rect]::new();[void][RealDocScreen]::GetWindowRect($h,[ref]$r)
 $bmp=[Drawing.Bitmap]::new(($r.Right-$r.Left),($r.Bottom-$r.Top))
 $g=[Drawing.Graphics]::FromImage($bmp)
 try{
  $g.CopyFromScreen($r.Left,$r.Top,0,0,$bmp.Size)
  $path=Join-Path $root ($Label+'.png')
  $bmp.Save($path,[Drawing.Imaging.ImageFormat]::Png)
  Write-Output ($path+' | '+$bmp.Width+'x'+$bmp.Height)
 }finally{$g.Dispose();$bmp.Dispose()}
}finally{
 [void][RealDocScreen]::SetWindowPos($h,[IntPtr](-2),0,0,0,0,0x0013)
 [void][Runtime.InteropServices.Marshal]::ReleaseComObject($d)
 [void][Runtime.InteropServices.Marshal]::ReleaseComObject($a)
}
