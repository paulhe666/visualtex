param(
    [Parameter(Mandatory = $true)]
    [string]$DocumentPath,
    [int]$ShapeIndex = 1
)

$ErrorActionPreference = "Stop"
$resolvedPath = (Resolve-Path $DocumentPath).Path
$sessionRoot = Join-Path ([Environment]::GetFolderPath("ApplicationData")) "com.visualtex.studio\office\sessions"
$tracePath = Join-Path $env:TEMP ("VisualTeX-Word-Hook-" + [Guid]::NewGuid().ToString("N") + ".log")
[Environment]::SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", $tracePath, "Process")

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class VisualTeXNativeMouse {
    [StructLayout(LayoutKind.Sequential)]
    public struct POINT { public int X; public int Y; }
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extraInfo);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hwnd, IntPtr insertAfter, int x, int y, int cx, int cy, uint flags);
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);
    [DllImport("user32.dll")] public static extern IntPtr GetDC(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern int ReleaseDC(IntPtr hwnd, IntPtr dc);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hwnd, int command);
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("gdi32.dll")] public static extern int GetDeviceCaps(IntPtr dc, int index);
    public const uint LEFTDOWN = 0x0002;
    public const uint LEFTUP = 0x0004;
    public const int HORZRES = 8;
    public const int VERTRES = 10;
    public const int DESKTOPVERTRES = 117;
    public const int DESKTOPHORZRES = 118;
}
"@

function Release-ComObject([object]$value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value)
    }
}

function Get-SessionIds {
    if (-not (Test-Path $sessionRoot)) { return @() }
    return @(Get-ChildItem $sessionRoot -Directory | ForEach-Object Name)
}

function Close-VisualTeXEditor([string]$sessionId) {
    $installCandidates = @(
        (Join-Path ([Environment]::GetFolderPath("ApplicationData")) "com.visualtex.studio\office\install.json"),
        (Join-Path ([Environment]::GetFolderPath("LocalApplicationData")) "com.visualtex.studio\office\install.json")
    )
    $installPath = $installCandidates | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not $installPath) { throw "VisualTeX install.json was not found." }
    $install = Get-Content $installPath -Raw | ConvertFrom-Json
    if (-not $install.installToken) { throw "VisualTeX install token is missing." }
    [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    $headers = @{ "X-VisualTeX-Install-Token" = [string]$install.installToken }
    Invoke-RestMethod -Method Post -Uri "https://127.0.0.1:43127/api/v1/app/sessions/$sessionId/close" -Headers $headers -ContentType "application/json" -Body "{}" | Out-Null
}

$word = $null
$document = $null
$shape = $null
$range = $null
$window = $null
$oleFormat = $null
$comAddIns = $null
$addIn = $null
$sessionId = $null
$consoleWindow = [VisualTeXNativeMouse]::GetConsoleWindow()
try {
    $before = @(Get-SessionIds)
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($resolvedPath, $false, $false)
    Start-Sleep -Milliseconds 800

    $comAddIns = $word.COMAddIns
    $addIn = $comAddIns.Item("VisualTeX.WordVsto")
    if (-not $addIn.Connect) {
        $addIn.Connect = $true
        Start-Sleep -Milliseconds 800
    }
    if (-not $addIn.Connect) { throw "VisualTeX.WordVsto is not connected." }

    if ($ShapeIndex -lt 1 -or $ShapeIndex -gt $document.InlineShapes.Count) {
        throw "Inline shape index $ShapeIndex is outside 1..$($document.InlineShapes.Count)."
    }
    $shape = $document.InlineShapes.Item($ShapeIndex)
    $oleFormat = $shape.OLEFormat
    $progId = [string]$oleFormat.ProgID
    if ($progId -ne "VisualTeX.Formula" -and $progId -ne "VisualTeX.Formula.1") {
        throw "Inline shape $ShapeIndex is not a VisualTeX OLE object; ProgID=$progId"
    }
    $range = $shape.Range
    $range.Select()
    $word.ActiveWindow.Activate()
    Start-Sleep -Milliseconds 800
    $window = $word.ActiveWindow
    $left = 0; $top = 0; $width = 0; $height = 0
    $window.GetPoint([ref]$left, [ref]$top, [ref]$width, [ref]$height, $range)
    if ($width -le 0 -or $height -le 0) {
        throw "Word did not return a visible screen rectangle for the OLE formula."
    }
    $dc = [VisualTeXNativeMouse]::GetDC([IntPtr]::Zero)
    try {
        $logicalWidth = [VisualTeXNativeMouse]::GetDeviceCaps($dc, [VisualTeXNativeMouse]::HORZRES)
        $logicalHeight = [VisualTeXNativeMouse]::GetDeviceCaps($dc, [VisualTeXNativeMouse]::VERTRES)
        $desktopWidth = [VisualTeXNativeMouse]::GetDeviceCaps($dc, [VisualTeXNativeMouse]::DESKTOPHORZRES)
        $desktopHeight = [VisualTeXNativeMouse]::GetDeviceCaps($dc, [VisualTeXNativeMouse]::DESKTOPVERTRES)
    }
    finally {
        [void][VisualTeXNativeMouse]::ReleaseDC([IntPtr]::Zero, $dc)
    }
    $scaleX = if ($logicalWidth -gt 0 -and $desktopWidth -gt 0) { $desktopWidth / [double]$logicalWidth } else { 1.0 }
    $scaleY = if ($logicalHeight -gt 0 -and $desktopHeight -gt 0) { $desktopHeight / [double]$logicalHeight } else { 1.0 }

    $wordHwnd = [IntPtr]([int64]$window.Hwnd)
    if ($consoleWindow -ne [IntPtr]::Zero) { [void][VisualTeXNativeMouse]::ShowWindow($consoleWindow, 0) }
    [void][VisualTeXNativeMouse]::SetWindowPos($wordHwnd, [IntPtr](-1), 0, 0, 0, 0, 0x0043)
    [void][VisualTeXNativeMouse]::SetForegroundWindow($wordHwnd)
    $wordRect = New-Object VisualTeXNativeMouse+RECT
    if ([VisualTeXNativeMouse]::GetWindowRect($wordHwnd, [ref]$wordRect)) {
        $titleX = [int](($wordRect.Left + $wordRect.Right) / 2)
        $titleY = [int]($wordRect.Top + 18)
        [VisualTeXNativeMouse]::SetCursorPos($titleX, $titleY) | Out-Null
        [VisualTeXNativeMouse]::mouse_event([VisualTeXNativeMouse]::LEFTDOWN, 0, 0, 0, [UIntPtr]::Zero)
        [VisualTeXNativeMouse]::mouse_event([VisualTeXNativeMouse]::LEFTUP, 0, 0, 0, [UIntPtr]::Zero)
    }
    Start-Sleep -Milliseconds 600

    $x = [int](($left + $width / 2) / $scaleX)
    $y = [int](($top + $height / 2) / $scaleY)
    Write-Output ("MOUSE_TARGET raw={0},{1},{2},{3} scale={4:0.###}x{5:0.###} physical={6},{7}" -f $left,$top,$width,$height,$scaleX,$scaleY,$x,$y)
    [VisualTeXNativeMouse]::SetCursorPos($x, $y) | Out-Null
    Start-Sleep -Milliseconds 150
    1..2 | ForEach-Object {
        [VisualTeXNativeMouse]::mouse_event([VisualTeXNativeMouse]::LEFTDOWN, 0, 0, 0, [UIntPtr]::Zero)
        [VisualTeXNativeMouse]::mouse_event([VisualTeXNativeMouse]::LEFTUP, 0, 0, 0, [UIntPtr]::Zero)
        Start-Sleep -Milliseconds 90
    }

    $deadline = [DateTime]::UtcNow.AddSeconds(20)
    while ([DateTime]::UtcNow -lt $deadline) {
        $newIds = @(Get-SessionIds | Where-Object { $before -notcontains $_ })
        if ($newIds.Count -gt 0) {
            $sessionId = $newIds |
                ForEach-Object { Get-Item (Join-Path $sessionRoot $_) } |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1 -ExpandProperty Name
            break
        }
        Start-Sleep -Milliseconds 150
    }
    if (-not $sessionId) {
        $trace = if (Test-Path $tracePath) { Get-Content $tracePath -Raw } else { "<no hook trace>" }
        throw "Real OLE double-click did not create a VisualTeX Session.`nHOOK_TRACE:`n$trace"
    }

    $sessionPath = Join-Path (Join-Path $sessionRoot $sessionId) "session.json"
    $session = Get-Content $sessionPath -Raw | ConvertFrom-Json
    $result = [ordered]@{
        SessionId = $sessionId
        Host = $session.host
        Mode = $session.mode
        ObjectMode = $session.objectMode
        ProgId = $progId
        ScreenLeft = $left
        ScreenTop = $top
        ScreenWidth = $width
        ScreenHeight = $height
    }
    Write-Output ("OLE_DOUBLE_CLICK " + ($result | ConvertTo-Json -Compress))
    if ($session.host -ne "word" -or $session.mode -ne "edit" -or $session.objectMode -ne "nativeOle") {
        throw "Real OLE double-click created the wrong Session routing."
    }

    Close-VisualTeXEditor $sessionId
    Write-Output "VisualTeX Word real-mouse OLE double-click interception probe passed."
}
finally {
    if ($consoleWindow -ne [IntPtr]::Zero) { [void][VisualTeXNativeMouse]::ShowWindow($consoleWindow, 5) }
    if ($sessionId) {
        try { Close-VisualTeXEditor $sessionId } catch { }
    }
    Release-ComObject $addIn
    Release-ComObject $comAddIns
    Release-ComObject $oleFormat
    Release-ComObject $window
    Release-ComObject $range
    Release-ComObject $shape
    if ($document) {
        try { $document.Close(0) } catch { }
    }
    Release-ComObject $document
    if ($word) {
        try { $word.Quit(0) } catch { }
    }
    Release-ComObject $word
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [Environment]::SetEnvironmentVariable("VISUALTEX_WORD_HOOK_TRACE_PATH", $null, "Process")
    if (Test-Path $tracePath) {
        Write-Output "HOOK_TRACE_BEGIN"
        Get-Content $tracePath
        Write-Output "HOOK_TRACE_END"
        Remove-Item $tracePath -Force -ErrorAction SilentlyContinue
    }
}
