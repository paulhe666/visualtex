[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$DocumentPath,
    [int]$EquationIndex = 1,
    [string]$ProbeText = "q",
    [switch]$KeepWordOpen
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $DocumentPath)) {
    throw "Word probe document does not exist: $DocumentPath"
}
$resolvedDocumentPath = (Resolve-Path -LiteralPath $DocumentPath).Path

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Threading;

public static class VisualTeXMouseProbe
{
    private const uint MouseEventLeftDown = 0x0002;
    private const uint MouseEventLeftUp = 0x0004;

    [DllImport("user32.dll")]
    private static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll")]
    private static extern void mouse_event(
        uint flags,
        uint dx,
        uint dy,
        uint data,
        UIntPtr extraInfo);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr windowHandle);

    public static void DoubleClick(int x, int y)
    {
        if (!SetCursorPos(x, y))
            throw new InvalidOperationException("Unable to move the real mouse cursor.");

        mouse_event(MouseEventLeftDown, 0, 0, 0, UIntPtr.Zero);
        mouse_event(MouseEventLeftUp, 0, 0, 0, UIntPtr.Zero);
        Thread.Sleep(90);
        mouse_event(MouseEventLeftDown, 0, 0, 0, UIntPtr.Zero);
        mouse_event(MouseEventLeftUp, 0, 0, 0, UIntPtr.Zero);
    }
}
"@

function Release-ComObject([object]$Value) {
    if ($null -eq $Value) { return }
    try {
        if ([Runtime.InteropServices.Marshal]::IsComObject($Value)) {
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)
        }
    }
    catch { }
}

$word = $null
$document = $null
$window = $null
$maths = $null
$math = $null
$equationRange = $null
$selection = $null
$beforeSelectionMaths = $null
$afterSelectionMaths = $null
$typedSelectionMaths = $null

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $word.DisplayAlerts = 0
    $document = $word.Documents.Open($resolvedDocumentPath, $false, $false)
    $window = $word.ActiveWindow
    $window.View.Zoom.Percentage = 160

    $maths = $document.OMaths
    if ($EquationIndex -lt 1 -or $EquationIndex -gt $maths.Count) {
        throw "Equation index $EquationIndex is outside the document OMath range 1..$($maths.Count)."
    }

    $math = $maths.Item($EquationIndex)
    $equationRange = $math.Range
    $equationRange.Select()
    Start-Sleep -Milliseconds 350

    $left = 0
    $top = 0
    $width = 0
    $height = 0
    $window.GetPoint(
        [ref]$left,
        [ref]$top,
        [ref]$width,
        [ref]$height,
        $equationRange)
    if ($width -le 0 -or $height -le 0) {
        throw "Word returned an invalid equation screen rectangle: $left,$top,$width,$height"
    }

    $selection = $word.Selection
    $beforeSelectionMaths = $selection.OMaths
    $before = [ordered]@{
        Start = $selection.Start
        End = $selection.End
        OMaths = $beforeSelectionMaths.Count
        ScreenLeft = $left
        ScreenTop = $top
        ScreenWidth = $width
        ScreenHeight = $height
    }
    Write-Host ("BEFORE " + ($before | ConvertTo-Json -Compress))

    [void][VisualTeXMouseProbe]::SetForegroundWindow([IntPtr]$window.Hwnd)
    Start-Sleep -Milliseconds 350
    [VisualTeXMouseProbe]::DoubleClick(
        [int]($left + ($width / 2)),
        [int]($top + ($height / 2)))
    Start-Sleep -Milliseconds 800

    Release-ComObject $afterSelectionMaths
    Release-ComObject $selection
    $selection = $word.Selection
    $afterSelectionMaths = $selection.OMaths
    $after = [ordered]@{
        Start = $selection.Start
        End = $selection.End
        OMaths = $afterSelectionMaths.Count
        Text = (($selection.Text -replace "[\r\n]", "|") -replace "\u0007", "<cell>")
    }
    Write-Host ("AFTER " + ($after | ConvertTo-Json -Compress))

    $beforeXml = $math.Range.WordOpenXML
    $selection.TypeText($ProbeText)
    Start-Sleep -Milliseconds 350
    $afterXml = $math.Range.WordOpenXML

    Release-ComObject $typedSelectionMaths
    Release-ComObject $selection
    $selection = $word.Selection
    $typedSelectionMaths = $selection.OMaths
    $typed = [ordered]@{
        Changed = ($beforeXml -ne $afterXml)
        Start = $selection.Start
        End = $selection.End
        OMaths = $typedSelectionMaths.Count
    }
    Write-Host ("TYPED " + ($typed | ConvertTo-Json -Compress))

    if ($after.OMaths -ne 1) {
        throw "Real double-click did not leave the Word selection inside one OMath."
    }
    if (-not $typed.Changed) {
        throw "Typing after the real double-click did not modify the OMML equation."
    }

    [void]$document.Undo(1)
    Start-Sleep -Milliseconds 250
    Write-Host "VisualTeX Word real-mouse OMML double-click probe passed."
}
finally {
    Release-ComObject $typedSelectionMaths
    Release-ComObject $afterSelectionMaths
    Release-ComObject $beforeSelectionMaths
    Release-ComObject $selection
    Release-ComObject $equationRange
    Release-ComObject $math
    Release-ComObject $maths
    if (-not $KeepWordOpen) {
        if ($null -ne $document) {
            try { $document.Close(0) } catch { }
        }
        if ($null -ne $word) {
            try { $word.Quit() } catch { }
        }
    }
    Release-ComObject $window
    Release-ComObject $document
    Release-ComObject $word
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
