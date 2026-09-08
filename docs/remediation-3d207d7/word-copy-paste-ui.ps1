param(
    [Parameter(Mandatory=$true)][int]$WordProcessId,
    [Parameter(Mandatory=$true)][string]$DocumentName,
    [ValidateSet('copy-ole','copy-math','paste-end','paste-marker','paste-math-end','undo')][string]$Action,
    [int]$Index = 1,
    [string]$Marker = '',
    [switch]$IncludeHost
)

$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$OutputEncoding=[Console]::OutputEncoding
. (Join-Path $PSScriptRoot 'word-connection.ps1')
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class VisualTeXWordCopyPasteUi {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern uint GetClipboardSequenceNumber();
}
'@

function Release-Com([object]$value) {
    if($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value) } catch { }
    }
}

function Focus-Word([IntPtr]$hwnd) {
    [void][VisualTeXWordCopyPasteUi]::ShowWindow($hwnd, 3)
    for($attempt=0;$attempt -lt 8;$attempt++) {
        [void][VisualTeXWordCopyPasteUi]::SetForegroundWindow($hwnd)
        Start-Sleep -Milliseconds 120
        if([VisualTeXWordCopyPasteUi]::GetForegroundWindow() -eq $hwnd) { return }
    }
    throw "Owned Word PID $WordProcessId did not obtain foreground; no keyboard input was sent."
}

$word=$null
$document=$null
$shape=$null
$math=$null
$range=$null
try {
    $word=Connect-RunningWord $WordProcessId
    $document=$word.ActiveDocument
    if($document.Name -ne $DocumentName) {
        throw "Unexpected active document '$($document.Name)'; expected '$DocumentName'."
    }
    $hwnd=[IntPtr]$word.ActiveWindow.Hwnd

    if($Action -eq 'copy-ole') {
        if($Index -lt 1 -or $Index -gt $document.InlineShapes.Count) {
            throw "OLE index $Index is outside 1..$($document.InlineShapes.Count)."
        }
        $shape=$document.InlineShapes.Item($Index)
        $range=$shape.Range.Duplicate
        if($IncludeHost){$range=$range.Paragraphs.Item(1).Range.Duplicate}
        $range.Select()
        Focus-Word $hwnd
        $before=[VisualTeXWordCopyPasteUi]::GetClipboardSequenceNumber()
        [System.Windows.Forms.SendKeys]::SendWait('^c')
        Start-Sleep -Milliseconds 500
        $after=[VisualTeXWordCopyPasteUi]::GetClipboardSequenceNumber()
        if($after -eq $before) { throw 'Ctrl+C did not change the Windows clipboard sequence.' }
        Write-Output ("REAL_COPY|ole|index={0}|range={1}:{2}|clipboard={3}->{4}" -f $Index,$range.Start,$range.End,$before,$after)
        exit
    }

    if($Action -eq 'copy-math') {
        if($Index -lt 1 -or $Index -gt $document.OMaths.Count) {
            throw "OMath index $Index is outside 1..$($document.OMaths.Count)."
        }
        $math=$document.OMaths.Item($Index)
        $range=$math.Range.Duplicate
        if($IncludeHost){
            if($range.Tables.Count -eq 1){$range=$range.Tables.Item(1).Range.Duplicate}
            else{$range=$range.Paragraphs.Item(1).Range.Duplicate}
        }
        $range.Select()
        Focus-Word $hwnd
        $before=[VisualTeXWordCopyPasteUi]::GetClipboardSequenceNumber()
        [System.Windows.Forms.SendKeys]::SendWait('^c')
        Start-Sleep -Milliseconds 500
        $after=[VisualTeXWordCopyPasteUi]::GetClipboardSequenceNumber()
        if($after -eq $before) { throw 'Ctrl+C did not change the Windows clipboard sequence.' }
        Write-Output ("REAL_COPY|math|index={0}|range={1}:{2}|clipboard={3}->{4}" -f $Index,$range.Start,$range.End,$before,$after)
        exit
    }

    if($Action -eq 'undo') {
        Focus-Word $hwnd
        $beforeOle=$document.InlineShapes.Count
        $beforeMath=$document.OMaths.Count
        $beforeEnd=$document.Content.End
        [System.Windows.Forms.SendKeys]::SendWait('^z')
        Start-Sleep -Milliseconds 700
        Write-Output ("REAL_UNDO|ole={0}->{1}|math={2}->{3}|docEnd={4}->{5}" -f $beforeOle,$document.InlineShapes.Count,$beforeMath,$document.OMaths.Count,$beforeEnd,$document.Content.End)
        exit
    }

    Focus-Word $hwnd
    $beforeOle=$document.InlineShapes.Count
    $beforeMath=$document.OMaths.Count
    $beforeEnd=$document.Content.End
    # Use the same user path as Ctrl+End -> Enter -> Ctrl+V so the paste always
    # starts in a fresh ordinary body paragraph rather than inside the preceding
    # OMath/OLE host.
    if($Action -eq 'paste-marker') {
        if(!$Marker){throw 'paste-marker requires an explicit fixture marker.'}
        $range=$document.Content.Duplicate
        if(!$range.Find.Execute($Marker+': ')){throw "Missing marker: $Marker"}
        $word.Selection.SetRange($range.End,$range.End)
        Focus-Word $hwnd
        [System.Windows.Forms.SendKeys]::SendWait('^v')
    } elseif($Action -eq 'paste-math-end') {
        if($Index -lt 1 -or $Index -gt $document.OMaths.Count){throw 'Invalid target OMath index'}
        $math=$document.OMaths.Item($Index)
        $range=$math.Range.Duplicate
        $word.Selection.SetRange($range.End,$range.End)
        Focus-Word $hwnd
        [System.Windows.Forms.SendKeys]::SendWait('^v')
    } else {
        [System.Windows.Forms.SendKeys]::SendWait('^{END}{ENTER}^v')
    }
    Start-Sleep -Milliseconds 1000
    $afterOle=$document.InlineShapes.Count
    $afterMath=$document.OMaths.Count
    $afterEnd=$document.Content.End
    if($afterOle -eq $beforeOle -and $afterMath -eq $beforeMath -and $afterEnd -eq $beforeEnd) {
        throw 'Ctrl+V did not change the owned Word document.'
    }
    Write-Output ("REAL_PASTE|new-paragraph|ole={0}->{1}|math={2}->{3}|docEnd={4}->{5}|selection={6}:{7}" -f $beforeOle,$afterOle,$beforeMath,$afterMath,$beforeEnd,$afterEnd,$word.Selection.Start,$word.Selection.End)
}
finally {
    Release-Com $range
    Release-Com $math
    Release-Com $shape
    Release-Com $document
    Release-Com $word
}
