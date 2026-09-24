param(
    [double]$BodyFontSizePt = 12.0,
    [string]$Prefix = "ABC",
    [string]$Suffix = "XYZ"
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class VisualTeXFontReproNative {
    public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int count);
    [DllImport("user32.dll")] public static extern int GetWindowTextLength(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hwnd);
}
"@

function Release-ComObject([object]$value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value) } catch { }
    }
}

function Find-ElementExact(
    [System.Windows.Automation.AutomationElement]$root,
    [string]$name,
    [System.Windows.Automation.ControlType]$controlType
) {
    $nameCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::NameProperty,
        $name)
    $typeCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        $controlType)
    return $root.FindFirst(
        [System.Windows.Automation.TreeScope]::Descendants,
        (New-Object System.Windows.Automation.AndCondition($nameCondition, $typeCondition)))
}

function Wait-ElementExact(
    [System.Windows.Automation.AutomationElement]$root,
    [string]$name,
    [System.Windows.Automation.ControlType]$controlType,
    [TimeSpan]$timeout
) {
    $deadline = [DateTime]::UtcNow + $timeout
    do {
        $element = Find-ElementExact $root $name $controlType
        if ($null -ne $element) { return $element }
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "Timed out locating UI element '$name' ($($controlType.ProgrammaticName))."
}

function Invoke-Element([System.Windows.Automation.AutomationElement]$element, [string]$description) {
    if ($null -eq $element) { throw "UI element not found: $description" }
    $patternObject = $null
    if (-not $element.TryGetCurrentPattern(
        [System.Windows.Automation.InvokePattern]::Pattern,
        [ref]$patternObject)) {
        throw "UI element does not expose InvokePattern: $description"
    }
    ([System.Windows.Automation.InvokePattern]$patternObject).Invoke()
}

function Text-FromCodePoints([int[]]$codePoints) {
    $builder = New-Object Text.StringBuilder
    foreach ($codePoint in $codePoints) { [void]$builder.Append([char]$codePoint) }
    return $builder.ToString()
}

function Select-VisualTeXRibbon([IntPtr]$wordHwnd) {
    [void][VisualTeXFontReproNative]::SetForegroundWindow($wordHwnd)
    Start-Sleep -Milliseconds 350
    $root = [System.Windows.Automation.AutomationElement]::FromHandle($wordHwnd)
    $tab = Wait-ElementExact $root "VisualTeX" ([System.Windows.Automation.ControlType]::TabItem) ([TimeSpan]::FromSeconds(15))
    $selectionObject = $null
    if (-not $tab.TryGetCurrentPattern(
        [System.Windows.Automation.SelectionItemPattern]::Pattern,
        [ref]$selectionObject)) {
        throw "VisualTeX Ribbon tab has no SelectionItemPattern."
    }
    ([System.Windows.Automation.SelectionItemPattern]$selectionObject).Select()
    Start-Sleep -Milliseconds 450
    return $root
}

function Snapshot-OfficeEditorHandles() {
    $set = @{}
    [VisualTeXFontReproNative]::EnumWindows({
        param([IntPtr]$hwnd, [IntPtr]$unused)
        if (-not [VisualTeXFontReproNative]::IsWindowVisible($hwnd)) { return $true }
        $length = [VisualTeXFontReproNative]::GetWindowTextLength($hwnd)
        if ($length -le 0) { return $true }
        $text = New-Object Text.StringBuilder ($length + 1)
        [void][VisualTeXFontReproNative]::GetWindowText($hwnd, $text, $text.Capacity)
        if ($text.ToString() -notmatch "Office") { return $true }
        [uint32]$pidValue = 0
        [void][VisualTeXFontReproNative]::GetWindowThreadProcessId($hwnd, [ref]$pidValue)
        if ($pidValue -eq 0) { return $true }
        try {
            $process = Get-Process -Id $pidValue -ErrorAction Stop
            if ($process.ProcessName -ne "visualtex") { return $true }
        } catch { return $true }
        $set[[int64]$hwnd] = $true
        return $true
    }, [IntPtr]::Zero) | Out-Null
    return $set
}

function Wait-NewOfficeEditorWindow([hashtable]$before, [TimeSpan]$timeout) {
    $deadline = [DateTime]::UtcNow + $timeout
    do {
        $script:found = [IntPtr]::Zero
        [VisualTeXFontReproNative]::EnumWindows({
            param([IntPtr]$hwnd, [IntPtr]$unused)
            if (-not [VisualTeXFontReproNative]::IsWindowVisible($hwnd)) { return $true }
            if ($before.ContainsKey([int64]$hwnd)) { return $true }
            $length = [VisualTeXFontReproNative]::GetWindowTextLength($hwnd)
            if ($length -le 0) { return $true }
            $text = New-Object Text.StringBuilder ($length + 1)
            [void][VisualTeXFontReproNative]::GetWindowText($hwnd, $text, $text.Capacity)
            if ($text.ToString() -notmatch "Office") { return $true }
            [uint32]$pidValue = 0
            [void][VisualTeXFontReproNative]::GetWindowThreadProcessId($hwnd, [ref]$pidValue)
            if ($pidValue -eq 0) { return $true }
            try {
                $process = Get-Process -Id $pidValue -ErrorAction Stop
                if ($process.ProcessName -ne "visualtex") { return $true }
            } catch { return $true }
            $script:found = $hwnd
            return $false
        }, [IntPtr]::Zero) | Out-Null
        if ($script:found -ne [IntPtr]::Zero) { return $script:found }
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "A new VisualTeX Office editor window did not appear."
}

function Select-MathTypeOle([System.Windows.Automation.AutomationElement]$editorRoot) {
    $objectFormatName = Text-FromCodePoints @(0x516C,0x5F0F,0x5BF9,0x8C61,0x683C,0x5F0F)
    $combo = Wait-ElementExact $editorRoot $objectFormatName ([System.Windows.Automation.ControlType]::ComboBox) ([TimeSpan]::FromSeconds(8))
    [void]$combo.SetFocus()
    $expandObject = $null
    if ($combo.TryGetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$expandObject)) {
        $expand = [System.Windows.Automation.ExpandCollapsePattern]$expandObject
        if ($expand.Current.ExpandCollapseState -ne [System.Windows.Automation.ExpandCollapseState]::Expanded) {
            $expand.Expand()
            Start-Sleep -Milliseconds 180
        }
    }
    [System.Windows.Forms.SendKeys]::SendWait("{HOME}{DOWN}")
    Start-Sleep -Milliseconds 300
    Write-Output "UI_SELECT|formula object format|MathType OLE"
}

$preexistingWordPids = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = $null
$document = $null
$addIns = $null
$addIn = $null
$shape = $null
$shapeRange = $null
$typedRange = $null
$tempPath = Join-Path $env:TEMP ("visualtex-font-repro-" + [guid]::NewGuid().ToString("N") + ".docx")
$newWordPid = 0
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $word.DisplayAlerts = 0
    $document = $word.Documents.Add()
    $document.SaveAs2($tempPath, 12)
    $document.Activate()
    $wordHwnd = [IntPtr]$word.ActiveWindow.Hwnd
    [uint32]$pidValue = 0
    [void][VisualTeXFontReproNative]::GetWindowThreadProcessId($wordHwnd, [ref]$pidValue)
    $newWordPid = [int]$pidValue
    if ($newWordPid -eq 0 -or $preexistingWordPids -contains $newWordPid) {
        throw "Safety guard: scratch Word did not receive a distinct WINWORD process."
    }
    Write-Output ("SCRATCH_WORD_PID=" + $newWordPid)
    Write-Output ("PREEXISTING_WORD_PIDS=" + ($preexistingWordPids -join ","))

    $selection = $word.Selection
    $selection.Font.Size = [single]$BodyFontSizePt
    [void][VisualTeXFontReproNative]::SetForegroundWindow($wordHwnd)
    Start-Sleep -Milliseconds 250
    [System.Windows.Forms.SendKeys]::SendWait($Prefix)
    Start-Sleep -Milliseconds 250
    $prefixEnd = $word.Selection.Start
    $prefixRange = $document.Range([Math]::Max(0, $prefixEnd - $Prefix.Length), $prefixEnd)
    try {
        Write-Output ("BEFORE_INSERT|selection="+$word.Selection.Start+"|selectionSize="+$word.Selection.Font.Size+"|prefixSize="+$prefixRange.Font.Size+"|prefixText="+$prefixRange.Text)
    } finally { Release-ComObject $prefixRange }

    $addIns = $word.COMAddIns
    $addIn = $addIns.Item("VisualTeX.WordVsto")
    if (-not $addIn.Connect) {
        $addIn.Connect = $true
        Start-Sleep -Milliseconds 800
    }
    if (-not $addIn.Connect) { throw "Installed VisualTeX.WordVsto add-in is not connected." }

    $beforeEditorHandles = Snapshot-OfficeEditorHandles
    $wordRoot = Select-VisualTeXRibbon $wordHwnd
    $inlineButtonName = "OLE " + (Text-FromCodePoints @(0x884C,0x5185,0x516C,0x5F0F))
    $inlineButton = Wait-ElementExact $wordRoot $inlineButtonName ([System.Windows.Automation.ControlType]::Button) ([TimeSpan]::FromSeconds(10))
    Invoke-Element $inlineButton "Word VisualTeX OLE inline formula button"
    Write-Output "UI_INVOKE|VisualTeX Ribbon|OLE inline formula"

    $editorHwnd = Wait-NewOfficeEditorWindow $beforeEditorHandles ([TimeSpan]::FromSeconds(20))
    [void][VisualTeXFontReproNative]::SetForegroundWindow($editorHwnd)
    Start-Sleep -Milliseconds 700
    $editorRoot = [System.Windows.Automation.AutomationElement]::FromHandle($editorHwnd)
    Select-MathTypeOle $editorRoot

    Start-Sleep -Milliseconds 350
    $editorRoot = [System.Windows.Automation.AutomationElement]::FromHandle($editorHwnd)
    $presetName = Text-FromCodePoints @(0x52FE,0x80A1,0x5B9A,0x7406)
    $preset = Wait-ElementExact $editorRoot $presetName ([System.Windows.Automation.ControlType]::Button) ([TimeSpan]::FromSeconds(8))
    Invoke-Element $preset "Pythagorean theorem preset"
    Write-Output "UI_INVOKE|preset|Pythagorean theorem"

    Start-Sleep -Milliseconds 500
    $finishName = Text-FromCodePoints @(0x5B8C,0x6210,0x5E76,0x63D2,0x5165)
    $finish = Wait-ElementExact $editorRoot $finishName ([System.Windows.Automation.ControlType]::Button) ([TimeSpan]::FromSeconds(8))
    Invoke-Element $finish "Finish and insert"
    Write-Output "UI_INVOKE|finish-and-insert"

    $deadline = [DateTime]::UtcNow.AddSeconds(45)
    do {
        Start-Sleep -Milliseconds 180
        if ($document.InlineShapes.Count -ge 1) { break }
    } while ([DateTime]::UtcNow -lt $deadline)
    if ($document.InlineShapes.Count -ne 1) {
        throw "Expected exactly one MathType InlineShape after real UI insertion; got $($document.InlineShapes.Count)."
    }

    Start-Sleep -Milliseconds 600
    $shape = $document.InlineShapes.Item(1)
    $shapeRange = $shape.Range
    $progId = $shape.OLEFormat.ProgID
    $caretBeforeTyping = $word.Selection.Start
    $selectionSizeBeforeTyping = $word.Selection.Font.Size
    Write-Output ("AFTER_INSERT|progId="+$progId+"|shapeRange="+$shapeRange.Start+":"+$shapeRange.End+"|shapeSize="+$shapeRange.Font.Size+"|selection="+$caretBeforeTyping+"|selectionSize="+$selectionSizeBeforeTyping)

    [void][VisualTeXFontReproNative]::SetForegroundWindow($wordHwnd)
    Start-Sleep -Milliseconds 350
    [System.Windows.Forms.SendKeys]::SendWait($Suffix)
    Start-Sleep -Milliseconds 450
    $caretAfterTyping = $word.Selection.Start
    $typedStart = [Math]::Min($caretBeforeTyping, $caretAfterTyping)
    $typedEnd = [Math]::Max($caretBeforeTyping, $caretAfterTyping)
    $typedRange = $document.Range($typedStart, $typedEnd)
    $paragraphMark = $document.Range($document.Paragraphs.Item(1).Range.End - 1, $document.Paragraphs.Item(1).Range.End)
    try {
        Write-Output ("AFTER_TYPING|typedRange="+$typedStart+":"+$typedEnd+"|typedText="+$typedRange.Text+"|typedSize="+$typedRange.Font.Size+"|selectionSize="+$word.Selection.Font.Size+"|paragraphMarkSize="+$paragraphMark.Font.Size)
    } finally { Release-ComObject $paragraphMark }

    if ($progId -notlike "Equation.DSMT4*") { throw "Inserted object is not MathType Equation.DSMT4: $progId" }
    if ([Math]::Abs([double]$typedRange.Font.Size - $BodyFontSizePt) -gt 0.1) {
        Write-Output ("REPRODUCED|expectedBodySize="+$BodyFontSizePt+"|actualTypedSize="+$typedRange.Font.Size+"|shapeSize="+$shapeRange.Font.Size)
        exit 23
    }
    Write-Output ("NOT_REPRODUCED|bodySizePreserved="+$typedRange.Font.Size)
}
finally {
    Release-ComObject $typedRange
    Release-ComObject $shapeRange
    Release-ComObject $shape
    Release-ComObject $addIn
    Release-ComObject $addIns
    if ($null -ne $document) {
        try { $document.Close(0) } catch { }
    }
    if ($null -ne $word) {
        try { $word.Quit(0) } catch { }
    }
    Release-ComObject $document
    Release-ComObject $word
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if (Test-Path -LiteralPath $tempPath) {
        try { Remove-Item -LiteralPath $tempPath -Force } catch { }
    }
}
