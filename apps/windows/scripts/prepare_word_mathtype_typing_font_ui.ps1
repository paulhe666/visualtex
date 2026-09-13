param(
    [double]$BodyFontSizePt = 12.0,
    [string]$Prefix = "ABC"
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class VisualTeXWordFontPrepNative {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
}
"@

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
    throw "Timed out locating UI element '$name'."
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

function Select-VisualTeXRibbon([IntPtr]$wordHwnd) {
    [void][VisualTeXWordFontPrepNative]::SetForegroundWindow($wordHwnd)
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

$preexistingWordPids = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$word.Visible = $true
$word.DisplayAlerts = 0
$document = $word.Documents.Add()
$tempPath = Join-Path $env:TEMP ("visualtex-font-repro-" + [guid]::NewGuid().ToString("N") + ".docx")
$document.SaveAs2($tempPath, 12)
$document.Activate()
$wordHwnd = [IntPtr]$word.ActiveWindow.Hwnd
[uint32]$pidValue = 0
[void][VisualTeXWordFontPrepNative]::GetWindowThreadProcessId($wordHwnd, [ref]$pidValue)
$wordPid = [int]$pidValue
if ($wordPid -eq 0 -or $preexistingWordPids -contains $wordPid) {
    try { $document.Close(0) } catch { }
    try { $word.Quit(0) } catch { }
    throw "Safety guard: scratch Word did not receive a distinct WINWORD process."
}

$selection = $word.Selection
$selection.Font.Size = [single]$BodyFontSizePt
[void][VisualTeXWordFontPrepNative]::SetForegroundWindow($wordHwnd)
Start-Sleep -Milliseconds 250
[System.Windows.Forms.SendKeys]::SendWait($Prefix)
Start-Sleep -Milliseconds 250
$prefixEnd = $word.Selection.Start
$prefixRange = $document.Range([Math]::Max(0, $prefixEnd - $Prefix.Length), $prefixEnd)
$prefixSize = $prefixRange.Font.Size
$prefixText = $prefixRange.Text
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($prefixRange)

$addIns = $word.COMAddIns
$addIn = $addIns.Item("VisualTeX.WordVsto")
if (-not $addIn.Connect) {
    $addIn.Connect = $true
    Start-Sleep -Milliseconds 800
}
if (-not $addIn.Connect) { throw "Installed VisualTeX.WordVsto add-in is not connected." }

$wordRoot = Select-VisualTeXRibbon $wordHwnd
$inlineButtonName = "OLE " + [char]0x884C + [char]0x5185 + [char]0x516C + [char]0x5F0F
$inlineButton = Wait-ElementExact $wordRoot $inlineButtonName ([System.Windows.Automation.ControlType]::Button) ([TimeSpan]::FromSeconds(10))
Invoke-Element $inlineButton "Word VisualTeX OLE inline formula button"
Start-Sleep -Milliseconds 800

Write-Output ("SCRATCH_WORD_PID=" + $wordPid)
Write-Output ("SCRATCH_WORD_HWND=" + [int64]$wordHwnd)
Write-Output ("SCRATCH_DOCUMENT=" + $tempPath)
Write-Output ("PREEXISTING_WORD_PIDS=" + ($preexistingWordPids -join ","))
Write-Output ("BEFORE_INSERT_PREFIX_TEXT=" + $prefixText)
Write-Output ("BEFORE_INSERT_PREFIX_SIZE=" + $prefixSize)
Write-Output ("BEFORE_INSERT_SELECTION_SIZE=" + $word.Selection.Font.Size)
Write-Output "RIBBON_INLINE_INVOKED=1"

# Intentionally leave this scratch Word document open for the next stage.
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($addIn)
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($addIns)
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($document)
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
[GC]::Collect()
[GC]::WaitForPendingFinalizers()
