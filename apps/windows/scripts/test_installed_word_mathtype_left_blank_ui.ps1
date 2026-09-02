param(
    [int]$TimeoutSeconds = 60
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class VTBlankUiNative {
    public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowTextLength(IntPtr hwnd);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int count);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hwnd, uint msg, UIntPtr wParam, IntPtr lParam);
}
"@

function Release-ComObject([object]$value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value) } catch { }
    }
}

function Text-FromCodePoints([int[]]$points) {
    $builder = New-Object Text.StringBuilder
    foreach ($point in $points) { [void]$builder.Append([char]$point) }
    return $builder.ToString()
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
    [int]$seconds = 15
) {
    $deadline = [DateTime]::UtcNow.AddSeconds($seconds)
    do {
        $element = Find-ElementExact $root $name $controlType
        if ($null -ne $element) { return $element }
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "Timed out locating UI element '$name' ($($controlType.ProgrammaticName))."
}

function Invoke-Element([System.Windows.Automation.AutomationElement]$element, [string]$description) {
    $patternObject = $null
    if (-not $element.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$patternObject)) {
        throw "UI element has no InvokePattern: $description"
    }
    ([System.Windows.Automation.InvokePattern]$patternObject).Invoke()
}

function Select-VisualTeXRibbon([IntPtr]$wordHwnd) {
    [void][VTBlankUiNative]::SetForegroundWindow($wordHwnd)
    Start-Sleep -Milliseconds 500
    $root = [System.Windows.Automation.AutomationElement]::FromHandle($wordHwnd)
    $tab = Wait-ElementExact $root "VisualTeX" ([System.Windows.Automation.ControlType]::TabItem) 15
    $selectionObject = $null
    if (-not $tab.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern, [ref]$selectionObject)) {
        throw "VisualTeX Ribbon tab has no SelectionItemPattern."
    }
    ([System.Windows.Automation.SelectionItemPattern]$selectionObject).Select()
    Start-Sleep -Milliseconds 600
    return $root
}

function Find-OfficeEditorWindow() {
    $script:foundEditor = [IntPtr]::Zero
    [VTBlankUiNative]::EnumWindows({
        param([IntPtr]$hwnd, [IntPtr]$unused)
        if (-not [VTBlankUiNative]::IsWindowVisible($hwnd)) { return $true }
        [uint32]$pidValue = 0
        [void][VTBlankUiNative]::GetWindowThreadProcessId($hwnd, [ref]$pidValue)
        if ($pidValue -eq 0) { return $true }
        try {
            $process = Get-Process -Id $pidValue -ErrorAction Stop
            if ($process.ProcessName -ne "visualtex") { return $true }
        } catch { return $true }
        $length = [VTBlankUiNative]::GetWindowTextLength($hwnd)
        if ($length -le 0) { return $true }
        $text = New-Object Text.StringBuilder ($length + 1)
        [void][VTBlankUiNative]::GetWindowText($hwnd, $text, $text.Capacity)
        if ($text.ToString() -notmatch "Office") { return $true }
        $script:foundEditor = $hwnd
        return $false
    }, [IntPtr]::Zero) | Out-Null
    return $script:foundEditor
}

function Wait-OfficeEditorWindow([int]$seconds) {
    $deadline = [DateTime]::UtcNow.AddSeconds($seconds)
    do {
        $hwnd = Find-OfficeEditorWindow
        if ($hwnd -ne [IntPtr]::Zero) { return $hwnd }
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "Office formula editor window did not appear."
}

function Wait-OfficeEditorHidden([int]$seconds) {
    $deadline = [DateTime]::UtcNow.AddSeconds($seconds)
    do {
        if ((Find-OfficeEditorWindow) -eq [IntPtr]::Zero) { return }
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 150
    } while ([DateTime]::UtcNow -lt $deadline)
    $vt = Get-Process visualtex -ErrorAction SilentlyContinue | Sort-Object StartTime -Descending | Select-Object -First 1
    $responding = if ($null -ne $vt) { $vt.Responding } else { $null }
    throw "EDITOR_TIMEOUT_AFTER_FINISH|visualtexResponding=$responding"
}

function Focus-And-SendKeys(
    [System.Windows.Automation.AutomationElement]$element,
    [string[]]$keys,
    [string]$description
) {
    $element.SetFocus()
    Start-Sleep -Milliseconds 150
    foreach ($key in $keys) {
        [System.Windows.Forms.SendKeys]::SendWait($key)
        Start-Sleep -Milliseconds 120
    }
    Write-Output "UI|$description"
}

function Ensure-Checked([System.Windows.Automation.AutomationElement]$checkbox) {
    $patternObject = $null
    if (-not $checkbox.TryGetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern, [ref]$patternObject)) {
        throw "Number checkbox has no TogglePattern."
    }
    $toggle = [System.Windows.Automation.TogglePattern]$patternObject
    if ($toggle.Current.ToggleState -eq [System.Windows.Automation.ToggleState]::Off) {
        $toggle.Toggle()
        Start-Sleep -Milliseconds 300
    }
    if ($toggle.Current.ToggleState -ne [System.Windows.Automation.ToggleState]::On) {
        throw "Number checkbox did not enter checked state."
    }
    Write-Output "UI|number=on"
}

function Read-ComboValue([System.Windows.Automation.AutomationElement]$combo) {
    $valueObject = $null
    if ($combo.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$valueObject)) {
        return ([System.Windows.Automation.ValuePattern]$valueObject).Current.Value
    }
    return ""
}

$word = $null
$document = $null
$addIns = $null
$addIn = $null
$shape = $null
$shapeRange = $null
$placeRef = $null
$placeRefCode = $null
$editorHwnd = [IntPtr]::Zero
$ownedWordPid = 0
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $word.DisplayAlerts = 0
    $ownedWordPid = (Get-Process WINWORD | Sort-Object StartTime -Descending | Select-Object -First 1).Id
    $document = $word.Documents.Add()
    $document.Activate()
    Write-Output "WORD|blank-document-created|pid=$ownedWordPid"

    $addIns = $word.COMAddIns
    $addIn = $addIns.Item("VisualTeX.WordVsto")
    if (-not $addIn.Connect) {
        $addIn.Connect = $true
        Start-Sleep -Milliseconds 1000
    }
    if (-not $addIn.Connect) { throw "Installed VisualTeX.WordVsto add-in is not connected." }
    Write-Output "WORD|addin-connected"

    $wordRoot = Select-VisualTeXRibbon ([IntPtr]$word.ActiveWindow.Hwnd)
    $insertButtonName = "OLE " + (Text-FromCodePoints @(0x884C,0x95F4,0x516C,0x5F0F))
    $insertButton = Wait-ElementExact $wordRoot $insertButtonName ([System.Windows.Automation.ControlType]::Button) 15
    Invoke-Element $insertButton "OLE display formula"
    Write-Output "UI|ribbon-ole-display-clicked"

    $editorHwnd = Wait-OfficeEditorWindow 20
    [void][VTBlankUiNative]::SetForegroundWindow($editorHwnd)
    Start-Sleep -Milliseconds 700
    $editorRoot = [System.Windows.Automation.AutomationElement]::FromHandle($editorHwnd)
    Write-Output "UI|editor-opened"

    $objectFormatName = Text-FromCodePoints @(0x516C,0x5F0F,0x5BF9,0x8C61,0x683C,0x5F0F)
    $objectCombo = Wait-ElementExact $editorRoot $objectFormatName ([System.Windows.Automation.ControlType]::ComboBox) 10
    Focus-And-SendKeys $objectCombo @("{HOME}","{DOWN}","{ENTER}") "save-as=MathType-OLE"
    Start-Sleep -Milliseconds 500
    $objectValue = Read-ComboValue $objectCombo
    Write-Output "UI_VALUE|object=$objectValue"

    $editorRoot = [System.Windows.Automation.AutomationElement]::FromHandle($editorHwnd)
    $numberName = Text-FromCodePoints @(0x7F16,0x53F7)
    $numberCheckbox = Wait-ElementExact $editorRoot $numberName ([System.Windows.Automation.ControlType]::CheckBox) 10
    Ensure-Checked $numberCheckbox

    Start-Sleep -Milliseconds 400
    $editorRoot = [System.Windows.Automation.AutomationElement]::FromHandle($editorHwnd)
    $numberSideName = "MathType " + (Text-FromCodePoints @(0x516C,0x5F0F,0x7F16,0x53F7,0x4F4D,0x7F6E))
    $sideCombo = Wait-ElementExact $editorRoot $numberSideName ([System.Windows.Automation.ControlType]::ComboBox) 10
    Focus-And-SendKeys $sideCombo @("{HOME}","{ENTER}") "number-side=left"
    Start-Sleep -Milliseconds 400
    $sideValue = Read-ComboValue $sideCombo
    Write-Output "UI_VALUE|side=$sideValue"

    $editorRoot = [System.Windows.Automation.AutomationElement]::FromHandle($editorHwnd)
    $presetName = Text-FromCodePoints @(0x52FE,0x80A1,0x5B9A,0x7406)
    $preset = Wait-ElementExact $editorRoot $presetName ([System.Windows.Automation.ControlType]::Button) 10
    Invoke-Element $preset "Pythagorean theorem preset"
    Write-Output "UI|preset-clicked"
    Start-Sleep -Milliseconds 1000

    $editorRoot = [System.Windows.Automation.AutomationElement]::FromHandle($editorHwnd)
    $finishName = Text-FromCodePoints @(0x5B8C,0x6210,0x5E76,0x63D2,0x5165)
    $finish = Wait-ElementExact $editorRoot $finishName ([System.Windows.Automation.ControlType]::Button) 10
    Invoke-Element $finish "Finish and insert"
    Write-Output "UI|finish-clicked"

    Wait-OfficeEditorHidden $TimeoutSeconds
    Write-Output "UI|editor-closed-after-finish"

    $deadline = [DateTime]::UtcNow.AddSeconds(20)
    do {
        try {
            if ($document.InlineShapes.Count -ge 1) { break }
        } catch { }
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 200
    } while ([DateTime]::UtcNow -lt $deadline)

    $shapeCount = $document.InlineShapes.Count
    $fieldCount = $document.Fields.Count
    Write-Output "WORD|after-commit|inlineShapes=$shapeCount|fields=$fieldCount"
    if ($shapeCount -ne 1) { throw "Expected exactly one inserted OLE in blank document; found $shapeCount." }

    $shape = $document.InlineShapes.Item(1)
    $progId = $shape.OLEFormat.ProgID
    $shapeRange = $shape.Range
    Write-Output "WORD|ole|progId=$progId|range=$($shapeRange.Start):$($shapeRange.End)"
    if ($progId -notlike "Equation.DSMT4*") { throw "Inserted object is not Equation.DSMT4: $progId" }

    for ($index = 1; $index -le $document.Fields.Count; $index++) {
        $field = $document.Fields.Item($index)
        $code = $null
        try {
            $code = $field.Code
            $fieldCodeText = if ($null -ne $code.Text) { [string]$code.Text } else { "" }
            if ($fieldCodeText -like "*MACROBUTTON MTPlaceRef*") {
                $placeRef = $field
                $field = $null
                break
            }
        } finally {
            Release-ComObject $code
            Release-ComObject $field
        }
    }
    if ($null -eq $placeRef) { throw "No MTPlaceRef field was created." }
    $placeRefCode = $placeRef.Code
    $nested = $placeRefCode.Fields.Count
    $codeText = $placeRefCode.Text
    Write-Output "WORD|mtplaceref|nested=$nested|code=$codeText"
    if ($nested -lt 2) { throw "MTPlaceRef nested field tree is incomplete: $nested" }
    if ($placeRefCode.Start -ge $shapeRange.Start) { throw "MTPlaceRef is not on the left side of the MathType OLE." }
    if ($codeText.Length -eq 0 -or [char]::IsWhiteSpace($codeText[$codeText.Length - 1])) {
        throw "MTPlaceRef code has trailing whitespace."
    }

    $tabRange = $document.Range([Math]::Max(0, $shapeRange.Start - 1), $shapeRange.Start)
    try {
        if ($tabRange.Text -ne "`t") { throw "MathType left number is not separated from the OLE by one tab." }
    } finally { Release-ComObject $tabRange }

    Write-Output "INSTALLED_BLANK_UI_TEST=PASS"
}
finally {
    if ($editorHwnd -ne [IntPtr]::Zero -and (Find-OfficeEditorWindow) -ne [IntPtr]::Zero) {
        try { [void][VTBlankUiNative]::PostMessage($editorHwnd, 0x0010, [UIntPtr]::Zero, [IntPtr]::Zero) } catch { }
        Start-Sleep -Milliseconds 300
    }
    Release-ComObject $placeRefCode
    Release-ComObject $placeRef
    Release-ComObject $shapeRange
    Release-ComObject $shape
    Release-ComObject $addIn
    Release-ComObject $addIns
    if ($document) { try { $document.Close(0) } catch { } }
    Release-ComObject $document
    if ($word) { try { $word.Quit() } catch { } }
    Release-ComObject $word
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($ownedWordPid -gt 0) {
        try { Stop-Process -Id $ownedWordPid -Force -ErrorAction SilentlyContinue } catch { }
    }
}
