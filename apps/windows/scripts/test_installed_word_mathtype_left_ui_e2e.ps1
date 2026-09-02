param(
    [string]$ArtifactRoot = "src-windows/artifacts/installed-word-mathtype-left-ui-e2e",
    [switch]$ProbeOnly,
    [switch]$SeedMalformedMathTypeRows,
    [int]$Iterations = 1
)

$ErrorActionPreference = "Stop"
$artifactPath = [IO.Path]::GetFullPath((Join-Path (Get-Location) $ArtifactRoot))
[IO.Directory]::CreateDirectory($artifactPath) | Out-Null

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition @"
using System;
using System.Text;
using System.Runtime.InteropServices;
public static class VisualTeXUiNative {
    public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr hwnd, StringBuilder text, int count);
    [DllImport("user32.dll")] public static extern int GetWindowTextLength(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr hwnd, uint msg, UIntPtr wParam, IntPtr lParam);
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
    foreach ($codePoint in $codePoints) {
        [void]$builder.Append([char]$codePoint)
    }
    return $builder.ToString()
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

function Select-ComboBoxOption(
    [System.Windows.Automation.AutomationElement]$combo,
    [string]$optionName,
    [string]$description
) {
    if ($null -eq $combo) { throw "ComboBox not found: $description" }
    [void]$combo.SetFocus()
    $expanded = $false
    $expandObject = $null
    if ($combo.TryGetCurrentPattern(
        [System.Windows.Automation.ExpandCollapsePattern]::Pattern,
        [ref]$expandObject)) {
        $expandPattern = [System.Windows.Automation.ExpandCollapsePattern]$expandObject
        if ($expandPattern.Current.ExpandCollapseState -ne [System.Windows.Automation.ExpandCollapseState]::Expanded) {
            $expandPattern.Expand()
        }
        $expanded = $true
        Start-Sleep -Milliseconds 200
    }

    $desktop = [System.Windows.Automation.AutomationElement]::RootElement
    $option = $null
    foreach ($type in @(
        [System.Windows.Automation.ControlType]::ListItem,
        [System.Windows.Automation.ControlType]::MenuItem,
        [System.Windows.Automation.ControlType]::Text
    )) {
        $option = Find-ElementExact $desktop $optionName $type
        if ($null -ne $option) { break }
    }
    if ($null -eq $option) {
        # Chromium/WebView2 exposes the HTML <select> itself as a UIA ComboBox,
        # but does not always surface its <option> nodes in the desktop UIA tree.
        # Keep this as a genuine UI action by driving the focused ComboBox with
        # keyboard navigation, then verify the resulting Session state below.
        if ($optionName -eq "MathType OLE") {
            [System.Windows.Forms.SendKeys]::SendWait("{HOME}{DOWN}")
        } elseif ($optionName -eq (Text-FromCodePoints @(0x5DE6,0x4FA7))) {
            [System.Windows.Forms.SendKeys]::SendWait("{HOME}")
        } elseif ($optionName -eq (Text-FromCodePoints @(0x53F3,0x4FA7))) {
            [System.Windows.Forms.SendKeys]::SendWait("{END}")
        } else {
            throw "Could not locate option '$optionName' and no keyboard fallback is defined for $description."
        }
        Start-Sleep -Milliseconds 350
    } else {
        $selectionObject = $null
        if ($option.TryGetCurrentPattern(
            [System.Windows.Automation.SelectionItemPattern]::Pattern,
            [ref]$selectionObject)) {
            ([System.Windows.Automation.SelectionItemPattern]$selectionObject).Select()
        } else {
            $invokeObject = $null
            if (-not $option.TryGetCurrentPattern(
                [System.Windows.Automation.InvokePattern]::Pattern,
                [ref]$invokeObject)) {
                throw "Option '$optionName' exposes neither SelectionItemPattern nor InvokePattern."
            }
            ([System.Windows.Automation.InvokePattern]$invokeObject).Invoke()
        }
        Start-Sleep -Milliseconds 250
    }

    if ($expanded -and $null -ne $expandObject) {
        try {
            $expandPattern = [System.Windows.Automation.ExpandCollapsePattern]$expandObject
            if ($expandPattern.Current.ExpandCollapseState -eq [System.Windows.Automation.ExpandCollapseState]::Expanded) {
                $expandPattern.Collapse()
            }
        } catch { }
    }
    Write-Output "UI_SELECT|$description|$optionName"
}

function Set-CheckBoxOn(
    [System.Windows.Automation.AutomationElement]$checkbox,
    [string]$description
) {
    if ($null -eq $checkbox) { throw "CheckBox not found: $description" }
    $toggleObject = $null
    if (-not $checkbox.TryGetCurrentPattern(
        [System.Windows.Automation.TogglePattern]::Pattern,
        [ref]$toggleObject)) {
        throw "CheckBox has no TogglePattern: $description"
    }
    $toggle = [System.Windows.Automation.TogglePattern]$toggleObject
    if ($toggle.Current.ToggleState -eq [System.Windows.Automation.ToggleState]::Off) {
        $toggle.Toggle()
        Start-Sleep -Milliseconds 250
    }
    if ($toggle.Current.ToggleState -ne [System.Windows.Automation.ToggleState]::On) {
        throw "CheckBox did not enter On state: $description"
    }
    Write-Output "UI_CHECK|$description|On"
}

$sessionRoot = Join-Path $env:APPDATA "com.visualtex.studio\office\sessions"
function Snapshot-SessionIds() {
    if (-not (Test-Path $sessionRoot)) { return @{} }
    $result = @{}
    foreach ($directory in Get-ChildItem -LiteralPath $sessionRoot -Directory -ErrorAction SilentlyContinue) {
        $result[$directory.Name] = $true
    }
    return $result
}

function Wait-NewWordSession([hashtable]$existing, [string]$expectedDocumentId, [TimeSpan]$timeout) {
    $deadline = [DateTime]::UtcNow + $timeout
    do {
        if (Test-Path $sessionRoot) {
            foreach ($directory in Get-ChildItem -LiteralPath $sessionRoot -Directory | Sort-Object LastWriteTimeUtc -Descending) {
                if ($existing.ContainsKey($directory.Name)) { continue }
                $sessionPath = Join-Path $directory.FullName "session.json"
                if (-not (Test-Path $sessionPath)) { continue }
                try {
                    $session = Get-Content -LiteralPath $sessionPath -Raw -Encoding UTF8 | ConvertFrom-Json
                    if ($session.host -eq "word" -and $session.sourceDocumentId -eq $expectedDocumentId) {
                        return $directory.Name
                    }
                } catch { }
            }
        }
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "No new Word Office Session appeared."
}

function Add-MalformedMathTypeNumberRow([object]$document) {
    $start = $document.Content.End - 1
    $range = $document.Range($start, $start)
    $outer = $null
    try {
        $outer = $document.Fields.Add($range, 51, "MTPlaceRef", $false)
        $outer.Code.Text = " MACROBUTTON MTPlaceRef \\* MERGEFORMAT (.)"
        $position = $outer.Result.End + 1
        foreach ($instruction in @(
            "MTEqn \\h \\* MERGEFORMAT",
            "MTChap \\c \\* Arabic \\* MERGEFORMAT",
            "MTEqn \\c \\* Arabic \\* MERGEFORMAT"
        )) {
            $insert = $document.Range($position, $position)
            try {
                $field = $document.Fields.Add($insert, 12, $instruction, $false)
                try { $field.Update() | Out-Null } catch { }
                $position = $field.Result.End + 1
                Release-ComObject $field
            } finally { Release-ComObject $insert }
        }
        $end = $document.Content.End - 1
        $paragraphMark = $document.Range($end, $end)
        try { $paragraphMark.InsertParagraphAfter() } finally { Release-ComObject $paragraphMark }
    } finally {
        Release-ComObject $outer
        Release-ComObject $range
    }
}

function Read-Session([string]$sessionId) {
    if ([string]::IsNullOrWhiteSpace($sessionId)) { return $null }
    $sessionPath = Join-Path (Join-Path $sessionRoot $sessionId) "session.json"
    if (-not (Test-Path $sessionPath)) { return $null }
    try { return Get-Content -LiteralPath $sessionPath -Raw -Encoding UTF8 | ConvertFrom-Json }
    catch { return $null }
}

function Select-VisualTeXRibbon([IntPtr]$wordHwnd) {
    [void][VisualTeXUiNative]::SetForegroundWindow($wordHwnd)
    Start-Sleep -Milliseconds 400
    $root = [System.Windows.Automation.AutomationElement]::FromHandle($wordHwnd)
    $deadline = [DateTime]::UtcNow.AddSeconds(15)
    $tab = $null
    do {
        $tab = Find-ElementExact $root "VisualTeX" ([System.Windows.Automation.ControlType]::TabItem)
        if ($null -ne $tab) { break }
        Start-Sleep -Milliseconds 200
    } while ([DateTime]::UtcNow -lt $deadline)
    if ($null -eq $tab) { throw "VisualTeX Ribbon tab was not found." }
    $selectionObject = $null
    if (-not $tab.TryGetCurrentPattern(
        [System.Windows.Automation.SelectionItemPattern]::Pattern,
        [ref]$selectionObject)) {
        throw "VisualTeX Ribbon tab has no SelectionItemPattern."
    }
    ([System.Windows.Automation.SelectionItemPattern]$selectionObject).Select()
    Start-Sleep -Milliseconds 500
    return $root
}

function Find-OfficeEditorWindow() {
    $script:foundEditor = [IntPtr]::Zero
    [VisualTeXUiNative]::EnumWindows({
        param([IntPtr]$hwnd, [IntPtr]$unused)
        if (-not [VisualTeXUiNative]::IsWindowVisible($hwnd)) { return $true }
        $length = [VisualTeXUiNative]::GetWindowTextLength($hwnd)
        if ($length -le 0) { return $true }
        $text = New-Object Text.StringBuilder ($length + 1)
        [void][VisualTeXUiNative]::GetWindowText($hwnd, $text, $text.Capacity)
        $title = $text.ToString()
        if ($title -notmatch "Office") { return $true }
        [uint32]$pidValue = 0
        [void][VisualTeXUiNative]::GetWindowThreadProcessId($hwnd, [ref]$pidValue)
        if ($pidValue -eq 0) { return $true }
        try {
            $process = Get-Process -Id $pidValue -ErrorAction Stop
            if ($process.ProcessName -ne "visualtex") { return $true }
        } catch { return $true }
        $script:foundEditor = $hwnd
        return $false
    }, [IntPtr]::Zero) | Out-Null
    return $script:foundEditor
}

function Wait-OfficeEditorWindow([TimeSpan]$timeout) {
    $deadline = [DateTime]::UtcNow + $timeout
    do {
        $hwnd = Find-OfficeEditorWindow
        if ($hwnd -ne [IntPtr]::Zero) { return $hwnd }
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "Office formula editor window did not appear."
}

function Dump-AutomationTree(
    [System.Windows.Automation.AutomationElement]$root,
    [string]$path
) {
    $lines = New-Object Collections.Generic.List[string]
    $lines.Add("ROOT|$($root.Current.ControlType.ProgrammaticName)|NAME=$($root.Current.Name)|ID=$($root.Current.AutomationId)")
    $elements = $root.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        [System.Windows.Automation.Condition]::TrueCondition)
    foreach ($element in $elements) {
        $name = $element.Current.Name
        $automationId = $element.Current.AutomationId
        if ([string]::IsNullOrEmpty($name) -and [string]::IsNullOrEmpty($automationId)) { continue }
        $lines.Add("EL|$($element.Current.ControlType.ProgrammaticName)|NAME=$name|ID=$automationId|ENABLED=$($element.Current.IsEnabled)")
    }
    [IO.File]::WriteAllLines($path, $lines, [Text.UTF8Encoding]::new($false))
    $lines | ForEach-Object { Write-Output $_ }
}

$word = $null
$document = $null
$addIns = $null
$addIn = $null
$editorHwnd = [IntPtr]::Zero
$sessionId = $null
$wordPid = 0
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $word.DisplayAlerts = 0
    $document = $word.Documents.Add()
    $document.Activate()
    $uniqueDocumentPath = Join-Path $artifactPath ("ui-e2e-source-" + [guid]::NewGuid().ToString("N") + ".docx")
    $document.SaveAs2($uniqueDocumentPath, 12)
    $expectedDocumentId = $document.FullName
    Write-Output "TEST_DOCUMENT_ID=$expectedDocumentId"
    if ($SeedMalformedMathTypeRows) {
        Add-MalformedMathTypeNumberRow $document
        Add-MalformedMathTypeNumberRow $document
        $document.Range($document.Content.End - 1, $document.Content.End - 1).Select()
        Write-Output ("SEEDED_MALFORMED_ROWS|paras="+$document.Paragraphs.Count+"|fields="+$document.Fields.Count)
    }

    $addIns = $word.COMAddIns
    $addIn = $addIns.Item("VisualTeX.WordVsto")
    if (-not $addIn.Connect) {
        $addIn.Connect = $true
        Start-Sleep -Milliseconds 800
    }
    if (-not $addIn.Connect) { throw "Installed VisualTeX.WordVsto add-in is not connected." }

    $existingSessions = Snapshot-SessionIds
    $wordHwnd = [IntPtr]$word.ActiveWindow.Hwnd
    [uint32]$wordPidValue = 0
    [void][VisualTeXUiNative]::GetWindowThreadProcessId($wordHwnd, [ref]$wordPidValue)
    $wordPid = [int]$wordPidValue
    Write-Output "TEST_WORD_PID=$wordPid"
    $wordRoot = Select-VisualTeXRibbon $wordHwnd
    $ribbonTreePath = Join-Path $artifactPath "word-ribbon-uia-tree.txt"
    if ($ProbeOnly) { Dump-AutomationTree $wordRoot $ribbonTreePath }
    else { Dump-AutomationTree $wordRoot $ribbonTreePath | Out-Null }
    $insertButtonName = "OLE " + [char]0x884C + [char]0x95F4 + [char]0x516C + [char]0x5F0F
    $insertButton = Find-ElementExact $wordRoot $insertButtonName ([System.Windows.Automation.ControlType]::Button)
    Invoke-Element $insertButton "Word VisualTeX OLE display formula button"

    $editorHwnd = Wait-OfficeEditorWindow ([TimeSpan]::FromSeconds(20))
    $sessionId = Wait-NewWordSession $existingSessions $expectedDocumentId ([TimeSpan]::FromSeconds(20))
    Write-Output "SESSION_ID=$sessionId"
    [void][VisualTeXUiNative]::SetForegroundWindow($editorHwnd)
    Start-Sleep -Milliseconds 800
    $editorRoot = [System.Windows.Automation.AutomationElement]::FromHandle($editorHwnd)
    $treePath = Join-Path $artifactPath "editor-uia-tree.txt"
    if ($ProbeOnly) { Dump-AutomationTree $editorRoot $treePath }
    else { Dump-AutomationTree $editorRoot $treePath | Out-Null }

    if ($ProbeOnly) {
        $cancelName = [string]([char]0x53D6) + [char]0x6D88
        $cancel = Find-ElementExact $editorRoot $cancelName ([System.Windows.Automation.ControlType]::Button)
        if ($null -ne $cancel) {
            Invoke-Element $cancel "Office editor Cancel button"
            Start-Sleep -Milliseconds 600
        } else {
            [void][VisualTeXUiNative]::PostMessage($editorHwnd, 0x0010, [UIntPtr]::Zero, [IntPtr]::Zero)
        }
        Write-Output "PROBE_ONLY_UI_TREE=$treePath"
        return
    }

    $objectFormatName = Text-FromCodePoints @(0x516C,0x5F0F,0x5BF9,0x8C61,0x683C,0x5F0F)
    $objectCombo = Wait-ElementExact $editorRoot $objectFormatName ([System.Windows.Automation.ControlType]::ComboBox) ([TimeSpan]::FromSeconds(5))
    Select-ComboBoxOption $objectCombo "MathType OLE" "formula object format"

    Start-Sleep -Milliseconds 400
    $editorRoot = [System.Windows.Automation.AutomationElement]::FromHandle($editorHwnd)
    $numberName = Text-FromCodePoints @(0x7F16,0x53F7)
    $numberCheckbox = Wait-ElementExact $editorRoot $numberName ([System.Windows.Automation.ControlType]::CheckBox) ([TimeSpan]::FromSeconds(5))
    Set-CheckBoxOn $numberCheckbox "equation numbering"

    Start-Sleep -Milliseconds 400
    $editorRoot = [System.Windows.Automation.AutomationElement]::FromHandle($editorHwnd)
    $numberSideName = "MathType " + (Text-FromCodePoints @(0x516C,0x5F0F,0x7F16,0x53F7,0x4F4D,0x7F6E))
    $numberSideCombo = Wait-ElementExact $editorRoot $numberSideName ([System.Windows.Automation.ControlType]::ComboBox) ([TimeSpan]::FromSeconds(5))
    $leftName = Text-FromCodePoints @(0x5DE6,0x4FA7)
    Select-ComboBoxOption $numberSideCombo $leftName "MathType equation number side"

    Start-Sleep -Milliseconds 400
    $editorRoot = [System.Windows.Automation.AutomationElement]::FromHandle($editorHwnd)
    $pythagoreanName = Text-FromCodePoints @(0x52FE,0x80A1,0x5B9A,0x7406)
    $preset = Wait-ElementExact $editorRoot $pythagoreanName ([System.Windows.Automation.ControlType]::Button) ([TimeSpan]::FromSeconds(5))
    Invoke-Element $preset "Pythagorean theorem preset"
    Write-Output "UI_INVOKE|preset|$pythagoreanName"

    Start-Sleep -Milliseconds 800
    $draft = Read-Session $sessionId
    if ($null -ne $draft) {
        Write-Output ("SESSION_BEFORE_COMMIT|status={0}|objectMode={1}|displayMode={2}|numbered={3}|side={4}|dirty={5}|error={6}" -f `
            $draft.status,$draft.objectMode,$draft.displayMode,$draft.numbered,$draft.mathTypeNumberPosition,$draft.dirty,$draft.error)
    }

    $editorRoot = [System.Windows.Automation.AutomationElement]::FromHandle($editorHwnd)
    $finishName = Text-FromCodePoints @(0x5B8C,0x6210,0x5E76,0x63D2,0x5165)
    $finish = Wait-ElementExact $editorRoot $finishName ([System.Windows.Automation.ControlType]::Button) ([TimeSpan]::FromSeconds(5))
    Invoke-Element $finish "Finish and insert"
    Write-Output "UI_INVOKE|finish-and-insert"

    $deadline = [DateTime]::UtcNow.AddSeconds(45)
    $completed = $false
    $lastSession = $null
    do {
        Start-Sleep -Milliseconds 200
        $lastSession = Read-Session $sessionId
        $editorStillVisible = (Find-OfficeEditorWindow) -ne [IntPtr]::Zero
        if ($null -ne $lastSession) {
            if ($lastSession.status -eq "failed") {
                throw ("UI_E2E_SESSION_FAILED|error=" + $lastSession.error)
            }
            if ($lastSession.status -eq "completed" -and -not $editorStillVisible) {
                $completed = $true
                break
            }
        }
    } while ([DateTime]::UtcNow -lt $deadline)

    if (-not $completed) {
        $visualTeXProcess = Get-Process visualtex -ErrorAction SilentlyContinue | Sort-Object StartTime -Descending | Select-Object -First 1
        $responding = if ($null -ne $visualTeXProcess) { $visualTeXProcess.Responding } else { $null }
        $status = if ($null -ne $lastSession) { $lastSession.status } else { "<missing>" }
        $errorText = if ($null -ne $lastSession) { $lastSession.error } else { "" }
        $timeoutTreePath = Join-Path $artifactPath "editor-uia-tree-timeout.txt"
        try {
            $liveEditor = Find-OfficeEditorWindow
            if ($liveEditor -ne [IntPtr]::Zero) {
                Dump-AutomationTree ([System.Windows.Automation.AutomationElement]::FromHandle($liveEditor)) $timeoutTreePath | Out-Null
            }
        } catch { }
        throw "UI_E2E_TIMEOUT|session=$sessionId|status=$status|visualtexResponding=$responding|error=$errorText|tree=$timeoutTreePath"
    }

    Start-Sleep -Milliseconds 500
    $shapeCount = $document.InlineShapes.Count
    $fieldCount = $document.Fields.Count
    Write-Output "WORD_AFTER_COMMIT|inlineShapes=$shapeCount|fields=$fieldCount"
    if ($shapeCount -ne 1) { throw "Expected exactly one inserted MathType OLE, found $shapeCount." }

    $shape = $document.InlineShapes.Item(1)
    try {
        $progId = $shape.OLEFormat.ProgID
        $shapeRange = $shape.Range
        try {
            Write-Output "WORD_OLE|progId=$progId|range=$($shapeRange.Start):$($shapeRange.End)"
            if ($progId -notlike "Equation.DSMT4*") { throw "Inserted OLE is not Equation.DSMT4: $progId" }

            $placeRef = $null
            for ($index = 1; $index -le $document.Fields.Count; $index++) {
                $field = $document.Fields.Item($index)
                try {
                    $code = $field.Code.Text
                    if ($code -like "*MACROBUTTON MTPlaceRef*") {
                        $placeRef = $field
                        break
                    }
                } finally {
                    if ($null -ne $field -and [Runtime.InteropServices.Marshal]::IsComObject($field) -and $field -ne $placeRef) {
                        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($field)
                    }
                }
            }
            if ($null -eq $placeRef) { throw "Inserted MathType OLE has no MTPlaceRef field." }
            try {
                $nestedCount = $placeRef.Code.Fields.Count
                $codeText = $placeRef.Code.Text
                Write-Output "WORD_MTPLACEREF|nested=$nestedCount|code=$codeText"
                if ($nestedCount -lt 2) { throw "MTPlaceRef nested field tree is incomplete: $nestedCount" }
                if ($placeRef.Code.Start -ge $shapeRange.Start) { throw "MTPlaceRef is not on the left side of the OLE." }
                $tabRange = $document.Range([Math]::Max(0,$shapeRange.Start - 1), $shapeRange.Start)
                try {
                    if ($tabRange.Text -ne "`t") { throw "Left MTPlaceRef is not separated from the OLE by one tab." }
                } finally { Release-ComObject $tabRange }
            } finally { Release-ComObject $placeRef }
        } finally { Release-ComObject $shapeRange }
    } finally { Release-ComObject $shape }

    Write-Output "UI_E2E_PASSED"
}
finally {
    if ($editorHwnd -ne [IntPtr]::Zero -and (Find-OfficeEditorWindow) -ne [IntPtr]::Zero) {
        try { [void][VisualTeXUiNative]::PostMessage($editorHwnd, 0x0010, [UIntPtr]::Zero, [IntPtr]::Zero) } catch { }
        Start-Sleep -Milliseconds 300
    }
    Release-ComObject $addIn
    Release-ComObject $addIns
    Release-ComObject $document
    Release-ComObject $word
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    if ($wordPid -gt 0) {
        try { Stop-Process -Id $wordPid -Force -ErrorAction SilentlyContinue } catch { }
    }
}
