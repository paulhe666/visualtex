param(
    [Parameter(Mandatory = $true)][int]$WordProcessId,
    [Parameter(Mandatory = $true)][string]$Label,
    [single]$BodyFontSize = 12,
    [single]$FormulaFontSize = 10.5,
    [string]$Suffix = "XYZ"
)

$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding

$root = $PSScriptRoot
$evidenceRoot = Join-Path $root "evidence"
$ui = Join-Path $root "word-ui.ps1"
$connection = Join-Path $root "word-connection.ps1"
$targetPath = Join-Path $evidenceRoot "target-word.json"
$reportPath = Join-Path $evidenceRoot ($Label + "-mathtype-typing-font.json")

. $connection

function Release-ComObject([object]$value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value) } catch { }
    }
}

function Invoke-WordUi([string[]]$Arguments) {
    $all = @(
        "-NoProfile",
        "-STA",
        "-ExecutionPolicy", "Bypass",
        "-File", $ui,
        "-WordProcessId", [string]$WordProcessId
    ) + $Arguments
    $output = & powershell.exe @all 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "word-ui failed: $($Arguments -join ' ')`n$($output -join "`n")"
    }
    $output | ForEach-Object { Write-Output $_ }
    return @($output)
}

function Wait-Editor([bool]$Visible, [int]$TimeoutSeconds = 20) {
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    do {
        $windows = Invoke-WordUi @("-Action", "windows", "-Label", $Label)
        $joined = $windows -join "`n"
        $hasEditor = $joined -match 'Office 公式编辑器'
        if ($hasEditor -eq $Visible) { return }
        Start-Sleep -Milliseconds 150
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "Timed out waiting for Office editor visible=$Visible"
}

if (-not (Test-Path -LiteralPath $targetPath)) {
    throw "target-word.json is missing. Start the stage Word with start-stage-word.ps1 first."
}
$target = Get-Content -LiteralPath $targetPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ([int]$target.pid -ne $WordProcessId) {
    throw "Pinned Word PID $WordProcessId does not match target-word.json PID $($target.pid)."
}
if ($target.stage -ne $Label) {
    throw "Pinned stage '$($target.stage)' does not match requested label '$Label'."
}

# A previous disposable stage can leave its Office editor open after its Word
# host is closed. Close only that VisualTeX editor through its real Cancel UI;
# never terminate a process or touch another Word document.
$initialWindows = Invoke-WordUi @("-Action", "windows", "-Label", $Label)
if (($initialWindows -join "`n") -match 'Office 公式编辑器') {
    Invoke-WordUi @("-Action", "click", "-Target", "editor", "-Name", "取消", "-Label", $Label) | Out-Null
    Wait-Editor $false 15
    Write-Output "CLOSED_PREEXISTING_OFFICE_EDITOR"
}

$word = $null
$document = $null
$normal = $null
$prefix = $null
try {
    $word = Connect-RunningWord $WordProcessId
    $document = $word.ActiveDocument
    if ($null -eq $document) { throw "Pinned Word has no active document." }

    # COM is used only to prepare the test document and insertion point.
    $normal = $document.Styles.Item(-1)
    $normal.Font.Size = 10.5
    $document.Content.Text = "DIRECT BODY " + [string][char]13
    $prefixEnd = "DIRECT BODY ".Length
    $prefix = $document.Range(0, $prefixEnd)
    $prefix.Font.Size = $BodyFontSize
    $word.Selection.SetRange($prefixEnd, $prefixEnd)
    $word.Selection.Font.Size = $BodyFontSize
    $document.Activate()
    try { $word.ActiveWindow.WindowState = 1 } catch { }

    Write-Output ("PREPARED|pid={0}|doc={1}|normal={2}|body={3}|selection={4}" -f `
        $WordProcessId, $document.Name, $normal.Font.Size, $prefix.Font.Size, $word.Selection.Font.Size)
}
finally {
    Release-ComObject $prefix
    Release-ComObject $normal
    Release-ComObject $document
    Release-ComObject $word
}

# From here through insertion, every VisualTeX operation is real UI.
Invoke-WordUi @("-Action", "ribbon", "-Label", $Label) | Out-Null
Invoke-WordUi @("-Action", "click", "-Name", "OLE 行内公式", "-Label", $Label) | Out-Null
Wait-Editor $true
Invoke-WordUi @("-Action", "combo", "-Target", "editor", "-Name", "公式对象格式", "-Choice", "MathType OLE", "-Label", $Label) | Out-Null

$formulaChoice = if ([Math]::Abs($FormulaFontSize - 10.5) -lt 0.01) {
    "五号（10.5 磅）"
} elseif ([Math]::Abs($FormulaFontSize - 12) -lt 0.01) {
    "小四（12 磅）"
} else {
    throw "This acceptance currently supports formula sizes 10.5 or 12 pt."
}
Invoke-WordUi @("-Action", "combo", "-Target", "editor", "-Name", "公式字号", "-Choice", $formulaChoice, "-Label", $Label) | Out-Null
Invoke-WordUi @("-Action", "click", "-Target", "editor", "-Name", "勾股定理", "-Label", $Label) | Out-Null
Invoke-WordUi @("-Action", "click", "-Target", "editor", "-Name", "完成并插入", "-Label", $Label) | Out-Null
Wait-Editor $false 60

$word = $null
$document = $null
$shape = $null
$shapeRange = $null
$beforeTyping = $null
$documentName = $null
$immediate = $null
try {
    $word = Connect-RunningWord $WordProcessId
    $document = $word.ActiveDocument
    if ($document.InlineShapes.Count -ne 1) {
        throw "Expected exactly one MathType OLE, found $($document.InlineShapes.Count)."
    }
    $documentName = $document.Name
    $shape = $document.InlineShapes.Item(1)
    if ($shape.OLEFormat.ProgID -notlike "Equation.DSMT4*") {
        throw "Inserted object is not MathType Equation.DSMT4: $($shape.OLEFormat.ProgID)"
    }
    $shapeRange = $shape.Range
    $beforeTyping = $word.Selection.Start
    $immediate = [ordered]@{
        selectionStart = $word.Selection.Start
        selectionEnd = $word.Selection.End
        selectionFontSize = [double]$word.Selection.Font.Size
        shapeStart = $shapeRange.Start
        shapeEnd = $shapeRange.End
        shapeFontSize = [double]$shapeRange.Font.Size
        progId = $shape.OLEFormat.ProgID
    }
    Write-Output ("AFTER_INSERT|selectionSize={0}|shapeSize={1}|range={2}:{3}" -f `
        $immediate.selectionFontSize, $immediate.shapeFontSize, $immediate.shapeStart, $immediate.shapeEnd)
}
finally {
    Release-ComObject $shapeRange
    Release-ComObject $shape
    Release-ComObject $document
    Release-ComObject $word
}

# Bring the pinned Word to the foreground by physically clicking its real UIA
# title bar. This does not move the document Selection or rewrite its typing
# format. Then send the suffix through the existing guarded real-keyboard path.
Invoke-WordUi @("-Action", "activate", "-Target", "word", "-Label", $Label) | Out-Null
Invoke-WordUi @("-Action", "keys", "-Target", "word", "-Value", $Suffix, "-Label", $Label) | Out-Null

$word = $null
$document = $null
$typed = $null
try {
    $word = Connect-RunningWord $WordProcessId
    $document = $word.ActiveDocument
    $afterTyping = $word.Selection.Start
    if ($afterTyping -le $beforeTyping) {
        throw "Real keyboard suffix did not advance the Word caret."
    }
    $typed = $document.Range($beforeTyping, $afterTyping)
    $typedText = $typed.Text
    $typedSize = [double]$typed.Font.Size
    $report = [ordered]@{
        label = $Label
        wordProcessId = $WordProcessId
        document = $document.Name
        normalFontSize = 10.5
        requestedBodyFontSize = [double]$BodyFontSize
        requestedFormulaFontSize = [double]$FormulaFontSize
        immediate = $immediate
        typed = [ordered]@{
            start = $beforeTyping
            end = $afterTyping
            text = $typedText
            fontSize = $typedSize
        }
        method = "independent /x Word; COM preparation/read-only verification; real Word Ribbon + VisualTeX editor UI + SendKeys typing"
    }
    [IO.File]::WriteAllText(
        $reportPath,
        ($report | ConvertTo-Json -Depth 6),
        [Text.UTF8Encoding]::new($false))
    Write-Output ("TYPED_RESULT|text={0}|fontSize={1}|expected={2}" -f $typedText, $typedSize, $BodyFontSize)
    if ([Math]::Abs($immediate.selectionFontSize - $BodyFontSize) -gt 0.01) {
        throw "Insertion caret font size changed to $($immediate.selectionFontSize) pt; expected $BodyFontSize pt."
    }
    if ([Math]::Abs($typedSize - $BodyFontSize) -gt 0.01) {
        throw "Text typed after MathType OLE is $typedSize pt; expected $BodyFontSize pt."
    }
    Write-Output "MATHTYPE_TYPING_FONT_UI_PASSED"
}
finally {
    Release-ComObject $typed
    Release-ComObject $document
    Release-ComObject $word
}
