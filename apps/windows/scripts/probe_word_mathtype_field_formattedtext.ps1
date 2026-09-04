param(
    [string]$SourceDocument = "src-windows/artifacts/mathtype-native-number-format-probe/03-native-continuous-with-reference/03-native-continuous-with-reference.docx",
    [string]$TargetDocument = "src-windows/artifacts/mathtype-native-number-format-probe/reference-completed/reference-completed.docx",
    [string]$ArtifactRoot = "src-windows/artifacts/mathtype-formattedtext-probe",
    [switch]$LiveOnly
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

function Release-ComObject([object]$value) {
    if ($null -eq $value -or -not [Runtime.InteropServices.Marshal]::IsComObject($value)) { return }
    try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) } catch { }
}

function ConvertTo-DebugText([string]$text) {
    if ($null -eq $text) { return "" }
    return $text.Replace([string][char]1, "<OLE>").Replace([string][char]9, "<TAB>").Replace([string][char]13, "<PARA>").Replace([string][char]19, "<FIELD_BEGIN>").Replace([string][char]20, "<FIELD_SEPARATOR>").Replace([string][char]21, "<FIELD_END>")
}

function Test-PlaceRefCode([string]$code) {
    return -not [string]::IsNullOrWhiteSpace($code) -and $code.TrimStart().StartsWith("MACROBUTTON MTPlaceRef", [StringComparison]::OrdinalIgnoreCase)
}

function Get-FieldEndExclusive([object]$document, [object]$field) {
    $code = $null
    $result = $null
    $probe = $null
    try {
        $code = $field.Code
        $result = $field.Result
        $probe = $document.Range($code.End, $code.End + 1)
        if ($probe.Text -eq [string][char]21) { return [int]$code.End + 1 }
        Release-ComObject $probe; $probe = $document.Range($result.End, $result.End + 1)
        if ($probe.Text -eq [string][char]21) { return [int]$result.End + 1 }
        throw "Unable to resolve MTPlaceRef end."
    } finally {
        Release-ComObject $probe
        Release-ComObject $result
        Release-ComObject $code
    }
}

function Get-PlaceRefEntries([object]$document) {
    $entries = New-Object Collections.Generic.List[object]
    $fields = $null
    $field = $null
    $code = $null
    $result = $null
    $paragraphs = $null
    $paragraph = $null
    $paragraphRange = $null
    $shapes = $null
    $shape = $null
    $shapeRange = $null
    try {
        $fields = $document.Fields
        for ($index = 1; $index -le $fields.Count; $index++) {
            Release-ComObject $shapeRange; $shapeRange = $null
            Release-ComObject $shape; $shape = $null
            Release-ComObject $shapes; $shapes = $null
            Release-ComObject $paragraphRange; $paragraphRange = $null
            Release-ComObject $paragraph; $paragraph = $null
            Release-ComObject $paragraphs; $paragraphs = $null
            Release-ComObject $result; $result = $null
            Release-ComObject $code; $code = $null
            Release-ComObject $field; $field = $fields.Item($index)
            $code = $field.Code
            if (-not (Test-PlaceRefCode ([string]$code.Text))) { continue }
            $result = $field.Result
            $fieldStart = [int]$code.Start - 1
            $fieldEnd = Get-FieldEndExclusive $document $field
            $paragraphs = $code.Paragraphs
            if ($paragraphs.Count -ne 1) { throw "MTPlaceRef spans multiple paragraphs." }
            $paragraph = $paragraphs.Item(1)
            $paragraphRange = $paragraph.Range
            $shapes = $paragraphRange.InlineShapes
            if ($shapes.Count -ne 1) { throw "MTPlaceRef paragraph has $($shapes.Count) shapes." }
            $shape = $shapes.Item(1)
            $shapeRange = $shape.Range
            $side = if ($fieldEnd -le $shapeRange.Start) { "left" } elseif ($fieldStart -ge $shapeRange.End) { "right" } else { "overlap" }
            $entries.Add([pscustomobject]@{
                Side = $side
                FieldStart = $fieldStart
                FieldEnd = $fieldEnd
                ParagraphStart = [int]$paragraphRange.Start
                ParagraphEnd = [int]$paragraphRange.End
                CodeText = [string]$code.Text
            })
        }
    } finally {
        Release-ComObject $shapeRange
        Release-ComObject $shape
        Release-ComObject $shapes
        Release-ComObject $paragraphRange
        Release-ComObject $paragraph
        Release-ComObject $paragraphs
        Release-ComObject $result
        Release-ComObject $code
        Release-ComObject $field
        Release-ComObject $fields
    }
    return $entries.ToArray()
}

function Dump-Document([object]$document, [string]$path) {
    $lines = New-Object Collections.Generic.List[string]
    $lines.Add("Paragraphs=$($document.Paragraphs.Count)")
    $lines.Add("Fields=$($document.Fields.Count)")
    $lines.Add("InlineShapes=$($document.InlineShapes.Count)")
    $lines.Add("Bookmarks=$($document.Bookmarks.Count)")
    $paragraphs = $null
    $paragraph = $null
    $range = $null
    $style = $null
    $format = $null
    $tabs = $null
    $tab = $null
    try {
        $paragraphs = $document.Paragraphs
        for ($index = 1; $index -le $paragraphs.Count; $index++) {
            Release-ComObject $tab; $tab = $null
            Release-ComObject $tabs; $tabs = $null
            Release-ComObject $format; $format = $null
            Release-ComObject $style; $style = $null
            Release-ComObject $range; $range = $null
            Release-ComObject $paragraph; $paragraph = $paragraphs.Item($index)
            $range = $paragraph.Range
            try { $style = $range.Style } catch { }
            $styleName = $(try { [string]$style.NameLocal } catch { "" })
            $tabEntries = New-Object Collections.Generic.List[string]
            try {
                $format = $paragraph.Format
                $tabs = $format.TabStops
                for ($tabIndex = 1; $tabIndex -le $tabs.Count; $tabIndex++) {
                    Release-ComObject $tab; $tab = $tabs.Item($tabIndex)
                    $tabEntries.Add("$([double]$tab.Position):$([int]$tab.Alignment)")
                }
            } catch { }
            $lines.Add("P$index|$($range.Start):$($range.End)|style=$styleName|text=$(ConvertTo-DebugText ([string]$range.Text))|fields=$($range.Fields.Count)|shapes=$($range.InlineShapes.Count)|tabs=$($tabEntries -join ',')")
        }
    } finally {
        Release-ComObject $tab
        Release-ComObject $tabs
        Release-ComObject $format
        Release-ComObject $style
        Release-ComObject $range
        Release-ComObject $paragraph
        Release-ComObject $paragraphs
    }

    $fields = $null
    $field = $null
    $code = $null
    $result = $null
    $nested = $null
    try {
        $fields = $document.Fields
        for ($index = 1; $index -le $fields.Count; $index++) {
            Release-ComObject $nested; $nested = $null
            Release-ComObject $result; $result = $null
            Release-ComObject $code; $code = $null
            Release-ComObject $field; $field = $fields.Item($index)
            $code = $field.Code
            $result = $field.Result
            $nested = $code.Fields
            $lines.Add("F$index|type=$([int]$field.Type)|code=$($code.Start):$($code.End)|result=$($result.Start):$($result.End)|nested=$($nested.Count)|showCodes=$($field.ShowCodes)|instruction=$(ConvertTo-DebugText ([string]$code.Text))|value=$(ConvertTo-DebugText ([string]$result.Text))")
        }
    } finally {
        Release-ComObject $nested
        Release-ComObject $result
        Release-ComObject $code
        Release-ComObject $field
        Release-ComObject $fields
    }

    $bookmarks = $null
    $bookmark = $null
    $bookmarkRange = $null
    try {
        $bookmarks = $document.Bookmarks
        $previousShowHidden = [bool]$bookmarks.ShowHidden
        $bookmarks.ShowHidden = $true
        try {
            for ($index = 1; $index -le $bookmarks.Count; $index++) {
                Release-ComObject $bookmarkRange; $bookmarkRange = $null
                Release-ComObject $bookmark; $bookmark = $bookmarks.Item($index)
                $bookmarkRange = $bookmark.Range
                $lines.Add("B$index|name=$($bookmark.Name)|range=$($bookmarkRange.Start):$($bookmarkRange.End)|text=$(ConvertTo-DebugText ([string]$bookmarkRange.Text))")
            }
        } finally { $bookmarks.ShowHidden = $previousShowHidden }
    } finally {
        Release-ComObject $bookmarkRange
        Release-ComObject $bookmark
        Release-ComObject $bookmarks
    }

    foreach ($entry in (Get-PlaceRefEntries $document)) {
        $lines.Add("MT|side=$($entry.Side)|field=$($entry.FieldStart):$($entry.FieldEnd)|paragraph=$($entry.ParagraphStart):$($entry.ParagraphEnd)|code=$(ConvertTo-DebugText $entry.CodeText)")
    }
    [IO.File]::WriteAllLines($path, $lines, [Text.UTF8Encoding]::new($false))
}

$sourcePath = [IO.Path]::GetFullPath((Join-Path (Get-Location) $SourceDocument))
$targetPath = [IO.Path]::GetFullPath((Join-Path (Get-Location) $TargetDocument))
$artifactPath = [IO.Path]::GetFullPath((Join-Path (Get-Location) $ArtifactRoot))
[IO.Directory]::CreateDirectory($artifactPath) | Out-Null
$workingPath = Join-Path $artifactPath "formattedtext-working.docx"
$outputPath = Join-Path $artifactPath "formattedtext-after.docx"
[IO.File]::Copy($targetPath, $workingPath, $true)

$beforeMathTypePids = @(Get-Process MathType* -ErrorAction SilentlyContinue | ForEach-Object Id)
$beforeWordPids = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = $null
$source = $null
$target = $null
try {
    Write-Output "STAGE=CREATE_WORD"
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    Write-Output "STAGE=OPEN_SOURCE"
    $source = $word.Documents.Open($sourcePath, $false, $true, $false)
    Write-Output "STAGE=OPEN_TARGET"
    $target = $word.Documents.Open($workingPath, $false, $false, $false)
    Write-Output "STAGE=DUMP_BEFORE"
    Dump-Document $target (Join-Path $artifactPath "before.txt")

    $sourceEntries = @(Get-PlaceRefEntries $source)
    $targetEntries = @(Get-PlaceRefEntries $target)
    foreach ($side in @("right", "left")) {
        Write-Output "STAGE=REPLACE_$($side.ToUpperInvariant())"
        $sourceEntry = $sourceEntries | Where-Object Side -eq $side | Select-Object -First 1
        $targetEntry = $targetEntries | Where-Object Side -eq $side | Select-Object -First 1
        if ($null -eq $sourceEntry -or $null -eq $targetEntry) { throw "Missing $side MTPlaceRef." }
        $sourceRange = $null
        $targetRange = $null
        $formatted = $null
        try {
            $sourceRange = $source.Range($sourceEntry.FieldStart, $sourceEntry.FieldEnd)
            $targetRange = $target.Range($targetEntry.FieldStart, $targetEntry.FieldEnd)
            $formatted = $sourceRange.FormattedText
            $targetRange.FormattedText = $formatted
        } finally {
            Release-ComObject $formatted
            Release-ComObject $targetRange
            Release-ComObject $sourceRange
        }
        $targetEntries = @(Get-PlaceRefEntries $target)
    }

    Write-Output "STAGE=DUMP_AFTER_LIVE"
    Dump-Document $target (Join-Path $artifactPath "after-live.txt")
    if ($LiveOnly) {
        Write-Output "FORMATTED_TEXT_LIVE_PROBE_COMPLETE=$artifactPath"
        return
    }

    Write-Output "STAGE=CLOSE_SOURCE_BEFORE_SAVE"
    $source.Close(0)
    Release-ComObject $source; $source = $null
    Write-Output "STAGE=SAVE_WORKING_COPY"
    $target.Save()
    Write-Output "STAGE=CLOSE_TARGET"
    $target.Close(0)
    Release-ComObject $target; $target = $null
    [IO.File]::Copy($workingPath, $outputPath, $true)
    Write-Output "STAGE=REOPEN_TARGET"
    $target = $word.Documents.Open($outputPath, $false, $false, $false)
    Write-Output "STAGE=DUMP_AFTER_REOPEN"
    Dump-Document $target (Join-Path $artifactPath "after-reopen.txt")
    Write-Output "FORMATTED_TEXT_PROBE_COMPLETE=$artifactPath"
} finally {
    if ($null -ne $target) { try { $target.Close(0) } catch { } }
    if ($null -ne $source) { try { $source.Close(0) } catch { } }
    Release-ComObject $target
    Release-ComObject $source
    if ($null -ne $word) { try { $word.Quit() } catch { } }
    Release-ComObject $word
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 500
    $afterWordProcesses = @(Get-CimInstance Win32_Process | Where-Object {
        $_.Name -eq "WINWORD.EXE" -and
        $beforeWordPids -notcontains [int]$_.ProcessId -and
        $_.CommandLine -like "*/Automation -Embedding*"
    })
    foreach ($ownedWordProcess in $afterWordProcesses) {
        try {
            Stop-Process -Id ([int]$ownedWordProcess.ProcessId) -Force
            Write-Output "FORCED_CLEANUP_PROBE_WORD=$($ownedWordProcess.ProcessId)"
        } catch { }
    }
    $afterMathTypePids = @(Get-Process MathType* -ErrorAction SilentlyContinue | ForEach-Object Id)
    $newMathTypePids = @($afterMathTypePids | Where-Object { $beforeMathTypePids -notcontains $_ })
    if ($newMathTypePids.Count -gt 0) { throw "Probe started MathType: $($newMathTypePids -join ',')." }
}
