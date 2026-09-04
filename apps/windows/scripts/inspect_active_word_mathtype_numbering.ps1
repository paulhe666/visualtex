param(
    [string]$ArtifactRoot = ""
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

if ([string]::IsNullOrWhiteSpace($ArtifactRoot)) {
    $ArtifactRoot = Join-Path $env:TEMP (
        "VisualTeX\active-word-mathtype-inspection-" +
        [DateTime]::Now.ToString("yyyyMMdd-HHmmss"))
}
$artifactPath = [IO.Path]::GetFullPath($ArtifactRoot)
[IO.Directory]::CreateDirectory($artifactPath) | Out-Null

function Release-ComObject([object]$value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value) } catch { }
    }
}

function Write-Utf8Text([string]$path, [AllowNull()][string]$text) {
    [IO.File]::WriteAllText(
        $path,
        $(if ($null -eq $text) { "" } else { $text }),
        [Text.UTF8Encoding]::new($false))
}

function ConvertTo-DebugText([AllowNull()][string]$text) {
    if ($null -eq $text) { return "" }
    $builder = New-Object Text.StringBuilder
    foreach ($character in $text.ToCharArray()) {
        $codePoint = [int][char]$character
        $token = switch ($codePoint) {
            1 { "<OLE>" }
            7 { "<CELL>" }
            9 { "<TAB>" }
            10 { "<LF>" }
            11 { "<LINE_BREAK>" }
            12 { "<PAGE_BREAK>" }
            13 { "<PARA>" }
            19 { "<FIELD_BEGIN>" }
            20 { "<FIELD_SEPARATOR>" }
            21 { "<FIELD_END>" }
            default { $null }
        }
        if ($null -ne $token) { [void]$builder.Append($token) }
        elseif ([char]::IsControl($character)) {
            [void]$builder.Append(("<U+{0:X4}>" -f $codePoint))
        } else {
            [void]$builder.Append($character)
        }
    }
    return $builder.ToString()
}

function Get-StyleName([object]$range) {
    $style = $null
    try {
        $style = $range.Style
        if ($style -is [string]) { return [string]$style }
        try { return [string]$style.NameLocal } catch { return [string]$style }
    } catch { return "<unreadable>" }
    finally { Release-ComObject $style }
}

$word = $null
$document = $null
$content = $null
$fields = $null
$field = $null
$code = $null
$result = $null
$nestedFields = $null
$paragraphs = $null
$paragraph = $null
$paragraphRange = $null
$paragraphFormat = $null
$tabStops = $null
$tabStop = $null
$inlineShapes = $null
$inlineShape = $null
$shapeRange = $null
$oleFormat = $null
$bookmarks = $null
$bookmark = $null
$bookmarkRange = $null

try {
    # This script is deliberately read-only: no Save, Close, Quit, Selection,
    # field Update, bookmark mutation or document text assignment is performed.
    $word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
    $document = $word.ActiveDocument
    if ($null -eq $document) { throw "Word has no active document." }
    $content = $document.Content

    Write-Output ("ACTIVE_DOCUMENT|name="+$document.Name+"|full="+$document.FullName+"|saved="+$document.Saved)
    Write-Output ("READ_ONLY_INSPECTION_ARTIFACT="+$artifactPath)

    $rawFlatOpc = [string]$content.WordOpenXML
    Write-Utf8Text (Join-Path $artifactPath "content-flat-opc.raw.xml") $rawFlatOpc
    $normalized = [regex]::Replace(
        $rawFlatOpc,
        '<pkg:binaryData>.*?</pkg:binaryData>',
        '<pkg:binaryData>[BINARY OMITTED]</pkg:binaryData>',
        [Text.RegularExpressions.RegexOptions]::Singleline)
    $normalized = [regex]::Replace(
        $normalized,
        '\s+(?:w:rsid[A-Za-z0-9]*|w14:paraId|w14:textId|w14:anchorId)="[^"]*"',
        '')
    Write-Utf8Text (Join-Path $artifactPath "content-flat-opc.normalized.xml") $normalized
    Write-Utf8Text (Join-Path $artifactPath "document-visible-text.txt") (ConvertTo-DebugText ([string]$content.Text))

    $fieldEntries = New-Object Collections.Generic.List[object]
    $fields = $document.Fields
    for ($index = 1; $index -le $fields.Count; $index++) {
        Release-ComObject $nestedFields; $nestedFields = $null
        Release-ComObject $result; $result = $null
        Release-ComObject $code; $code = $null
        Release-ComObject $field; $field = $fields.Item($index)
        $code = $field.Code
        $result = $field.Result
        $nestedFields = $code.Fields
        $fieldEntries.Add([pscustomobject][ordered]@{
            Index = $index
            Type = [int]$field.Type
            Locked = [bool]$field.Locked
            ShowCodes = [bool]$field.ShowCodes
            FieldStart = [int]$code.Start - 1
            CodeStart = [int]$code.Start
            CodeEnd = [int]$code.End
            ResultStart = [int]$result.Start
            ResultEnd = [int]$result.End
            CodeText = ConvertTo-DebugText ([string]$code.Text)
            ResultText = ConvertTo-DebugText ([string]$result.Text)
            NestedCodeFieldCount = [int]$nestedFields.Count
        })
    }
    Write-Utf8Text (Join-Path $artifactPath "fields.json") ($fieldEntries | ConvertTo-Json -Depth 6)

    $paragraphEntries = New-Object Collections.Generic.List[object]
    $paragraphs = $document.Paragraphs
    for ($index = 1; $index -le $paragraphs.Count; $index++) {
        Release-ComObject $tabStop; $tabStop = $null
        Release-ComObject $tabStops; $tabStops = $null
        Release-ComObject $paragraphFormat; $paragraphFormat = $null
        Release-ComObject $paragraphRange; $paragraphRange = $null
        Release-ComObject $paragraph; $paragraph = $paragraphs.Item($index)
        $paragraphRange = $paragraph.Range
        $paragraphFormat = $paragraph.Format
        $tabStops = $paragraphFormat.TabStops
        $tabs = New-Object Collections.Generic.List[object]
        for ($tabIndex = 1; $tabIndex -le $tabStops.Count; $tabIndex++) {
            Release-ComObject $tabStop; $tabStop = $tabStops.Item($tabIndex)
            $tabs.Add([pscustomobject][ordered]@{
                Position = [double]$tabStop.Position
                Alignment = [int]$tabStop.Alignment
                Leader = [int]$tabStop.Leader
            })
        }
        $paragraphEntries.Add([pscustomobject][ordered]@{
            Index = $index
            Start = [int]$paragraphRange.Start
            End = [int]$paragraphRange.End
            Style = Get-StyleName $paragraphRange
            Text = ConvertTo-DebugText ([string]$paragraphRange.Text)
            FieldCount = [int]$paragraphRange.Fields.Count
            InlineShapeCount = [int]$paragraphRange.InlineShapes.Count
            Alignment = [int]$paragraphFormat.Alignment
            Tabs = $tabs.ToArray()
            ContainsVisiblePlaceRefInstruction =
                (([string]$paragraphRange.Text).IndexOf(
                    "MACROBUTTON MTPlaceRef",
                    [StringComparison]::OrdinalIgnoreCase) -ge 0)
        })
    }
    Write-Utf8Text (Join-Path $artifactPath "paragraphs.json") ($paragraphEntries | ConvertTo-Json -Depth 8)

    $shapeEntries = New-Object Collections.Generic.List[object]
    $inlineShapes = $document.InlineShapes
    for ($index = 1; $index -le $inlineShapes.Count; $index++) {
        Release-ComObject $oleFormat; $oleFormat = $null
        Release-ComObject $shapeRange; $shapeRange = $null
        Release-ComObject $inlineShape; $inlineShape = $inlineShapes.Item($index)
        $shapeRange = $inlineShape.Range
        $progId = ""
        try {
            $oleFormat = $inlineShape.OLEFormat
            $progId = [string]$oleFormat.ProgID
        } catch { }
        $shapeEntries.Add([pscustomobject][ordered]@{
            Index = $index
            Type = [int]$inlineShape.Type
            Start = [int]$shapeRange.Start
            End = [int]$shapeRange.End
            ProgId = $progId
            Width = [double]$inlineShape.Width
            Height = [double]$inlineShape.Height
        })
    }
    Write-Utf8Text (Join-Path $artifactPath "inline-shapes.json") ($shapeEntries | ConvertTo-Json -Depth 5)

    $bookmarkEntries = New-Object Collections.Generic.List[object]
    $bookmarks = $document.Bookmarks
    for ($index = 1; $index -le $bookmarks.Count; $index++) {
        Release-ComObject $bookmarkRange; $bookmarkRange = $null
        Release-ComObject $bookmark; $bookmark = $bookmarks.Item($index)
        $bookmarkRange = $bookmark.Range
        $bookmarkEntries.Add([pscustomobject][ordered]@{
            Index = $index
            Name = [string]$bookmark.Name
            Start = [int]$bookmarkRange.Start
            End = [int]$bookmarkRange.End
            Text = ConvertTo-DebugText ([string]$bookmarkRange.Text)
        })
    }
    Write-Utf8Text (Join-Path $artifactPath "bookmarks.json") ($bookmarkEntries | ConvertTo-Json -Depth 5)

    $summary = [pscustomobject][ordered]@{
        Name = [string]$document.Name
        FullName = [string]$document.FullName
        Saved = [bool]$document.Saved
        Start = [int]$content.Start
        End = [int]$content.End
        Fields = [int]$document.Fields.Count
        Paragraphs = [int]$document.Paragraphs.Count
        InlineShapes = [int]$document.InlineShapes.Count
        Bookmarks = [int]$document.Bookmarks.Count
        VisiblePlaceRefInstructionCount = @(
            $paragraphEntries |
                Where-Object { $_.ContainsVisiblePlaceRefInstruction }).Count
        CapturedAt = [DateTimeOffset]::Now.ToString("O")
    }
    Write-Utf8Text (Join-Path $artifactPath "summary.json") ($summary | ConvertTo-Json -Depth 5)
    Write-Output ("READ_ONLY_INSPECTION_COMPLETE|fields="+$summary.Fields+"|paragraphs="+$summary.Paragraphs+"|shapes="+$summary.InlineShapes+"|bookmarks="+$summary.Bookmarks+"|visiblePlaceRefParagraphs="+$summary.VisiblePlaceRefInstructionCount)
}
finally {
    Release-ComObject $bookmarkRange
    Release-ComObject $bookmark
    Release-ComObject $bookmarks
    Release-ComObject $oleFormat
    Release-ComObject $shapeRange
    Release-ComObject $inlineShape
    Release-ComObject $inlineShapes
    Release-ComObject $tabStop
    Release-ComObject $tabStops
    Release-ComObject $paragraphFormat
    Release-ComObject $paragraphRange
    Release-ComObject $paragraph
    Release-ComObject $paragraphs
    Release-ComObject $nestedFields
    Release-ComObject $result
    Release-ComObject $code
    Release-ComObject $field
    Release-ComObject $fields
    Release-ComObject $content
    Release-ComObject $document
    Release-ComObject $word
}
