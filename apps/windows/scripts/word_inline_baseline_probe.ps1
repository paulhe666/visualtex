[CmdletBinding()]
param(
    [string]$SourceDocumentName = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [Text.UTF8Encoding]::new()
Add-Type -AssemblyName System.Drawing

function Release-Com([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) } catch {}
    }
}

function Insert-Text([object]$Document, [object]$State, [string]$Text) {
    $start = [int]$State.Position
    $range = $Document.Range($start, $start)
    try {
        $range.Text = $Text
        $State.Position = [int]$range.End
        return [pscustomobject]@{ Start = $start; End = [int]$range.End }
    }
    finally {
        Release-Com $range
    }
}

function Format-BodyText([object]$Document, [int]$Start, [int]$End) {
    $range = $Document.Range($Start, $End)
    try {
        $range.Font.Name = "Times New Roman"
        $range.Font.Size = 10.5
        $range.Font.Position = 0
    }
    finally {
        Release-Com $range
    }
}

function Clone-FormattedRange(
    [object]$Document,
    [object]$State,
    [object]$SourceRange
) {
    $start = [int]$State.Position
    $destination = $Document.Range($start, $start)
    try {
        $destination.FormattedText = $SourceRange.FormattedText
        $State.Position = [int]$destination.End
        return [pscustomobject]@{ Start = $start; End = [int]$destination.End }
    }
    finally {
        Release-Com $destination
    }
}

function Set-OleObjectPosition(
    [object]$Document,
    [int]$Start,
    [int]$End,
    [int]$Position
) {
    $whole = $Document.Range($Start, $End)
    try {
        $whole.Font.Position = 0
    }
    finally {
        Release-Com $whole
    }

    $found = $false
    for ($cursor = $Start; $cursor -lt $End; $cursor++) {
        $character = $Document.Range($cursor, $cursor + 1)
        try {
            if ($character.Text -eq [char]1) {
                $character.Font.Position = $Position
                $found = $true
            }
        }
        finally {
            Release-Com $character
        }
    }
    if (-not $found) {
        throw "Copied OLE range has no U+0001 object character."
    }
}

function Read-X([object]$Document, [int]$Position) {
    $probe = $Document.Range($Position, $Position)
    try {
        return [double]$probe.Information(5) # wdHorizontalPositionRelativeToPage
    }
    finally {
        Release-Com $probe
    }
}

function Measure-Crop(
    [Drawing.Bitmap]$Bitmap,
    [double]$LeftPixels,
    [double]$RightPixels
) {
    $left = [Math]::Max(0, [int][Math]::Floor($LeftPixels))
    $right = [Math]::Min($Bitmap.Width, [int][Math]::Ceiling($RightPixels))
    if ($right -le $left) {
        throw "Invalid crop $left..$right for bitmap width $($Bitmap.Width)."
    }

    $minX = $right
    $maxX = -1
    $minY = $Bitmap.Height
    $maxY = -1
    [double]$weightedY = 0
    [double]$totalWeight = 0
    [int]$darkPixels = 0
    for ($y = 0; $y -lt $Bitmap.Height; $y++) {
        for ($x = $left; $x -lt $right; $x++) {
            $pixel = $Bitmap.GetPixel($x, $y)
            $darkness = 255.0 - (($pixel.R + $pixel.G + $pixel.B) / 3.0)
            if ($darkness -le 18.0) { continue }
            $darkPixels++
            if ($x -lt $minX) { $minX = $x }
            if ($x -gt $maxX) { $maxX = $x }
            if ($y -lt $minY) { $minY = $y }
            if ($y -gt $maxY) { $maxY = $y }
            $weightedY += $y * $darkness
            $totalWeight += $darkness
        }
    }
    if ($maxY -lt 0 -or $totalWeight -le 0) {
        throw "No visible ink in crop $left..$right."
    }

    return [pscustomobject]@{
        Left = $minX
        Right = $maxX
        Top = $minY
        Bottom = $maxY
        Height = $maxY - $minY + 1
        CentroidY = $weightedY / $totalWeight
        DarkPixels = $darkPixels
    }
}

function Get-Spread([double[]]$Values) {
    $measurement = $Values | Measure-Object -Minimum -Maximum
    return [double]$measurement.Maximum - [double]$measurement.Minimum
}

$word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
$source = $null
$temp = $null
$screenUpdating = [bool]$word.ScreenUpdating
try {
    if ([string]::IsNullOrWhiteSpace($SourceDocumentName)) {
        $source = $word.ActiveDocument
    }
    else {
        foreach ($document in @($word.Documents)) {
            if ([string]::Equals(
                    [string]$document.Name,
                    $SourceDocumentName,
                    [StringComparison]::OrdinalIgnoreCase)) {
                $source = $document
                break
            }
            Release-Com $document
        }
        if ($null -eq $source) {
            throw "Requested Word source document is not open."
        }
    }
    if ($source.InlineShapes.Count -lt 2 -or $source.OMaths.Count -lt 2) {
        throw "Source document does not contain the two required OLE/OMML pairs."
    }

    $sourceState = [pscustomobject]@{
        Saved = [bool]$source.Saved
        Characters = [int]$source.Content.End
        InlineShapes = [int]$source.InlineShapes.Count
        OMaths = [int]$source.OMaths.Count
    }

    $word.ScreenUpdating = $false
    $temp = $word.Documents.Add()
    $state = [pscustomobject]@{ Position = 0 }
    $records = [Collections.Generic.List[object]]::new()
    $cases = @(
        [pscustomobject]@{ Name = "L_old"; Pair = 1; Position = -1 },
        [pscustomobject]@{ Name = "L_new"; Pair = 1; Position = -2 },
        [pscustomobject]@{ Name = "Integral_old"; Pair = 2; Position = -3 },
        [pscustomobject]@{ Name = "Integral_new"; Pair = 2; Position = -4 }
    )

    foreach ($case in $cases) {
        $lineStart = [int]$state.Position

        $h1 = Insert-Text $temp $state "H  "
        Format-BodyText $temp $h1.Start $h1.End

        $sourceShape = $source.InlineShapes.Item([int]$case.Pair)
        $sourceShapeRange = $sourceShape.Range
        try {
            [void](Clone-FormattedRange $temp $state $sourceShapeRange)
        }
        finally {
            Release-Com $sourceShapeRange
            Release-Com $sourceShape
        }

        $copiedShape = $temp.InlineShapes.Item($temp.InlineShapes.Count)
        $copiedShapeRange = $copiedShape.Range
        try {
            $oleStart = [int]$copiedShapeRange.Start
            $oleEnd = [int]$copiedShapeRange.End
            $oleWidthPoints = [double]$copiedShape.Width
            Set-OleObjectPosition $temp $oleStart $oleEnd ([int]$case.Position)
        }
        finally {
            Release-Com $copiedShapeRange
            Release-Com $copiedShape
        }

        $h2 = Insert-Text $temp $state "  H  "
        Format-BodyText $temp $h2.Start $h2.End

        $sourceMath = $source.OMaths.Item([int]$case.Pair)
        $sourceMathRange = $sourceMath.Range
        try {
            [void](Clone-FormattedRange $temp $state $sourceMathRange)
        }
        finally {
            Release-Com $sourceMathRange
            Release-Com $sourceMath
        }

        $copiedMath = $temp.OMaths.Item($temp.OMaths.Count)
        $copiedMathRange = $copiedMath.Range
        try {
            $ommlStart = [int]$copiedMathRange.Start
            $ommlEnd = [int]$copiedMathRange.End
        }
        finally {
            Release-Com $copiedMathRange
            Release-Com $copiedMath
        }

        $h3 = Insert-Text $temp $state "  H"
        Format-BodyText $temp $h3.Start $h3.End

        $lineEnd = [int]$state.Position
        [void](Insert-Text $temp $state "`r")
        $paragraphRange = $temp.Range($lineStart, [int]$state.Position)
        try {
            $paragraphRange.ParagraphFormat.SpaceBefore = 0
            $paragraphRange.ParagraphFormat.SpaceAfter = 0
            $paragraphRange.ParagraphFormat.LineSpacingRule = 0 # wdLineSpaceSingle
        }
        finally {
            Release-Com $paragraphRange
        }

        $records.Add([pscustomobject]@{
            Name = [string]$case.Name
            Position = [int]$case.Position
            LineStart = $lineStart
            LineEnd = $lineEnd
            H1Start = [int]$h1.Start
            H1End = [int]($h1.Start + 1)
            H2Start = [int]($h2.Start + 2)
            H2End = [int]($h2.Start + 3)
            H3Start = [int]($h3.Start + 2)
            H3End = [int]($h3.Start + 3)
            OleStart = $oleStart
            OleEnd = $oleEnd
            OleWidthPoints = $oleWidthPoints
            OmmlStart = $ommlStart
            OmmlEnd = $ommlEnd
        })
    }

    $temp.Repaginate()
    Start-Sleep -Milliseconds 120

    $metrics = [Collections.Generic.List[object]]::new()
    foreach ($record in $records) {
        $lineRange = $temp.Range([int]$record.LineStart, [int]$record.LineEnd)
        try {
            [byte[]]$bytes = $lineRange.EnhMetaFileBits
        }
        finally {
            Release-Com $lineRange
        }
        if ($null -eq $bytes -or $bytes.Length -le 0) {
            throw "Word returned no EMF for $($record.Name)."
        }

        $stream = [IO.MemoryStream]::new($bytes, $false)
        $metafile = [Drawing.Imaging.Metafile]::new($stream)
        $bitmap = [Drawing.Bitmap]::new(
            $metafile.Width,
            $metafile.Height,
            [Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        try {
            $graphics.Clear([Drawing.Color]::White)
            $graphics.DrawImage($metafile, 0, 0, $bitmap.Width, $bitmap.Height)
        }
        finally {
            $graphics.Dispose()
        }

        try {
            $dpiX = [double]$metafile.HorizontalResolution
            $dpiY = [double]$metafile.VerticalResolution
            $originX = Read-X $temp ([int]$record.LineStart)
            $oleX = Read-X $temp ([int]$record.OleStart)
            $ommlX1 = Read-X $temp ([int]$record.OmmlStart)
            $ommlX2 = Read-X $temp ([int]$record.OmmlEnd)
            $h1X1 = Read-X $temp ([int]$record.H1Start)
            $h1X2 = Read-X $temp ([int]$record.H1End)
            $h2X1 = Read-X $temp ([int]$record.H2Start)
            $h2X2 = Read-X $temp ([int]$record.H2End)
            $h3X1 = Read-X $temp ([int]$record.H3Start)
            $h3X2 = Read-X $temp ([int]$record.H3End)

            $ptToPxX = $dpiX / 72.0
            $marginPx = 1.0 * $ptToPxX
            $toPixel = {
                param([double]$PageX)
                return ($PageX - $originX) * $ptToPxX
            }

            $oleLeftPx = & $toPixel $oleX
            $oleRightPx = & $toPixel ($oleX + [double]$record.OleWidthPoints)
            $ommlLeftPx = & $toPixel $ommlX1
            $ommlRightPx = & $toPixel $ommlX2
            $ole = Measure-Crop `
                $bitmap `
                ($oleLeftPx - $marginPx) `
                ($oleRightPx + $marginPx)
            $omml = Measure-Crop `
                $bitmap `
                ($ommlLeftPx - $marginPx) `
                ($ommlRightPx + $marginPx)
            # The right 55% excludes the integral/operator and its limits while
            # retaining ordinary baseline-bearing glyphs. It also selects the
            # second L in L_zL^2. This is measurement-only; production never
            # classifies formulas or applies a per-expression adjustment.
            $oleBody = Measure-Crop `
                $bitmap `
                ($oleLeftPx + ($oleRightPx - $oleLeftPx) * 0.45) `
                $oleRightPx
            $ommlBody = Measure-Crop `
                $bitmap `
                ($ommlLeftPx + ($ommlRightPx - $ommlLeftPx) * 0.45) `
                $ommlRightPx
            $h1Ink = Measure-Crop `
                $bitmap `
                ((& $toPixel $h1X1) - $marginPx) `
                ((& $toPixel $h1X2) + $marginPx)
            $h2Ink = Measure-Crop `
                $bitmap `
                ((& $toPixel $h2X1) - $marginPx) `
                ((& $toPixel $h2X2) + $marginPx)
            $h3Ink = Measure-Crop `
                $bitmap `
                ((& $toPixel $h3X1) - $marginPx) `
                ((& $toPixel $h3X2) + $marginPx)

            $pxToPtY = 72.0 / $dpiY
            $hBottom = ($h1Ink.Bottom + $h2Ink.Bottom + $h3Ink.Bottom) / 3.0
            $metrics.Add([pscustomobject]@{
                Name = $record.Name
                Position = $record.Position
                EmfPixels = "$($bitmap.Width)x$($bitmap.Height)"
                Dpi = [Math]::Round($dpiY, 2)
                OleTopPt = [Math]::Round($ole.Top * $pxToPtY, 3)
                OleBottomPt = [Math]::Round($ole.Bottom * $pxToPtY, 3)
                OleHeightPt = [Math]::Round($ole.Height * $pxToPtY, 3)
                OleCentroidPt = [Math]::Round($ole.CentroidY * $pxToPtY, 3)
                OmmlTopPt = [Math]::Round($omml.Top * $pxToPtY, 3)
                OmmlBottomPt = [Math]::Round($omml.Bottom * $pxToPtY, 3)
                OmmlHeightPt = [Math]::Round($omml.Height * $pxToPtY, 3)
                OmmlCentroidPt = [Math]::Round($omml.CentroidY * $pxToPtY, 3)
                DeltaTopPt = [Math]::Round(($ole.Top - $omml.Top) * $pxToPtY, 3)
                DeltaBottomPt = [Math]::Round(($ole.Bottom - $omml.Bottom) * $pxToPtY, 3)
                DeltaCentroidPt = [Math]::Round(($ole.CentroidY - $omml.CentroidY) * $pxToPtY, 3)
                BodyDeltaTopPt = [Math]::Round(($oleBody.Top - $ommlBody.Top) * $pxToPtY, 3)
                BodyDeltaBottomPt = [Math]::Round(($oleBody.Bottom - $ommlBody.Bottom) * $pxToPtY, 3)
                BodyDeltaCentroidPt = [Math]::Round(
                    ($oleBody.CentroidY - $ommlBody.CentroidY) * $pxToPtY,
                    3)
                OleBottomVsHPt = [Math]::Round(($ole.Bottom - $hBottom) * $pxToPtY, 3)
                OmmlBottomVsHPt = [Math]::Round(($omml.Bottom - $hBottom) * $pxToPtY, 3)
                OleBodyBottomVsHPt = [Math]::Round(($oleBody.Bottom - $hBottom) * $pxToPtY, 3)
                OmmlBodyBottomVsHPt = [Math]::Round(($ommlBody.Bottom - $hBottom) * $pxToPtY, 3)
                HTopSpreadPt = [Math]::Round(
                    (Get-Spread @($h1Ink.Top, $h2Ink.Top, $h3Ink.Top)) * $pxToPtY,
                    3)
                HBottomSpreadPt = [Math]::Round(
                    (Get-Spread @($h1Ink.Bottom, $h2Ink.Bottom, $h3Ink.Bottom)) * $pxToPtY,
                    3)
            })
        }
        finally {
            $bitmap.Dispose()
            $metafile.Dispose()
            $stream.Dispose()
        }
    }

    $currentSourceState = [pscustomobject]@{
        Saved = [bool]$source.Saved
        Characters = [int]$source.Content.End
        InlineShapes = [int]$source.InlineShapes.Count
        OMaths = [int]$source.OMaths.Count
    }
    if (
        $currentSourceState.Saved -ne $sourceState.Saved -or
        $currentSourceState.Characters -ne $sourceState.Characters -or
        $currentSourceState.InlineShapes -ne $sourceState.InlineShapes -or
        $currentSourceState.OMaths -ne $sourceState.OMaths
    ) {
        throw "Source document state changed during the read-only probe."
    }

    [pscustomobject]@{
        Source = [string]$source.Name
        SourceUnchanged = $true
        TemporaryDocumentSaved = [bool]$temp.Saved
        Metrics = $metrics
    } | ConvertTo-Json -Depth 6
}
finally {
    if ($null -ne $temp) {
        try { $temp.Close(0) } catch {} # wdDoNotSaveChanges
        Release-Com $temp
    }
    if ($null -ne $source) {
        try { $source.Activate() } catch {}
        Release-Com $source
    }
    try { $word.ScreenUpdating = $screenUpdating } catch {}
    Release-Com $word
}
