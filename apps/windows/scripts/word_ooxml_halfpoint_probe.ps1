[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourcePath,
    [Parameter(Mandatory = $true)]
    [string]$PatchedPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
Add-Type -AssemblyName System.Drawing

function Release-Com([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) } catch {}
    }
}

function Read-X([object]$Document, [int]$Position) {
    $range = $Document.Range($Position, $Position)
    try { return [double]$range.Information(5) }
    finally { Release-Com $range }
}

function Measure-Crop(
    [Drawing.Bitmap]$Bitmap,
    [double]$LeftPixels,
    [double]$RightPixels
) {
    $left = [Math]::Max(0, [int][Math]::Floor($LeftPixels))
    $right = [Math]::Min($Bitmap.Width, [int][Math]::Ceiling($RightPixels))
    if ($right -le $left) { throw "Invalid crop $left..$right." }
    $minY = $Bitmap.Height
    $maxY = -1
    [double]$weightedY = 0
    [double]$totalWeight = 0
    for ($y = 0; $y -lt $Bitmap.Height; $y++) {
        for ($x = $left; $x -lt $right; $x++) {
            $pixel = $Bitmap.GetPixel($x, $y)
            $darkness = 255.0 - (($pixel.R + $pixel.G + $pixel.B) / 3.0)
            if ($darkness -le 18.0) { continue }
            $minY = [Math]::Min($minY, $y)
            $maxY = [Math]::Max($maxY, $y)
            $weightedY += $y * $darkness
            $totalWeight += $darkness
        }
    }
    if ($maxY -lt 0 -or $totalWeight -le 0) { throw "No visible ink in crop." }
    [pscustomobject]@{
        Top = $minY
        Bottom = $maxY
        Centroid = $weightedY / $totalWeight
    }
}

function Measure-Document([object]$Word, [string]$Path, [string]$Label) {
    $doc = $null
    $shape = $null
    $shapeRange = $null
    $paragraph = $null
    $paragraphRange = $null
    $stream = $null
    $metafile = $null
    $bitmap = $null
    try {
        $doc = $Word.Documents.Open(
            [IO.Path]::GetFullPath($Path),
            $false,
            $true,
            $false)
        $doc.Repaginate()
        Start-Sleep -Milliseconds 150
        if ([int]$doc.InlineShapes.Count -lt 1) {
            throw "$Label contains no inline shape."
        }
        $shape = $doc.InlineShapes.Item(1)
        $shapeRange = $shape.Range
        $shapeStart = [int]$shapeRange.Start
        $shapeEnd = [int]$shapeRange.End
        $position = [double]$shapeRange.Font.Position
        $wordOpenXml = [string]$shapeRange.WordOpenXML
        $positionMatch = [regex]::Match(
            $wordOpenXml,
            '<w:position\b[^>]*\bw:val="([^"]+)"[^>]*/?>',
            [Text.RegularExpressions.RegexOptions]::IgnoreCase)
        $xmlHalfPoints = if ($positionMatch.Success) {
            [int]$positionMatch.Groups[1].Value
        } else { $null }

        $paragraph = $shapeRange.Paragraphs.Item(1)
        $paragraphRange = $paragraph.Range.Duplicate
        if ($paragraphRange.End -gt $paragraphRange.Start) {
            $paragraphRange.End = $paragraphRange.End - 1
        }
        [byte[]]$bytes = $paragraphRange.EnhMetaFileBits
        if ($null -eq $bytes -or $bytes.Length -eq 0) {
            throw "$Label returned no line EMF."
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
        finally { $graphics.Dispose() }

        $dpiX = [double]$metafile.HorizontalResolution
        $dpiY = [double]$metafile.VerticalResolution
        $originX = Read-X $doc ([int]$paragraphRange.Start)
        $shapeX = Read-X $doc $shapeStart
        $textStart = [int]$paragraphRange.Start
        $hX1 = Read-X $doc $textStart
        $hX2 = Read-X $doc ($textStart + 1)
        $ptToPxX = $dpiX / 72.0
        $pxToPtY = 72.0 / $dpiY
        $toPixel = {
            param([double]$PageX)
            ($PageX - $originX) * $ptToPxX
        }
        $marginPx = $ptToPxX
        $oleInk = Measure-Crop `
            $bitmap `
            ((& $toPixel $shapeX) - $marginPx) `
            ((& $toPixel ($shapeX + [double]$shape.Width)) + $marginPx)
        $hInk = Measure-Crop `
            $bitmap `
            ((& $toPixel $hX1) - $marginPx) `
            ((& $toPixel $hX2) + $marginPx)

        [pscustomobject]@{
            Label = $Label
            Path = [IO.Path]::GetFullPath($Path)
            ComFontPositionPt = $position
            XmlPositionHalfPoints = $xmlHalfPoints
            ObjectHeightPt = [Math]::Round([double]$shape.Height, 4)
            OleTopPt = [Math]::Round($oleInk.Top * $pxToPtY, 4)
            OleBottomPt = [Math]::Round($oleInk.Bottom * $pxToPtY, 4)
            OleCentroidPt = [Math]::Round($oleInk.Centroid * $pxToPtY, 4)
            HTopPt = [Math]::Round($hInk.Top * $pxToPtY, 4)
            HBottomPt = [Math]::Round($hInk.Bottom * $pxToPtY, 4)
            HCentroidPt = [Math]::Round($hInk.Centroid * $pxToPtY, 4)
            OleBottomVsHPt = [Math]::Round(($oleInk.Bottom - $hInk.Bottom) * $pxToPtY, 4)
            OleCentroidVsHPt = [Math]::Round(($oleInk.Centroid - $hInk.Centroid) * $pxToPtY, 4)
            DpiY = [Math]::Round($dpiY, 3)
            EmfPixels = "$($bitmap.Width)x$($bitmap.Height)"
            ShapeStart = $shapeStart
            ShapeEnd = $shapeEnd
        }
    }
    finally {
        if ($null -ne $bitmap) { $bitmap.Dispose() }
        if ($null -ne $metafile) { $metafile.Dispose() }
        if ($null -ne $stream) { $stream.Dispose() }
        Release-Com $paragraphRange
        Release-Com $paragraph
        Release-Com $shapeRange
        Release-Com $shape
        if ($null -ne $doc) {
            try { $doc.Close(0) } catch {}
        }
        Release-Com $doc
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$word.ScreenUpdating = $false
try {
    $source = Measure-Document $word $SourcePath "source"
    $patched = Measure-Document $word $PatchedPath "minus-half-point"
    [pscustomobject]@{
        Source = $source
        Patched = $patched
        DeltaOleTopPt = [Math]::Round($patched.OleTopPt - $source.OleTopPt, 4)
        DeltaOleBottomPt = [Math]::Round($patched.OleBottomPt - $source.OleBottomPt, 4)
        DeltaOleCentroidPt = [Math]::Round($patched.OleCentroidPt - $source.OleCentroidPt, 4)
        DeltaHBottomPt = [Math]::Round($patched.HBottomPt - $source.HBottomPt, 4)
    } | ConvertTo-Json -Depth 6
}
finally {
    try { $word.Quit(0) } catch {}
    Release-Com $word
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
