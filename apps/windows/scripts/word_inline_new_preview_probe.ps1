[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestPath,
    [string]$SourceDocumentName = "",
    [string]$OutputPath = "",
    [switch]$CompactOutput
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

function Test-IsRejectedOfficeCall([Exception]$Exception) {
    $current = $Exception
    while ($null -ne $current) {
        if (
            $current.HResult -eq [int]0x80010001 -or
            $current.HResult -eq [int]0x8001010A
        ) {
            return $true
        }
        $current = $current.InnerException
    }
    return $false
}

function Invoke-WithOfficeRetry(
    [scriptblock]$Action,
    [int]$TimeoutSeconds = 30
) {
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    while ($true) {
        try {
            return & $Action
        }
        catch {
            if (
                -not (Test-IsRejectedOfficeCall $_.Exception) -or
                [DateTime]::UtcNow -ge $deadline
            ) {
                throw
            }
            [Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 100
        }
    }
}

function Release-Com([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) } catch {}
    }
}

function Add-TrackingBookmark(
    [object]$Document,
    [string]$Name,
    [int]$Start,
    [int]$End
) {
    $range = $null
    $bookmarks = $null
    $bookmark = $null
    try {
        $range = $Document.Range($Start, $End)
        $bookmarks = $Document.Bookmarks
        if ($bookmarks.Exists($Name)) {
            $bookmark = $bookmarks.Item($Name)
            $bookmark.Delete()
            Release-Com $bookmark
            $bookmark = $null
        }
        $bookmark = $bookmarks.Add($Name, $range)
        return $Name
    }
    finally {
        Release-Com $bookmark
        Release-Com $bookmarks
        Release-Com $range
    }
}

function Get-TrackingBookmarkBounds([object]$Document, [string]$Name) {
    $bookmarks = $null
    $bookmark = $null
    $range = $null
    try {
        $bookmarks = $Document.Bookmarks
        if (-not $bookmarks.Exists($Name)) {
            throw "Tracking bookmark '$Name' no longer exists."
        }
        $bookmark = $bookmarks.Item($Name)
        $range = $bookmark.Range
        return [pscustomobject]@{
            Start = [int]$range.Start
            End = [int]$range.End
        }
    }
    finally {
        Release-Com $range
        Release-Com $bookmark
        Release-Com $bookmarks
    }
}

function Formula-BookmarkName([string]$Prefix, [string]$FormulaId) {
    return $Prefix + ([guid]$FormulaId).ToString("N")
}

function Insert-Text([object]$Document, [object]$State, [string]$Text) {
    $start = [int]$State.Position
    $range = $Document.Range($start, $start)
    try {
        $range.Text = $Text
        $State.Position = [int]$range.End
        return [pscustomobject]@{ Start = $start; End = [int]$range.End }
    }
    finally { Release-Com $range }
}

function Format-BodyText(
    [object]$Document,
    [int]$Start,
    [int]$End,
    [double]$FontSizePt
) {
    $range = $Document.Range($Start, $End)
    try {
        $range.Font.Name = "Times New Roman"
        $range.Font.Size = $FontSizePt
        $range.Font.Position = 0
    }
    finally { Release-Com $range }
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
    finally { Release-Com $destination }
}

function Read-X([object]$Document, [int]$Position) {
    $probe = $Document.Range($Position, $Position)
    try {
        return [double]$probe.Information(5) # wdHorizontalPositionRelativeToPage
    }
    finally { Release-Com $probe }
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

function Measure-RightmostInkCluster(
    [Drawing.Bitmap]$Bitmap,
    [double]$LeftPixels,
    [double]$RightPixels,
    [int]$MinimumBlankGapPixels = 4
) {
    $left = [Math]::Max(0, [int][Math]::Floor($LeftPixels))
    $right = [Math]::Min($Bitmap.Width, [int][Math]::Ceiling($RightPixels))
    if ($right -le $left) {
        throw "Invalid cluster crop $left..$right for bitmap width $($Bitmap.Width)."
    }

    $occupied = [bool[]]::new($right - $left)
    for ($x = $left; $x -lt $right; $x++) {
        for ($y = 0; $y -lt $Bitmap.Height; $y++) {
            $pixel = $Bitmap.GetPixel($x, $y)
            $darkness = 255.0 - (($pixel.R + $pixel.G + $pixel.B) / 3.0)
            if ($darkness -gt 18.0) {
                $occupied[$x - $left] = $true
                break
            }
        }
    }

    $rightmost = -1
    for ($index = $occupied.Length - 1; $index -ge 0; $index--) {
        if ($occupied[$index]) {
            $rightmost = $index
            break
        }
    }
    if ($rightmost -lt 0) {
        throw "No visible ink in rightmost-cluster crop $left..$right."
    }

    $clusterLeft = 0
    $blankRun = 0
    for ($index = $rightmost - 1; $index -ge 0; $index--) {
        if ($occupied[$index]) {
            $blankRun = 0
            continue
        }
        $blankRun++
        if ($blankRun -ge $MinimumBlankGapPixels) {
            $clusterLeft = $index + $blankRun
            break
        }
    }

    return Measure-Crop `
        $Bitmap `
        ($left + $clusterLeft) `
        ($left + $rightmost + 1)
}

function Get-HorizontalInkClusters(
    [Drawing.Bitmap]$Bitmap,
    [int]$MinimumBlankGapPixels = 4
) {
    $occupied = [bool[]]::new($Bitmap.Width)
    for ($x = 0; $x -lt $Bitmap.Width; $x++) {
        for ($y = 0; $y -lt $Bitmap.Height; $y++) {
            $pixel = $Bitmap.GetPixel($x, $y)
            $darkness = 255.0 - (($pixel.R + $pixel.G + $pixel.B) / 3.0)
            if ($darkness -gt 18.0) {
                $occupied[$x] = $true
                break
            }
        }
    }

    $clusters = [Collections.Generic.List[object]]::new()
    $start = -1
    $lastInk = -1
    $blankRun = 0
    for ($x = 0; $x -lt $occupied.Length; $x++) {
        if ($occupied[$x]) {
            if ($start -lt 0) { $start = $x }
            $lastInk = $x
            $blankRun = 0
            continue
        }
        if ($start -lt 0) { continue }
        $blankRun++
        if ($blankRun -ge $MinimumBlankGapPixels) {
            $clusters.Add([pscustomobject]@{ Left = $start; Right = $lastInk })
            $start = -1
            $lastInk = -1
            $blankRun = 0
        }
    }
    if ($start -ge 0) {
        $clusters.Add([pscustomobject]@{ Left = $start; Right = $lastInk })
    }
    return @($clusters)
}

function Measure-PngInkBounds([string]$PngPath) {
    $bitmap = [Drawing.Bitmap]::new($PngPath)
    try {
        $minX = $bitmap.Width
        $maxX = -1
        $minY = $bitmap.Height
        $maxY = -1
        for ($y = 0; $y -lt $bitmap.Height; $y++) {
            for ($x = 0; $x -lt $bitmap.Width; $x++) {
                if ($bitmap.GetPixel($x, $y).A -le 8) { continue }
                if ($x -lt $minX) { $minX = $x }
                if ($x -gt $maxX) { $maxX = $x }
                if ($y -lt $minY) { $minY = $y }
                if ($y -gt $maxY) { $maxY = $y }
            }
        }
        if ($maxX -lt 0 -or $maxY -lt 0) {
            throw "Generated preview PNG contains no visible ink."
        }
        return [pscustomobject]@{
            Width = $bitmap.Width
            Height = $bitmap.Height
            Left = $minX
            Top = $minY
            Right = $maxX
            Bottom = $maxY
            InkWidth = $maxX - $minX + 1
            InkHeight = $maxY - $minY + 1
            LeftMargin = $minX
            TopMargin = $minY
            RightMargin = $bitmap.Width - 1 - $maxX
            BottomMargin = $bitmap.Height - 1 - $maxY
        }
    }
    finally { $bitmap.Dispose() }
}

function Get-Spread([double[]]$Values) {
    $measurement = $Values | Measure-Object -Minimum -Maximum
    return [double]$measurement.Maximum - [double]$measurement.Minimum
}

function Get-OptionalProperty(
    [object]$Value,
    [string]$Name,
    [object]$DefaultValue = $null
) {
    $property = $Value.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) {
        return $DefaultValue
    }
    return $property.Value
}

function Get-InternalProperty([object]$Value, [string]$Name) {
    $flags = [Reflection.BindingFlags]"Instance,Public,NonPublic"
    $property = $Value.GetType().GetProperty($Name, $flags)
    if ($null -eq $property) { throw "Internal preview property '$Name' is unavailable." }
    return $property.GetValue($Value)
}

function Read-InlineObjectPosition([object]$Document, [int]$Start, [int]$End) {
    $deadline = [DateTime]::UtcNow.AddSeconds(5)
    do {
        $probeEnd = [Math]::Min([int]$Document.Content.End, $End + 8)
        for ($cursor = $Start; $cursor -lt $probeEnd; $cursor++) {
            $character = $Document.Range($cursor, $cursor + 1)
            try {
                if ($character.Text -eq [char]1) {
                    return [double]$character.Font.Position
                }
            }
            finally { Release-Com $character }
        }
        [Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "Inserted OLE range has no U+0001 object character after Word materialization settled."
}

function Set-InlineObjectPosition(
    [object]$Document,
    [int]$Start,
    [int]$End,
    [int]$Position
) {
    $range = $Document.Range($Start, $End)
    try { $range.Font.Position = 0 }
    finally { Release-Com $range }
    $deadline = [DateTime]::UtcNow.AddSeconds(5)
    do {
        $probeEnd = [Math]::Min([int]$Document.Content.End, $End + 8)
        for ($cursor = $Start; $cursor -lt $probeEnd; $cursor++) {
            $character = $Document.Range($cursor, $cursor + 1)
            try {
                if ($character.Text -eq [char]1) {
                    $character.Font.Position = $Position
                    return
                }
            }
            finally { Release-Com $character }
        }
        [Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 100
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "Inserted OLE range has no U+0001 object character after Word materialization settled."
}

$manifest = Get-Content (Resolve-Path $ManifestPath) -Raw | ConvertFrom-Json
if ($null -eq $manifest.items -or @($manifest.items).Count -lt 1) {
    throw "The preview manifest contains no formula items."
}

$root = Split-Path -Parent $PSScriptRoot
$vstoOutput = Join-Path $root "src-windows\VisualTeX.WordVsto\bin\x64\Release\net472"
$contractsPath = Join-Path $vstoOutput "VisualTeX.WindowsOffice.Contracts.dll"
$wordVstoPath = Join-Path $vstoOutput "VisualTeX.WordVsto.dll"
if (-not (Test-Path $contractsPath) -or -not (Test-Path $wordVstoPath)) {
    throw "The current x64 Release Word VSTO output is unavailable."
}

foreach ($dependency in @(
    "System.Runtime.CompilerServices.Unsafe.dll",
    "System.Memory.dll",
    "System.Buffers.dll",
    "Microsoft.Bcl.AsyncInterfaces.dll",
    "System.Threading.Tasks.Extensions.dll",
    "System.Text.Encodings.Web.dll",
    "System.Text.Json.dll",
    "Microsoft.Office.Interop.Word.dll"
)) {
    $dependencyPath = Join-Path $vstoOutput $dependency
    if (Test-Path $dependencyPath) {
        [void][Reflection.Assembly]::LoadFrom($dependencyPath)
    }
}
$assemblyResolver = [ResolveEventHandler]{
    param($sender, $eventArgs)
    try {
        $simpleName = ([Reflection.AssemblyName]$eventArgs.Name).Name
        $candidate = Join-Path $vstoOutput ($simpleName + ".dll")
        if (Test-Path $candidate) {
            return [Reflection.Assembly]::LoadFrom($candidate)
        }
    }
    catch {}
    return $null
}
[AppDomain]::CurrentDomain.add_AssemblyResolve($assemblyResolver)

$contractsAssembly = [Reflection.Assembly]::LoadFrom($contractsPath)
$wordAssembly = [Reflection.Assembly]::LoadFrom($wordVstoPath)
$sessionType = $contractsAssembly.GetType(
    "VisualTeX.WindowsOffice.Contracts.OfficeSessionDocument",
    $true)
$lineType = $contractsAssembly.GetType(
    "VisualTeX.WindowsOffice.Contracts.FormulaLine",
    $true)
$exportType = $contractsAssembly.GetType(
    "VisualTeX.WindowsOffice.Contracts.OfficeExportDocument",
    $true)
$metadataCodecType = $contractsAssembly.GetType(
    "VisualTeX.WindowsOffice.Contracts.FormulaMetadataCodec",
    $true)
$wordServiceType = $wordAssembly.GetType(
    "VisualTeX.WordVsto.WordFormulaService",
    $true)
$previewType = $wordAssembly.GetType(
    "VisualTeX.WindowsOffice.VstoShared.OfficeOlePreview",
    $true)
$previewMethod = $previewType.GetMethod(
    "CreateInkSafePreviewFromSvg",
    [Reflection.BindingFlags]"Static,NonPublic")
if ($null -eq $previewMethod) {
    throw "The current VSTO output does not expose the ink-safe preview helper."
}
$decodeMethod = $metadataCodecType.GetMethod(
    "Decode",
    [Reflection.BindingFlags]"Static,Public")
$insertOleMethod = $wordServiceType.GetMethods() |
    Where-Object { $_.Name -eq "InsertOle" -and $_.GetParameters().Count -eq 6 } |
    Select-Object -First 1
if ($null -eq $insertOleMethod) {
    throw "The current WordFormulaService.InsertOle overload is unavailable."
}
$insertOmmlMethod = $wordServiceType.GetMethods() |
    Where-Object { $_.Name -eq "InsertOmml" -and $_.GetParameters().Count -eq 8 } |
    Select-Object -First 1
if ($null -eq $insertOmmlMethod) {
    throw "The current WordFormulaService.InsertOmml overload is unavailable."
}
$listType = [Collections.Generic.List``1].MakeGenericType($lineType)

$ownsWord = [string]::IsNullOrWhiteSpace($SourceDocumentName)
$wordCom = if ($ownsWord) {
    New-Object -ComObject Word.Application
} else {
    [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
}
$wordApplicationType = [Type]::GetType(
    "Microsoft.Office.Interop.Word.Application, Microsoft.Office.Interop.Word, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c",
    $true)
$wordUnknown = [Runtime.InteropServices.Marshal]::GetIUnknownForObject($wordCom)
try {
    $word = [Runtime.InteropServices.Marshal]::GetTypedObjectForIUnknown(
        $wordUnknown,
        $wordApplicationType)
}
finally {
    [void][Runtime.InteropServices.Marshal]::Release($wordUnknown)
}
$source = $null
$temp = $null
$service = $null
$screenUpdating = [bool]$word.ScreenUpdating
if ($ownsWord) {
    $word.Visible = $false
    $word.DisplayAlerts = 0
}
$generatedFiles = [Collections.Generic.List[string]]::new()
$probePreviewRoot = Join-Path (
    [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) (
        "VisualTeX\office\temp\inline-baseline-probe-" + [guid]::NewGuid().ToString("N"))
[void](New-Item -Path $probePreviewRoot -ItemType Directory -Force)
try {
    if (-not $ownsWord) {
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
        if ($null -eq $source) { throw "Requested Word source document is not open." }
    }

    $items = @($manifest.items)
    $requiredClonedPairs = @($items | Where-Object {
        [string]::IsNullOrWhiteSpace(
            [string](Get-OptionalProperty $_ "mathMl" ""))
    }).Count
    if ($requiredClonedPairs -gt 0 -and $null -eq $source) {
        throw "Generated-only probe mode requires mathMl for every formula item."
    }
    if (
        $requiredClonedPairs -gt 0 -and
        ($source.InlineShapes.Count -lt $requiredClonedPairs -or
         $source.OMaths.Count -lt $requiredClonedPairs)
    ) {
        throw "Source document does not contain enough OLE/OMML pairs for cloned-reference items."
    }
    $sourceState = if ($null -ne $source) {
        [pscustomobject]@{
            Saved = [bool]$source.Saved
            Characters = [int]$source.Content.End
            InlineShapes = [int]$source.InlineShapes.Count
            OMaths = [int]$source.OMaths.Count
        }
    } else { $null }

    $word.ScreenUpdating = $false
    $temp = $word.Documents.Add()
    # Keep each diagnostic OLE/OMML/anchor row on one physical line. Complex
    # structures can exceed a normal portrait page when both representations
    # are placed side by side, which makes Word return an undefined horizontal
    # position and invalidates pixel crops without saying anything about the
    # baseline algorithm itself.
    # Word can silently keep the printer's portrait text width unless the paper
    # kind is switched to custom before PageWidth/PageHeight are assigned.
    $temp.PageSetup.PaperSize = 41 # wdPaperCustom
    $temp.PageSetup.Orientation = 1 # wdOrientLandscape
    $temp.PageSetup.PageWidth = 1584 # Word's 22-inch custom-paper limit
    $temp.PageSetup.PageHeight = 792 # 11 inches
    $temp.PageSetup.LeftMargin = 18
    $temp.PageSetup.RightMargin = 18
    $temp.PageSetup.TopMargin = 18
    $temp.PageSetup.BottomMargin = 18
    $temp.Activate()
    $wordServiceConstructor = $wordServiceType.GetConstructor(@($wordApplicationType))
    if ($null -eq $wordServiceConstructor) {
        throw "WordFormulaService constructor is unavailable."
    }
    $service = $wordServiceConstructor.Invoke(@($word))
    $state = [pscustomobject]@{ Position = 0 }
    $records = [Collections.Generic.List[object]]::new()
    $cloneSourcePair = 0

    for ($index = 0; $index -lt $items.Count; $index++) {
        $item = $items[$index]
        $mathMl = [string](Get-OptionalProperty $item "mathMl" "")
        $usesGeneratedOmml = -not [string]::IsNullOrWhiteSpace($mathMl)
        if ($usesGeneratedOmml) {
            $sourcePair = 1
        }
        else {
            $cloneSourcePair++
            $sourcePair = $cloneSourcePair
        }
        $sourceMetadata = $null
        if ($null -ne $source) {
            $sourceShape = $source.InlineShapes.Item($sourcePair)
            try {
                $sourceMetadata = $decodeMethod.Invoke(
                    $null,
                    @([string]$sourceShape.AlternativeText))
            }
            finally { Release-Com $sourceShape }
            if ($null -eq $sourceMetadata) {
                throw "Unable to decode source metadata template."
            }
        }
        $defaultFontSizePt = if ($null -ne $sourceMetadata -and $sourceMetadata.FontSizePt) {
            [double]$sourceMetadata.FontSizePt
        } else { 10.5 }
        $fontSizePt = [double](Get-OptionalProperty `
            $item `
            "fontSizePt" `
            $defaultFontSizePt)
        $defaultFormulaLetterFont = if ($null -ne $sourceMetadata -and $sourceMetadata.FormulaLetterFont) {
            [string]$sourceMetadata.FormulaLetterFont
        } else { "katex" }
        $defaultFormulaChineseFont = if ($null -ne $sourceMetadata -and $sourceMetadata.FormulaChineseFont) {
            [string]$sourceMetadata.FormulaChineseFont
        } else { "system" }
        $formulaLetterFont = [string](Get-OptionalProperty $item "formulaLetterFont" $defaultFormulaLetterFont)
        $ommlFormulaLetterFont = [string](Get-OptionalProperty `
            $item `
            "ommlFormulaLetterFont" `
            $formulaLetterFont)
        $formulaChineseFont = [string](Get-OptionalProperty $item "formulaChineseFont" $defaultFormulaChineseFont)
        $anchorStartFraction = [double](Get-OptionalProperty $item "anchorStartFraction" 0.45)
        $compactLine = [bool](Get-OptionalProperty $item "compactLine" $false)
        if ($anchorStartFraction -lt 0 -or $anchorStartFraction -ge 1) {
            throw "anchorStartFraction must be in [0, 1)."
        }

        $localSvgPath = Join-Path $probePreviewRoot (
            ("{0:D2}-" -f ($index + 1)) + [IO.Path]::GetFileName([string]$item.svgPath))
        Copy-Item -LiteralPath ([string]$item.svgPath) -Destination $localSvgPath -Force
        $generatedFiles.Add($localSvgPath)
        $previewArguments = [object[]]::new(6)
        $previewArguments[0] = [string]$localSvgPath
        $previewArguments[1] = [single]$item.width
        $previewArguments[2] = [single]$item.height
        $previewArguments[3] = [Nullable[single]]([single]$item.baseline)
        $previewArguments[4] = $null
        $previewArguments[5] = [single](Get-OptionalProperty $item "safetyPaddingPixels" 1)
        $preview = $previewMethod.Invoke($null, $previewArguments)
        $emfPath = [string](Get-InternalProperty $preview "EmfPath")
        $pngPath = [string](Get-InternalProperty $preview "PngPath")
        $renderWidth = [single](Get-InternalProperty $preview "WidthPixels")
        $renderHeight = [single](Get-InternalProperty $preview "HeightPixels")
        $renderBaseline = [single](Get-InternalProperty $preview "BaselinePixels")
        $previewInk = Measure-PngInkBounds $pngPath
        $generatedFiles.Add($emfPath)
        $generatedFiles.Add($pngPath)

        $lineStart = [int]$state.Position
        $h1Text = if ($compactLine) { "H" } else { "H  " }
        $h1FontSizePt = if ($compactLine) { 12.0 } else { $fontSizePt }
        $h1 = Insert-Text $temp $state $h1Text
        Format-BodyText $temp $h1.Start $h1.End $h1FontSizePt
        $h1BookmarkName = "VTPH1_{0:D4}" -f ($index + 1)
        [void](Add-TrackingBookmark `
            $temp `
            $h1BookmarkName `
            ([int]$h1.Start) `
            ([int]$h1.Start + 1))

        $selection = $word.Selection
        try {
            $selection.SetRange([int]$state.Position, [int]$state.Position)
            $selection.Font.Name = "Times New Roman"
            $selection.Font.Size = $fontSizePt
            $selection.Font.Position = 0
        }
        finally { Release-Com $selection }

        $formulaId = [guid]::NewGuid().ToString("D")
        $oleBookmarkName = Formula-BookmarkName "VTO_" $formulaId
        $line = [Activator]::CreateInstance($lineType)
        $line.Id = [guid]::NewGuid().ToString("D")
        $line.Latex = [string]$item.latex
        $lines = [Activator]::CreateInstance($listType)
        $lines.Add($line)
        $export = [Activator]::CreateInstance($exportType)
        $export.Width = $renderWidth
        $export.Height = $renderHeight
        $exportType.GetProperty("Baseline").SetValue(
            $export,
            [single]$renderBaseline,
            $null)
        $export.FormulaLetterFont = $formulaLetterFont
        $export.FormulaChineseFont = $formulaChineseFont
        $session = [Activator]::CreateInstance($sessionType)
        $session.Id = [guid]::NewGuid().ToString("D")
        $session.Mode = "create"
        $session.Host = "word"
        $session.FormulaId = $formulaId
        $session.SourceDocumentId = $null
        $session.SourceObjectId = $null
        $session.Title = "VisualTeX inline baseline probe"
        $session.Lines = $lines
        $session.CodeFormat = "latex"
        $session.DisplayMode = "inline"
        $session.ObjectMode = "nativeOle"
        $session.Numbered = $false
        $session.FontSizePt = $fontSizePt
        $session.Status = "ready"
        $session.Dirty = $false
        $session.ExplicitCancel = $false
        $session.OriginalMetadata = $null
        $session.ExportResult = $export
        $temp.Activate()
        Write-Host ("[inline-baseline-probe] {0}/{1} {2}" -f ($index + 1), $items.Count, [string]$item.name)
        [void](Invoke-WithOfficeRetry {
            $insertOleMethod.Invoke(
                $service,
                @($session, $pngPath, $emfPath, $false, $false, $false))
        })

        $insertedShape = $temp.InlineShapes.Item($temp.InlineShapes.Count)
        $insertedRange = $insertedShape.Range
        try {
            $oleStart = [int]$insertedRange.Start
            $oleEnd = [int]$insertedRange.End
            $olePosition = Read-InlineObjectPosition $temp $oleStart $oleEnd

            # AddOLEObject can return while Word is still expanding the EMBED
            # field around its U+0001 result character. Refresh the live range
            # after materialization before inserting following prose; otherwise a
            # stale End can place that prose inside the hidden field instruction.
            Release-Com $insertedRange
            $insertedRange = $null
            [Windows.Forms.Application]::DoEvents()
            $insertedRange = $insertedShape.Range
            $oleStart = [int]$insertedRange.Start
            $oleEnd = [int]$insertedRange.End
            $oleWidthPoints = [double]$insertedShape.Width
            $oleHeightPoints = [double]$insertedShape.Height

            $positionAdjustment = [int](Get-OptionalProperty $item "positionAdjustment" 0)
            if ($positionAdjustment -ne 0) {
                $adjustedPosition = [int]$olePosition + $positionAdjustment
                Set-InlineObjectPosition $temp $oleStart $oleEnd $adjustedPosition
                $olePosition = Read-InlineObjectPosition $temp $oleStart $oleEnd
            }
            $state.Position = $oleEnd
        }
        finally {
            Release-Com $insertedRange
            Release-Com $insertedShape
        }

        $h2Text = if ($compactLine) { " " } else { "  H  " }
        $h2 = Insert-Text $temp $state $h2Text
        Format-BodyText `
            $temp `
            $h2.Start `
            $h2.End `
            $(if ($compactLine) { 1.0 } else { $fontSizePt })
        $h2BookmarkName = "VTPH2_{0:D4}" -f ($index + 1)
        $h2BookmarkStart = if ($compactLine) {
            [int]$h2.Start
        } else {
            [int]$h2.Start + 2
        }
        $h2BookmarkEnd = if ($compactLine) {
            [int]$h2.End
        } else {
            [int]$h2.Start + 3
        }
        [void](Add-TrackingBookmark `
            $temp `
            $h2BookmarkName `
            $h2BookmarkStart `
            $h2BookmarkEnd)

        if ($usesGeneratedOmml) {
            $selection = $word.Selection
            try {
                $selection.SetRange([int]$state.Position, [int]$state.Position)
                $selection.Font.Name = "Times New Roman"
                $selection.Font.Size = $fontSizePt
                $selection.Font.Position = 0
            }
            finally { Release-Com $selection }

            $ommlExport = [Activator]::CreateInstance($exportType)
            $ommlExport.FormulaLetterFont = $ommlFormulaLetterFont
            $ommlExport.FormulaChineseFont = $formulaChineseFont
            $ommlSession = [Activator]::CreateInstance($sessionType)
            $ommlSession.Id = [guid]::NewGuid().ToString("D")
            $ommlSession.Mode = "create"
            $ommlSession.Host = "word"
            $ommlFormulaId = [guid]::NewGuid().ToString("D")
            $ommlBookmarkName = Formula-BookmarkName "VTOMML_" $ommlFormulaId
            $ommlSession.FormulaId = $ommlFormulaId
            $ommlSession.SourceDocumentId = $null
            $ommlSession.SourceObjectId = $null
            $ommlSession.Title = "VisualTeX generated OMML baseline probe"
            $ommlSession.Lines = $lines
            $ommlSession.CodeFormat = "latex"
            $ommlSession.DisplayMode = "inline"
            $ommlSession.ObjectMode = "wordOmml"
            $ommlSession.Numbered = $false
            $ommlSession.FontSizePt = $fontSizePt
            $ommlSession.Status = "ready"
            $ommlSession.Dirty = $false
            $ommlSession.ExplicitCancel = $false
            $ommlSession.OriginalMetadata = $null
            $ommlSession.ExportResult = $ommlExport
            $ommlArguments = [object[]]::new(8)
            $ommlArguments[0] = $ommlSession
            $ommlArguments[1] = [string]$mathMl
            $ommlArguments[2] = $false
            $ommlArguments[3] = $false
            $ommlArguments[4] = $false
            $ommlArguments[5] = $null
            $ommlArguments[6] = $false
            $ommlArguments[7] = $false
            [void](Invoke-WithOfficeRetry {
                $insertOmmlMethod.Invoke($service, $ommlArguments)
            })

            if ($temp.OMaths.Count -lt 1) {
                throw "Generated MathML did not produce a Word OMath object."
            }
            $insertedMath = $temp.OMaths.Item($temp.OMaths.Count)
            $insertedMathRange = $insertedMath.Range
            try {
                $ommlStart = [int]$insertedMathRange.Start
                $ommlEnd = [int]$insertedMathRange.End
            }
            finally {
                Release-Com $insertedMathRange
                Release-Com $insertedMath
            }
            $state.Position = [int]$temp.Content.End - 1
        }
        else {
            $sourceMath = $source.OMaths.Item($sourcePair)
            $sourceMathRange = $sourceMath.Range
            try { [void](Clone-FormattedRange $temp $state $sourceMathRange) }
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
        }

        $ommlTrackingBookmarkName = "VTPOM_{0:D4}" -f ($index + 1)
        [void](Add-TrackingBookmark `
            $temp `
            $ommlTrackingBookmarkName `
            $ommlStart `
            $ommlEnd)

        $anchorMathMl = [string](Get-OptionalProperty $item "anchorMathMl" "")
        $anchorTrackingBookmarkName = ""
        $anchorOmmlStart = -1
        $anchorOmmlEnd = -1
        $h4Start = -1
        $h4End = -1
        if (-not [string]::IsNullOrWhiteSpace($anchorMathMl)) {
            $h3 = Insert-Text $temp $state "  H  "
            Format-BodyText $temp $h3.Start $h3.End $fontSizePt

            $selection = $word.Selection
            try {
                $selection.SetRange([int]$state.Position, [int]$state.Position)
                $selection.Font.Name = "Times New Roman"
                $selection.Font.Size = $fontSizePt
                $selection.Font.Position = 0
            }
            finally { Release-Com $selection }

            $anchorLine = [Activator]::CreateInstance($lineType)
            $anchorLine.Id = [guid]::NewGuid().ToString("D")
            $anchorLine.Latex = "x"
            $anchorLines = [Activator]::CreateInstance($listType)
            $anchorLines.Add($anchorLine)
            $anchorExport = [Activator]::CreateInstance($exportType)
            $anchorExport.FormulaLetterFont = $ommlFormulaLetterFont
            $anchorExport.FormulaChineseFont = $formulaChineseFont
            $anchorSession = [Activator]::CreateInstance($sessionType)
            $anchorSession.Id = [guid]::NewGuid().ToString("D")
            $anchorSession.Mode = "create"
            $anchorSession.Host = "word"
            $anchorSession.FormulaId = [guid]::NewGuid().ToString("D")
            $anchorSession.SourceDocumentId = $null
            $anchorSession.SourceObjectId = $null
            $anchorSession.Title = "VisualTeX OMML baseline anchor"
            $anchorSession.Lines = $anchorLines
            $anchorSession.CodeFormat = "latex"
            $anchorSession.DisplayMode = "inline"
            $anchorSession.ObjectMode = "wordOmml"
            $anchorSession.Numbered = $false
            $anchorSession.FontSizePt = $fontSizePt
            $anchorSession.Status = "ready"
            $anchorSession.Dirty = $false
            $anchorSession.ExplicitCancel = $false
            $anchorSession.OriginalMetadata = $null
            $anchorSession.ExportResult = $anchorExport
            $anchorArguments = [object[]]::new(8)
            $anchorArguments[0] = $anchorSession
            $anchorArguments[1] = [string]$anchorMathMl
            $anchorArguments[2] = $false
            $anchorArguments[3] = $false
            $anchorArguments[4] = $false
            $anchorArguments[5] = $null
            $anchorArguments[6] = $false
            $anchorArguments[7] = $false
            [void](Invoke-WithOfficeRetry {
                $insertOmmlMethod.Invoke($service, $anchorArguments)
            })

            $anchorMath = $temp.OMaths.Item($temp.OMaths.Count)
            $anchorMathRange = $anchorMath.Range
            try {
                $anchorOmmlStart = [int]$anchorMathRange.Start
                $anchorOmmlEnd = [int]$anchorMathRange.End
            }
            finally {
                Release-Com $anchorMathRange
                Release-Com $anchorMath
            }
            $anchorTrackingBookmarkName = "VTPAN_{0:D4}" -f ($index + 1)
            [void](Add-TrackingBookmark `
                $temp `
                $anchorTrackingBookmarkName `
                $anchorOmmlStart `
                $anchorOmmlEnd)
            $state.Position = [int]$temp.Content.End - 1

            $h4 = Insert-Text $temp $state "  H"
            Format-BodyText $temp $h4.Start $h4.End $fontSizePt
            $h4Start = [int]($h4.Start + 2)
            $h4End = [int]($h4.Start + 3)
        }
        else {
            $h3 = Insert-Text $temp $state "  H"
            Format-BodyText $temp $h3.Start $h3.End $fontSizePt
        }
        $lineEnd = [int]$state.Position
        [void](Insert-Text $temp $state "`r")
        $paragraphRange = $temp.Range($lineStart, [int]$state.Position)
        try {
            $paragraphRange.ParagraphFormat.SpaceBefore = 0
            $paragraphRange.ParagraphFormat.SpaceAfter = 0
            $paragraphRange.ParagraphFormat.LineSpacingRule = 0
        }
        finally { Release-Com $paragraphRange }

        $records.Add([pscustomobject]@{
            Name = [string]$item.name
            Latex = [string]$item.latex
            FontSizePt = $fontSizePt
            FormulaLetterFont = $formulaLetterFont
            OmmlFormulaLetterFont = $ommlFormulaLetterFont
            UsesGeneratedOmml = $usesGeneratedOmml
            ComparisonMode = [string](Get-OptionalProperty $item "comparisonMode" "full-formula")
            CompactLine = $compactLine
            AnchorStartFraction = $anchorStartFraction
            Position = $olePosition
            RenderWidthPx = $renderWidth
            RenderHeightPx = $renderHeight
            RenderBaselinePx = $renderBaseline
            PreviewPixels = "$($previewInk.Width)x$($previewInk.Height)"
            PreviewInkPixels = "$($previewInk.InkWidth)x$($previewInk.InkHeight)"
            PreviewInkLeftMarginPx = $previewInk.LeftMargin
            PreviewInkTopMarginPx = $previewInk.TopMargin
            PreviewInkRightMarginPx = $previewInk.RightMargin
            PreviewInkBottomMarginPx = $previewInk.BottomMargin
            H1BookmarkName = $h1BookmarkName
            H2BookmarkName = $h2BookmarkName
            OleBookmarkName = $oleBookmarkName
            OmmlBookmarkName = $ommlTrackingBookmarkName
            AnchorBookmarkName = $anchorTrackingBookmarkName
            LineStart = $lineStart
            LineEnd = $lineEnd
            H1Start = [int]$h1.Start
            H1End = [int]($h1.Start + 1)
            H2Start = [int]($h2.Start + 2)
            H2End = [int]($h2.Start + 3)
            H3Start = [int]($h3.Start + 2)
            H3End = [int]($h3.Start + 3)
            H4Start = $h4Start
            H4End = $h4End
            OleStart = $oleStart
            OleEnd = $oleEnd
            OleWidthPoints = $oleWidthPoints
            OleHeightPoints = $oleHeightPoints
            OmmlStart = $ommlStart
            OmmlEnd = $ommlEnd
            AnchorOmmlStart = $anchorOmmlStart
            AnchorOmmlEnd = $anchorOmmlEnd
        })
    }

    $temp.Repaginate()
    Start-Sleep -Milliseconds 180
    $metrics = [Collections.Generic.List[object]]::new()
    foreach ($record in $records) {
        $h1Bounds = Get-TrackingBookmarkBounds $temp ([string]$record.H1BookmarkName)
        $h2Bounds = Get-TrackingBookmarkBounds $temp ([string]$record.H2BookmarkName)
        $oleBounds = Get-TrackingBookmarkBounds $temp ([string]$record.OleBookmarkName)
        $ommlBounds = Get-TrackingBookmarkBounds $temp ([string]$record.OmmlBookmarkName)
        $hasIndependentAnchor = -not [string]::IsNullOrWhiteSpace(
            [string]$record.AnchorBookmarkName)
        $anchorBounds = if ($hasIndependentAnchor) {
            Get-TrackingBookmarkBounds $temp ([string]$record.AnchorBookmarkName)
        } else { $null }
        $lineStart = [int]$h1Bounds.Start
        $lineEnd = if ($hasIndependentAnchor) {
            [int]$anchorBounds.End
        } else {
            [int]$ommlBounds.End
        }
        $lineRange = $temp.Range($lineStart, $lineEnd)
        try { [byte[]]$bytes = $lineRange.EnhMetaFileBits }
        finally { Release-Com $lineRange }
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
        finally { $graphics.Dispose() }

        try {
            $dpiX = [double]$metafile.HorizontalResolution
            $dpiY = [double]$metafile.VerticalResolution
            $originX = Read-X $temp $lineStart
            $oleX = Read-X $temp ([int]$oleBounds.Start)
            $ommlX1 = Read-X $temp ([int]$ommlBounds.Start)
            $ommlX2 = Read-X $temp ([int]$ommlBounds.End)
            $h1X1 = Read-X $temp ([int]$h1Bounds.Start)
            $h1X2 = Read-X $temp ([int]$h1Bounds.End)
            $h2X1 = Read-X $temp ([int]$h2Bounds.Start)
            $h2X2 = Read-X $temp ([int]$h2Bounds.End)
            if ($hasIndependentAnchor) {
                $anchorOmmlX1 = Read-X $temp ([int]$anchorBounds.Start)
                $anchorOmmlX2 = Read-X $temp ([int]$anchorBounds.End)
            }
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
            $ole = Measure-Crop $bitmap `
                ($oleLeftPx - $marginPx) `
                ($oleRightPx + $marginPx)
            $omml = Measure-Crop $bitmap `
                ($ommlLeftPx - $marginPx) `
                ($ommlRightPx + $marginPx)
            # Diagnostic manifests append a separated baseline-bearing anchor
            # glyph. Locate the final ink cluster by its preceding blank gap;
            # this remains stable when OLE and OMML have different total widths.
            $minimumAnchorGapPx = [Math]::Max(
                4,
                [int][Math]::Ceiling([double]$record.FontSizePt * $dpiX / 72.0 * 0.15))
            $oleBody = Measure-RightmostInkCluster `
                $bitmap `
                $oleLeftPx `
                $oleRightPx `
                $minimumAnchorGapPx
            if ($hasIndependentAnchor) {
                $ommlBody = Measure-Crop `
                    $bitmap `
                    ((& $toPixel $anchorOmmlX1) - $marginPx) `
                    ((& $toPixel $anchorOmmlX2) + $marginPx)
            }
            else {
                $ommlBody = Measure-RightmostInkCluster `
                    $bitmap `
                    $ommlLeftPx `
                    $ommlRightPx `
                    $minimumAnchorGapPx
            }
            if ([double]$record.FontSizePt -ge 42) {
                $clusterSummary = @(Get-HorizontalInkClusters $bitmap 6 | ForEach-Object {
                    "{0}-{1}" -f $_.Left, $_.Right
                }) -join ","
                $coordinateFormat =
                    "[inline-baseline-coordinates] {0} bitmap={1}x{2} origin={3:0.###} " +
                    "ole={4:0.###}-{5:0.###} omml={6:0.###}-{7:0.###} " +
                    "h1={8:0.###}-{9:0.###} h2={10:0.###}-{11:0.###} clusters={12}"
                Write-Host ($coordinateFormat -f
                    $record.Name,
                    $bitmap.Width,
                    $bitmap.Height,
                    $originX,
                    $oleLeftPx,
                    $oleRightPx,
                    $ommlLeftPx,
                    $ommlRightPx,
                    (& $toPixel $h1X1),
                    (& $toPixel $h1X2),
                    (& $toPixel $h2X1),
                    (& $toPixel $h2X2),
                    $clusterSummary)
            }
            $h1Ink = Measure-Crop $bitmap `
                ((& $toPixel $h1X1) - $marginPx) `
                ((& $toPixel $h1X2) + $marginPx)
            if ([bool]$record.CompactLine) {
                # The compact 72 pt diagnostic uses a single 12 pt baseline
                # marker and an ink-free 1 pt separator to stay within Word's
                # fixed 432 pt range-to-EMF width.
                $hTopValues = @($h1Ink.Top)
                $hBottomValues = @($h1Ink.Bottom)
            }
            else {
                $h2Ink = Measure-Crop $bitmap `
                    ((& $toPixel $h2X1) - $marginPx) `
                    ((& $toPixel $h2X2) + $marginPx)
                # At large font sizes Word may place the trailing post-OMML
                # probe on the next visual line. The two prose anchors before
                # OMML define the current line baseline without mixing boxes.
                $hTopValues = @($h1Ink.Top, $h2Ink.Top)
                $hBottomValues = @($h1Ink.Bottom, $h2Ink.Bottom)
            }

            $pxToPtY = 72.0 / $dpiY
            $geometricDescentPt = [double]$record.OleHeightPoints * (
                [double]$record.RenderHeightPx - [double]$record.RenderBaselinePx
            ) / [double]$record.RenderHeightPx
            $hBottom = [double](
                $hBottomValues | Measure-Object -Average
            ).Average
            $metrics.Add([pscustomobject]@{
                Name = $record.Name
                Latex = $record.Latex
                FontSizePt = $record.FontSizePt
                FormulaLetterFont = $record.FormulaLetterFont
                OmmlFormulaLetterFont = $record.OmmlFormulaLetterFont
                UsesGeneratedOmml = $record.UsesGeneratedOmml
                ComparisonMode = $record.ComparisonMode
                CompactLine = $record.CompactLine
                AnchorStartFraction = $record.AnchorStartFraction
                Position = $record.Position
                GeometricDescentPt = [Math]::Round($geometricDescentPt, 3)
                FractionalDescentPt = [Math]::Round(
                    $geometricDescentPt - [Math]::Floor($geometricDescentPt),
                    3)
                ObjectWidthPt = [Math]::Round([double]$record.OleWidthPoints, 3)
                ObjectHeightPt = [Math]::Round([double]$record.OleHeightPoints, 3)
                RenderHeightPx = [Math]::Round([double]$record.RenderHeightPx, 4)
                RenderBaselinePx = [Math]::Round([double]$record.RenderBaselinePx, 4)
                PreviewPixels = $record.PreviewPixels
                PreviewInkPixels = $record.PreviewInkPixels
                PreviewInkLeftMarginPx = $record.PreviewInkLeftMarginPx
                PreviewInkTopMarginPx = $record.PreviewInkTopMarginPx
                PreviewInkRightMarginPx = $record.PreviewInkRightMarginPx
                PreviewInkBottomMarginPx = $record.PreviewInkBottomMarginPx
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
                OleBodyBottomVsHPt = [Math]::Round(($oleBody.Bottom - $hBottom) * $pxToPtY, 3)
                OmmlBodyBottomVsHPt = [Math]::Round(($ommlBody.Bottom - $hBottom) * $pxToPtY, 3)
                OmmlBottomVsHPt = [Math]::Round(($omml.Bottom - $hBottom) * $pxToPtY, 3)
                HTopSpreadPt = [Math]::Round(
                    (Get-Spread $hTopValues) * $pxToPtY,
                    3)
                HBottomSpreadPt = [Math]::Round(
                    (Get-Spread $hBottomValues) * $pxToPtY,
                    3)
            })
        }
        finally {
            $bitmap.Dispose()
            $metafile.Dispose()
            $stream.Dispose()
        }
    }

    $sourceUnchanged = $true
    $sourceName = "independent-word-instance"
    if ($null -ne $source) {
        $currentSourceState = [pscustomobject]@{
            Saved = [bool]$source.Saved
            Characters = [int]$source.Content.End
            InlineShapes = [int]$source.InlineShapes.Count
            OMaths = [int]$source.OMaths.Count
        }
        $sourceUnchanged = (
            $currentSourceState.Saved -eq $sourceState.Saved -and
            $currentSourceState.Characters -eq $sourceState.Characters -and
            $currentSourceState.InlineShapes -eq $sourceState.InlineShapes -and
            $currentSourceState.OMaths -eq $sourceState.OMaths
        )
        if (-not $sourceUnchanged) {
            throw "Source document state changed during the read-only probe."
        }
        $sourceName = [string]$source.Name
    }

    $outputMetrics = if ($CompactOutput) {
        @($metrics | Select-Object `
            Name,
            FontSizePt,
            FormulaLetterFont,
            OmmlFormulaLetterFont,
            ComparisonMode,
            CompactLine,
            Position,
            GeometricDescentPt,
            FractionalDescentPt,
            ObjectHeightPt,
            PreviewInkLeftMarginPx,
            PreviewInkTopMarginPx,
            PreviewInkRightMarginPx,
            PreviewInkBottomMarginPx,
            DeltaTopPt,
            DeltaBottomPt,
            DeltaCentroidPt,
            BodyDeltaTopPt,
            BodyDeltaBottomPt,
            BodyDeltaCentroidPt,
            HTopSpreadPt,
            HBottomSpreadPt)
    } else {
        $metrics
    }
    $result = [pscustomobject]@{
        Source = $sourceName
        SourceUnchanged = $sourceUnchanged
        TemporaryDocumentSaved = [bool]$temp.Saved
        Metrics = $outputMetrics
    }
    $json = $result | ConvertTo-Json -Depth 7
    if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
        $resolvedOutputPath = [IO.Path]::GetFullPath($OutputPath)
        $outputDirectory = [IO.Path]::GetDirectoryName($resolvedOutputPath)
        if (-not [string]::IsNullOrWhiteSpace($outputDirectory)) {
            [void][IO.Directory]::CreateDirectory($outputDirectory)
        }
        [IO.File]::WriteAllText(
            $resolvedOutputPath,
            $json + [Environment]::NewLine,
            [Text.UTF8Encoding]::new($false))
    }
    $json
}
finally {
    if ($null -ne $temp) {
        try { $temp.Close(0) } catch {}
        Release-Com $temp
    }
    if ($null -ne $source) {
        try { $source.Activate() } catch {}
        Release-Com $source
    }
    try { $word.ScreenUpdating = $screenUpdating } catch {}
    if ($ownsWord -and $null -ne $word) {
        try { $word.Quit(0) } catch {}
    }
    Release-Com $word
    # The typed Word RCW and the raw ProgID RCW can wrap the same COM identity.
    # Final-releasing both can over-release Word during PowerShell teardown.
    $wordCom = $null
    foreach ($path in $generatedFiles) {
        try { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue } catch {}
    }
    try { Remove-Item -LiteralPath $probePreviewRoot -Recurse -Force -ErrorAction SilentlyContinue } catch {}
    if ($null -ne $assemblyResolver) {
        try { [AppDomain]::CurrentDomain.remove_AssemblyResolve($assemblyResolver) } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
