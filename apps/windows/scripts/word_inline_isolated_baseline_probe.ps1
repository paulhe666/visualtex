[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestPath,
    [string]$OutputPath = "",
    [string]$SaveFirstDocumentPath = "",
    [int]$PositionAdjustmentFilter = 9999,
    [int]$StartIndex = 0,
    [int]$MaximumItems = 0
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

function Release-Com([object]$Value) {
    if ($null -ne $Value -and [Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value) } catch {}
    }
}

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
    if ($null -eq $property) {
        throw "Internal preview property '$Name' is unavailable."
    }
    return $property.GetValue($Value)
}

function Insert-Text(
    [object]$Document,
    [object]$State,
    [string]$Text,
    [double]$FontSizePt
) {
    $start = [int]$State.Position
    $range = $Document.Range($start, $start)
    try {
        $range.Font.Name = "Times New Roman"
        $range.Font.Size = $FontSizePt
        $range.Font.Position = 0
        $range.Text = $Text
        $State.Position = [int]$range.End
        return [pscustomobject]@{ Start = $start; End = [int]$range.End }
    }
    finally { Release-Com $range }
}

function Prepare-InsertionCaret(
    [object]$Word,
    [object]$Document,
    [int]$Position,
    [double]$FontSizePt
) {
    [void](Invoke-WithOfficeRetry { $Document.Activate() })
    $selection = Invoke-WithOfficeRetry { $Word.Selection }
    try {
        [void](Invoke-WithOfficeRetry { $selection.SetRange($Position, $Position) })
        $selection.Font.Name = "Times New Roman"
        $selection.Font.Size = $FontSizePt
        $selection.Font.Position = 0
    }
    finally { Release-Com $selection }
}

function Read-X([object]$Document, [int]$Position) {
    $probe = $Document.Range($Position, $Position)
    try {
        return [double]$probe.Information(5) # wdHorizontalPositionRelativeToPage
    }
    finally { Release-Com $probe }
}

function Read-InlineObjectPosition([object]$Document, [int]$Start, [int]$End) {
    for ($cursor = $Start; $cursor -lt $End; $cursor++) {
        $character = $Document.Range($cursor, $cursor + 1)
        try {
            if ($character.Text -eq [char]1) {
                return [int]$character.Font.Position
            }
        }
        finally { Release-Com $character }
    }
    throw "Inserted OLE range has no U+0001 object character."
}

function Set-InlineObjectPosition(
    [object]$Document,
    [int]$Start,
    [int]$End,
    [int]$Position
) {
    $whole = $Document.Range($Start, $End)
    try { $whole.Font.Position = 0 }
    finally { Release-Com $whole }

    $found = $false
    for ($cursor = $Start; $cursor -lt $End; $cursor++) {
        $character = $Document.Range($cursor, $cursor + 1)
        try {
            if ($character.Text -eq [char]1) {
                $character.Font.Position = [Math]::Max(-256, [Math]::Min(256, $Position))
                $found = $true
                break
            }
        }
        finally { Release-Com $character }
    }
    if (-not $found) {
        throw "Inserted OLE range has no U+0001 object character."
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

function Measure-RightmostInkCluster(
    [Drawing.Bitmap]$Bitmap,
    [double]$LeftPixels,
    [double]$RightPixels,
    [int]$MinimumBlankGapPixels
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
            LeftMargin = $minX
            TopMargin = $minY
            RightMargin = $bitmap.Width - 1 - $maxX
            BottomMargin = $bitmap.Height - 1 - $maxY
            InkWidth = $maxX - $minX + 1
            InkHeight = $maxY - $minY + 1
        }
    }
    finally { $bitmap.Dispose() }
}

function New-LineList(
    [Type]$LineType,
    [Type]$ListType,
    [string]$Latex
) {
    $line = [Activator]::CreateInstance($LineType)
    $line.Id = [guid]::NewGuid().ToString("D")
    $line.Latex = $Latex
    $lines = [Activator]::CreateInstance($ListType)
    $lines.Add($line)
    return $lines
}

function New-Session(
    [Type]$SessionType,
    [Type]$ExportType,
    [object]$Lines,
    [string]$ObjectMode,
    [string]$DocumentId,
    [double]$FontSizePt,
    [string]$FormulaLetterFont,
    [string]$FormulaChineseFont,
    [single]$Width,
    [single]$Height,
    [object]$Baseline
) {
    $export = [Activator]::CreateInstance($ExportType)
    if ($Width -gt 0) { $export.Width = $Width }
    if ($Height -gt 0) { $export.Height = $Height }
    if ($null -ne $Baseline) {
        $ExportType.GetProperty("Baseline").SetValue(
            $export,
            [single]$Baseline,
            $null)
    }
    $export.FormulaLetterFont = $FormulaLetterFont
    $export.FormulaChineseFont = $FormulaChineseFont

    $session = [Activator]::CreateInstance($SessionType)
    $session.Id = [guid]::NewGuid().ToString("D")
    $session.Mode = "create"
    $session.Host = "word"
    $session.FormulaId = [guid]::NewGuid().ToString("D")
    $session.SourceDocumentId = $DocumentId
    $session.SourceObjectId = $null
    $session.Title = "VisualTeX isolated inline baseline probe"
    $session.Lines = $Lines
    $session.CodeFormat = "latex"
    $session.DisplayMode = "inline"
    $session.ObjectMode = $ObjectMode
    $session.Numbered = $false
    $session.FontSizePt = $FontSizePt
    $session.Status = "ready"
    $session.Dirty = $false
    $session.ExplicitCancel = $false
    $session.OriginalMetadata = $null
    $session.ExportResult = $export
    return $session
}

function Capture-DocumentSnapshot([object]$Word) {
    $snapshot = @()
    foreach ($candidate in @($Word.Documents)) {
        $content = $null
        try {
            $content = $candidate.Content
            $snapshot += [pscustomobject]@{
                Name = [string]$candidate.Name
                End = [int]$content.End
                InlineShapes = [int]$candidate.InlineShapes.Count
                OMaths = [int]$candidate.OMaths.Count
                Saved = [bool]$candidate.Saved
            }
        }
        finally {
            Release-Com $content
            Release-Com $candidate
        }
    }
    return @($snapshot)
}

$manifest = Get-Content (Resolve-Path $ManifestPath) -Raw | ConvertFrom-Json
$items = @($manifest.items)
if ($PositionAdjustmentFilter -ne 9999) {
    $items = @($items | Where-Object {
        [int](Get-OptionalProperty $_ "positionAdjustment" 0) -eq
            $PositionAdjustmentFilter
    })
}
if ($StartIndex -lt 0) {
    throw "StartIndex cannot be negative."
}
if ($StartIndex -gt 0) {
    $items = @($items | Select-Object -Skip $StartIndex)
}
if ($MaximumItems -gt 0) {
    $items = @($items | Select-Object -First $MaximumItems)
}
if ($items.Count -lt 1) {
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
$wordServiceType = $wordAssembly.GetType(
    "VisualTeX.WordVsto.WordFormulaService",
    $true)
$previewType = $wordAssembly.GetType(
    "VisualTeX.WindowsOffice.VstoShared.OfficeOlePreview",
    $true)
$previewMethod = $previewType.GetMethod(
    "CreateInkSafePreviewFromSvg",
    [Reflection.BindingFlags]"Static,NonPublic")
$readDocumentIdMethod = $wordServiceType.GetMethod(
    "ReadActiveDocumentId",
    [Reflection.BindingFlags]"Instance,Public")
$insertOleMethod = $wordServiceType.GetMethods() |
    Where-Object { $_.Name -eq "InsertOle" -and $_.GetParameters().Count -eq 6 } |
    Select-Object -First 1
$insertOmmlMethod = $wordServiceType.GetMethods() |
    Where-Object { $_.Name -eq "InsertOmml" -and $_.GetParameters().Count -eq 8 } |
    Select-Object -First 1
if (
    $null -eq $previewMethod -or
    $null -eq $readDocumentIdMethod -or
    $null -eq $insertOleMethod -or
    $null -eq $insertOmmlMethod
) {
    throw "The current Word integration does not expose the required probe methods."
}
$listType = [Collections.Generic.List``1].MakeGenericType($lineType)

$word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
$screenUpdating = [bool]$word.ScreenUpdating
$displayAlerts = $word.DisplayAlerts
$initialDocumentSnapshot = Capture-DocumentSnapshot $word
$initialDocumentSnapshotJson = $initialDocumentSnapshot | ConvertTo-Json -Compress
$generatedFiles = [Collections.Generic.List[string]]::new()
$probePreviewRoot = Join-Path (
    [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)) (
        "VisualTeX\office\temp\inline-baseline-probe-" + [guid]::NewGuid().ToString("N"))
[void](New-Item -Path $probePreviewRoot -ItemType Directory -Force)
$results = [Collections.Generic.List[object]]::new()
$savedFirstDocument = $false
try {
    $word.ScreenUpdating = $false
    $word.DisplayAlerts = 0

    foreach ($item in $items) {
        $document = $null
        $lineRange = $null
        $metafile = $null
        $stream = $null
        $bitmap = $null
        try {
            $stage = "read-fixture"
            $name = [string]$item.name
            $baseName = $name -replace "-pos[+-]\d+$", ""
            $latex = [string]$item.latex
            $mathMl = [string]$item.mathMl
            $fontSizePt = [double](Get-OptionalProperty $item "fontSizePt" 10.5)
            $formulaLetterFont = [string](Get-OptionalProperty $item "formulaLetterFont" "katex")
            $formulaChineseFont = [string](Get-OptionalProperty $item "formulaChineseFont" "system")
            $positionAdjustment = [int](Get-OptionalProperty $item "positionAdjustment" 0)
            if ([string]::IsNullOrWhiteSpace($latex) -or [string]::IsNullOrWhiteSpace($mathMl)) {
                throw "$name has no complete SVG/MathML comparison fixture."
            }

            $sourceSvgPath = [string]$item.svgPath
            $localSvgPath = Join-Path $probePreviewRoot (
                [guid]::NewGuid().ToString("N") + ".svg")
            Copy-Item -LiteralPath $sourceSvgPath -Destination $localSvgPath -Force
            $generatedFiles.Add($localSvgPath)

            $previewArguments = [object[]]::new(6)
            $previewArguments[0] = [string]$localSvgPath
            $previewArguments[1] = [single]$item.width
            $previewArguments[2] = [single]$item.height
            $previewArguments[3] = [Nullable[single]]([single]$item.baseline)
            $previewArguments[4] = $null
            $previewArguments[5] = [single](Get-OptionalProperty $item "safetyPaddingPixels" 1)
            $stage = "create-ink-safe-preview"
            $preview = $previewMethod.Invoke($null, $previewArguments)
            $emfPath = [string](Get-InternalProperty $preview "EmfPath")
            $pngPath = [string](Get-InternalProperty $preview "PngPath")
            $renderWidth = [single](Get-InternalProperty $preview "WidthPixels")
            $renderHeight = [single](Get-InternalProperty $preview "HeightPixels")
            $renderBaseline = [single](Get-InternalProperty $preview "BaselinePixels")
            $generatedFiles.Add($emfPath)
            $generatedFiles.Add($pngPath)
            $previewInk = Measure-PngInkBounds $pngPath

            $document = Invoke-WithOfficeRetry { $word.Documents.Add() }
            [void](Invoke-WithOfficeRetry { $document.Activate() })
            $document.PageSetup.Orientation = 1 # wdOrientLandscape
            $document.PageSetup.LeftMargin = 36
            $document.PageSetup.RightMargin = 36
            $document.PageSetup.TopMargin = 36
            $document.PageSetup.BottomMargin = 36
            $service = [Activator]::CreateInstance($wordServiceType, @($word))
            $serviceTarget = $service.PSObject.BaseObject
            $stage = "read-active-document-id"
            $documentId = [string](Invoke-WithOfficeRetry {
                $readDocumentIdMethod.Invoke($serviceTarget, [object[]]@())
            })
            if ([string]::IsNullOrWhiteSpace($documentId)) {
                throw "$name test document has no stable VisualTeX identity."
            }

            $state = [pscustomobject]@{ Position = 0 }
            $lineStart = 0
            $h1 = Insert-Text $document $state "H  " $fontSizePt

            Prepare-InsertionCaret $word $document $state.Position $fontSizePt
            $oleLines = New-LineList $lineType $listType $latex
            $oleSession = New-Session `
                $sessionType `
                $exportType `
                $oleLines `
                "nativeOle" `
                $documentId `
                $fontSizePt `
                $formulaLetterFont `
                $formulaChineseFont `
                $renderWidth `
                $renderHeight `
                ([Nullable[single]]$renderBaseline)
            $oleArguments = [object[]]::new(6)
            $oleArguments[0] = $oleSession.PSObject.BaseObject
            $oleArguments[1] = [string]$pngPath
            $oleArguments[2] = [string]$emfPath
            $oleArguments[3] = [bool]$false
            $oleArguments[4] = [bool]$false
            $oleArguments[5] = [bool]$false
            $stage = "insert-native-ole"
            [void](Invoke-WithOfficeRetry {
                $insertOleMethod.Invoke($serviceTarget, $oleArguments)
            })
            if ([int]$document.InlineShapes.Count -ne 1) {
                throw "$name did not create exactly one isolated OLE object."
            }
            $oleShape = $document.InlineShapes.Item(1)
            $oleRange = $oleShape.Range
            try {
                $oleStart = [int]$oleRange.Start
                $oleEnd = [int]$oleRange.End
                $oleWidthPt = [double]$oleShape.Width
                $oleHeightPt = [double]$oleShape.Height
                $calculatedPosition = Read-InlineObjectPosition $document $oleStart $oleEnd
                $appliedPosition = $calculatedPosition + $positionAdjustment
                if ($positionAdjustment -ne 0) {
                    Set-InlineObjectPosition `
                        $document `
                        $oleStart `
                        $oleEnd `
                        $appliedPosition
                }
                $appliedPosition = Read-InlineObjectPosition $document $oleStart $oleEnd
                $state.Position = $oleEnd
            }
            finally {
                Release-Com $oleRange
                Release-Com $oleShape
            }

            $h2 = Insert-Text $document $state "  H  " $fontSizePt
            Prepare-InsertionCaret $word $document $state.Position $fontSizePt
            $ommlLines = New-LineList $lineType $listType $latex
            $ommlSession = New-Session `
                $sessionType `
                $exportType `
                $ommlLines `
                "wordOmml" `
                $documentId `
                $fontSizePt `
                $formulaLetterFont `
                $formulaChineseFont `
                0 `
                0 `
                ([Nullable[single]]$null)
            $ommlArguments = [object[]]::new(8)
            $ommlArguments[0] = $ommlSession.PSObject.BaseObject
            $ommlArguments[1] = [string]$mathMl
            $ommlArguments[2] = $false
            $ommlArguments[3] = $false
            $ommlArguments[4] = $false
            $ommlArguments[5] = $null
            $ommlArguments[6] = $false
            $ommlArguments[7] = $false
            $stage = "insert-omml"
            [void](Invoke-WithOfficeRetry {
                $insertOmmlMethod.Invoke($serviceTarget, $ommlArguments)
            })
            if ([int]$document.OMaths.Count -ne 1) {
                throw "$name MathML did not create exactly one Word OMath."
            }
            $omml = $document.OMaths.Item(1)
            $ommlRange = $omml.Range
            try {
                $ommlStart = [int]$ommlRange.Start
                $ommlEnd = [int]$ommlRange.End
                $state.Position = $ommlEnd
            }
            finally {
                Release-Com $ommlRange
                Release-Com $omml
            }

            $h3 = Insert-Text $document $state "  H" $fontSizePt
            $lineEnd = [int]$state.Position
            [void](Insert-Text $document $state "`r" $fontSizePt)
            $paragraph = $document.Range($lineStart, [int]$state.Position)
            try {
                $paragraph.ParagraphFormat.SpaceBefore = 0
                $paragraph.ParagraphFormat.SpaceAfter = 0
                $paragraph.ParagraphFormat.LineSpacingRule = 0
            }
            finally { Release-Com $paragraph }

            [void](Invoke-WithOfficeRetry { $document.Repaginate() })
            Start-Sleep -Milliseconds 120
            $lineRange = $document.Range($lineStart, $lineEnd)
            $stage = "capture-word-line"
            [byte[]]$bytes = Invoke-WithOfficeRetry { $lineRange.EnhMetaFileBits }
            if ($null -eq $bytes -or $bytes.Length -le 0) {
                throw "$name line produced no Word EMF snapshot."
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
            $originX = Read-X $document $lineStart
            $oleX = Read-X $document $oleStart
            $ommlX1 = Read-X $document $ommlStart
            $ommlX2 = Read-X $document $ommlEnd
            $h1X1 = Read-X $document $h1.Start
            $h1X2 = Read-X $document ($h1.Start + 1)
            $h2X1 = Read-X $document ($h2.Start + 2)
            $h2X2 = Read-X $document ($h2.Start + 3)
            $h3X1 = Read-X $document ($h3.Start + 2)
            $h3X2 = Read-X $document ($h3.Start + 3)
            $ptToPxX = $dpiX / 72.0
            $pxToPtY = 72.0 / $dpiY
            $marginPx = 1.0 * $ptToPxX
            $toPixel = {
                param([double]$PageX)
                return ($PageX - $originX) * $ptToPxX
            }

            $oleLeftPx = & $toPixel $oleX
            $oleRightPx = & $toPixel ($oleX + $oleWidthPt)
            $ommlLeftPx = & $toPixel $ommlX1
            $ommlRightPx = & $toPixel $ommlX2
            $oleInk = Measure-Crop $bitmap ($oleLeftPx - $marginPx) ($oleRightPx + $marginPx)
            $ommlInk = Measure-Crop $bitmap ($ommlLeftPx - $marginPx) ($ommlRightPx + $marginPx)
            $minimumAnchorGapPx = [Math]::Max(
                4,
                [int][Math]::Ceiling($fontSizePt * $dpiX / 72.0 * 0.15))
            $oleAnchorInk = Measure-RightmostInkCluster `
                $bitmap `
                $oleLeftPx `
                $oleRightPx `
                $minimumAnchorGapPx
            $ommlAnchorInk = Measure-RightmostInkCluster `
                $bitmap `
                $ommlLeftPx `
                $ommlRightPx `
                $minimumAnchorGapPx
            $h1Ink = Measure-Crop $bitmap ((& $toPixel $h1X1) - $marginPx) ((& $toPixel $h1X2) + $marginPx)
            $h2Ink = Measure-Crop $bitmap ((& $toPixel $h2X1) - $marginPx) ((& $toPixel $h2X2) + $marginPx)
            $h3Ink = Measure-Crop $bitmap ((& $toPixel $h3X1) - $marginPx) ((& $toPixel $h3X2) + $marginPx)
            $textTops = @($h1Ink.Top, $h2Ink.Top, $h3Ink.Top)
            $textBottoms = @($h1Ink.Bottom, $h2Ink.Bottom, $h3Ink.Bottom)
            $textBottom = ($textBottoms | Measure-Object -Average).Average
            $topStats = $textTops | Measure-Object -Minimum -Maximum
            $bottomStats = $textBottoms | Measure-Object -Minimum -Maximum

            $anchorBottomDeltaPt = ($oleAnchorInk.Bottom - $ommlAnchorInk.Bottom) * $pxToPtY
            $anchorCentroidDeltaPt = ($oleAnchorInk.CentroidY - $ommlAnchorInk.CentroidY) * $pxToPtY
            $results.Add([pscustomobject]@{
                Name = $name
                BaseName = $baseName
                FontSizePt = $fontSizePt
                FormulaLetterFont = $formulaLetterFont
                PositionAdjustment = $positionAdjustment
                CalculatedPositionPt = $calculatedPosition
                AppliedPositionPt = $appliedPosition
                ObjectWidthPt = [Math]::Round($oleWidthPt, 3)
                ObjectHeightPt = [Math]::Round($oleHeightPt, 3)
                RenderHeightPx = [Math]::Round([double]$renderHeight, 4)
                RenderBaselinePx = [Math]::Round([double]$renderBaseline, 4)
                PreviewPixels = "$($previewInk.Width)x$($previewInk.Height)"
                PreviewInkPixels = "$($previewInk.InkWidth)x$($previewInk.InkHeight)"
                PreviewLeftMarginPx = $previewInk.LeftMargin
                PreviewTopMarginPx = $previewInk.TopMargin
                PreviewRightMarginPx = $previewInk.RightMargin
                PreviewBottomMarginPx = $previewInk.BottomMargin
                FormulaDeltaTopPt = [Math]::Round(($oleInk.Top - $ommlInk.Top) * $pxToPtY, 3)
                FormulaDeltaBottomPt = [Math]::Round(($oleInk.Bottom - $ommlInk.Bottom) * $pxToPtY, 3)
                FormulaDeltaCentroidPt = [Math]::Round(
                    ($oleInk.CentroidY - $ommlInk.CentroidY) * $pxToPtY,
                    3)
                AnchorDeltaTopPt = [Math]::Round(
                    ($oleAnchorInk.Top - $ommlAnchorInk.Top) * $pxToPtY,
                    3)
                AnchorDeltaBottomPt = [Math]::Round($anchorBottomDeltaPt, 3)
                AnchorDeltaCentroidPt = [Math]::Round($anchorCentroidDeltaPt, 3)
                AnchorScore = [Math]::Round(
                    [Math]::Abs($anchorBottomDeltaPt) + [Math]::Abs($anchorCentroidDeltaPt),
                    3)
                OleBottomVsTextPt = [Math]::Round(($oleInk.Bottom - $textBottom) * $pxToPtY, 3)
                OmmlBottomVsTextPt = [Math]::Round(($ommlInk.Bottom - $textBottom) * $pxToPtY, 3)
                TextTopSpreadPt = [Math]::Round(
                    ([double]$topStats.Maximum - [double]$topStats.Minimum) * $pxToPtY,
                    3)
                TextBottomSpreadPt = [Math]::Round(
                    ([double]$bottomStats.Maximum - [double]$bottomStats.Minimum) * $pxToPtY,
                    3)
                LineEmfPixels = "$($bitmap.Width)x$($bitmap.Height)"
                DpiY = [Math]::Round($dpiY, 3)
            })
            if (-not $savedFirstDocument -and -not [string]::IsNullOrWhiteSpace($SaveFirstDocumentPath)) {
                $stage = "save-diagnostic-docx"
                $resolvedSavePath = [IO.Path]::GetFullPath($SaveFirstDocumentPath)
                $saveDirectory = [IO.Path]::GetDirectoryName($resolvedSavePath)
                if (-not [string]::IsNullOrWhiteSpace($saveDirectory)) {
                    [void](New-Item -Path $saveDirectory -ItemType Directory -Force)
                }
                [void](Invoke-WithOfficeRetry {
                    $document.SaveAs2($resolvedSavePath, 16) # wdFormatDocumentDefault (.docx)
                })
                $savedFirstDocument = $true
            }
        }
        catch {
            throw [InvalidOperationException]::new(
                "Isolated Word baseline probe failed for '$([string]$item.name)' at stage '$stage'.",
                $_.Exception)
        }
        finally {
            if ($null -ne $bitmap) { $bitmap.Dispose() }
            if ($null -ne $metafile) { $metafile.Dispose() }
            if ($null -ne $stream) { $stream.Dispose() }
            Release-Com $lineRange
            if ($null -ne $document) {
                try { $document.Close(0) } catch {}
            }
            Release-Com $document
            $service = $null
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
    }

    $finalDocumentSnapshot = Capture-DocumentSnapshot $word
    $finalDocumentSnapshotJson = $finalDocumentSnapshot | ConvertTo-Json -Compress
    if ($finalDocumentSnapshotJson -ne $initialDocumentSnapshotJson) {
        throw "The pre-existing Word document set changed during the isolated baseline probe."
    }

    $positionSummary = @($results |
        Group-Object PositionAdjustment |
        Sort-Object { [int]$_.Name } |
        ForEach-Object {
            $rows = @($_.Group)
            [pscustomobject]@{
                PositionAdjustment = [int]$_.Name
                Count = $rows.Count
                AverageAbsoluteAnchorBottomDeltaPt = [Math]::Round(
                    ($rows | ForEach-Object {
                        [Math]::Abs([double]$_.AnchorDeltaBottomPt)
                    } | Measure-Object -Average).Average,
                    3)
                AverageAbsoluteAnchorCentroidDeltaPt = [Math]::Round(
                    ($rows | ForEach-Object {
                        [Math]::Abs([double]$_.AnchorDeltaCentroidPt)
                    } | Measure-Object -Average).Average,
                    3)
                AverageAnchorScore = [Math]::Round(
                    ($rows | Measure-Object AnchorScore -Average).Average,
                    3)
                MaximumAnchorScore = [Math]::Round(
                    ($rows | Measure-Object AnchorScore -Maximum).Maximum,
                    3)
            }
        })
    $bestAdjustmentByFormula = @($results |
        Group-Object BaseName |
        ForEach-Object {
            $_.Group |
                Sort-Object AnchorScore, {
                    [Math]::Abs([int]$_.PositionAdjustment)
                } |
                Select-Object -First 1 BaseName,PositionAdjustment,AppliedPositionPt,AnchorScore,AnchorDeltaBottomPt,AnchorDeltaCentroidPt
        } |
        Sort-Object BaseName)
    $unsafePreviewCount = @($results | Where-Object {
        $_.PreviewLeftMarginPx -lt 1 -or
        $_.PreviewTopMarginPx -lt 1 -or
        $_.PreviewRightMarginPx -lt 1 -or
        $_.PreviewBottomMarginPx -lt 1
    }).Count
    $maximumTextBaselineSpread = ($results | ForEach-Object {
        [Math]::Max(
            [Math]::Abs([double]$_.TextTopSpreadPt),
            [Math]::Abs([double]$_.TextBottomSpreadPt))
    } | Measure-Object -Maximum).Maximum

    $report = [pscustomobject]@{
        Manifest = (Resolve-Path $ManifestPath).Path
        ExistingDocumentsUnchanged = $true
        Count = $results.Count
        Summary = [pscustomobject]@{
            UnsafePreviewMarginCount = $unsafePreviewCount
            MaximumFollowingTextBaselineSpreadPt = [Math]::Round(
                [double]$maximumTextBaselineSpread,
                3)
            PositionAdjustments = $positionSummary
            BestAdjustmentByFormula = $bestAdjustmentByFormula
        }
        Results = $results
    }
    $json = $report | ConvertTo-Json -Depth 9
    if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
        $resolvedOutput = [IO.Path]::GetFullPath($OutputPath)
        $outputDirectory = [IO.Path]::GetDirectoryName($resolvedOutput)
        if (-not [string]::IsNullOrWhiteSpace($outputDirectory)) {
            [IO.Directory]::CreateDirectory($outputDirectory) | Out-Null
        }
        [IO.File]::WriteAllText(
            $resolvedOutput,
            $json + [Environment]::NewLine,
            [Text.UTF8Encoding]::new($false))
    }

    [pscustomobject]@{
        OutputPath = if ([string]::IsNullOrWhiteSpace($OutputPath)) { $null } else { [IO.Path]::GetFullPath($OutputPath) }
        ExistingDocumentsUnchanged = $true
        Count = $results.Count
        UnsafePreviewMarginCount = $unsafePreviewCount
        MaximumFollowingTextBaselineSpreadPt = [Math]::Round(
            [double]$maximumTextBaselineSpread,
            3)
        PositionAdjustments = $positionSummary
        BestAdjustmentByFormula = $bestAdjustmentByFormula
        WorstAnchorScores = @($results |
            Sort-Object AnchorScore -Descending |
            Select-Object -First 12 Name,CalculatedPositionPt,AppliedPositionPt,AnchorScore,AnchorDeltaBottomPt,AnchorDeltaCentroidPt)
    } | ConvertTo-Json -Depth 8
}
finally {
    try { $word.ScreenUpdating = $screenUpdating } catch {}
    try { $word.DisplayAlerts = $displayAlerts } catch {}
    Release-Com $word
    foreach ($path in $generatedFiles) {
        try { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue } catch {}
    }
    try {
        Remove-Item -LiteralPath $probePreviewRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
    catch {}
    if ($null -ne $assemblyResolver) {
        try { [AppDomain]::CurrentDomain.remove_AssemblyResolve($assemblyResolver) } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
