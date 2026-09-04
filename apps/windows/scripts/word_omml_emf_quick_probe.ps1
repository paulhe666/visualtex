[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestPath,
    [int]$Count = 2
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

function Load-AssemblyDependencies([string]$OutputRoot) {
    foreach ($name in @(
        "System.Runtime.CompilerServices.Unsafe.dll",
        "System.Memory.dll",
        "System.Buffers.dll",
        "Microsoft.Bcl.AsyncInterfaces.dll",
        "System.Threading.Tasks.Extensions.dll",
        "System.Text.Encodings.Web.dll",
        "System.Text.Json.dll",
        "Microsoft.Office.Interop.Word.dll"
    )) {
        $path = Join-Path $OutputRoot $name
        if (Test-Path $path) { [void][Reflection.Assembly]::LoadFrom($path) }
    }
}

function Measure-Emf([byte[]]$Bytes) {
    $stream = [IO.MemoryStream]::new($Bytes, $false)
    $metafile = [Drawing.Imaging.Metafile]::new($stream)
    try {
        $width = [Math]::Max(1, [int]$metafile.Width)
        $height = [Math]::Max(1, [int]$metafile.Height)
        $bitmap = [Drawing.Bitmap]::new($width, $height, [Drawing.Imaging.PixelFormat]::Format32bppArgb)
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        try {
            $graphics.Clear([Drawing.Color]::White)
            $graphics.DrawImage($metafile, 0, 0, $width, $height)
        }
        finally { $graphics.Dispose() }
        try {
            $minX = $width
            $minY = $height
            $maxX = -1
            $maxY = -1
            for ($y = 0; $y -lt $height; $y++) {
                for ($x = 0; $x -lt $width; $x++) {
                    $p = $bitmap.GetPixel($x, $y)
                    if ($p.R -ge 245 -and $p.G -ge 245 -and $p.B -ge 245) { continue }
                    if ($x -lt $minX) { $minX = $x }
                    if ($x -gt $maxX) { $maxX = $x }
                    if ($y -lt $minY) { $minY = $y }
                    if ($y -gt $maxY) { $maxY = $y }
                }
            }
            return [pscustomobject]@{
                WidthPx = $width
                HeightPx = $height
                DpiX = [double]$metafile.HorizontalResolution
                DpiY = [double]$metafile.VerticalResolution
                InkLeftPx = $minX
                InkTopPx = $minY
                InkRightPx = $maxX
                InkBottomPx = $maxY
                InkWidthPx = if ($maxX -ge $minX) { $maxX - $minX + 1 } else { 0 }
                InkHeightPx = if ($maxY -ge $minY) { $maxY - $minY + 1 } else { 0 }
            }
        }
        finally { $bitmap.Dispose() }
    }
    finally {
        $metafile.Dispose()
        $stream.Dispose()
    }
}

$manifest = Get-Content -LiteralPath (Resolve-Path $ManifestPath) -Raw | ConvertFrom-Json
$items = @(@($manifest.items) | Select-Object -First $Count)
if ($items.Count -lt 1) { throw "Manifest has no items." }

$root = Split-Path -Parent $PSScriptRoot
$vstoOutput = Join-Path $root "src-windows\VisualTeX.WordVsto\bin\x64\Release\net472"
Load-AssemblyDependencies $vstoOutput
$assemblyResolver = [ResolveEventHandler]{
    param($sender, $eventArgs)
    try {
        $simpleName = ([Reflection.AssemblyName]$eventArgs.Name).Name
        $candidate = Join-Path $vstoOutput ($simpleName + ".dll")
        if (Test-Path $candidate) { return [Reflection.Assembly]::LoadFrom($candidate) }
    }
    catch {}
    return $null
}
[AppDomain]::CurrentDomain.add_AssemblyResolve($assemblyResolver)
$contracts = [Reflection.Assembly]::LoadFrom((Join-Path $vstoOutput "VisualTeX.WindowsOffice.Contracts.dll"))
$wordAssembly = [Reflection.Assembly]::LoadFrom((Join-Path $vstoOutput "VisualTeX.WordVsto.dll"))
$sessionType = $contracts.GetType("VisualTeX.WindowsOffice.Contracts.OfficeSessionDocument", $true)
$lineType = $contracts.GetType("VisualTeX.WindowsOffice.Contracts.FormulaLine", $true)
$exportType = $contracts.GetType("VisualTeX.WindowsOffice.Contracts.OfficeExportDocument", $true)
$serviceType = $wordAssembly.GetType("VisualTeX.WordVsto.WordFormulaService", $true)
$storeType = $wordAssembly.GetType("VisualTeX.WordVsto.WordOmmlFormulaStore", $true)
$listType = [Collections.Generic.List``1].MakeGenericType($lineType)
$insertOmml = $serviceType.GetMethods() | Where-Object { $_.Name -eq "InsertOmml" -and $_.GetParameters().Count -eq 8 } | Select-Object -First 1
$findById = $storeType.GetMethod("FindByFormulaId", [Reflection.BindingFlags]"Static,Public,NonPublic")
$getEquationRange = $storeType.GetMethod("GetEquationRange", [Reflection.BindingFlags]"Static,Public,NonPublic")
if ($null -eq $insertOmml -or $null -eq $findById -or $null -eq $getEquationRange) { throw "Required Word methods unavailable." }

$word = New-Object -ComObject Word.Application
$wordInteropType = ([Reflection.Assembly]::LoadFrom((Join-Path $vstoOutput "Microsoft.Office.Interop.Word.dll"))).GetType("Microsoft.Office.Interop.Word.Application", $true)
$typedWord = $word -as $wordInteropType
if ($null -eq $typedWord) { throw "Unable to cast Word COM object to interop Application." }
$word.Visible = $false
$word.DisplayAlerts = 0
$word.ScreenUpdating = $false
$doc = $null
$service = $null
$results = [Collections.Generic.List[object]]::new()
try {
    $doc = $word.Documents.Add()
    $doc.PageSetup.PageWidth = 1008
    $doc.PageSetup.PageHeight = 612
    $doc.PageSetup.LeftMargin = 18
    $doc.PageSetup.RightMargin = 18
    $service = [Activator]::CreateInstance($serviceType, @($typedWord))

    foreach ($item in $items) {
        $selection = $word.Selection
        try {
            $selection.SetRange([int]$doc.Content.End - 1, [int]$doc.Content.End - 1)
            $selection.Font.Name = "Times New Roman"
            $selection.Font.Size = [single]$item.fontSizePt
            $selection.Font.Position = 0
        }
        finally { Release-Com $selection }

        $line = [Activator]::CreateInstance($lineType)
        $line.Id = [guid]::NewGuid().ToString("D")
        $line.Latex = [string]$item.latex
        $lines = [Activator]::CreateInstance($listType)
        $lines.Add($line)
        $export = [Activator]::CreateInstance($exportType)
        $export.Width = [single]$item.width
        $export.Height = [single]$item.height
        $exportType.GetProperty("Baseline").SetValue($export, [single]$item.baseline, $null)
        $export.FormulaLetterFont = [string]$item.formulaLetterFont
        $export.FormulaChineseFont = [string]$item.formulaChineseFont
        $session = [Activator]::CreateInstance($sessionType)
        $session.Id = [guid]::NewGuid().ToString("D")
        $session.Mode = "create"
        $session.Host = "word"
        $session.FormulaId = [guid]::NewGuid().ToString("D")
        $session.Title = "VisualTeX Word OMML EMF probe"
        $session.Lines = $lines
        $session.CodeFormat = "latex"
        $session.DisplayMode = "inline"
        $session.ObjectMode = "wordOmml"
        $session.Numbered = $false
        $session.FontSizePt = [double]$item.fontSizePt
        $session.Status = "ready"
        $session.Dirty = $true
        $session.ExportResult = $export

        $sw = [Diagnostics.Stopwatch]::StartNew()
        $args = [object[]]::new(8)
        $args[0] = $session
        $args[1] = [string]$item.mathMl
        $args[2] = $false
        $args[3] = $false
        $args[4] = $false
        $args[5] = $null
        $args[6] = $false
        $args[7] = $false
        [void]$insertOmml.Invoke($service, $args)
        $sw.Stop()

        $bookmark = $findById.Invoke($null, @($doc, [string]$session.FormulaId))
        if ($null -eq $bookmark) { throw "Inserted OMML bookmark missing for $($item.name)." }
        $range = $getEquationRange.Invoke($null, @($bookmark))
        try {
            $extract = [Diagnostics.Stopwatch]::StartNew()
            [byte[]]$bytes = $range.EnhMetaFileBits
            $extract.Stop()
            if ($null -eq $bytes -or $bytes.Length -lt 88) { throw "OMML returned no usable EMF for $($item.name)." }
            $metrics = Measure-Emf $bytes
            $copySw = [Diagnostics.Stopwatch]::StartNew()
            $range.CopyAsPicture()
            $copySw.Stop()
            $clipboardType = ""
            $clipboardWidth = 0
            $clipboardHeight = 0
            $clipboardDpiX = 0
            $clipboardDpiY = 0
            try {
                $dataObject = [Windows.Forms.Clipboard]::GetDataObject()
                if ($null -ne $dataObject) {
                    $clipboard = $dataObject.GetData([Windows.Forms.DataFormats]::EnhancedMetafile)
                    if ($null -ne $clipboard) {
                        $clipboardType = $clipboard.GetType().FullName
                        if ($clipboard -is [Drawing.Imaging.Metafile]) {
                            $clipboardWidth = [int]$clipboard.Width
                            $clipboardHeight = [int]$clipboard.Height
                            $clipboardDpiX = [double]$clipboard.HorizontalResolution
                            $clipboardDpiY = [double]$clipboard.VerticalResolution
                            $clipboard.Dispose()
                        }
                    }
                }
            }
            catch { $clipboardType = "clipboard-error: $($_.Exception.Message)" }
            $x1 = [double]$range.Information(5)
            $endRange = $doc.Range([int]$range.End, [int]$range.End)
            try { $x2 = [double]$endRange.Information(5) } finally { Release-Com $endRange }
            $results.Add([pscustomobject]@{
                Name = [string]$item.name
                InsertOmmlMs = $sw.Elapsed.TotalMilliseconds
                ExtractEmfMs = $extract.Elapsed.TotalMilliseconds
                EmfBytes = $bytes.Length
                RangeWidthPt = [Math]::Round($x2 - $x1, 3)
                EmfWidthPx = $metrics.WidthPx
                EmfHeightPx = $metrics.HeightPx
                EmfDpiX = [Math]::Round($metrics.DpiX, 2)
                EmfDpiY = [Math]::Round($metrics.DpiY, 2)
                InkWidthPx = $metrics.InkWidthPx
                InkHeightPx = $metrics.InkHeightPx
                InkTopPx = $metrics.InkTopPx
                InkBottomPx = $metrics.InkBottomPx
                CopyAsPictureMs = $copySw.Elapsed.TotalMilliseconds
                ClipboardType = $clipboardType
                ClipboardWidthPx = $clipboardWidth
                ClipboardHeightPx = $clipboardHeight
                ClipboardDpiX = [Math]::Round($clipboardDpiX, 2)
                ClipboardDpiY = [Math]::Round($clipboardDpiY, 2)
            })
        }
        finally {
            Release-Com $range
            Release-Com $bookmark
        }

        $selection = $word.Selection
        try {
            $selection.SetRange([int]$doc.Content.End - 1, [int]$doc.Content.End - 1)
            $selection.TypeParagraph()
        }
        finally { Release-Com $selection }
    }

    [pscustomobject]@{
        Count = $results.Count
        Results = $results
    } | ConvertTo-Json -Depth 6
}
finally {
    if ($null -ne $doc) { try { $doc.Close(0) } catch {}; Release-Com $doc }
    if ($null -ne $word) { try { $word.Quit(0) } catch {}; Release-Com $word }
    if ($null -ne $assemblyResolver) {
        try { [AppDomain]::CurrentDomain.remove_AssemblyResolve($assemblyResolver) } catch {}
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
