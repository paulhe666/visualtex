[CmdletBinding()]
param(
    [string]$SidecarPath,
    [ValidateRange(20, 200)]
    [int]$FormulaCount = 20,
    [switch]$KeepDocuments,
    [switch]$TestModeSwitch,
    [string]$VstoMsiPath
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

if (-not ("VisualTeXWordAutomation" -as [type])) {
    Add-Type -TypeDefinition @"
public static class VisualTeXWordAutomation
{
    public static void SaveAsDocx(object document, string path)
    {
        dynamic value = document;
        value.SaveAs2(FileName: path, FileFormat: 12);
    }
}
"@ -ReferencedAssemblies Microsoft.CSharp
}

if ($PSVersionTable.PSEdition -eq "Core" -and -not $IsWindows) {
    throw "This acceptance suite must run on Windows with desktop Microsoft Word and PowerPoint installed."
}

$root = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($SidecarPath)) {
    $SidecarPath = Join-Path $root "src-tauri\binaries\visualtex-windows-office-bridge-x86_64-pc-windows-msvc.exe"
}
if (-not (Test-Path $SidecarPath)) {
    throw "Windows Office sidecar is missing: $SidecarPath"
}

function Assert-True([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw "ACCEPTANCE FAILURE: $Message" }
}

function Release-ComObject($Value) {
    if ($null -eq $Value) { return }
    try {
        if ([Runtime.InteropServices.Marshal]::IsComObject($Value)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
        }
    } catch { }
}

function New-FormulaMetadata([string]$FormulaId, [string]$Latex, [string]$DisplayMode = "block") {
    $lineId = [Guid]::NewGuid().ToString()
    $now = [DateTimeOffset]::UtcNow.ToString("O")
    return [ordered]@{
        schema = "visualtex-formula"
        schemaVersion = 1
        formulaId = $FormulaId
        title = "Acceptance $FormulaId"
        latex = $Latex
        lines = @([ordered]@{ id = $lineId; latex = $Latex })
        codeFormat = "latex"
        displayMode = $DisplayMode
        createdWithVersion = "1.0.18"
        updatedWithVersion = "1.0.18"
        createdAt = $now
        updatedAt = $now
    }
}

function New-FormulaParams(
    [string]$FormulaId,
    [string]$ImagePath,
    [string]$Latex,
    [string]$DisplayMode = "block",
    [string]$SourceDocumentId = "",
    [string]$SourceObjectId = "",
    $Baseline = $null,
    [double]$Width = 240.0,
    [double]$Height = 72.0
) {
    return [ordered]@{
        sessionId = [Guid]::NewGuid().ToString()
        formulaId = $FormulaId
        imagePath = $ImagePath
        metadata = New-FormulaMetadata $FormulaId $Latex $DisplayMode
        width = $Width
        height = $Height
        baseline = $Baseline
        sourceDocumentId = $(if ($SourceDocumentId) { $SourceDocumentId } else { $null })
        sourceObjectId = $(if ($SourceObjectId) { $SourceObjectId } else { $null })
    }
}

function Decode-FormulaMetadata([string]$Value) {
    $prefix = "visualtex:v1:deflate:"
    $index = $Value.IndexOf($prefix, [StringComparison]::Ordinal)
    if ($index -lt 0) { return $null }
    $encoded = $Value.Substring($index + $prefix.Length).Trim()
    $separator = $encoded.IndexOfAny([char[]]@("`r", "`n", " ", "`t"))
    if ($separator -ge 0) { $encoded = $encoded.Substring(0, $separator) }
    $normalized = $encoded.Replace('-', '+').Replace('_', '/')
    while (($normalized.Length % 4) -ne 0) { $normalized += "=" }
    $bytes = [Convert]::FromBase64String($normalized)
    $input = [IO.MemoryStream]::new($bytes)
    try {
        $deflate = [IO.Compression.DeflateStream]::new(
            $input,
            [IO.Compression.CompressionMode]::Decompress)
        try {
            $reader = [IO.StreamReader]::new($deflate, [Text.Encoding]::UTF8)
            try { return ($reader.ReadToEnd() | ConvertFrom-Json) }
            finally { $reader.Dispose() }
        } finally { $deflate.Dispose() }
    } finally { $input.Dispose() }
}

function Invoke-Bridge(
    [string]$Method,
    $Params,
    [switch]$AllowFailure
) {
    $client = [IO.Pipes.NamedPipeClientStream]::new(
        ".",
        $script:PipeLeaf,
        [IO.Pipes.PipeDirection]::InOut,
        [IO.Pipes.PipeOptions]::Asynchronous)
    try {
        $client.Connect(5000)
        $encoding = [Text.UTF8Encoding]::new($false)
        $reader = [IO.StreamReader]::new($client, $encoding, $false, 16384, $true)
        $writer = [IO.StreamWriter]::new($client, $encoding, 16384, $true)
        try {
            $writer.NewLine = "`n"
            $writer.AutoFlush = $true
            $handshakeId = [Guid]::NewGuid().ToString()
            $handshake = [ordered]@{
                protocolVersion = 1
                id = $handshakeId
                method = "handshake"
                params = [ordered]@{ token = $script:Token }
            } | ConvertTo-Json -Compress -Depth 40
            $writer.WriteLine($handshake)
            $handshakeResponse = $reader.ReadLine() | ConvertFrom-Json
            Assert-True ($handshakeResponse.ok -eq $true) "Named-pipe token handshake failed."

            $requestId = [Guid]::NewGuid().ToString()
            $request = [ordered]@{
                protocolVersion = 1
                id = $requestId
                method = $Method
                params = $(if ($null -eq $Params) { [ordered]@{} } else { $Params })
            } | ConvertTo-Json -Compress -Depth 40
            $writer.WriteLine($request)
            $responseLine = $reader.ReadLine()
            Assert-True (-not [string]::IsNullOrWhiteSpace($responseLine)) "$Method returned no response."
            $response = $responseLine | ConvertFrom-Json
            Assert-True ($response.id -eq $requestId) "$Method returned a mismatched response id."
            if (-not $AllowFailure -and $response.ok -ne $true) {
                throw "Bridge method $Method failed: $($response.error.code): $($response.error.message)"
            }
            return $response
        } finally {
            $writer.Dispose()
            $reader.Dispose()
        }
    } finally { $client.Dispose() }
}

function Start-BridgeProcess {
    $argumentLine = "--parent-pid $PID --pipe-name `"$script:FullPipeName`" --token $script:Token --temp-root `"$script:TempRoot`" --log-root `"$script:LogRoot`" --acceptance true"
    $process = Start-Process -FilePath $SidecarPath -ArgumentList $argumentLine -PassThru -WindowStyle Hidden
    $sidecarFullPath = [IO.Path]::GetFullPath($SidecarPath)
    $deadline = [DateTimeOffset]::UtcNow.AddSeconds(15)
    do {
        Start-Sleep -Milliseconds 150
        try {
            $health = Invoke-Bridge "health" ([ordered]@{})
            if ($health.ok) {
                # A self-contained single-file .NET executable may hand off from
                # its short-lived launcher to the extracted app process. Track
                # the process that owns this acceptance token instead of treating
                # a successful launcher exit as a bridge failure.
                $candidate = Get-CimInstance Win32_Process | Where-Object {
                    -not [string]::IsNullOrWhiteSpace($_.ExecutablePath) -and
                    [string]::Equals(
                        [IO.Path]::GetFullPath($_.ExecutablePath),
                        $sidecarFullPath,
                        [StringComparison]::OrdinalIgnoreCase) -and
                    -not [string]::IsNullOrWhiteSpace($_.CommandLine) -and
                    $_.CommandLine.Contains($script:Token)
                } | Sort-Object CreationDate -Descending | Select-Object -First 1
                if ($null -ne $candidate) {
                    try { return Get-Process -Id $candidate.ProcessId -ErrorAction Stop } catch { }
                }
                return $process
            }
        } catch {
            if ($process.HasExited -and $process.ExitCode -ne 0) {
                throw "Windows Office sidecar exited before becoming healthy (exit $($process.ExitCode))."
            }
        }
    } while ([DateTimeOffset]::UtcNow -lt $deadline)
    try { Stop-Process -Id $process.Id -Force } catch { }
    Get-CimInstance Win32_Process | Where-Object {
        -not [string]::IsNullOrWhiteSpace($_.CommandLine) -and
        $_.CommandLine.Contains($script:Token)
    } | ForEach-Object { try { Stop-Process -Id $_.ProcessId -Force } catch { } }
    throw "Windows Office sidecar did not become healthy in 15 seconds."
}

function Stop-BridgeProcess($Process) {
    $sidecarFullPath = [IO.Path]::GetFullPath($SidecarPath)

    function Find-TokenBridgeProcesses {
        return @(Get-CimInstance Win32_Process | Where-Object {
            $_.Name -like "*windows-office-bridge*" -and
            -not [string]::IsNullOrWhiteSpace($_.ExecutablePath) -and
            [string]::Equals(
                [IO.Path]::GetFullPath($_.ExecutablePath),
                $sidecarFullPath,
                [StringComparison]::OrdinalIgnoreCase) -and
            -not [string]::IsNullOrWhiteSpace($_.CommandLine) -and
            $_.CommandLine.Contains($script:Token)
        })
    }

    $processIds = New-Object System.Collections.Generic.HashSet[int]
    foreach ($candidateProcess in @($Process)) {
        if ($null -eq $candidateProcess) { continue }
        try { [void]$processIds.Add([int]$candidateProcess.Id) } catch { }
    }
    foreach ($match in @(Find-TokenBridgeProcesses)) {
        [void]$processIds.Add([int]$match.ProcessId)
    }

    foreach ($processId in $processIds) {
        try {
            $runtimeProcess = [Diagnostics.Process]::GetProcessById($processId)
            $runtimeProcess.Kill()
            [void]$runtimeProcess.WaitForExit(3000)
            $runtimeProcess.Dispose()
        } catch {
            try {
                & "$env:SystemRoot\System32\taskkill.exe" /PID $processId /T /F | Out-Null
            } catch { }
        }
    }

    $deadline = [DateTimeOffset]::UtcNow.AddSeconds(5)
    do {
        $remaining = @(Find-TokenBridgeProcesses)
        if ($remaining.Count -eq 0) { return }
        foreach ($match in $remaining) {
            try {
                & "$env:SystemRoot\System32\taskkill.exe" /PID $match.ProcessId /T /F | Out-Null
            } catch { }
        }
        Start-Sleep -Milliseconds 100
    } while ([DateTimeOffset]::UtcNow -lt $deadline)
    throw "Windows Office bridge process did not stop within 5 seconds."
}

function New-TestPng(
    [string]$Path,
    [int]$Width,
    [int]$Height,
    [string]$Label,
    [switch]$Transparent
) {
    Add-Type -AssemblyName System.Drawing
    $bitmap = [Drawing.Bitmap]::new($Width, $Height)
    try {
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        try {
            $graphics.Clear($(if ($Transparent) { [Drawing.Color]::Transparent } else { [Drawing.Color]::White }))
            $font = [Drawing.Font]::new("Arial", [Math]::Max(10, [Math]::Min(32, $Height / 3)))
            try {
                $graphics.DrawString($Label, $font, [Drawing.Brushes]::Black, 8, 8)
            } finally { $font.Dispose() }
        } finally { $graphics.Dispose() }
        $bitmap.Save($Path, [Drawing.Imaging.ImageFormat]::Png)
    } finally { $bitmap.Dispose() }
}

function Get-WordFormulaIds($Document) {
    $ids = New-Object System.Collections.Generic.List[string]
    $shapes = $null
    try {
        $shapes = $Document.InlineShapes
        for ($index = 1; $index -le $shapes.Count; $index += 1) {
            $shape = $null
            try {
                $shape = $shapes.Item($index)
                $fromAlt = Decode-FormulaMetadata ([string]$shape.AlternativeText)
                $fromTitle = Decode-FormulaMetadata ([string]$shape.Title)
                Assert-True ($null -ne $fromAlt) "Word InlineShape $index has no AlternativeText metadata."
                Assert-True ($null -ne $fromTitle) "Word InlineShape $index has no Title metadata."
                Assert-True ($fromAlt.formulaId -eq $fromTitle.formulaId) "Word Title and AlternativeText metadata disagree."
                $ids.Add([string]$fromAlt.formulaId)
            } finally { Release-ComObject $shape }
        }
    } finally { Release-ComObject $shapes }
    return @($ids)
}

function Get-WordFormulaFontPosition($Document, [string]$FormulaId) {
    $shapes = $null
    try {
        $shapes = $Document.InlineShapes
        for ($index = 1; $index -le $shapes.Count; $index += 1) {
            $shape = $null
            $range = $null
            $font = $null
            try {
                $shape = $shapes.Item($index)
                $metadata = Decode-FormulaMetadata ([string]$shape.AlternativeText)
                if ($metadata.formulaId -ne $FormulaId) { continue }
                $range = $shape.Range
                $font = $range.Font
                return [int]$font.Position
            } finally {
                Release-ComObject $font
                Release-ComObject $range
                Release-ComObject $shape
            }
        }
    } finally { Release-ComObject $shapes }
    throw "Word formula was not found while checking baseline alignment: $FormulaId"
}

function Get-WordFormulaSize($Document, [string]$FormulaId) {
    $shapes = $null
    try {
        $shapes = $Document.InlineShapes
        for ($index = 1; $index -le $shapes.Count; $index += 1) {
            $shape = $null
            try {
                $shape = $shapes.Item($index)
                $metadata = Decode-FormulaMetadata ([string]$shape.AlternativeText)
                if ($metadata.formulaId -ne $FormulaId) { continue }
                return [PSCustomObject]@{
                    Width = [double]$shape.Width
                    Height = [double]$shape.Height
                }
            } finally { Release-ComObject $shape }
        }
    } finally { Release-ComObject $shapes }
    throw "Word formula was not found while checking size: $FormulaId"
}

function Assert-WordFormulaSize(
    $Document,
    [string]$FormulaId,
    [double]$ExpectedWidth,
    [double]$ExpectedHeight,
    [string]$Stage
) {
    $size = Get-WordFormulaSize $Document $FormulaId
    $tolerance = 0.25
    Assert-True ([Math]::Abs($size.Width - $ExpectedWidth) -le $tolerance) `
        "$Stage width was $($size.Width) pt instead of $ExpectedWidth pt."
    Assert-True ([Math]::Abs($size.Height - $ExpectedHeight) -le $tolerance) `
        "$Stage height was $($size.Height) pt instead of $ExpectedHeight pt."
}

function Add-WordFormulaContextText($Document, [string]$FormulaId) {
    $shapes = $null
    try {
        $shapes = $Document.InlineShapes
        for ($index = 1; $index -le $shapes.Count; $index += 1) {
            $shape = $null
            $shapeRange = $null
            $before = $null
            $after = $null
            try {
                $shape = $shapes.Item($index)
                $metadata = Decode-FormulaMetadata ([string]$shape.AlternativeText)
                if ($metadata.formulaId -ne $FormulaId) { continue }
                $shapeRange = $shape.Range
                $before = $Document.Range($shapeRange.Start, $shapeRange.Start)
                $after = $Document.Range($shapeRange.End, $shapeRange.End)
                $before.InsertBefore("Inline text A ")
                $after.InsertAfter(" B aligned text")
                $before.Font.Size = 14
                $after.Font.Size = 14
                return
            } finally {
                Release-ComObject $after
                Release-ComObject $before
                Release-ComObject $shapeRange
                Release-ComObject $shape
            }
        }
    } finally { Release-ComObject $shapes }
    throw "Word formula was not found while adding alignment context: $FormulaId"
}

function Get-PowerPointFormulaIds($Slide) {
    $ids = New-Object System.Collections.Generic.List[string]
    $shapes = $null
    try {
        $shapes = $Slide.Shapes
        for ($index = 1; $index -le $shapes.Count; $index += 1) {
            $shape = $null
            $tags = $null
            try {
                $shape = $shapes.Item($index)
                if (-not ([string]$shape.Name).StartsWith("VisualTeX_")) { continue }
                $tags = $shape.Tags
                $formulaId = [string]$tags.Item("VisualTeXFormulaId")
                $metadata = Decode-FormulaMetadata ([string]$shape.AlternativeText)
                Assert-True (-not [string]::IsNullOrWhiteSpace($formulaId)) "PowerPoint shape $index has no formula-id tag."
                Assert-True ($null -ne $metadata) "PowerPoint shape $index has no AlternativeText metadata."
                Assert-True ($formulaId -eq $metadata.formulaId) "PowerPoint Tags and AlternativeText metadata disagree."
                Assert-True ($shape.Name -eq "VisualTeX_$formulaId") "PowerPoint object name does not match its UUID."
                $ids.Add($formulaId)
            } finally {
                Release-ComObject $tags
                Release-ComObject $shape
            }
        }
    } finally { Release-ComObject $shapes }
    return @($ids)
}

function Assert-SameIdSet([string[]]$Expected, [string[]]$Actual, [string]$Context) {
    $expectedSorted = @($Expected | Sort-Object)
    $actualSorted = @($Actual | Sort-Object)
    Assert-True ($expectedSorted.Count -eq $actualSorted.Count) "$Context formula count changed."
    for ($index = 0; $index -lt $expectedSorted.Count; $index += 1) {
        Assert-True ($expectedSorted[$index] -eq $actualSorted[$index]) "$Context formula UUID set changed."
    }
}

$sid = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
# Acceptance mode keeps the production SID binding while using a PID-scoped
# pipe and mutex, so it can coexist with an installed VisualTeX bridge.
$script:PipeLeaf = "VisualTeX.OfficeBridge.$sid.Acceptance.$PID"
$script:FullPipeName = "\\.\pipe\$($script:PipeLeaf)"
$script:Token = -join ((1..32) | ForEach-Object { '{0:x2}' -f (Get-Random -Minimum 0 -Maximum 256) })
$script:TempRoot = Join-Path $env:LOCALAPPDATA "VisualTeX\office\temp"
$script:LogRoot = Join-Path $env:LOCALAPPDATA "VisualTeX\office\acceptance-logs"
$artifactRoot = Join-Path $env:TEMP ("VisualTeX-Office-Acceptance-" + [Guid]::NewGuid().ToString("N"))
New-Item $script:TempRoot -ItemType Directory -Force | Out-Null
New-Item $script:LogRoot -ItemType Directory -Force | Out-Null
New-Item $artifactRoot -ItemType Directory -Force | Out-Null
$widePng = Join-Path $script:TempRoot ([Guid]::NewGuid().ToString() + ".png")
$squarePng = Join-Path $script:TempRoot ([Guid]::NewGuid().ToString() + ".png")
$longPng = Join-Path $script:TempRoot ([Guid]::NewGuid().ToString() + ".png")
$inlinePng = Join-Path $script:TempRoot ([Guid]::NewGuid().ToString() + ".png")
New-TestPng $widePng 480 120 "VisualTeX wide formula"
New-TestPng $squarePng 240 240 "matrix"
New-TestPng $longPng 720 90 "long integral and summation"
New-TestPng $inlinePng 160 40 "x + y" -Transparent

$bridgeProcess = $null
$word = $null
$wordDocument = $null
$powerPoint = $null
$presentation = $null
$slide = $null
$wordPath = Join-Path $artifactRoot "VisualTeX-Word-Acceptance.docx"
$powerPointPath = Join-Path $artifactRoot "VisualTeX-PowerPoint-Acceptance.pptx"
$preexistingOfficeProcessIds = @(
    Get-Process WINWORD, POWERPNT -ErrorAction SilentlyContinue |
        Select-Object -ExpandProperty Id
)
$acceptanceOfficeProcessIds = @()

try {
    Write-Host "[1/10] Starting authenticated current-user named-pipe bridge..."
    $bridgeProcess = Start-BridgeProcess
    $detectBefore = Invoke-Bridge "office.detect" ([ordered]@{})

    Write-Host "[2/10] Creating Word and PowerPoint test documents..."
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $word.DisplayAlerts = 0
    $wordDocument = $word.Documents.Add()
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1
    $powerPoint.DisplayAlerts = 1
    $presentation = $powerPoint.Presentations.Add()
    $slide = $presentation.Slides.Add(1, 12)
    Start-Sleep -Milliseconds 500
    $acceptanceOfficeProcessIds = @(
        Get-Process WINWORD, POWERPNT -ErrorAction SilentlyContinue |
            Where-Object { $_.Id -notin $preexistingOfficeProcessIds } |
            Select-Object -ExpandProperty Id
    )

    Write-Host "[3/10] Inserting $FormulaCount independent Word formulas..."
    $wordIds = New-Object System.Collections.Generic.List[string]
    for ($index = 0; $index -lt $FormulaCount; $index += 1) {
        $formulaId = [Guid]::NewGuid().ToString()
        $wordIds.Add($formulaId)
        $content = $null
        try {
            $content = $wordDocument.Content
            $word.Selection.SetRange([Math]::Max(0, $content.End - 1), [Math]::Max(0, $content.End - 1))
        } finally { Release-ComObject $content }
        $displayMode = $(if (($index % 2) -eq 0) { "inline" } else { "block" })
        $method = $(if ($displayMode -eq "inline") { "word.insertInlineFormula" } else { "word.insertDisplayFormula" })
        $latex = $(if (($index % 3) -eq 0) { "\sum_{i=0}^{n}x_i" } elseif (($index % 3) -eq 1) { "\prod_{k=1}^{m}a_k" } else { "\int_0^1 f(x)\,\mathrm{d}x" })
        $selection = (Invoke-Bridge "word.getSelection" ([ordered]@{})).result
        if ($index -eq 0) {
            $params = New-FormulaParams `
                -FormulaId $formulaId `
                -ImagePath $inlinePng `
                -Latex "x+y" `
                -DisplayMode "inline" `
                -SourceDocumentId ([string]$selection.documentId) `
                -SourceObjectId ([string]$selection.objectId) `
                -Baseline 12.0 `
                -Width 60.0 `
                -Height 15.0
            [void](Invoke-Bridge $method $params)
        } else {
            [void](Invoke-Bridge $method (New-FormulaParams $formulaId $widePng $latex $displayMode ([string]$selection.documentId) ([string]$selection.objectId)))
        }
    }
    $wordActual = Get-WordFormulaIds $wordDocument
    Assert-True ($wordActual.Count -eq $FormulaCount) "Word did not contain $FormulaCount formulas after insertion."
    Assert-SameIdSet @($wordIds) $wordActual "Word insert"
    $alignmentWordId = $wordIds[0]
    $alignmentPosition = Get-WordFormulaFontPosition $wordDocument $alignmentWordId
    Assert-True ($alignmentPosition -eq -3) "Word inline baseline alignment was $alignmentPosition pt instead of -3 pt."
    Add-WordFormulaContextText $wordDocument $alignmentWordId

    Write-Host "[4/10] Randomly editing Word formulas without increasing the count..."
    $wordEditIds = @($wordIds | Where-Object { $_ -ne $alignmentWordId } | Sort-Object { Get-Random } | Select-Object -First ([Math]::Min(10, $FormulaCount - 1)))
    foreach ($formulaId in $wordEditIds) {
        $params = New-FormulaParams $formulaId $squarePng "\begin{bmatrix}a&b\\c&d\end{bmatrix}" "inline"
        [void](Invoke-Bridge "word.replaceFormula" $params)
    }
    # Rapid repeated update of one persistent UUID.
    $rapidWordId = $wordEditIds[0]
    [void](Invoke-Bridge "word.replaceFormula" (New-FormulaParams $rapidWordId $longPng "\sum_{i=0}^{n}\int_0^1 f_i(x)\,\mathrm{d}x" "inline"))
    Assert-WordFormulaSize $wordDocument $rapidWordId 240.0 30.0 "Word wider edit"
    [void](Invoke-Bridge "word.replaceFormula" (New-FormulaParams $rapidWordId $widePng "a_i+b_i" "inline"))
    Assert-WordFormulaSize $wordDocument $rapidWordId 240.0 60.0 "Word edit restored to original aspect"
    $wordAfterEdit = Get-WordFormulaIds $wordDocument
    Assert-SameIdSet @($wordIds) $wordAfterEdit "Word edit"

    $outsideImage = Join-Path $artifactRoot "outside.png"
    Copy-Item $widePng $outsideImage
    $failedWord = Invoke-Bridge "word.replaceFormula" (New-FormulaParams $rapidWordId $outsideImage "failure" "inline") -AllowFailure
    Assert-True ($failedWord.ok -eq $false) "Word replacement outside the dedicated temp directory unexpectedly succeeded."
    Assert-SameIdSet @($wordIds) (Get-WordFormulaIds $wordDocument) "Word failed replacement"

    Write-Host "[5/10] Inserting $FormulaCount independent PowerPoint formulas..."
    $powerPointIds = New-Object System.Collections.Generic.List[string]
    for ($index = 0; $index -lt $FormulaCount; $index += 1) {
        $formulaId = [Guid]::NewGuid().ToString()
        $powerPointIds.Add($formulaId)
        $latex = $(if (($index % 2) -eq 0) { "\sum_{i=0}^{n}x_i" } else { "\begin{bmatrix}1&0\\0&1\end{bmatrix}" })
        $slide.Select()
        $selection = (Invoke-Bridge "powerpoint.getSelection" ([ordered]@{})).result
        [void](Invoke-Bridge "powerpoint.insertFormula" (New-FormulaParams $formulaId $widePng $latex "block" ([string]$selection.documentId) ([string]$selection.objectId)))
    }
    $pptActual = Get-PowerPointFormulaIds $slide
    Assert-True ($pptActual.Count -eq $FormulaCount) "PowerPoint did not contain $FormulaCount formulas after insertion."
    Assert-SameIdSet @($powerPointIds) $pptActual "PowerPoint insert"

    Write-Host "[6/10] Randomly editing PowerPoint formulas in place..."
    $pptEditIds = @($powerPointIds | Sort-Object { Get-Random } | Select-Object -First ([Math]::Min(10, $FormulaCount)))
    foreach ($formulaId in $pptEditIds) {
        $objectId = "VisualTeX_$formulaId"
        $params = New-FormulaParams $formulaId $squarePng "\begin{pmatrix}a&b\\c&d\end{pmatrix}" "block" ([string]$presentation.Name) $objectId
        [void](Invoke-Bridge "powerpoint.replaceFormula" $params)
    }
    $rapidPptId = $pptEditIds[0]
    [void](Invoke-Bridge "powerpoint.replaceFormula" (New-FormulaParams $rapidPptId $longPng "\int_{-\infty}^{\infty}e^{-x^2}\,\mathrm{d}x" "block" ([string]$presentation.Name) "VisualTeX_$rapidPptId"))
    [void](Invoke-Bridge "powerpoint.replaceFormula" (New-FormulaParams $rapidPptId $widePng "E=mc^2" "block" ([string]$presentation.Name) "VisualTeX_$rapidPptId"))
    Assert-SameIdSet @($powerPointIds) (Get-PowerPointFormulaIds $slide) "PowerPoint edit"

    $failedPpt = Invoke-Bridge "powerpoint.replaceFormula" (New-FormulaParams $rapidPptId $outsideImage "failure" "block" ([string]$presentation.Name) "VisualTeX_$rapidPptId") -AllowFailure
    Assert-True ($failedPpt.ok -eq $false) "PowerPoint replacement outside the dedicated temp directory unexpectedly succeeded."
    Assert-SameIdSet @($powerPointIds) (Get-PowerPointFormulaIds $slide) "PowerPoint failed replacement"

    Write-Host "[7/10] Testing delete, multiple documents/windows, undo and redo..."
    $deleteId = $powerPointIds[$powerPointIds.Count - 1]
    [void](Invoke-Bridge "powerpoint.deleteFormula" ([ordered]@{ formulaId = $deleteId }))
    $remainingPptIds = @($powerPointIds | Where-Object { $_ -ne $deleteId })
    Assert-SameIdSet $remainingPptIds (Get-PowerPointFormulaIds $slide) "PowerPoint delete"

    $secondDocument = $null
    $sameDeckSlide = $null
    $secondPresentation = $null
    $secondSlide = $null
    try {
        $wordDocument.Activate()
        $content = $null
        try {
            $content = $wordDocument.Content
            $word.Selection.SetRange([Math]::Max(0, $content.End - 1), [Math]::Max(0, $content.End - 1))
        } finally { Release-ComObject $content }
        $originalWordSource = (Invoke-Bridge "word.getSelection" ([ordered]@{})).result

        $secondDocument = $word.Documents.Add()
        $secondDocument.Activate()
        $wrongDocumentId = [Guid]::NewGuid().ToString()
        $wrongDocumentResult = Invoke-Bridge "word.insertInlineFormula" (
            New-FormulaParams $wrongDocumentId $widePng "must-not-insert" "inline" ([string]$originalWordSource.documentId) ([string]$originalWordSource.objectId)
        ) -AllowFailure
        Assert-True ($wrongDocumentResult.ok -eq $false) "Word accepted a formula after the active document changed."
        Assert-True (@(Get-WordFormulaIds $secondDocument).Count -eq 0) "Word wrote into the wrong active document."
        Assert-SameIdSet @($wordIds) (Get-WordFormulaIds $wordDocument) "Original Word document after rejected switch"

        $secondSource = (Invoke-Bridge "word.getSelection" ([ordered]@{})).result
        $secondWordId = [Guid]::NewGuid().ToString()
        [void](Invoke-Bridge "word.insertInlineFormula" (New-FormulaParams $secondWordId $widePng "x+y" "inline" ([string]$secondSource.documentId) ([string]$secondSource.objectId)))
        Assert-SameIdSet @($secondWordId) (Get-WordFormulaIds $secondDocument) "Second Word document"
        Assert-SameIdSet @($wordIds) (Get-WordFormulaIds $wordDocument) "Original Word document after window switch"
        [void]$secondDocument.Undo()
        Assert-True (@(Get-WordFormulaIds $secondDocument).Count -eq 0) "Word undo did not remove the latest formula."
        [void]$secondDocument.Redo()
        Assert-SameIdSet @($secondWordId) (Get-WordFormulaIds $secondDocument) "Word redo"

        $presentation.Windows.Item(1).Activate()
        $slide.Select()
        $originalSlideSource = (Invoke-Bridge "powerpoint.getSelection" ([ordered]@{})).result
        $sameDeckSlide = $presentation.Slides.Add(2, 12)
        $sameDeckSlide.Select()
        $sourceSlideFormulaId = [Guid]::NewGuid().ToString()
        [void](Invoke-Bridge "powerpoint.insertFormula" (
            New-FormulaParams $sourceSlideFormulaId $widePng "source-slide" "block" ([string]$originalSlideSource.documentId) ([string]$originalSlideSource.objectId)
        ))
        Assert-True (@(Get-PowerPointFormulaIds $sameDeckSlide).Count -eq 0) "PowerPoint wrote into the newly active slide instead of the recorded source slide."
        Assert-SameIdSet (@($remainingPptIds) + @($sourceSlideFormulaId)) (Get-PowerPointFormulaIds $slide) "PowerPoint source-slide insertion"
        [void](Invoke-Bridge "powerpoint.deleteFormula" ([ordered]@{ formulaId = $sourceSlideFormulaId }))
        Assert-SameIdSet $remainingPptIds (Get-PowerPointFormulaIds $slide) "PowerPoint source-slide cleanup"

        $secondPresentation = $powerPoint.Presentations.Add()
        $secondSlide = $secondPresentation.Slides.Add(1, 12)
        $secondPresentation.Windows.Item(1).Activate()
        $wrongPresentationId = [Guid]::NewGuid().ToString()
        $wrongPresentationResult = Invoke-Bridge "powerpoint.insertFormula" (
            New-FormulaParams $wrongPresentationId $widePng "must-not-insert" "block" ([string]$originalSlideSource.documentId) ([string]$originalSlideSource.objectId)
        ) -AllowFailure
        Assert-True ($wrongPresentationResult.ok -eq $false) "PowerPoint accepted a formula after the active presentation changed."
        Assert-True (@(Get-PowerPointFormulaIds $secondSlide).Count -eq 0) "PowerPoint wrote into the wrong presentation."

        $secondSlide.Select()
        $secondPptSource = (Invoke-Bridge "powerpoint.getSelection" ([ordered]@{})).result
        $secondPptId = [Guid]::NewGuid().ToString()
        [void](Invoke-Bridge "powerpoint.insertFormula" (New-FormulaParams $secondPptId $widePng "p=q" "block" ([string]$secondPptSource.documentId) ([string]$secondPptSource.objectId)))
        Assert-SameIdSet @($secondPptId) (Get-PowerPointFormulaIds $secondSlide) "Second PowerPoint presentation"
        Assert-SameIdSet $remainingPptIds (Get-PowerPointFormulaIds $slide) "Original PowerPoint after window switch"
        try {
            $powerPoint.CommandBars.ExecuteMso("Undo")
            Assert-True (@(Get-PowerPointFormulaIds $secondSlide).Count -eq 0) "PowerPoint undo did not remove the latest formula."
            $powerPoint.CommandBars.ExecuteMso("Redo")
            Assert-SameIdSet @($secondPptId) (Get-PowerPointFormulaIds $secondSlide) "PowerPoint redo"
        } catch {
            throw "PowerPoint undo/redo acceptance failed: $($_.Exception.Message)"
        }
    } finally {
        if ($null -ne $secondPresentation) { try { $secondPresentation.Close() } catch { } }
        if ($null -ne $secondDocument) { try { $secondDocument.Close(0) } catch { } }
        Release-ComObject $secondSlide
        Release-ComObject $secondPresentation
        Release-ComObject $sameDeckSlide
        Release-ComObject $secondDocument
        try { $wordDocument.Activate() } catch { }
        try { $presentation.Windows.Item(1).Activate() } catch { }
    }

    Write-Host "[8/10] Testing PowerPoint slide-show and read-only failure paths..."
    $slideShowWindow = $null
    try {
        $slideShowWindow = $presentation.SlideShowSettings.Run()
        Start-Sleep -Milliseconds 300
        $slideShowFailure = Invoke-Bridge "powerpoint.insertFormula" (New-FormulaParams ([Guid]::NewGuid().ToString()) $widePng "blocked" "block") -AllowFailure
        Assert-True ($slideShowFailure.ok -eq $false) "PowerPoint slide-show insertion unexpectedly succeeded."
    } finally {
        if ($null -ne $slideShowWindow) { try { $slideShowWindow.View.Exit() } catch { } }
        Release-ComObject $slideShowWindow
    }

    Write-Host "[8a/10] Saving DOCX and PPTX test artifacts with explicit Open XML formats..."
    [VisualTeXWordAutomation]::SaveAsDocx($wordDocument, $wordPath)
    Write-Host "[8b/10] Word DOCX saved."
    $presentation.SaveAs($powerPointPath, 24)
    Write-Host "[8c/10] PowerPoint PPTX saved."
    $wordDocument.Close(0)
    Release-ComObject $wordDocument
    $wordDocument = $word.Documents.Open($wordPath, $false, $true)
    $readOnlyWordFailure = Invoke-Bridge "word.insertInlineFormula" (New-FormulaParams ([Guid]::NewGuid().ToString()) $widePng "blocked" "inline") -AllowFailure
    Assert-True ($readOnlyWordFailure.ok -eq $false) "Read-only Word insertion unexpectedly succeeded."

    $presentation.Close()
    Release-ComObject $slide
    Release-ComObject $presentation
    $presentation = $powerPoint.Presentations.Open($powerPointPath, -1, 0, -1)
    $slide = $presentation.Slides.Item(1)
    $readOnlyPptFailure = Invoke-Bridge "powerpoint.insertFormula" (New-FormulaParams ([Guid]::NewGuid().ToString()) $widePng "blocked" "block") -AllowFailure
    Assert-True ($readOnlyPptFailure.ok -eq $false) "Read-only PowerPoint insertion unexpectedly succeeded."

    Write-Host "[9/10] Testing Office-not-running and bridge crash/restart..."
    $wordDocument.Close(0)
    $presentation.Close()
    $word.Quit()
    $powerPoint.Quit()
    Release-ComObject $slide
    Release-ComObject $presentation
    Release-ComObject $wordDocument
    Release-ComObject $powerPoint
    Release-ComObject $word
    $slide = $null
    $presentation = $null
    $wordDocument = $null
    $powerPoint = $null
    $word = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    $officeExitDeadline = [DateTimeOffset]::UtcNow.AddSeconds(15)
    do {
        Start-Sleep -Milliseconds 250
        $detectAfter = Invoke-Bridge "office.detect" ([ordered]@{})
        if ($detectAfter.result.wordRunning -eq $false -and
            $detectAfter.result.powerPointRunning -eq $false) {
            break
        }
    } while ([DateTimeOffset]::UtcNow -lt $officeExitDeadline)

    if ($detectAfter.result.wordRunning -or $detectAfter.result.powerPointRunning) {
        $lingeringAcceptanceProcesses = @(
            Get-Process WINWORD, POWERPNT -ErrorAction SilentlyContinue |
                Where-Object { $_.Id -in $acceptanceOfficeProcessIds }
        )
        foreach ($process in $lingeringAcceptanceProcesses) {
            Write-Host "Stopping acceptance-owned Office background process $($process.ProcessName) ($($process.Id))."
            Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
        }
        $forcedExitDeadline = [DateTimeOffset]::UtcNow.AddSeconds(10)
        do {
            Start-Sleep -Milliseconds 250
            $detectAfter = Invoke-Bridge "office.detect" ([ordered]@{})
            if ($detectAfter.result.wordRunning -eq $false -and
                $detectAfter.result.powerPointRunning -eq $false) {
                break
            }
        } while ([DateTimeOffset]::UtcNow -lt $forcedExitDeadline)
    }

    Assert-True ($detectAfter.result.wordRunning -eq $false) "Word remained detectable after Quit and acceptance-owned background cleanup."
    Assert-True ($detectAfter.result.powerPointRunning -eq $false) "PowerPoint remained detectable after Quit and acceptance-owned background cleanup."
    $notRunning = Invoke-Bridge "word.getSelection" ([ordered]@{}) -AllowFailure
    Assert-True ($notRunning.ok -eq $false) "Word selection unexpectedly succeeded while Word was stopped."

    Stop-BridgeProcess $bridgeProcess
    $bridgeProcess = Start-BridgeProcess
    $restartHealth = Invoke-Bridge "health" ([ordered]@{})
    Assert-True ($restartHealth.ok -eq $true) "Bridge did not recover after a forced crash/restart."

    Write-Host "[10/10] Optional OLE/VSTO mode mutual-exclusion test..."
    if ($TestModeSwitch) {
        & (Join-Path $PSScriptRoot "install_windows_ole.ps1")
        $catalogKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\{69C6A866-755B-4C5A-BACB-EEA28B03C724}"
        Assert-True (Test-Path $catalogKey) "OLE Trusted Catalog was not registered."
        foreach ($key in @(
            "HKCU:\Software\Microsoft\Office\Word\Addins\VisualTeX.WordVsto",
            "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\VisualTeX.PowerPointVsto"
        )) {
            if (Test-Path $key) {
                $loadBehavior = (Get-ItemProperty $key -Name LoadBehavior).LoadBehavior
                Assert-True ($loadBehavior -ne 3) "VSTO remained enabled while OLE mode was active."
            }
        }
        if (-not [string]::IsNullOrWhiteSpace($VstoMsiPath)) {
            & (Join-Path $PSScriptRoot "install_windows_vsto.ps1") -MsiPath $VstoMsiPath
            Assert-True (-not (Test-Path $catalogKey)) "OLE Trusted Catalog remained registered while VSTO mode was active."
            foreach ($key in @(
                "HKCU:\Software\Microsoft\Office\Word\Addins\VisualTeX.WordVsto",
                "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\VisualTeX.PowerPointVsto"
            )) {
                Assert-True ((Get-ItemProperty $key -Name LoadBehavior).LoadBehavior -eq 3) "VSTO add-in was not enabled."
            }
        }
    }

    Write-Host "VisualTeX Windows Office acceptance passed."
    Write-Host "Word: inserted $FormulaCount, edited $($wordEditIds.Count), persistent UUID set preserved."
    Write-Host "PowerPoint: inserted $FormulaCount, edited $($pptEditIds.Count), deleted one target, all other UUIDs preserved."
} finally {
    if ($null -ne $bridgeProcess) {
        try { [void](Invoke-Bridge "shutdown" ([ordered]@{})) } catch { }
        Start-Sleep -Milliseconds 200
        try { Stop-BridgeProcess $bridgeProcess } catch { }
    }
    if ($null -ne $presentation) { try { $presentation.Close() } catch { } }
    if ($null -ne $wordDocument) { try { $wordDocument.Close(0) } catch { } }
    if ($null -ne $powerPoint) { try { $powerPoint.Quit() } catch { } }
    if ($null -ne $word) { try { $word.Quit() } catch { } }
    Release-ComObject $slide
    Release-ComObject $presentation
    Release-ComObject $wordDocument
    Release-ComObject $powerPoint
    Release-ComObject $word
    Remove-Item $widePng, $squarePng, $longPng, $inlinePng -Force -ErrorAction SilentlyContinue
    if (-not $KeepDocuments) { Remove-Item $artifactRoot -Recurse -Force -ErrorAction SilentlyContinue }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
