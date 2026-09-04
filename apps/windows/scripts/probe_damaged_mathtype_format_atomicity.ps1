param(
    [string]$FlatOpcPath = "src-windows/artifacts/active-document1-mathtype-numbering-inspection/content-flat-opc.raw.xml",
    [string]$AssemblyPath = "src-windows/VisualTeX.WordVsto/bin/x64/Release/net472/VisualTeX.WordVsto.dll",
    [string]$ArtifactRoot = "src-windows/artifacts/document1-damaged-format-atomicity",
    [string]$FormatId = "continuous",
    [ValidateSet("Core", "Service")]
    [string]$Mode = "Core"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

function Release-ComObject([object]$value) {
    if ($null -eq $value -or -not [Runtime.InteropServices.Marshal]::IsComObject($value)) { return }
    try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value) } catch { }
}

function Normalize-FlatOpc([string]$xml) {
    $normalized = [regex]::Replace(
        $xml,
        'w:rsid(?:R|RDefault|P|RPr|Del|Sect)="[0-9A-Fa-f]+"',
        '')
    $normalized = [regex]::Replace(
        $normalized,
        'ObjectID="_\d+"',
        'ObjectID="_VOLATILE"')
    $normalized = [regex]::Replace(
        $normalized,
        '<w:rsids>.*?</w:rsids>',
        '<w:rsids/>',
        [Text.RegularExpressions.RegexOptions]::Singleline)
    return $normalized
}

$flatOpcFile = [IO.Path]::GetFullPath((Join-Path (Get-Location) $FlatOpcPath))
$assemblyFile = [IO.Path]::GetFullPath((Join-Path (Get-Location) $AssemblyPath))
$artifactPath = [IO.Path]::GetFullPath((Join-Path (Get-Location) $ArtifactRoot))
[IO.Directory]::CreateDirectory($artifactPath) | Out-Null
$flatOpc = [IO.File]::ReadAllText($flatOpcFile, [Text.Encoding]::UTF8)

$beforeWordPids = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$beforeMathTypePids = @(Get-Process MathType* -ErrorAction SilentlyContinue | ForEach-Object Id)
$operationMathTypePids = $null
$word = $null
$document = $null
$content = $null
$assembly = $null
try {
    $assembly = [Reflection.Assembly]::LoadFrom($assemblyFile)
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $document = $word.Documents.Add()
    $content = $document.Content
    $content.InsertXML($flatOpc)
    Release-ComObject $content; $content = $null

    $content = $document.Content
    $beforeXml = [string]$content.WordOpenXML
    $beforeSummary = [pscustomobject]@{
        Paragraphs = [int]$document.Paragraphs.Count
        Fields = [int]$document.Fields.Count
        InlineShapes = [int]$document.InlineShapes.Count
        Bookmarks = [int]$document.Bookmarks.Count
        ContentStart = [int]$content.Start
        ContentEnd = [int]$content.End
    }
    [IO.File]::WriteAllText((Join-Path $artifactPath "before.xml"), $beforeXml, [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText((Join-Path $artifactPath "before-summary.json"), ($beforeSummary | ConvertTo-Json), [Text.UTF8Encoding]::new($false))
    Release-ComObject $content; $content = $null

    # Opening an isolated Word instance can load the user's MathType startup
    # template before the VisualTeX operation begins. Remove only processes that
    # appeared during this probe, then establish a zero-dependency operation
    # baseline. Any process relaunched by VisualTeX is a test failure.
    $startupMathTypePids = @(Get-Process MathType* -ErrorAction SilentlyContinue | ForEach-Object Id)
    foreach ($processId in @($startupMathTypePids | Where-Object { $beforeMathTypePids -notcontains $_ })) {
        try { Stop-Process -Id ([int]$processId) -Force } catch { }
    }
    Start-Sleep -Milliseconds 300
    $operationMathTypePids = @(Get-Process MathType* -ErrorAction SilentlyContinue | ForEach-Object Id)

    $caught = $null
    try {
        if ($Mode -eq "Core") {
            $type = $assembly.GetType("VisualTeX.WordVsto.MathTypeEquationNumbering", $true)
            $method = $type.GetMethod("SetEquationNumberFormat", [Reflection.BindingFlags]"Static,NonPublic")
            if ($null -eq $method) { throw "MathTypeEquationNumbering.SetEquationNumberFormat was not found." }
            [void]$method.Invoke($null, @($document, $FormatId))
        } else {
            $serviceType = $assembly.GetType("VisualTeX.WordVsto.WordFormulaService", $true)
            $constructor = $serviceType.GetConstructors(
                [Reflection.BindingFlags]"Instance,Public,NonPublic") |
                Where-Object { $_.GetParameters().Count -eq 1 } |
                Select-Object -First 1
            if ($null -eq $constructor) { throw "WordFormulaService constructor was not found." }
            $wordComObject = $word.PSObject.BaseObject
            $service = $constructor.Invoke([object[]](,$wordComObject))
            $serviceMethod = $serviceType.GetMethod(
                "SetEquationNumberFormat",
                [Reflection.BindingFlags]"Instance,Public,NonPublic")
            if ($null -eq $serviceMethod) { throw "WordFormulaService.SetEquationNumberFormat was not found." }
            [void]$serviceMethod.Invoke($service, @($FormatId))
        }
    } catch {
        $caught = $_.Exception
        while ($null -ne $caught.InnerException) { $caught = $caught.InnerException }
    }
    if ($null -eq $caught) {
        throw "Damaged MathType numbering unexpectedly accepted format '$FormatId'."
    }

    $content = $document.Content
    $afterXml = [string]$content.WordOpenXML
    $afterSummary = [pscustomobject]@{
        Paragraphs = [int]$document.Paragraphs.Count
        Fields = [int]$document.Fields.Count
        InlineShapes = [int]$document.InlineShapes.Count
        Bookmarks = [int]$document.Bookmarks.Count
        ContentStart = [int]$content.Start
        ContentEnd = [int]$content.End
        ExceptionType = $caught.GetType().FullName
        ExceptionMessage = $caught.Message
        ExactXmlEqual = [string]::Equals($beforeXml, $afterXml, [StringComparison]::Ordinal)
        NormalizedXmlEqual = [string]::Equals((Normalize-FlatOpc $beforeXml), (Normalize-FlatOpc $afterXml), [StringComparison]::Ordinal)
    }
    [IO.File]::WriteAllText((Join-Path $artifactPath "after.xml"), $afterXml, [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText((Join-Path $artifactPath "after-summary.json"), ($afterSummary | ConvertTo-Json), [Text.UTF8Encoding]::new($false))

    if ($afterSummary.Paragraphs -ne $beforeSummary.Paragraphs -or
        $afterSummary.Fields -ne $beforeSummary.Fields -or
        $afterSummary.InlineShapes -ne $beforeSummary.InlineShapes -or
        $afterSummary.Bookmarks -ne $beforeSummary.Bookmarks -or
        -not $afterSummary.NormalizedXmlEqual) {
        throw "Damaged-document rejection was not atomic. See $artifactPath."
    }
    Write-Output "DAMAGED_MATHTYPE_REJECTED_ATOMICALLY|mode=$Mode|exception=$($caught.GetType().FullName)|message=$($caught.Message)|artifact=$artifactPath"
} finally {
    Release-ComObject $content
    if ($null -ne $document) { try { $document.Close(0) } catch { } }
    Release-ComObject $document
    if ($null -ne $word) { try { $word.Quit() } catch { } }
    Release-ComObject $word
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 500
    $ownedWord = @(Get-CimInstance Win32_Process | Where-Object {
        $_.Name -eq "WINWORD.EXE" -and
        $beforeWordPids -notcontains [int]$_.ProcessId -and
        $_.CommandLine -like "*/Automation -Embedding*"
    })
    foreach ($process in $ownedWord) {
        try { Stop-Process -Id ([int]$process.ProcessId) -Force } catch { }
    }
    $afterMathTypePids = @(Get-Process MathType* -ErrorAction SilentlyContinue | ForEach-Object Id)
    $baselineMathTypePids = if ($null -ne $operationMathTypePids) {
        @($operationMathTypePids)
    } else {
        @($beforeMathTypePids)
    }
    $newMathTypePids = @($afterMathTypePids | Where-Object { $baselineMathTypePids -notcontains $_ })
    if ($newMathTypePids.Count -gt 0) {
        throw "Damaged-document probe started MathType process(es): $($newMathTypePids -join ',')."
    }
}
