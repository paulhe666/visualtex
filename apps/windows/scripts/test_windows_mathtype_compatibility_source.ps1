$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$productionFiles = @(
    'src-windows/VisualTeX.WordVsto/MathTypeOleInterop.cs',
    'src-windows/VisualTeX.WordVsto/MathTypeOleStorage.cs',
    'src-windows/VisualTeX.WordVsto/MathTypeOleClipboardProxy.cs',
    'src-windows/VisualTeX.WordVsto/MathTypeWordCommandsBridge.cs',
    'src-windows/VisualTeX.WindowsOleBridge/MathTypeNativePreviewCommand.cs',
    'src-windows/VisualTeX.WindowsOffice.VstoShared/WordOmmlConverter.cs',
    'src-windows/VisualTeX.WordVsto/WordFormulaService.cs',
    'src-windows/VisualTeX.PowerPointVsto/PowerPointFormulaService.cs'
)

$contents = @{}
foreach ($relativePath in $productionFiles) {
    $path = Join-Path $root $relativePath
    if (-not (Test-Path -LiteralPath $path)) {
        throw "Missing compatibility-audit source file: $relativePath"
    }
    $contents[$relativePath] = Get-Content -LiteralPath $path -Raw
}

$allText = ($contents.Values -join "`n")
$forbiddenAbsolutePatterns = @(
    'C:\\Users\\',
    'C:/Users/',
    '\\pojian_liao\\',
    '/pojian_liao/'
)
foreach ($pattern in $forbiddenAbsolutePatterns) {
    if ($allText -match $pattern) {
        throw "A production Office/MathType source file contains a machine-specific path matching '$pattern'."
    }
}

$interop = $contents['src-windows/VisualTeX.WordVsto/MathTypeOleInterop.cs']
$storage = $contents['src-windows/VisualTeX.WordVsto/MathTypeOleStorage.cs']
$preview = $contents['src-windows/VisualTeX.WindowsOleBridge/MathTypeNativePreviewCommand.cs']
$omml = $contents['src-windows/VisualTeX.WindowsOffice.VstoShared/WordOmmlConverter.cs']
$powerPoint = $contents['src-windows/VisualTeX.PowerPointVsto/PowerPointFormulaService.cs']

$requiredInteropTokens = @(
    'ResolvePreferredStorageIdentity',
    'OpenClassesSubKey',
    'OleGetAutoConvert',
    'LooksLikeMathTypeCompoundFile',
    'MaximumCapabilityCacheEntries',
    'RegistryHive.CurrentUser',
    'RegistryHive.LocalMachine',
    'RegistryView.Registry32',
    'RegistryView.Registry64',
    'RunForConversionVerb',
    'RegisteredMathMlGetSet'
)
foreach ($token in $requiredInteropTokens) {
    if (-not $interop.Contains($token)) {
        throw "MathType interop is missing required compatibility mechanism '$token'."
    }
}

$requiredStorageTokens = @(
    'LooksLikeMathTypeCompoundFile',
    'ResolvePreferredStorageIdentity',
    '64 * 1024 * 1024'
)
foreach ($token in $requiredStorageTokens) {
    if (-not $storage.Contains($token)) {
        throw "MathType storage is missing required compatibility mechanism '$token'."
    }
}

if ($interop -match 'return\s+fragment\.ProgId\.StartsWith\("Equation\.' ) {
    throw 'MathType detection still accepts every Equation.* ProgID without validating the embedded storage.'
}
if ($powerPoint -match 'attempts\s*<\s*512') {
    throw 'PowerPoint Z-order restoration still has the legacy fixed 512-shape ceiling.'
}
if (-not $powerPoint.Contains('CompleteGeneration')) {
    throw 'PowerPoint OLE geometry generation state is not explicitly retired.'
}
if (-not $preview.Contains('MathType')) {
    throw 'MathType native preview bridge source is unexpectedly empty or disconnected.'
}
if ($omml -match 'Office16\\OMML2MML\.XSL' -and -not $omml.Contains('RegistryView')) {
    throw 'OMML converter appears tied to one Office installation path without registry-view probing.'
}

Write-Host 'VisualTeX MathType/Office compatibility source audit passed.'
