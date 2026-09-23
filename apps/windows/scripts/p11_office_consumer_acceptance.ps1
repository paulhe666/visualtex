[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

Add-Type -TypeDefinition @'
using System;
using System.Reflection;
public static class VisualTeXP11ReflectionBridge
{
    public static object InvokeStatic(MethodInfo method, object[] arguments)
    {
        for (var index = 0; index < arguments.Length; index++)
        {
            var value = arguments[index];
            while (value != null && value.GetType().FullName == "System.Management.Automation.PSObject")
            {
                var property = value.GetType().GetProperty("BaseObject");
                value = property == null ? value : property.GetValue(value, null);
                if (property == null) break;
            }
            arguments[index] = value;
        }
        return method.Invoke(null, arguments);
    }
}
'@

$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$tsx = Join-Path $root "node_modules\.bin\tsx.cmd"
$fixtureScript = Join-Path $root "scripts\p11_emit_office_consumer_fixture.mts"
if (-not (Test-Path -LiteralPath $tsx -PathType Leaf)) { throw "tsx is missing: $tsx" }
if (-not (Test-Path -LiteralPath $fixtureScript -PathType Leaf)) { throw "P11 fixture generator is missing." }

$fixtureText = (& $tsx $fixtureScript) -join "`n"
if ($LASTEXITCODE -ne 0) { throw "P11 core fixture generator failed with exit code $LASTEXITCODE" }
$fixture = $fixtureText | ConvertFrom-Json
if ($fixture.schema -ne "visualtex-p11-office-consumer-fixture-v1") {
    throw "Unexpected P11 fixture schema."
}

$vstoOutput = Join-Path $root "src-windows\VisualTeX.WordVsto\bin\x64\Release\net472"
$wordVstoPath = Join-Path $vstoOutput "VisualTeX.WordVsto.dll"
if (-not (Test-Path -LiteralPath $wordVstoPath -PathType Leaf)) {
    throw "Current x64 Word VSTO output is missing: $wordVstoPath"
}
foreach ($dependency in @(
    "VisualTeX.WindowsOffice.Contracts.dll",
    "System.Memory.dll",
    "System.Buffers.dll",
    "System.Numerics.Vectors.dll",
    "System.Runtime.CompilerServices.Unsafe.dll",
    "System.ValueTuple.dll",
    "Microsoft.Bcl.AsyncInterfaces.dll",
    "System.Threading.Tasks.Extensions.dll",
    "System.Text.Encodings.Web.dll",
    "System.Text.Json.dll",
    "Microsoft.Office.Interop.Word.dll"
)) {
    $path = Join-Path $vstoOutput $dependency
    if (Test-Path -LiteralPath $path -PathType Leaf) {
        [void][Reflection.Assembly]::LoadFrom($path)
    }
}
$resolver = [ResolveEventHandler]{
    param($sender, $eventArgs)
    try {
        $simpleName = ([Reflection.AssemblyName]$eventArgs.Name).Name
        $candidate = Join-Path $vstoOutput ($simpleName + ".dll")
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return [Reflection.Assembly]::LoadFrom($candidate)
        }
    } catch {}
    return $null
}
[AppDomain]::CurrentDomain.add_AssemblyResolve($resolver)

$tempRoot = Join-Path $env:TEMP ("visualtex-p11-consumer-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
try {
    $wordAssembly = [Reflection.Assembly]::LoadFrom($wordVstoPath)

    $codecType = $wordAssembly.GetType("VisualTeX.WordVsto.MathTypeMtefCodec", $true)
    $flags = [Reflection.BindingFlags]"Static,NonPublic"
    $createNative = $codecType.GetMethod("CreateEquationNativeAtFontSize", $flags)
    $readNative = $codecType.GetMethod("ReadEquationNativeMathMl", $flags)
    $semanticSignature = $codecType.GetMethod("SemanticSignature", $flags)
    $readFullSize = $codecType.GetMethod("ReadEquationNativeFullFontSize", $flags)
    foreach ($method in @($createNative, $readNative, $semanticSignature, $readFullSize)) {
        if ($null -eq $method) { throw "Current MathTypeMtefCodec consumer method is missing." }
    }

    $inputMathMl = [string]$fixture.mathType.mathMl
    $createArgs = [object[]]::new(3)
    $createArgs[0] = $inputMathMl
    $createArgs[1] = [bool]$false
    $createArgs[2] = [double]$fixture.mathType.fontSizePt
    try { $nativeResult = $createNative.Invoke($null, $createArgs) }
    catch [Reflection.TargetInvocationException] { if ($null -ne $_.Exception.InnerException) { throw $_.Exception.InnerException }; throw }
    $resultType = $nativeResult.GetType()
    $equationNative = [byte[]]$resultType.GetProperty("EquationNative").GetValue($nativeResult)
    $mtef = [byte[]]$resultType.GetProperty("Mtef").GetValue($nativeResult)
    if ($equationNative.Length -le 28 -or $mtef.Length -le 0) {
        throw "MathType consumer returned an empty native equation."
    }
    $singleBytesArg = [object[]]::new(1); $singleBytesArg[0] = $equationNative
    try { $roundTripMathMl = [string]$readNative.Invoke($null, $singleBytesArg) }
    catch [Reflection.TargetInvocationException] { if ($null -ne $_.Exception.InnerException) { throw $_.Exception.InnerException }; throw }
    $singleStringArg = [object[]]::new(1); $singleStringArg[0] = $inputMathMl
    $sourceSignature = [string]$semanticSignature.Invoke($null, $singleStringArg)
    $singleStringArg[0] = $roundTripMathMl
    $roundTripSignature = [string]$semanticSignature.Invoke($null, $singleStringArg)
    if ($sourceSignature -ne $roundTripSignature) {
        throw "Core Presentation MathML changed semantic signature through the existing MathType MTEF consumer."
    }
    $rulerMatch = [regex]::Match($roundTripMathMl, 'data-visualtex-mtef-ruler-stops="([0-9,]+)"')
    if (-not $rulerMatch.Success) {
        throw "MathType MTEF round-trip lost the core ruler stops."
    }
    if ($rulerMatch.Groups[1].Value -ne [string]$fixture.mathType.rulerStops) {
        throw "MathType MTEF ruler stops changed: expected $($fixture.mathType.rulerStops), got $($rulerMatch.Groups[1].Value)"
    }
    if ($roundTripMathMl -notmatch 'data-mtef-tabs="true"') {
        throw "MathType MTEF round-trip lost aligned tab semantics."
    }
    $fullSize = [double]$readFullSize.Invoke($null, $singleBytesArg)
    if ([Math]::Abs($fullSize - [double]$fixture.mathType.fontSizePt) -gt 0.000001) {
        throw "MathType consumer changed semantic font size: $fullSize"
    }

    $svgPath = Join-Path $tempRoot "core-office.svg"
    [IO.File]::WriteAllText($svgPath, [string]$fixture.vector.svg, [Text.UTF8Encoding]::new($false))
    $previewType = $wordAssembly.GetType("VisualTeX.WindowsOffice.VstoShared.OfficeOlePreview", $true)
    $previewMethod = $previewType.GetMethod(
        "CreateInkSafeVectorPreviewFromSvg",
        [Reflection.BindingFlags]"Static,NonPublic")
    if ($null -eq $previewMethod) {
        throw "Current OfficeOlePreview vector consumer is missing."
    }
    $nullableBaseline = [Nullable[Single]]([single]$fixture.vector.baseline)
    $previewArgs = [object[]]::new(5)
    $previewArgs[0] = $svgPath
    $previewArgs[1] = [single]$fixture.vector.width
    $previewArgs[2] = [single]$fixture.vector.height
    $previewArgs[3] = $nullableBaseline
    $previewArgs[4] = [single]1.0
    try { $preview = [VisualTeXP11ReflectionBridge]::InvokeStatic($previewMethod, $previewArgs) }
    catch [Reflection.TargetInvocationException] { if ($null -ne $_.Exception.InnerException) { throw $_.Exception.InnerException }; throw }
    $previewTypeResult = $preview.GetType()
    $instanceFlags = [Reflection.BindingFlags]"Instance,NonPublic"
    $emfPath = [string]$previewTypeResult.GetProperty("EmfPath", $instanceFlags).GetValue($preview)
    $widthPixels = [single]$previewTypeResult.GetProperty("WidthPixels", $instanceFlags).GetValue($preview)
    $heightPixels = [single]$previewTypeResult.GetProperty("HeightPixels", $instanceFlags).GetValue($preview)
    $baselinePixels = [single]$previewTypeResult.GetProperty("BaselinePixels", $instanceFlags).GetValue($preview)
    if (-not (Test-Path -LiteralPath $emfPath -PathType Leaf) -or (Get-Item -LiteralPath $emfPath).Length -le 0) {
        throw "Existing Office SVG consumer did not produce a vector EMF."
    }
    # CreateInkSafeVectorPreviewFromSvg itself calls ValidateVectorEmf, which rejects
    # raster blits/images. For path-only core SVG, no geometry expansion is expected.
    if ([Math]::Abs($widthPixels - [single]$fixture.vector.width) -gt 0.01 -or
        [Math]::Abs($heightPixels - [single]$fixture.vector.height) -gt 0.01 -or
        [Math]::Abs($baselinePixels - [single]$fixture.vector.baseline) -gt 0.01) {
        throw "Existing Office vector consumer changed path-only core natural geometry."
    }
    $emfBytes = (Get-Item -LiteralPath $emfPath).Length
    $emfSha256 = (Get-FileHash -LiteralPath $emfPath -Algorithm SHA256).Hash

    [pscustomobject]@{
        schema = "visualtex-p11-office-consumer-acceptance-v1"
        status = "PASS"
        wordAutomation = 0
        officeSessionCalls = 0
        sourceWordVsto = $wordVstoPath
        sourceWordVstoSha256 = (Get-FileHash -LiteralPath $wordVstoPath -Algorithm SHA256).Hash
        mathTypeSemanticRoundTrip = $true
        mathTypeRulerStops = $rulerMatch.Groups[1].Value
        mathTypeFullSizePt = $fullSize
        equationNativeBytes = $equationNative.Length
        mtefBytes = $mtef.Length
        vectorEmfValidated = $true
        vectorEmfBytes = $emfBytes
        vectorEmfSha256 = $emfSha256
        vectorWidthPx = [double]$widthPixels
        vectorHeightPx = [double]$heightPixels
        vectorBaselinePx = [double]$baselinePixels
    } | ConvertTo-Json -Depth 4
} finally {
    [AppDomain]::CurrentDomain.remove_AssemblyResolve($resolver)
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}
