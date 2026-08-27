[CmdletBinding()]
param(
    [string]$Configuration = "Release",
    [ValidateSet("x86", "x64")]
    [string]$Platform = "x64"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot

function Invoke-MatchingPowerShell {
    $requires32Bit = $Platform -eq "x86"
    if ($requires32Bit -eq (-not [Environment]::Is64BitProcess)) {
        return $false
    }

    $hostPath = if ($requires32Bit) {
        Join-Path $env:WINDIR "SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    } else {
        Join-Path $env:WINDIR "Sysnative\WindowsPowerShell\v1.0\powershell.exe"
    }
    if (-not (Test-Path $hostPath)) {
        throw "Matching Windows PowerShell host is missing: $hostPath"
    }

    & $hostPath `
        -NoProfile `
        -ExecutionPolicy Bypass `
        -File $PSCommandPath `
        -Configuration $Configuration `
        -Platform $Platform |
        ForEach-Object { Write-Host $_ }
    if ($LASTEXITCODE -ne 0) {
        throw "VSTO Ribbon COM dispatch verification failed in the $Platform PowerShell host."
    }
    return $true
}

if (Invoke-MatchingPowerShell) {
    return
}

function Test-RibbonComDispatch(
    [string]$AssemblyPath,
    [string]$ClassName,
    [string]$InterfaceName,
    [Guid]$InterfaceId,
    [string]$ExpectedTabId,
    [string[]]$Methods
) {
    if (-not (Test-Path $AssemblyPath)) {
        throw "VSTO assembly is missing: $AssemblyPath"
    }

    $resolvedAssemblyPath = (Resolve-Path $AssemblyPath).Path
    $assemblyDirectory = Split-Path -Parent $resolvedAssemblyPath
    $resolver = [ResolveEventHandler]{
        param($sender, $eventArgs)
        $dependencyName = ([Reflection.AssemblyName]$eventArgs.Name).Name + ".dll"
        $candidate = Join-Path $assemblyDirectory $dependencyName
        if (Test-Path $candidate) {
            return [Reflection.Assembly]::LoadFrom($candidate)
        }
        return $null
    }

    [AppDomain]::CurrentDomain.add_AssemblyResolve($resolver)
    try {
        $assembly = [Reflection.Assembly]::LoadFrom($resolvedAssemblyPath)
        $classType = $assembly.GetType($ClassName, $true)
        $callbackInterface = $assembly.GetType($InterfaceName, $true)

        $defaultInterface = @($classType.GetCustomAttributes(
            [Runtime.InteropServices.ComDefaultInterfaceAttribute],
            $false))
        if ($defaultInterface.Count -ne 1 -or $defaultInterface[0].Value -ne $callbackInterface) {
            throw "$ClassName does not expose $InterfaceName as its default COM interface."
        }

        $interfaceType = @($callbackInterface.GetCustomAttributes(
            [Runtime.InteropServices.InterfaceTypeAttribute],
            $false))
        if ($interfaceType.Count -ne 1 -or
            $interfaceType[0].Value -ne [Runtime.InteropServices.ComInterfaceType]::InterfaceIsIDispatch) {
            throw "$InterfaceName is not an IDispatch callback interface."
        }

        foreach ($methodName in $Methods) {
            $method = $callbackInterface.GetMethod($methodName)
            if ($null -eq $method) {
                throw "$InterfaceName is missing Ribbon callback $methodName."
            }
            if (@($method.GetCustomAttributes(
                    [Runtime.InteropServices.DispIdAttribute],
                    $false)).Count -ne 1) {
                throw "$InterfaceName.$methodName is missing a DispId."
            }
        }

        $instance = [Activator]::CreateInstance($classType)
        $ribbonXml = [string]$classType.GetMethod("GetCustomUI").Invoke(
            $instance,
            @("Microsoft.Office.Document"))
        $expectedTabMarker = 'id="{0}"' -f $ExpectedTabId
        if ($ribbonXml -notmatch [regex]::Escape($expectedTabMarker)) {
            throw "$ClassName Ribbon XML is missing the independent $ExpectedTabId tab."
        }
        if ($ribbonXml -match '<tab\s+idMso="TabHome"') {
            throw "$ClassName still injects VisualTeX directly into the built-in Home tab."
        }
        $isWordAddIn = $ClassName -eq "VisualTeX.WordVsto.ThisAddIn"
        $missingBulkImport = $ribbonXml -notmatch 'OnBulkImport' -or
            $ribbonXml -notmatch 'tag="batchImport"'
        if ($isWordAddIn -and $missingBulkImport) {
            throw "$ClassName Ribbon XML is missing the bulk import callback or high-DPI icon tag."
        }
        if ($isWordAddIn) {
            if ($ribbonXml -match 'VisualTeX.WordVsto.ExportPicture' -or
                $ribbonXml -match 'OnExportSelectedAsPicture') {
                throw "$ClassName Ribbon XML still exposes the removed picture export command."
            }
            foreach ($requiredFormatControl in @(
                'VisualTeX.WordVsto.VisualTeXToMathType',
                'VisualTeX.WordVsto.MathTypeToVisualTeX',
                'VisualTeX.WordVsto.OmmlToMathType',
                'VisualTeX.WordVsto.MathTypeToOmml',
                'VisualTeX.WordVsto.VisualTeXToOmml',
                'VisualTeX.WordVsto.OmmlToVisualTeX'
            )) {
                $controlMarker = 'id="{0}"' -f $requiredFormatControl
                if ($ribbonXml -notmatch [regex]::Escape($controlMarker)) {
                    throw "$ClassName Ribbon XML is missing format conversion control $requiredFormatControl."
                }
            }

            $formulaServiceType = $assembly.GetType(
                "VisualTeX.WordVsto.WordFormulaService",
                $true)
            $pairValidator = $formulaServiceType.GetMethod(
                "ValidateSimpleFormatConversionPair",
                [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic)
            if ($null -eq $pairValidator) {
                throw "$ClassName is missing the format conversion pair validator."
            }
            foreach ($pair in @(
                @("nativeOle", "wordOmml"),
                @("wordOmml", "nativeOle")
            )) {
                try {
                    [void]$pairValidator.Invoke($null, @($pair[0], $pair[1]))
                } catch {
                    throw "$ClassName rejects supported format conversion $($pair[0]) -> $($pair[1]): $($_.Exception.InnerException.Message)"
                }
            }
        }

        $unknown = [Runtime.InteropServices.Marshal]::GetIUnknownForObject($instance)
        $callbackPointer = [IntPtr]::Zero
        try {
            $iid = $InterfaceId
            $hresult = [Runtime.InteropServices.Marshal]::QueryInterface(
                $unknown,
                [ref]$iid,
                [ref]$callbackPointer)
            if ($hresult -ne 0 -or $callbackPointer -eq [IntPtr]::Zero) {
                throw ("QueryInterface failed for {0}: 0x{1:X8}" -f $InterfaceName, $hresult)
            }
        } finally {
            if ($callbackPointer -ne [IntPtr]::Zero) {
                [void][Runtime.InteropServices.Marshal]::Release($callbackPointer)
            }
            [void][Runtime.InteropServices.Marshal]::Release($unknown)
        }

        Write-Host "$ClassName independent Ribbon tab and COM callbacks passed ($Platform)."
    } finally {
        [AppDomain]::CurrentDomain.remove_AssemblyResolve($resolver)
    }
}

$wordAssembly = Join-Path $root "src-windows\VisualTeX.WordVsto\bin\$Platform\$Configuration\net472\VisualTeX.WordVsto.dll"
$powerPointAssembly = Join-Path $root "src-windows\VisualTeX.PowerPointVsto\bin\$Platform\$Configuration\net472\VisualTeX.PowerPointVsto.dll"

Test-RibbonComDispatch `
    $wordAssembly `
    "VisualTeX.WordVsto.ThisAddIn" `
    "VisualTeX.WordVsto.IWordRibbonCallbacks" `
    ([Guid]"D4A1A3CB-0ED7-4B2F-8A2B-5CB0B1E25421") `
    "VisualTeX.WordVsto.Tab" `
    @(
        "OnRibbonLoad",
        "OnInsertInline",
        "OnInsertDisplay",
        "OnEditSelected",
        "OnConvertSelected",
        "OnUpdateEquationNumbers",
        "GetEquationNumberFormatPressed",
        "OnEquationNumberFormatChanged",
        "OnDeleteSelected",
        "OnOpenDesktop",
        "OnInsertEquationReference",
        "OnBulkImport",
        "OnRedrawSelectionOleToLatex",
        "OnRedrawSelectionOmmlToLatex",
        "OnRedrawDocumentOleToLatex",
        "OnRedrawDocumentOmmlToLatex",
        "OnConvertVisualTeXToMathTypeSelection",
        "OnConvertVisualTeXToMathTypeDocument",
        "OnConvertMathTypeToVisualTeXSelection",
        "OnConvertMathTypeToVisualTeXDocument",
        "OnConvertOmmlToMathTypeSelection",
        "OnConvertOmmlToMathTypeDocument",
        "OnConvertMathTypeToOmmlSelection",
        "OnConvertMathTypeToOmmlDocument",
        "OnConvertVisualTeXToOmmlSelection",
        "OnConvertVisualTeXToOmmlDocument",
        "OnConvertOmmlToVisualTeXSelection",
        "OnConvertOmmlToVisualTeXDocument"
    )

Test-RibbonComDispatch `
    $powerPointAssembly `
    "VisualTeX.PowerPointVsto.ThisAddIn" `
    "VisualTeX.PowerPointVsto.IPowerPointRibbonCallbacks" `
    ([Guid]"29C64025-AB17-4F25-9B89-6E1D8D22C2D7") `
    "VisualTeX.PowerPointVsto.Tab" `
    @(
        "OnRibbonLoad",
        "OnNewFormula",
        "OnEditSelected",
        "OnConvertSelected",
        "OnExportSelectedAsPicture",
        "OnDeleteSelected",
        "OnOpenDesktop"
    )
