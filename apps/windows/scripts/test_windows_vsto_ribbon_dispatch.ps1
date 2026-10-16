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
            # Every conversion row has its own Selection and Document buttons.
            foreach ($row in @(
                'VisualTeXToMathType',
                'MathTypeToVisualTeX',
                'OmmlToMathType',
                'MathTypeToOmml',
                'VisualTeXToOmml',
                'OmmlToVisualTeX'
            )) {
                foreach ($scope in @('Selection', 'Document')) {
                    $wired = '<button id="VisualTeX.WordVsto.{0}{1}"[^>]*onAction="OnConvert{0}{1}"' -f $row, $scope
                    if ($ribbonXml -notmatch $wired) {
                        throw "$ClassName Ribbon XML has no $scope button for $row wired to OnConvert$row$scope."
                    }
                }
            }
            foreach ($retired in @(
                'OnConvertFormulaFormatDefault',
                'GetFormatConversionScopePressed',
                'OnFormatConversionScopeChanged',
                '<splitButton id="VisualTeX.WordVsto.VisualTeXToMathType"'
            )) {
                if ($ribbonXml -match [regex]::Escape($retired)) {
                    throw "$ClassName Ribbon XML still exposes the retired default-scope control $retired."
                }
            }

            $formulaServiceType = $assembly.GetType(
                "VisualTeX.WordVsto.WordFormulaService",
                $true)
            $pairValidator = $formulaServiceType.GetMethod(
                "ValidateCoreConversionPair",
                [Reflection.BindingFlags]::Static -bor [Reflection.BindingFlags]::NonPublic)
            if ($null -eq $pairValidator) {
                throw "$ClassName is missing the Host Core format conversion pair validator."
            }
            foreach ($pair in @(
                @("nativeOle", "mathTypeOle"),
                @("mathTypeOle", "nativeOle"),
                @("nativeOle", "wordOmml"),
                @("wordOmml", "nativeOle"),
                @("wordOmml", "mathTypeOle"),
                @("mathTypeOle", "wordOmml")
            )) {
                try {
                    [void]$pairValidator.Invoke($null, @($pair[0], $pair[1]))
                } catch {
                    throw "$ClassName rejects supported format conversion $($pair[0]) -> $($pair[1]): $($_.Exception.InnerException.Message)"
                }
            }
        }

        # Both language variants: every callback the XML names must be a
        # DispId member of the callback interface, ids must be unique, and the
        # two languages must expose the same controls.
        $languageIds = @{}
        foreach ($constName in @('RibbonXml', 'RibbonXmlEnglish')) {
            $constField = $classType.GetField($constName, [Reflection.BindingFlags]'Static,NonPublic,Public')
            if ($null -eq $constField) {
                throw "$ClassName is missing its $constName Ribbon definition."
            }
            $parsed = New-Object Xml.XmlDocument
            $parsed.LoadXml([string]$constField.GetRawConstantValue())
            foreach ($node in @($parsed.SelectNodes('//*'))) {
                foreach ($attribute in @('onLoad', 'onAction', 'getImage', 'getPressed', 'getEnabled', 'getText', 'onChange', 'getLabel')) {
                    $callback = $node.GetAttribute($attribute)
                    if ([string]::IsNullOrEmpty($callback)) { continue }
                    $method = $callbackInterface.GetMethod($callback)
                    if ($null -eq $method -or @($method.GetCustomAttributes(
                            [Runtime.InteropServices.DispIdAttribute], $false)).Count -ne 1) {
                        throw "$ClassName $constName names Ribbon callback $callback, which $InterfaceName does not expose with a DispId."
                    }
                }
            }
            $ids = @($parsed.SelectNodes('//*[@id]') | ForEach-Object { $_.GetAttribute('id') })
            $duplicate = @($ids | Group-Object | Where-Object { $_.Count -gt 1 }) | Select-Object -First 1
            if ($duplicate) {
                throw "$ClassName $constName repeats Ribbon id $($duplicate.Name)."
            }
            $languageIds[$constName] = ($ids | Sort-Object) -join '|'

            if ($ClassName -eq 'VisualTeX.PowerPointVsto.ThisAddIn') {
                $requestType = $assembly.GetType('VisualTeX.PowerPointVsto.PowerPointConversionRequest', $true)
                $tryParse = $requestType.GetMethod('TryParse', [Reflection.BindingFlags]'Static,NonPublic,Public')
                $conversionButtons = @($parsed.SelectNodes('//*[@onAction="OnConvertFormulaFormat"]'))
                if ($conversionButtons.Count -ne 24) {
                    throw "$ClassName $constName exposes $($conversionButtons.Count) format conversion buttons; expected 18 directional and 6 unify buttons."
                }
                foreach ($button in $conversionButtons) {
                    $parseArguments = @($button.GetAttribute('tag'), 'nativeOle', $null)
                    if (-not $tryParse.Invoke($null, $parseArguments)) {
                        throw "$ClassName $constName conversion button $($button.GetAttribute('id')) has an invalid tag '$($button.GetAttribute('tag'))'."
                    }
                }
                foreach ($row in @('VisualTeXToMathType', 'MathTypeToVisualTeX', 'OmmlToMathType', 'MathTypeToOmml', 'VisualTeXToOmml', 'OmmlToVisualTeX')) {
                    foreach ($scope in @('Selection', 'Slide', 'All')) {
                        if ($null -eq $parsed.SelectSingleNode("//*[@id='VisualTeX.PowerPointVsto.$row$scope']")) {
                            throw "$ClassName $constName has no $scope button for $row."
                        }
                    }
                }
                # Retired single-object buttons: the conversion group, the editor's
                # object-format choice and PowerPoint's Delete key replace them.
                foreach ($retired in @('ConvertOmml', 'ConvertSelected', 'ExportPicture', 'Delete', 'OpenDesktop')) {
                    if ($null -ne $parsed.SelectSingleNode("//*[@id='VisualTeX.PowerPointVsto.$retired']")) {
                        throw "$ClassName $constName still exposes the retired $retired button."
                    }
                }
            }
        }
        if ($languageIds['RibbonXml'] -ne $languageIds['RibbonXmlEnglish']) {
            throw "$ClassName Chinese and English Ribbon definitions expose different controls."
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
        "OnOpenDesktop",
        "OnNewSvgFormula",
        "OnNewOmmlFormula",
        "OnNewOleFormula",
        "OnInsertInlineOmml",
        "OnInsertDisplayOmml",
        "OnConvertFormulaFormat",
        "GetVisualTeXTargetFormPressed",
        "OnVisualTeXTargetFormChanged",
        "GetVisualTeXTargetFormLabel"
    )
