<#
.SYNOPSIS
    Phase 0 compatibility probe for the VisualTeX WPS add-in.

.DESCRIPTION
    Exercises the WPS Writer COM surface the WPS add-in depends on and writes a
    machine-readable capability matrix. Run this on a Windows machine with WPS
    Office installed BEFORE relying on any fallback path.

    The probe never modifies the user's documents: it creates a scratch document,
    runs each check, then closes the document without saving.

.PARAMETER OutputDirectory
    Where the JSON report is written. Defaults to
    %LOCALAPPDATA%\VisualTeX\office\logs.

.PARAMETER IncludeOleInsert
    Also test Equation.DSMT4 AddOLEObject / OLEFormat.Object. These tests can
    block while the OLE server opens its editor window, so they are opt-in.
    Only run them when MathType (or the WPS equation server) can be closed
    interactively.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File wps_compat_probe.ps1

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File wps_compat_probe.ps1 -IncludeOleInsert
#>
[CmdletBinding()]
param(
    [string]$OutputDirectory = '',
    [switch]$IncludeOleInsert,
    [switch]$KeepOpen
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $env:LOCALAPPDATA 'VisualTeX\office\logs'
}
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null

$script:report = [ordered]@{
    probe = 'visualtex-wps-compat'
    generatedAt = (Get-Date).ToString('o')
    wps = [ordered]@{}
    tests = [ordered]@{}
}

function Add-Probe([string]$name, [scriptblock]$body) {
    try {
        $detail = & $body
        $script:report.tests[$name] = [ordered]@{
            ok = $true
            detail = (($detail | Out-String).Trim())
        }
    }
    catch {
        $script:report.tests[$name] = [ordered]@{
            ok = $false
            detail = $_.Exception.Message
        }
    }
    Write-Host ("[{0}] {1}" -f $(if ($script:report.tests[$name].ok) { ' OK ' } else { 'FAIL' }), $name)
}

function Get-WpsApplication {
    $progIds = @('KWPS.Application', 'WPS.Application', 'Word.Application')
    foreach ($progId in $progIds) {
        try {
            $app = [Runtime.InteropServices.Marshal]::GetActiveObject($progId)
            if ($null -ne $app) {
                return [pscustomobject]@{ ProgId = $progId; App = $app; Reused = $true }
            }
        }
        catch { }
    }
    foreach ($progId in @('KWPS.Application', 'WPS.Application')) {
        try {
            $app = New-Object -ComObject $progId
            if ($null -ne $app) {
                return [pscustomobject]@{ ProgId = $progId; App = $app; Reused = $false }
            }
        }
        catch { }
    }
    throw 'WPS Writer COM automation is not available (KWPS.Application). Confirm WPS Office is installed.'
}

function Release-Com($value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value) } catch { }
    }
}

function Get-ErrorCode($errorRecord) {
    $exception = $errorRecord.Exception
    while ($null -ne $exception.InnerException) { $exception = $exception.InnerException }
    if ($exception -is [Runtime.InteropServices.COMException]) {
        return ('0x{0:X8} {1}' -f $exception.ErrorCode, $exception.Message)
    }
    return $exception.Message
}

$wps = $null
$document = $null
try {
    $wps = Get-WpsApplication
    $app = $wps.App
    $script:report.wps.progId = $wps.ProgId
    $script:report.wps.reusedRunningInstance = $wps.Reused
    try { $script:report.wps.name = [string]$app.Name } catch { }
    try { $script:report.wps.version = [string]$app.Version } catch { }
    try { $script:report.wps.path = [string]$app.Path } catch { }
    try { $app.Visible = $true } catch { }

    Add-Probe 'application.identity' {
        "Name=$($app.Name); Version=$($app.Version); Path=$($app.Path); ProgId=$($wps.ProgId)"
    }

    $document = $app.Documents.Add()
    $script:report.wps.documentCount = $app.Documents.Count

    Add-Probe 'range.insertXml' {
        $flatOpc = @'
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml" />
      </Relationships>
    </pkg:xmlData>
  </pkg:part>
  <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">
    <pkg:xmlData>
      <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:body><w:p><w:r><w:t>VisualTeX WPS probe</w:t></w:r></w:p></w:body>
      </w:document>
    </pkg:xmlData>
  </pkg:part>
</pkg:package>
'@
        $range = $document.Range(0, 0)
        try { $range.InsertXML($flatOpc) } finally { Release-Com $range }
        $text = [string]$document.Range(0, [Math]::Min(60, [int]$document.Content.End)).Text
        if ($text -notmatch 'VisualTeX WPS probe') {
            throw "InsertXML returned without raising but the text was not inserted: '$text'"
        }
        'Inserted Flat OPC paragraph successfully.'
    }

    Add-Probe 'range.wordOpenXml' {
        $range = $document.Range(0, [Math]::Min(20, [int]$document.Content.End))
        try {
            $xml = [string]$range.WordOpenXML
            if ([string]::IsNullOrWhiteSpace($xml)) { throw 'WordOpenXML returned an empty string.' }
            "Length=$($xml.Length)"
        }
        finally { Release-Com $range }
    }

    Add-Probe 'document.variables' {
        $name = 'VT_Probe'
        $value = 'visualtex-probe-value'
        try {
            try { $document.Variables.Add($name, $value) }
            catch { $document.Variables[$name].Value = $value }
            $read = [string]$document.Variables[$name].Value
            if ($read -ne $value) { throw "Read back '$read' instead of '$value'." }
            "Persisted and read '$name'."
        }
        finally {
            try { $document.Variables[$name].Delete() } catch { }
        }
    }

    Add-Probe 'fields.seqCreateAndUpdate' {
        $document.Content.InsertAfter("`r")
        $range = $document.Range([Math]::Max(0, [int]$document.Content.End - 1), [Math]::Max(0, [int]$document.Content.End - 1))
        try {
            $field = $document.Fields.Add($range, -1, 'SEQ VisualTeXEquation \* ARABIC', $false)
            try {
                $field.Update()
                $result = [string]$field.Result.Text
                if ([string]::IsNullOrWhiteSpace($result)) { throw 'SEQ field produced an empty result.' }
                "SEQ field result='$result'."
            }
            finally { Release-Com $field }
        }
        finally { Release-Com $range }
    }

    Add-Probe 'event.windowSelectionChange.subscription' {
        $subscription = $null
        try {
            $subscription = Register-ObjectEvent -InputObject $app -EventName WindowSelectionChange -Action { }
            'COM event subscription succeeded.'
        }
        finally {
            if ($null -ne $subscription) {
                try { Unregister-Event -SubscriptionId $subscription.SubscriptionId } catch { }
                try { Get-Job | Remove-Job -Force } catch { }
            }
        }
    }

    Add-Probe 'event.windowBeforeDoubleClick.subscription' {
        $subscription = $null
        try {
            $subscription = Register-ObjectEvent -InputObject $app -EventName WindowBeforeDoubleClick -Action { }
            'COM event subscription succeeded.'
        }
        finally {
            if ($null -ne $subscription) {
                try { Unregister-Event -SubscriptionId $subscription.SubscriptionId } catch { }
                try { Get-Job | Remove-Job -Force } catch { }
            }
        }
    }

    if ($IncludeOleInsert) {
        Add-Probe 'inlineshapes.addOleObject.classType' {
            $shape = $null
            try {
                $range = $document.Range(0, 0)
                try {
                    $shape = $document.InlineShapes.AddOLEObject(
                        'Equation.DSMT4', $null, $false, $false, $null, $null, $null, $range)
                }
                finally { Release-Com $range }
                if ($null -eq $shape) { throw 'AddOLEObject returned no shape.' }
                $progId = ''
                try { $progId = [string]$shape.OLEFormat.ProgID } catch { }
                "Inserted OLE object; OLEFormat.ProgID='$progId'."
            }
            catch {
                throw (Get-ErrorCode $_)
            }
            finally {
                if ($null -ne $shape) {
                    try { $shape.Delete() } catch { }
                    Release-Com $shape
                }
            }
        }

        Add-Probe 'inlineshapes.oleFormat.object' {
            $shape = $null
            try {
                $range = $document.Range(0, 0)
                try {
                    $shape = $document.InlineShapes.AddOLEObject(
                        'Equation.DSMT4', $null, $false, $false, $null, $null, $null, $range)
                }
                finally { Release-Com $range }
                if ($null -eq $shape) { throw 'AddOLEObject returned no shape.' }
                $object = $shape.OLEFormat.Object
                if ($null -eq $object) { throw 'OLEFormat.Object returned null.' }
                'OLEFormat.Object is reachable.'
            }
            finally {
                if ($null -ne $shape) {
                    try { $shape.Delete() } catch { }
                    Release-Com $shape
                }
            }
        }
    }

    Add-Probe 'saveReopen.roundTrip' {
        $path = Join-Path $env:TEMP ("visualtex-wps-probe-{0}.docx" -f ([Guid]::NewGuid().ToString('N')))
        try {
            try { $document.SaveAs2($path) } catch { $document.SaveAs($path) }
            $document.Close(0)
            Release-Com $document
            $document = $app.Documents.Open($path, $false, $false, $false)
            $count = [int]$document.InlineShapes.Count
            $fieldText = ''
            for ($index = 1; $index -le $document.Fields.Count; $index++) {
                $field = $document.Fields[$index]
                try { $fieldText += [string]$field.Code.Text + ';' } finally { Release-Com $field }
            }
            "Reopened; InlineShapes=$count; Fields='$fieldText'."
        }
        finally {
            try { if (Test-Path $path) { Remove-Item $path -Force } } catch { }
        }
    }
}
catch {
    $script:report.fatal = $_.Exception.Message
    Write-Host ("FATAL: {0}" -f $_.Exception.Message) -ForegroundColor Red
}
finally {
    if ($null -ne $document) {
        try { $document.Close(0) } catch { }
        Release-Com $document
    }
    if ($null -ne $wps -and -not $KeepOpen) {
        try { $wps.App.Quit() } catch { }
        Release-Com $wps.App
    }
}

$reportPath = Join-Path $OutputDirectory 'wps-compat-probe.json'
$script:report | ConvertTo-Json -Depth 6 | Set-Content -Path $reportPath -Encoding UTF8
Write-Host ""
Write-Host "Report: $reportPath"
Write-Host "Summary:"
$script:report.tests.GetEnumerator() | ForEach-Object {
    $state = if ($_.Value.ok) { 'PASS' } else { 'FAIL' }
    Write-Host ("  {0}  {1}" -f $state, $_.Key)
}
if ($script:report.Contains('fatal')) {
    exit 1
}
$failed = @($script:report.tests.GetEnumerator() | Where-Object { -not $_.Value.ok })
if ($failed.Count -gt 0) { exit 2 }
exit 0
