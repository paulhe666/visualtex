[CmdletBinding()]
param(
    [switch]$EnableBackgroundStart
)

$ErrorActionPreference = "Stop"

function ConvertFrom-VisualTeXExtendedPath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $Path }
    if ($Path.StartsWith("\\?\UNC\", [StringComparison]::OrdinalIgnoreCase)) {
        return "\\" + $Path.Substring(8)
    }
    if ($Path.StartsWith("\\?\", [StringComparison]::OrdinalIgnoreCase)) {
        return $Path.Substring(4)
    }
    return $Path
}

$scriptRoot = ConvertFrom-VisualTeXExtendedPath $PSScriptRoot

function Test-VisualTeXOnlyRibbonCache(
    [IO.FileInfo]$File,
    [string[]]$ManifestIds
) {
    try {
        $ascii = [Text.Encoding]::ASCII.GetString(
            [IO.File]::ReadAllBytes($File.FullName))
        $records = $ascii -split [char]0x1e
        $header = [regex]::Match($records[0], '^0:(\d+):')
        if (-not $header.Success) { return $false }

        $presentVisualTeXIds = @($ManifestIds | Where-Object {
            $ascii.IndexOf($_, [StringComparison]::OrdinalIgnoreCase) -ge 0
        })
        if ($presentVisualTeXIds.Count -eq 0) { return $false }

        # RibbonCache is host-wide. Delete it only when its declared add-in
        # count equals the number of VisualTeX entries in the file. This keeps
        # unrelated Office add-ins intact when they share the same cache.
        $declaredAddinCount = [int]$header.Groups[1].Value
        return $declaredAddinCount -eq $presentVisualTeXIds.Count
    } catch {
        return $false
    }
}

function Clear-VisualTeXWefCache([string[]]$ManifestIds) {
    $wefRoot = Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\Wef"
    if (-not (Test-Path $wefRoot)) { return }
    $markers = @($ManifestIds) + @("VisualTeX.WindowsOle", "VisualTeX OLE", "127.0.0.1:43127")
    foreach ($file in Get-ChildItem $wefRoot -File -Recurse -ErrorAction SilentlyContinue) {
        $isVisualTeX = $false
        $relativePath = $file.FullName.Substring($wefRoot.Length).TrimStart('\')
        foreach ($marker in $markers) {
            if ($relativePath -match [regex]::Escape($marker)) { $isVisualTeX = $true; break }
        }
        # Host-wide RibbonCache files can contain several unrelated add-ins.
        # Never remove one just because its contents mention VisualTeX.
        if (-not $isVisualTeX -and
            $relativePath -match '^AppCommands\\11\.0\\(?:Word|PowerPoint)\.RibbonCache\.') {
            $containsVisualTeX = $false
            try {
                $ascii = [Text.Encoding]::ASCII.GetString(
                    [IO.File]::ReadAllBytes($file.FullName))
                foreach ($manifestId in $ManifestIds) {
                    if ($ascii.IndexOf($manifestId, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                        $containsVisualTeX = $true
                        break
                    }
                }
            } catch { }

            if ($containsVisualTeX) {
                if (Test-VisualTeXOnlyRibbonCache $file $ManifestIds) {
                    $isVisualTeX = $true
                } else {
                    Write-Warning "Preserved shared Office Ribbon cache containing non-VisualTeX add-ins: $($file.FullName)"
                }
            }
        }
        if ($isVisualTeX) {
            Remove-Item $file.FullName -Force -ErrorAction Stop
            Write-Host "Removed VisualTeX Office.js cache file: $($file.FullName)"
        }
    }
}

function Clear-VisualTeXCustomUiCache([string[]]$ManifestIds) {
    $cacheKey = "HKCU:\Software\Microsoft\Office\16.0\Common\CustomUIValidationCache"
    if (-not (Test-Path $cacheKey)) { return }
    $properties = (Get-ItemProperty $cacheKey).PSObject.Properties |
        Where-Object { $_.Name -notmatch '^PS(Path|ParentPath|ChildName|Drive|Provider)$' }
    foreach ($property in $properties) {
        foreach ($manifestId in $ManifestIds) {
            if ($property.Name.StartsWith($manifestId, [StringComparison]::OrdinalIgnoreCase)) {
                Remove-ItemProperty $cacheKey -Name $property.Name -Force
                Write-Host "Removed VisualTeX Office CustomUI cache value: $($property.Name)"
                break
            }
        }
    }
}

function Read-AndValidateManifest([string]$Path, [string]$ExpectedHost, [string]$ExpectedId) {
    try {
        $utf8 = [Text.UTF8Encoding]::new($false, $true)
        [xml]$manifest = [IO.File]::ReadAllText($Path, $utf8)
    }
    catch { throw "Office.js manifest is not valid UTF-8 XML: $Path. $($_.Exception.Message)" }
    $manager = [Xml.XmlNamespaceManager]::new($manifest.NameTable)
    $manager.AddNamespace("o", "http://schemas.microsoft.com/office/appforoffice/1.1")
    $id = $manifest.SelectSingleNode("/o:OfficeApp/o:Id", $manager).InnerText
    $manifestHost = $manifest.SelectSingleNode("/o:OfficeApp/o:Hosts/o:Host", $manager).GetAttribute("Name")
    $source = $manifest.SelectSingleNode("/o:OfficeApp/o:DefaultSettings/o:SourceLocation", $manager).GetAttribute("DefaultValue")
    if ($id -ne $ExpectedId) { throw "Unexpected manifest GUID in ${Path}: $id" }
    if ($manifestHost -ne $ExpectedHost) { throw "Unexpected manifest host in ${Path}: $manifestHost" }
    if ($source -ne "https://127.0.0.1:43127/bridge/index.html") { throw "Unexpected SourceLocation in ${Path}: $source" }
    $requiredCommands = @("Commands.Url", "newFormula", "editFormula", "openDesktop")
    if ($ExpectedHost -eq "Document") { $requiredCommands += "updateEquationNumbers" }
    foreach ($required in $requiredCommands) {
        if ($manifest.OuterXml -notmatch [regex]::Escape($required)) { throw "Manifest is missing required command/resource '$required': $Path" }
    }
    return $id
}

function Assert-HttpsResource([string]$Path) {
    $uri = "https://127.0.0.1:43127$Path"
    $response = Invoke-WebRequest $uri -UseBasicParsing -TimeoutSec 10
    if ($response.StatusCode -ne 200) { throw "Office resource returned HTTP $($response.StatusCode): $uri" }
    Write-Host "Schannel HTTP 200: $uri"
}

function Invoke-ElevatedEncodedPowerShell([string]$Script) {
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($Script))
    $process = Start-Process powershell.exe -Verb RunAs -WindowStyle Hidden -Wait -PassThru `
        -ArgumentList "-NoProfile", "-EncodedCommand", $encoded
    if ($process.ExitCode -ne 0) {
        throw "Elevated VisualTeX Office setup failed with exit code $($process.ExitCode)."
    }
}

function Ensure-VisualTeXCatalogShare([string]$Name, [string]$Path) {
    $account = "$env:USERDOMAIN\$env:USERNAME"
    $existing = Get-SmbShare -Name $Name -ErrorAction SilentlyContinue
    if ($existing -and $existing.Path -eq $Path -and (Test-Path "\\127.0.0.1\$Name")) {
        return "\\127.0.0.1\$Name"
    }
    $escapedName = $Name.Replace("'", "''")
    $escapedPath = $Path.Replace("'", "''")
    $escapedAccount = $account.Replace("'", "''")
    $script = @"
`$ErrorActionPreference = 'Stop'
`$old = Get-SmbShare -Name '$escapedName' -ErrorAction SilentlyContinue
if (`$old) { Remove-SmbShare -Name '$escapedName' -Force }
New-SmbShare -Name '$escapedName' -Path '$escapedPath' -ReadAccess '$escapedAccount' | Out-Null
"@
    Invoke-ElevatedEncodedPowerShell $script
    $unc = "\\127.0.0.1\$Name"
    if (-not (Test-Path $unc)) { throw "VisualTeX Office catalog share is not readable: $unc" }
    return $unc
}

function Ensure-OfficeWebViewLoopbackExemption {
    $sid = "S-1-15-2-1310292540-1029022339-4008023048-2190398717-53961996-4257829345-603366646"
    $existing = (& CheckNetIsolation.exe LoopbackExempt -s 2>&1 | Out-String)
    if ($existing -match [regex]::Escape($sid)) { return }
    Invoke-ElevatedEncodedPowerShell "& CheckNetIsolation.exe LoopbackExempt -a -p=$sid | Out-Null; if (`$LASTEXITCODE -ne 0) { exit `$LASTEXITCODE }"
    $verified = (& CheckNetIsolation.exe LoopbackExempt -s 2>&1 | Out-String)
    if ($verified -notmatch [regex]::Escape($sid)) {
        throw "Microsoft Edge WebView host loopback exemption was not installed."
    }
}

function Test-TrustedCatalogCommandCache([string]$ManifestId) {
    $root = Join-Path $env:LOCALAPPDATA "Microsoft\Office\16.0\Wef\AppCommands\11.0\TrustedCatalog"
    if (-not (Test-Path $root)) { return $false }
    return $null -ne (Get-ChildItem $root -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -match [regex]::Escape($ManifestId) } | Select-Object -First 1)
}

function Install-TrustedCatalogAddin([ValidateSet("Word", "PowerPoint")][string]$OfficeHost, [string]$ManifestId) {
    if (Test-TrustedCatalogCommandCache $ManifestId) { return }
    $processName = if ($OfficeHost -eq "Word") { "WINWORD" } else { "POWERPNT" }
    if (Get-Process $processName -ErrorAction SilentlyContinue) {
        throw "Close all $OfficeHost windows before installing the VisualTeX OLE add-in."
    }

    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes
    if (-not ("VisualTeXOfficeInput" -as [type])) {
        Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class VisualTeXOfficeInput {
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr handle);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr handle, int command);
  [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
  [DllImport("user32.dll")] public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr info);
  public static void Click(int x, int y) {
    SetCursorPos(x, y);
    mouse_event(2, 0, 0, 0, UIntPtr.Zero);
    mouse_event(4, 0, 0, 0, UIntPtr.Zero);
  }
}
'@
    }

    $application = $null
    $document = $null
    $startedProcessId = $null
    try {
        if ($OfficeHost -eq "Word") {
            $application = New-Object -ComObject Word.Application
            $application.Visible = $true
            $document = $application.Documents.Add()
        } else {
            $application = New-Object -ComObject PowerPoint.Application
            $application.Visible = -1
            $document = $application.Presentations.Add()
            $document.Slides.Add(1, 12) | Out-Null
            $document.Windows.Item(1).Activate()
        }
        Start-Sleep -Seconds 5
        $process = Get-Process $processName | Sort-Object StartTime -Descending | Select-Object -First 1
        $startedProcessId = $process.Id
        $shell = New-Object -ComObject WScript.Shell
        $shell.AppActivate([int]$process.Id) | Out-Null
        Start-Sleep -Seconds 1
        $application.CommandBars.ExecuteMso("OfficeExtensionsShowAddinFlyout")
        Start-Sleep -Seconds 2

        $root = [System.Windows.Automation.AutomationElement]::RootElement
        $processCondition = [System.Windows.Automation.PropertyCondition]::new(
            [System.Windows.Automation.AutomationElement]::ProcessIdProperty, [int]$process.Id)
        $moreAddins = $null
        $moreAddinsZh = ([string][char]0x66F4) + [char]0x591A + [char]0x52A0 + [char]0x8F7D + [char]0x9879
        $moreDeadline = [DateTimeOffset]::UtcNow.AddSeconds(15)
        do {
            $windows = $root.FindAll([System.Windows.Automation.TreeScope]::Children, $processCondition)
            foreach ($window in $windows) {
                $elements = $window.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
                foreach ($element in $elements) {
                    if ($element.Current.ControlType -eq [System.Windows.Automation.ControlType]::Button -and
                        ($element.Current.Name -eq $moreAddinsZh -or $element.Current.Name -match 'More.*Add-ins')) {
                        $moreAddins = $element
                        break
                    }
                }
                if ($moreAddins) { break }
            }
            if (-not $moreAddins) {
                $shell.AppActivate([int]$process.Id) | Out-Null
                try { $application.CommandBars.ExecuteMso("OfficeExtensionsShowAddinFlyout") } catch { }
                Start-Sleep -Seconds 2
            }
        } while (-not $moreAddins -and [DateTimeOffset]::UtcNow -lt $moreDeadline)
        if (-not $moreAddins) { throw "$OfficeHost did not expose the More Add-ins command." }
        $moreAddins.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern).Invoke()
        Start-Sleep -Seconds 8

        $windows = $root.FindAll([System.Windows.Automation.TreeScope]::Children, $processCondition)
        $dialog = $null
        $officeAddinsZh = "Office " + [char]0x52A0 + [char]0x8F7D + [char]0x9879
        foreach ($window in $windows) {
            $elements = $window.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
            foreach ($element in $elements) {
                if ($element.Current.Name -eq $officeAddinsZh -or $element.Current.Name -eq "Office Add-ins") {
                    $dialog = $element
                    break
                }
            }
            if ($dialog) { break }
        }
        if (-not $dialog) { throw "$OfficeHost did not open the Office Add-ins dialog." }
        $bounds = $dialog.Current.BoundingRectangle
        $handle = [IntPtr]$dialog.Current.NativeWindowHandle
        [VisualTeXOfficeInput]::ShowWindow($handle, 5) | Out-Null
        [VisualTeXOfficeInput]::SetForegroundWindow($handle) | Out-Null
        Start-Sleep -Seconds 1
        [VisualTeXOfficeInput]::Click([int]($bounds.Left + 0.158 * $bounds.Width), [int]($bounds.Top + 0.061 * $bounds.Height))
        Start-Sleep -Seconds 5
        [VisualTeXOfficeInput]::Click([int]($bounds.Left + 0.158 * $bounds.Width), [int]($bounds.Top + 0.104 * $bounds.Height))
        Start-Sleep -Seconds 2
        [VisualTeXOfficeInput]::Click([int]($bounds.Left + 0.847 * $bounds.Width), [int]($bounds.Top + 0.973 * $bounds.Height))

        $deadline = [DateTimeOffset]::UtcNow.AddSeconds(25)
        do {
            Start-Sleep -Milliseconds 500
            if (Test-TrustedCatalogCommandCache $ManifestId) { return }
        } while ([DateTimeOffset]::UtcNow -lt $deadline)
        throw "$OfficeHost did not persist the VisualTeX trusted-catalog command cache."
    } finally {
        try {
            if ($OfficeHost -eq "Word" -and $document) { $document.Close(0) }
            elseif ($document) { $document.Close() }
        } catch { }
        try { if ($application) { $application.Quit() } } catch { }
        if ($document) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) | Out-Null }
        if ($application) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($application) | Out-Null }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        if ($startedProcessId) {
            Start-Sleep -Seconds 2
            Get-Process -Id $startedProcessId -ErrorAction SilentlyContinue |
                Stop-Process -Force -ErrorAction SilentlyContinue
        }
    }
}

function Install-TrustedCatalogAddinWithRetry(
    [ValidateSet("Word", "PowerPoint")][string]$OfficeHost,
    [string]$ManifestId
) {
    $processName = if ($OfficeHost -eq "Word") { "WINWORD" } else { "POWERPNT" }
    $errors = [Collections.Generic.List[string]]::new()
    for ($attempt = 1; $attempt -le 2; $attempt++) {
        try {
            Write-Host "Configuring the VisualTeX $OfficeHost add-in (attempt $attempt of 2)."
            Install-TrustedCatalogAddin $OfficeHost $ManifestId
            return
        } catch {
            $errors.Add($_.Exception.Message)
            Write-Warning "VisualTeX $OfficeHost add-in setup attempt $attempt failed: $($_.Exception.Message)"
            Get-Process $processName -ErrorAction SilentlyContinue |
                Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 2
        }
    }
    throw "VisualTeX $OfficeHost add-in setup was interrupted or could not complete after two attempts. Do not close the Office Add-ins window while setup is running. $($errors -join ' | ')"
}

$root = Split-Path -Parent $scriptRoot
if ([string]::IsNullOrWhiteSpace($root)) {
    throw "Unable to determine the VisualTeX installation root from script path: $PSScriptRoot"
}
$manifestRoot = Join-Path $root "office\windows\ole\manifests"
if (-not (Test-Path $manifestRoot)) {
    $manifestRoot = Join-Path $root "office-manifests\windows-ole"
}
if (-not (Test-Path $manifestRoot)) {
    $manifestRoot = Join-Path $root "office\manifests"
}
$catalog = Join-Path $env:LOCALAPPDATA "VisualTeX\OfficeCatalog"
$catalogShareName = "VisualTeXOfficeCatalog"
$catalogId = "{69C6A866-755B-4C5A-BACB-EEA28B03C724}"
$catalogKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\$catalogId"
$legacyCatalogKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\VisualTeX"
$developerKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
$modeKey = "HKCU:\Software\VisualTeX\OfficeIntegration"
$certificateScript = Join-Path $scriptRoot "ensure_windows_office_certificate.ps1"
if (-not (Test-Path $certificateScript)) {
    throw "VisualTeX Windows Office certificate helper is missing: $certificateScript"
}
& $certificateScript
Ensure-OfficeWebViewLoopbackExemption

$wordManifest = Join-Path $manifestRoot "VisualTeX.Word.xml"
$powerPointManifest = Join-Path $manifestRoot "VisualTeX.PowerPoint.xml"
if (-not (Test-Path $wordManifest) -or -not (Test-Path $powerPointManifest)) {
    throw "Windows OLE manifests are missing. Run npm run prepare:office:windows-ole first."
}
$wordManifestId = Read-AndValidateManifest $wordManifest "Document" "7c7d3b35-56b2-4c40-88d9-c9eb836d6021"
$powerPointManifestId = Read-AndValidateManifest $powerPointManifest "Presentation" "fdc8d615-7e60-4586-bff4-5a1d728f9f6c"
Clear-VisualTeXWefCache @($wordManifestId, $powerPointManifestId)
Clear-VisualTeXCustomUiCache @($wordManifestId, $powerPointManifestId)

New-Item $catalog -ItemType Directory -Force | Out-Null
Copy-Item $wordManifest (Join-Path $catalog "VisualTeX.WindowsOle.Word.xml") -Force
Copy-Item $powerPointManifest (Join-Path $catalog "VisualTeX.WindowsOle.PowerPoint.xml") -Force
$installedWordManifest = Join-Path $catalog "VisualTeX.WindowsOle.Word.xml"
$installedPowerPointManifest = Join-Path $catalog "VisualTeX.WindowsOle.PowerPoint.xml"
$catalogUnc = Ensure-VisualTeXCatalogShare $catalogShareName $catalog

Remove-Item $legacyCatalogKey -Recurse -Force -ErrorAction SilentlyContinue
New-Item $catalogKey -Force | Out-Null
New-ItemProperty $catalogKey -Name "Id" -PropertyType String -Value $catalogId -Force | Out-Null
New-ItemProperty $catalogKey -Name "Url" -PropertyType String -Value $catalogUnc -Force | Out-Null
New-ItemProperty $catalogKey -Name "Flags" -PropertyType DWord -Value 1 -Force | Out-Null

# Trusted-catalog installation is the persistent, document-independent Office
# registration. Remove developer sideload mappings to avoid duplicate ribbons.
New-Item $developerKey -Force | Out-Null
Remove-ItemProperty $developerKey -Name $wordManifestId -ErrorAction SilentlyContinue
Remove-ItemProperty $developerKey -Name $powerPointManifestId -ErrorAction SilentlyContinue

# Keep installed VSTO files intact, but disable both add-ins so Office cannot load
# the native and OLE ribbons simultaneously.
foreach ($key in @(
    "HKCU:\Software\Microsoft\Office\Word\Addins\VisualTeX.WordVsto",
    "HKCU:\Software\Microsoft\Office\PowerPoint\Addins\VisualTeX.PowerPointVsto"
)) {
    if (Test-Path $key) {
        New-ItemProperty $key -Name "LoadBehavior" -PropertyType DWord -Value 0 -Force | Out-Null
        if ((Get-ItemProperty $key -Name LoadBehavior).LoadBehavior -ne 0) {
            throw "VSTO remained enabled in OLE mode: $key"
        }
    }
}

if (-not (Test-Path $modeKey)) { New-Item $modeKey -Force | Out-Null }
New-ItemProperty $modeKey -Name "Mode" -PropertyType String -Value "ole" -Force | Out-Null
New-ItemProperty $modeKey -Name "OleManifestEnabled" -PropertyType DWord -Value 1 -Force | Out-Null

$runKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
if ($EnableBackgroundStart) {
    $visualTeX = @(
        (Join-Path $env:LOCALAPPDATA "Programs\VisualTeX\VisualTeX.exe"),
        (Join-Path $env:LOCALAPPDATA "VisualTeX\visualtex.exe")
    ) | Where-Object { Test-Path $_ } | Select-Object -First 1
    if (-not [string]::IsNullOrWhiteSpace($visualTeX)) {
        New-ItemProperty $runKey -Name "VisualTeXOffice" -PropertyType String -Value ('"{0}" --office-background' -f $visualTeX) -Force | Out-Null
    }
} else {
    Remove-ItemProperty $runKey -Name "VisualTeXOffice" -ErrorAction SilentlyContinue
}

$certificateThumbprint = (Get-ItemProperty $modeKey -Name CertificateThumbprint -ErrorAction Stop).CertificateThumbprint
if (-not (Get-ChildItem Cert:\CurrentUser\Root | Where-Object { $_.Thumbprint -eq $certificateThumbprint })) {
    throw "VisualTeX HTTPS certificate is not in the current-user Root store."
}
$visualTeX = @(
    (Join-Path (Split-Path -Parent $root) "VisualTeX.exe"),
    (Join-Path $env:LOCALAPPDATA "Programs\VisualTeX\VisualTeX.exe"),
    (Join-Path $env:LOCALAPPDATA "VisualTeX\visualtex.exe")
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $visualTeX) { throw "VisualTeX executable was not found; the OLE HTTPS/Bridge host cannot start." }
Start-Process -FilePath $visualTeX -ArgumentList "--office-background" -WindowStyle Hidden | Out-Null
$healthDeadline = [DateTimeOffset]::UtcNow.AddSeconds(20)
do {
    try {
        $health = Invoke-WebRequest "https://127.0.0.1:43127/health" -UseBasicParsing -TimeoutSec 2
        if ($health.StatusCode -eq 200) { break }
    } catch { }
    Start-Sleep -Milliseconds 250
} while ([DateTimeOffset]::UtcNow -lt $healthDeadline)
if ($null -eq $health -or $health.StatusCode -ne 200) {
    throw "VisualTeX Office HTTPS companion did not become healthy on port 43127."
}
foreach ($resource in @("/health", "/bridge/index.html", "/icons/icon-16.png", "/icons/icon-32.png", "/icons/icon-80.png")) {
    Assert-HttpsResource $resource
}

Install-TrustedCatalogAddinWithRetry Word $wordManifestId
Install-TrustedCatalogAddinWithRetry PowerPoint $powerPointManifestId
if (-not (Test-TrustedCatalogCommandCache $wordManifestId) -or
    -not (Test-TrustedCatalogCommandCache $powerPointManifestId)) {
    throw "VisualTeX OLE Ribbon commands were not persisted for both Word and PowerPoint."
}

Write-Host "VisualTeX Windows OLE Office integration installed for the current user."
Write-Host "Trusted catalog: $catalogUnc"
