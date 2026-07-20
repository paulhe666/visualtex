param(
    [string]$ArtifactRoot = "src-windows/artifacts/installed-ribbon-icons"
)

$ErrorActionPreference = "Stop"
$artifactPath = [IO.Path]::GetFullPath((Join-Path (Get-Location) $ArtifactRoot))
New-Item -ItemType Directory -Path $artifactPath -Force | Out-Null

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class VisualTeXRibbonNative {
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hwnd);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hwnd, int command);
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
}
"@

function Release-ComObject([object]$value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value) } catch { }
    }
}

function Find-AutomationElementByName(
    [System.Windows.Automation.AutomationElement]$root,
    [string]$name,
    [System.Windows.Automation.ControlType]$controlType
) {
    $nameCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::NameProperty,
        $name
    )
    $typeCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
        $controlType
    )
    $condition = New-Object System.Windows.Automation.AndCondition($nameCondition, $typeCondition)
    return $root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $condition)
}

function Wait-OfficeProcessWindowHandle([string]$processName) {
    $deadline = [DateTime]::UtcNow.AddSeconds(15)
    do {
        $process = Get-Process $processName -ErrorAction SilentlyContinue |
            Sort-Object StartTime -Descending |
            Select-Object -First 1
        if ($null -ne $process -and $process.MainWindowHandle -ne 0) {
            return [IntPtr]$process.MainWindowHandle
        }
        Start-Sleep -Milliseconds 250
    } while ([DateTime]::UtcNow -lt $deadline)
    throw "$processName did not expose a non-zero main window handle."
}

function Select-VisualTeXTab([IntPtr]$windowHandle) {
    [void][VisualTeXRibbonNative]::SetForegroundWindow($windowHandle)
    Start-Sleep -Milliseconds 500
    $root = [System.Windows.Automation.AutomationElement]::FromHandle($windowHandle)
    $deadline = [DateTime]::UtcNow.AddSeconds(15)
    do {
        $tab = Find-AutomationElementByName $root "VisualTeX" ([System.Windows.Automation.ControlType]::TabItem)
        if ($null -ne $tab) { break }
        Start-Sleep -Milliseconds 250
    } while ([DateTime]::UtcNow -lt $deadline)
    if ($null -eq $tab) { throw "VisualTeX Ribbon tab was not found." }

    $selectionPattern = $tab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
    $selectionPattern.Select()
    Start-Sleep -Milliseconds 800
    return $root
}

function Save-WindowRibbonScreenshot(
    [System.Windows.Automation.AutomationElement]$root,
    [string]$path
) {
    $bounds = $root.Current.BoundingRectangle
    if ($bounds.Width -le 0 -or $bounds.Height -le 0) {
        throw "Office window has invalid screen bounds."
    }
    $height = [Math]::Min([int][Math]::Ceiling($bounds.Height), 330)
    $width = [int][Math]::Ceiling($bounds.Width)
    $bitmap = New-Object System.Drawing.Bitmap($width, $height)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    try {
        $graphics.CopyFromScreen(
            [int][Math]::Floor($bounds.Left),
            [int][Math]::Floor($bounds.Top),
            0,
            0,
            (New-Object System.Drawing.Size($width, $height)),
            [System.Drawing.CopyPixelOperation]::SourceCopy
        )
        $bitmap.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally {
        $graphics.Dispose()
        $bitmap.Dispose()
    }
}

function Measure-ButtonArtwork(
    [System.Windows.Automation.AutomationElement]$root,
    [string]$screenshotPath,
    [string[]]$buttonNames
) {
    $windowBounds = $root.Current.BoundingRectangle
    $bitmap = New-Object System.Drawing.Bitmap($screenshotPath)
    try {
        $results = @()
        foreach ($name in $buttonNames) {
            $button = Find-AutomationElementByName $root $name ([System.Windows.Automation.ControlType]::Button)
            if ($null -eq $button) { throw "Ribbon button '$name' was not found." }
            $bounds = $button.Current.BoundingRectangle
            $left = [Math]::Max(0, [int][Math]::Floor($bounds.Left - $windowBounds.Left))
            $top = [Math]::Max(0, [int][Math]::Floor($bounds.Top - $windowBounds.Top))
            $right = [Math]::Min($bitmap.Width, [int][Math]::Ceiling($bounds.Right - $windowBounds.Left))
            $bottom = [Math]::Min($bitmap.Height, [int][Math]::Ceiling($bounds.Bottom - $windowBounds.Top))
            if ($right -le $left -or $bottom -le $top) {
                throw "Ribbon button '$name' has invalid screenshot bounds."
            }

            $darkPixels = 0
            $nonWhitePixels = 0
            for ($y = $top; $y -lt $bottom; $y++) {
                for ($x = $left; $x -lt $right; $x++) {
                    $pixel = $bitmap.GetPixel($x, $y)
                    $sum = [int]$pixel.R + [int]$pixel.G + [int]$pixel.B
                    if ($sum -lt 690) { $nonWhitePixels++ }
                    if ($sum -lt 480) { $darkPixels++ }
                }
            }
            if ($nonWhitePixels -lt 30 -or $darkPixels -lt 8) {
                throw "Ribbon button '$name' appears blank: nonWhite=$nonWhitePixels dark=$darkPixels."
            }
            $results += [pscustomobject]@{
                Name = $name
                Left = [int]$bounds.Left
                Top = [int]$bounds.Top
                Width = [int]$bounds.Width
                Height = [int]$bounds.Height
                NonWhitePixels = $nonWhitePixels
                DarkPixels = $darkPixels
            }
        }
        return $results
    }
    finally {
        $bitmap.Dispose()
    }
}

$consoleWindow = [VisualTeXRibbonNative]::GetConsoleWindow()
$word = $null
$wordDocument = $null
$wordAddIns = $null
$wordAddIn = $null
$powerPoint = $null
$presentation = $null
$slide = $null
$powerPointAddIns = $null
$powerPointAddIn = $null
try {
    if ($consoleWindow -ne [IntPtr]::Zero) { [void][VisualTeXRibbonNative]::ShowWindow($consoleWindow, 0) }

    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $word.DisplayAlerts = 0
    $wordDocument = $word.Documents.Add()
    $wordAddIns = $word.COMAddIns
    $wordAddIn = $wordAddIns.Item("VisualTeX.WordVsto")
    if (-not $wordAddIn.Connect) {
        $wordAddIn.Connect = $true
        Start-Sleep -Milliseconds 800
    }
    if (-not $wordAddIn.Connect) { throw "Installed Word add-in is not connected." }
    $wordHwnd = Wait-OfficeProcessWindowHandle "WINWORD"
    $wordRoot = Select-VisualTeXTab $wordHwnd
    $wordScreenshot = Join-Path $artifactPath "VisualTeX-Word-Ribbon.png"
    Save-WindowRibbonScreenshot $wordRoot $wordScreenshot
    $wordResults = Measure-ButtonArtwork $wordRoot $wordScreenshot @(
        "OLE 行内公式",
        "OLE 行间公式",
        "OMML 行内公式",
        "OMML 行间公式",
        "编辑所选公式",
        "转为原生 OLE",
        "转为 Word OMML",
        "更新公式编号"
    )

    $wordDocument.Close(0)
    Release-ComObject $wordDocument
    $wordDocument = $null
    $word.Quit()
    Release-ComObject $wordAddIn
    $wordAddIn = $null
    Release-ComObject $wordAddIns
    $wordAddIns = $null
    Release-ComObject $word
    $word = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPoint.Visible = -1
    $presentation = $powerPoint.Presentations.Add(-1)
    $slide = $presentation.Slides.Add(1, 12)
    $powerPointAddIns = $powerPoint.COMAddIns
    $powerPointAddIn = $powerPointAddIns.Item("VisualTeX.PowerPointVsto")
    if (-not $powerPointAddIn.Connect) {
        $powerPointAddIn.Connect = $true
        Start-Sleep -Milliseconds 800
    }
    if (-not $powerPointAddIn.Connect) { throw "Installed PowerPoint add-in is not connected." }
    $powerPointHwnd = Wait-OfficeProcessWindowHandle "POWERPNT"
    $powerPointRoot = Select-VisualTeXTab $powerPointHwnd
    $powerPointScreenshot = Join-Path $artifactPath "VisualTeX-PowerPoint-Ribbon.png"
    Save-WindowRibbonScreenshot $powerPointRoot $powerPointScreenshot
    $powerPointResults = Measure-ButtonArtwork $powerPointRoot $powerPointScreenshot @(
        "新建公式",
        "编辑所选公式",
        "转为原生 OLE"
    )

    Write-Output "WORD_RIBBON_SCREENSHOT=$wordScreenshot"
    $wordResults | Format-Table -AutoSize
    Write-Output "POWERPOINT_RIBBON_SCREENSHOT=$powerPointScreenshot"
    $powerPointResults | Format-Table -AutoSize
    Write-Output "Installed VisualTeX Ribbon icon UI probe passed."
}
finally {
    if ($consoleWindow -ne [IntPtr]::Zero) { [void][VisualTeXRibbonNative]::ShowWindow($consoleWindow, 5) }
    Release-ComObject $powerPointAddIn
    Release-ComObject $powerPointAddIns
    Release-ComObject $slide
    if ($presentation) { try { $presentation.Close() } catch { } }
    Release-ComObject $presentation
    if ($powerPoint) { try { $powerPoint.Quit() } catch { } }
    Release-ComObject $powerPoint
    Release-ComObject $wordAddIn
    Release-ComObject $wordAddIns
    if ($wordDocument) { try { $wordDocument.Close(0) } catch { } }
    Release-ComObject $wordDocument
    if ($word) { try { $word.Quit() } catch { } }
    Release-ComObject $word
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
