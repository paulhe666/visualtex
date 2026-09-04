[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestPath,
    [Parameter(Mandatory = $true)]
    [string]$ArtifactRoot,
    [int]$StartIndex = 0,
    [int]$Count = 0,
    [ValidateRange(30, 600)]
    [int]$TimeoutSeconds = 240
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)

function Stop-AutomationWordProcesses {
    Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -eq "WINWORD.EXE" -and
            $_.CommandLine -match "(?i)/(Automation|Embedding)|-Embedding"
        } |
        ForEach-Object {
            Stop-Process -Id ([int]$_.ProcessId) -Force -ErrorAction SilentlyContinue
        }
}

$resolvedManifest = (Resolve-Path $ManifestPath).Path
$manifest = Get-Content -LiteralPath $resolvedManifest -Raw | ConvertFrom-Json
$items = @($manifest.items)
if ($items.Count -eq 0) {
    throw "The manifest contains no formula items."
}
if ($StartIndex -lt 0 -or $StartIndex -ge $items.Count) {
    throw "StartIndex is outside the manifest item range."
}
$endExclusive = if ($Count -gt 0) {
    [Math]::Min($items.Count, $StartIndex + $Count)
} else {
    $items.Count
}
$selected = @($items | Select-Object -Skip $StartIndex -First ($endExclusive - $StartIndex))
if ($selected.Count -eq 0) {
    throw "The selected manifest slice is empty."
}

$root = Split-Path -Parent $PSScriptRoot
$acceptancePath = Join-Path $root (
    "src-windows\VisualTeX.NativeOfficeOleAcceptance\bin\x64\Release\net48\" +
    "VisualTeX.NativeOfficeOleAcceptance.exe")
if (-not (Test-Path $acceptancePath)) {
    throw "The x64 Release native Office acceptance executable is missing: $acceptancePath"
}

$resolvedArtifactRoot = [IO.Path]::GetFullPath($ArtifactRoot)
[IO.Directory]::CreateDirectory($resolvedArtifactRoot) | Out-Null
$sliceName = "{0:D3}-{1:D3}" -f $StartIndex, ($endExclusive - 1)
$subsetPath = Join-Path $env:TEMP (
    "visualtex-word-inline-baseline-$sliceName-" + [guid]::NewGuid().ToString("N") + ".json")
$stdoutPath = Join-Path $resolvedArtifactRoot "stdout.log"
$stderrPath = Join-Path $resolvedArtifactRoot "stderr.log"

try {
    $subset = [pscustomobject]@{
        root = $manifest.root
        suite = "{0}-{1}" -f ([string]$manifest.suite), $sliceName
        items = $selected
    }
    [IO.File]::WriteAllText(
        $subsetPath,
        ($subset | ConvertTo-Json -Depth 30),
        [Text.UTF8Encoding]::new($false))

    $startInfo = [Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = $acceptancePath
    $startInfo.Arguments = [string]::Join(" ", @(
        "--word-inline-baseline-manifest",
        ('"' + $subsetPath + '"'),
        ('"' + $resolvedArtifactRoot + '"')
    ))
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.StandardOutputEncoding = [Text.UTF8Encoding]::new($false)
    $startInfo.StandardErrorEncoding = [Text.UTF8Encoding]::new($false)

    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $startInfo
    if (-not $process.Start()) {
        throw "Unable to start the native Word inline-baseline acceptance."
    }
    $stdoutTask = $process.StandardOutput.ReadToEndAsync()
    $stderrTask = $process.StandardError.ReadToEndAsync()
    if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
        Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
        Stop-AutomationWordProcesses
        throw "Word inline-baseline acceptance slice $sliceName exceeded $TimeoutSeconds seconds."
    }
    $process.WaitForExit()
    $stdout = $stdoutTask.GetAwaiter().GetResult()
    $stderr = $stderrTask.GetAwaiter().GetResult()
    [IO.File]::WriteAllText($stdoutPath, $stdout, [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText($stderrPath, $stderr, [Text.UTF8Encoding]::new($false))
    $exitCode = $process.ExitCode
    $process.Dispose()

    $suiteStem = (([string]$subset.suite).ToCharArray() | ForEach-Object {
        if ([char]::IsLetterOrDigit($_) -or $_ -in @('-', '_')) { $_ } else { '-' }
    }) -join ''
    $reportPath = Join-Path $resolvedArtifactRoot (
        "word-inline-baseline-" + $suiteStem.Trim('-') + ".json")
    if (-not (Test-Path $reportPath)) {
        throw "The native acceptance report was not produced. ExitCode=$exitCode. STDERR=$stderr"
    }
    $report = Get-Content -LiteralPath $reportPath -Raw | ConvertFrom-Json
    $worst = @($report.Metrics |
        Sort-Object {
            [Math]::Max(
                [Math]::Abs([double]$_.AnchorDeltaBottomPt),
                [Math]::Abs([double]$_.AnchorDeltaCentroidPt))
        } -Descending |
        Select-Object -First 6 `
            Name,FontSizePt,FormulaLetterFont,PositionPt,
            AnchorDeltaBottomPt,AnchorDeltaCentroidPt,
            PreviewInkTopMarginPx,PreviewInkBottomMarginPx,
            TextTopSpreadPt,TextBottomSpreadPt)

    $summary = [pscustomobject]@{
        Manifest = $resolvedManifest
        Slice = $sliceName
        ExitCode = $exitCode
        Passed = [int]$report.Passed
        Failed = [int]$report.Failed
        Report = $reportPath
        Worst = $worst
        Failures = @($report.Failures)
    }
    $summary | ConvertTo-Json -Depth 8
    if ($exitCode -ne 0 -or [int]$report.Failed -ne 0) {
        exit 1
    }
}
finally {
    Remove-Item -LiteralPath $subsetPath -Force -ErrorAction SilentlyContinue
}
