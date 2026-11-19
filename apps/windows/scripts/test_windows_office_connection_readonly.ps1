[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$VisualTeXPath,
    [string]$ReportPath
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
Add-Type -AssemblyName System.Net.Http

$VisualTeXPath = (Resolve-Path -LiteralPath $VisualTeXPath).Path
$integrationKey = "HKCU:\Software\VisualTeX\OfficeIntegration"
$integration = Get-ItemProperty -LiteralPath $integrationKey -ErrorAction Stop
$registeredExecutable = [string]$integration.ExecutablePath
if (-not [string]::Equals(
    [IO.Path]::GetFullPath($registeredExecutable),
    [IO.Path]::GetFullPath($VisualTeXPath),
    [StringComparison]::OrdinalIgnoreCase)) {
    throw "Registered companion executable '$registeredExecutable' does not match '$VisualTeXPath'."
}
$port = [int]$integration.CompanionPort
$expectedProtocol = [int]$integration.ProtocolVersion
if ($port -le 0 -or $expectedProtocol -le 0) {
    throw "OfficeIntegration companion port/protocol configuration is invalid."
}

# CompanionPort is the standard port, which install.json names too. When it is taken
# or reserved, VisualTeX listens on another one and writes the port it bound to
# CompanionActivePort at every start.
function Get-CompanionListeningPort([int]$RegisteredPort) {
    $active = 0
    try {
        $key = Get-Item -LiteralPath "HKCU:\Software\VisualTeX\OfficeIntegration" -ErrorAction SilentlyContinue
        if ($null -ne $key) { $active = [int]$key.GetValue("CompanionActivePort", 0) }
    } catch { $active = 0 }
    if ($active -gt 0 -and $active -le 65535) { return $active }
    return $RegisteredPort
}
if ([string]::IsNullOrWhiteSpace($ReportPath)) {
    $logRoot = Join-Path $env:LOCALAPPDATA "VisualTeX\office\install-logs"
    New-Item -ItemType Directory -Path $logRoot -Force | Out-Null
    $ReportPath = Join-Path $logRoot ("office-readonly-connection-{0}.json" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
}

function Get-ExactProcess([string]$Path) {
    foreach ($process in @(Get-Process visualtex -ErrorAction SilentlyContinue)) {
        try {
            if ([string]::Equals(
                [IO.Path]::GetFullPath($process.Path),
                [IO.Path]::GetFullPath($Path),
                [StringComparison]::OrdinalIgnoreCase)) {
                return $process
            }
        } catch { }
    }
    return $null
}

function Wait-CompanionHealth([int]$Port, [int]$ProtocolVersion, [int]$TimeoutSeconds = 25) {
    $handler = [System.Net.Http.HttpClientHandler]::new()
    $handler.UseProxy = $false
    $handler.Proxy = $null
    $client = [System.Net.Http.HttpClient]::new($handler)
    $client.Timeout = [TimeSpan]::FromSeconds(3)
    try {
        $deadline = [DateTimeOffset]::UtcNow.AddSeconds($TimeoutSeconds)
        $lastError = ""
        do {
            try {
                # the port is known once VisualTeX has bound it
                $listeningPort = Get-CompanionListeningPort $Port
                $raw = $client.GetStringAsync("https://127.0.0.1:$listeningPort/health").GetAwaiter().GetResult()
                $health = $raw | ConvertFrom-Json
                if ([bool]$health.ok -and [int]$health.protocolVersion -eq $ProtocolVersion) {
                    return [pscustomobject]@{ raw = $raw; health = $health }
                }
                $lastError = "Unexpected health response: $raw"
            } catch {
                $lastError = $_.Exception.Message
            }
            Start-Sleep -Milliseconds 250
        } while ([DateTimeOffset]::UtcNow -lt $deadline)
        throw "Companion did not become healthy: $lastError"
    } finally {
        $client.Dispose()
        $handler.Dispose()
    }
}

function Get-OrStartOfficeApplication([string]$ProgId) {
    try {
        $application = [Runtime.InteropServices.Marshal]::GetActiveObject($ProgId)
        return [pscustomobject]@{ Application = $application; Started = $false }
    } catch {
        $application = New-Object -ComObject $ProgId
        return [pscustomobject]@{ Application = $application; Started = $true }
    }
}

function Wait-AddinConnected([object]$Application, [string]$ProgId, [int]$TimeoutSeconds = 20) {
    $deadline = [DateTimeOffset]::UtcNow.AddSeconds($TimeoutSeconds)
    $lastError = ""
    do {
        $addin = $null
        try {
            $addin = $Application.COMAddIns.Item($ProgId)
            if ([bool]$addin.Connect) {
                return [pscustomobject]@{
                    connected = $true
                    progId = $ProgId
                    description = [string]$addin.Description
                }
            }
            $lastError = "COMAddIn.Connect is false"
        } catch {
            $lastError = $_.Exception.Message
        } finally {
            if ($null -ne $addin -and [Runtime.InteropServices.Marshal]::IsComObject($addin)) {
                [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($addin)
            }
        }
        Start-Sleep -Milliseconds 300
    } while ([DateTimeOffset]::UtcNow -lt $deadline)
    throw "$ProgId did not report COMAddIn.Connect=true: $lastError"
}

$companionProcess = Get-ExactProcess $VisualTeXPath
$startedCompanion = $false
if ($null -eq $companionProcess) {
    $companionProcess = Start-Process -FilePath $VisualTeXPath -ArgumentList "--office-background" -PassThru
    $startedCompanion = $true
}
$wordState = $null
$powerPointState = $null
$wordResult = $null
$powerPointResult = $null
try {
    $health = Wait-CompanionHealth $port $expectedProtocol

    $wordState = Get-OrStartOfficeApplication "Word.Application"
    $wordResult = Wait-AddinConnected $wordState.Application "VisualTeX.WordVsto"

    $powerPointState = Get-OrStartOfficeApplication "PowerPoint.Application"
    $powerPointResult = Wait-AddinConnected $powerPointState.Application "VisualTeX.PowerPointVsto"

    $report = [ordered]@{
        schemaVersion = 1
        verifiedAt = [DateTimeOffset]::Now.ToString("o")
        executable = $VisualTeXPath
        companionPid = $companionProcess.Id
        companionStartedByProbe = $startedCompanion
        health = $health.health
        word = [ordered]@{
            usedExistingInstance = -not $wordState.Started
            connected = $wordResult.connected
            progId = $wordResult.progId
            description = $wordResult.description
        }
        powerPoint = [ordered]@{
            usedExistingInstance = -not $powerPointState.Started
            connected = $powerPointResult.connected
            progId = $powerPointResult.progId
            description = $powerPointResult.description
        }
    }
    $report | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $ReportPath -Encoding UTF8
    Write-Host "Non-destructive Office connection verification passed."
    Write-Host "Word existing instance: $(-not $wordState.Started); Connect=$($wordResult.connected)"
    Write-Host "PowerPoint existing instance: $(-not $powerPointState.Started); Connect=$($powerPointResult.connected)"
    Write-Host "Report: $ReportPath"
} finally {
    if ($null -ne $powerPointState) {
        if ($powerPointState.Started) {
            try { $powerPointState.Application.Quit() } catch { }
        }
        if ([Runtime.InteropServices.Marshal]::IsComObject($powerPointState.Application)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($powerPointState.Application)
        }
    }
    if ($null -ne $wordState) {
        if ($wordState.Started) {
            try { $wordState.Application.Quit(0) } catch { }
        }
        if ([Runtime.InteropServices.Marshal]::IsComObject($wordState.Application)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wordState.Application)
        }
    }
    if ($startedCompanion -and $null -ne $companionProcess) {
        try {
            $companionProcess.Refresh()
            if (-not $companionProcess.HasExited) {
                Stop-Process -Id $companionProcess.Id -Force -ErrorAction SilentlyContinue
            }
        } catch { }
    }
    if ($null -ne $companionProcess) { $companionProcess.Dispose() }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
