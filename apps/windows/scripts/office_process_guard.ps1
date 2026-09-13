param(
    [ValidateSet('Check','Stop')]
    [string]$Mode = 'Check'
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding

$officeNames = @(
    'WINWORD', 'POWERPNT', 'EXCEL', 'OUTLOOK', 'ONENOTE',
    'MSACCESS', 'MSPUB', 'VISIO', 'MSPROJECT'
)

function Get-VisualTeXOfficeProcesses {
    $result = @()
    foreach ($process in @(Get-Process -Name $officeNames -ErrorAction SilentlyContinue)) {
        $commandLine = ''
        try {
            $cim = Get-CimInstance Win32_Process -Filter ("ProcessId={0}" -f $process.Id) -ErrorAction Stop
            if ($null -ne $cim -and $null -ne $cim.CommandLine) {
                $commandLine = [string]$cim.CommandLine
            }
        } catch { }

        $backgroundCom = $commandLine -match '(?i)(?:^|\s)-Embedding(?:\s|$)'
        $hasWindow = $false
        try { $hasWindow = [int64]$process.MainWindowHandle -ne 0 } catch { }
        $kind = if ($backgroundCom) {
            'background COM/OLE (-Embedding, no visible Office window required)'
        } elseif ($hasWindow) {
            'visible Office application'
        } else {
            'background Office process (no visible main window)'
        }
        $result += [pscustomobject]@{
            ProcessName = $process.ProcessName.ToUpperInvariant() + '.EXE'
            ProcessId = [int]$process.Id
            Kind = $kind
            CommandLine = $commandLine
        }
    }
    return @($result | Sort-Object ProcessName, ProcessId)
}

function Write-ProcessSummary([object[]]$Processes) {
    foreach ($item in $Processes) {
        Write-Output ("{0} (PID {1}) - {2}" -f $item.ProcessName, $item.ProcessId, $item.Kind)
    }
}

$running = @(Get-VisualTeXOfficeProcesses)
if ($running.Count -eq 0) {
    Write-Output 'No blocking Microsoft Office processes are running.'
    exit 0
}

if ($Mode -eq 'Check') {
    Write-ProcessSummary $running
    exit 10
}

Write-Output 'Stopping the Office processes explicitly approved by the installer user:'
Write-ProcessSummary $running
foreach ($item in $running) {
    try {
        Stop-Process -Id $item.ProcessId -Force -ErrorAction Stop
    } catch {
        Write-Output ("Failed to stop {0} (PID {1}): {2}" -f $item.ProcessName, $item.ProcessId, $_.Exception.Message)
    }
}

$deadline = [DateTime]::UtcNow.AddSeconds(4)
do {
    Start-Sleep -Milliseconds 250
    $remaining = @(Get-VisualTeXOfficeProcesses)
    if ($remaining.Count -eq 0) {
        Write-Output 'All blocking Microsoft Office processes are closed.'
        exit 0
    }
} while ([DateTime]::UtcNow -lt $deadline)

Write-Output 'The following Microsoft Office processes are still running:'
Write-ProcessSummary $remaining
exit 11
