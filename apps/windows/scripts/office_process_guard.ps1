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
    # The installer only needs to know whether a blocking Office process exists
    # and whether it has a visible main window. Querying Win32_Process once per
    # PID made this guard scale linearly with the number of open Word instances;
    # a heavily used/acceptance machine could exceed the NSIS 20 s timeout even
    # though process inspection itself was healthy. Get-Process already reports
    # hidden COM/OLE Office instances, so avoid CIM/WMI entirely here.
    $result = @()
    foreach ($process in @(Get-Process -Name $officeNames -ErrorAction SilentlyContinue)) {
        $hasWindow = $false
        try { $hasWindow = [int64]$process.MainWindowHandle -ne 0 } catch { }
        $kind = if ($hasWindow) {
            'visible Office application'
        } else {
            'background Office process (no visible main window; may be COM/OLE)'
        }
        $result += [pscustomobject]@{
            ProcessName = $process.ProcessName.ToUpperInvariant() + '.EXE'
            ProcessId = [int]$process.Id
            Kind = $kind
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
