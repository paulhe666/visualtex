[CmdletBinding()]
param()

$processes = @(Get-Process VisualTeX -ErrorAction SilentlyContinue)
if ($processes.Count -eq 0) {
    Write-Host "VisualTeX is not running."
    exit 0
}

foreach ($process in $processes) {
    $closed = $process.CloseMainWindow()
    Write-Host ("VisualTeX PID {0}: CloseMainWindow={1}" -f $process.Id, $closed)
}

$deadline = [DateTime]::UtcNow.AddSeconds(10)
do {
    Start-Sleep -Milliseconds 250
    $remaining = @(Get-Process VisualTeX -ErrorAction SilentlyContinue)
} while ($remaining.Count -gt 0 -and [DateTime]::UtcNow -lt $deadline)

if ($remaining.Count -gt 0) {
    throw "VisualTeX is still running after a normal window-close request; no forced termination was attempted."
}
Write-Host "VisualTeX closed normally."
