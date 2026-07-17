[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
for ($attempt = 1; $attempt -le 8; $attempt++) {
    $processes = @(Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'")
    if ($processes.Count -eq 0) {
        Write-Host "No Word processes remain."
        exit 0
    }
    $unsafe = @($processes | Where-Object { $_.CommandLine -notmatch '/Automation\s+-Embedding' })
    if ($unsafe.Count -gt 0) {
        throw "A non-automation Word process is running; refusing to close Word."
    }

    try {
        $word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
    }
    catch {
        Write-Host "Word automation processes remain but no COM instance is accessible."
        break
    }
    try {
        for ($index = $word.Documents.Count; $index -ge 1; $index--) {
            $document = $word.Documents.Item($index)
            $content = $null
            $inlineShapes = $null
            $shapes = $null
            $maths = $null
            try {
                $content = $document.Content
                $inlineShapes = $document.InlineShapes
                $shapes = $document.Shapes
                $maths = $document.OMaths
                $text = [string]$content.Text
                $isBlankText = [string]::IsNullOrWhiteSpace($text.Replace("`r", "").Replace("`a", ""))
                if (-not $isBlankText -or $inlineShapes.Count -ne 0 -or $shapes.Count -ne 0 -or $maths.Count -ne 0) {
                    throw "An automation Word document contains content; refusing to close it."
                }
                Write-Host ("Closing blank automation document '{0}' without saving." -f $document.Name)
                $document.Close(0)
            }
            finally {
                foreach ($value in @($maths, $shapes, $inlineShapes, $content, $document)) {
                    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
                        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value)
                    }
                }
            }
        }
        Write-Host ("Closing empty Word automation instance; Windows={0}." -f $word.Windows.Count)
        $word.Quit(0)
    }
    finally {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Start-Sleep -Milliseconds 500
}

Start-Sleep -Seconds 2
$remaining = @(Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE'")
if ($remaining.Count -gt 0) {
    throw "Empty automation Word processes still remain after normal COM Quit; no forced termination was attempted."
}
