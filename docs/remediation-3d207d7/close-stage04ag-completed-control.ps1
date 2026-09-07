$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$evidence = Join-Path $PSScriptRoot 'evidence'
$stage = Get-Content (Join-Path $evidence 'stage04ag-word-process.json') -Raw | ConvertFrom-Json
$expected = Get-Content (Join-Path $evidence 'stage04ag-first-omml-__3.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$process = Get-Process -Id $stage.pid
if ($process.ProcessName -ne 'WINWORD' -or $stage.pid -ne 111272 -or [Math]::Abs(($process.StartTime.ToUniversalTime() - [DateTimeOffset]::Parse($stage.started).UtcDateTime).TotalSeconds) -gt 10) { throw 'Stage process identity changed.' }
$word = Connect-RunningWord $stage.pid
$records = @()
try {
    for ($index = $word.Documents.Count; $index -ge 1; $index--) {
        $document = $word.Documents.Item($index)
        try {
            $name = [string]$document.Name
            if ($document.FullName -ne $name) { continue }
            $completed = $name -eq $expected.name
            $emptyStarter = $name -match '^[^0-9]+2$' -and $document.Saved -and $document.Content.End -eq 1 -and $document.OMaths.Count -eq 0 -and $document.InlineShapes.Count -eq 0 -and $document.Fields.Count -eq 0
            if (!$completed -and !$emptyStarter) { continue }
            if ($completed -and ($document.Content.End -ne $expected.end -or $document.OMaths.Count -ne $expected.maths.Count -or $document.InlineShapes.Count -ne $expected.shapes.Count -or $document.Fields.Count -ne $expected.fields.Count -or $document.Tables.Count -ne $expected.tables.Count)) { throw 'Completed control changed; preserving document.' }
            $document.Close(0)
            $records += @{ action = 'closed-without-save'; pid = $stage.pid; document = $name; completedControl = $completed; emptyStarter = $emptyStarter; time = [DateTime]::UtcNow.ToString('o') }
        } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document) }
    }
    $remaining = @()
    foreach ($document in $word.Documents) {
        try { $remaining += @{ document = $document.Name; saved = $document.Saved; end = $document.Content.End } }
        finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document) }
    }
    $result = @{ closed = $records; remaining = $remaining; activeDocument = $word.ActiveDocument.Name }
    $result | ConvertTo-Json -Depth 6 | Set-Content (Join-Path $evidence 'stage04ag-close-completed-control.json') -Encoding UTF8
    $result | ConvertTo-Json -Depth 6
} finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
