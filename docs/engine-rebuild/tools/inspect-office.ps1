# Read-only. Never creates, selects, saves, closes, or quits an Office document.
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$result = [ordered]@{ capturedAt = [DateTime]::UtcNow.ToString('o'); ownership = 'No documents owned by engine rebuild'; processes = @(); word = $null; powerpoint = $null }
$processes = Get-Process | Where-Object { $_.ProcessName -match '^(WINWORD|POWERPNT|visualtex|VisualTeX\.FormulaOleServer)$' }
foreach ($p in $processes) {
    $modules = @()
    try {
        $modules = @($p.Modules | Where-Object { $_.ModuleName -match 'VisualTeX|MathPage|MathType' } | ForEach-Object {
            [ordered]@{ path = $_.FileName; version = $_.FileVersionInfo.FileVersion; sha256 = (Get-FileHash -LiteralPath $_.FileName -Algorithm SHA256).Hash }
        })
    } catch { $modules = @([ordered]@{ readError = $_.Exception.Message }) }
    $result.processes += [ordered]@{ pid = $p.Id; name = $p.ProcessName; path = $p.Path; version = $p.MainModule.FileVersionInfo.FileVersion; modules = $modules }
}
if (@($processes | Where-Object ProcessName -eq 'WINWORD').Count -gt 0) {
    try {
        $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
        $documents = @()
        foreach ($d in $word.Documents) {
            $documents += [ordered]@{ name = $d.Name; fullName = $d.FullName; saved = [bool]$d.Saved; readOnly = [bool]$d.ReadOnly; ownedByThisRun = $false }
        }
        $result.word = [ordered]@{ version = $word.Version; build = $word.Build; hwnd = $word.Hwnd; documents = $documents }
    } catch { $result.word = [ordered]@{ readError = $_.Exception.Message } }
}
if (@($processes | Where-Object ProcessName -eq 'POWERPNT').Count -gt 0) {
    try {
        $ppt = [Runtime.InteropServices.Marshal]::GetActiveObject('PowerPoint.Application')
        $presentations = @()
        foreach ($d in $ppt.Presentations) {
            $presentations += [ordered]@{ name = $d.Name; fullName = $d.FullName; saved = $d.Saved; ownedByThisRun = $false }
        }
        $result.powerpoint = [ordered]@{ version = $ppt.Version; presentations = $presentations }
    } catch { $result.powerpoint = [ordered]@{ readError = $_.Exception.Message } }
}
$result | ConvertTo-Json -Depth 8
