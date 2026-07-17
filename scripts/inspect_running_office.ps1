[CmdletBinding()]
param()

$ErrorActionPreference = "Continue"
Get-CimInstance Win32_Process -Filter "Name='WINWORD.EXE' OR Name='POWERPNT.EXE'" |
    Select-Object Name, ProcessId, CreationDate, CommandLine, ExecutablePath |
    Format-List

try {
    $word = [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
    Write-Host ("Active Word: Documents={0}; Windows={1}; Visible={2}" -f $word.Documents.Count, $word.Windows.Count, $word.Visible)
    for ($index = 1; $index -le $word.Documents.Count; $index++) {
        $document = $word.Documents.Item($index)
        Write-Host ("Word document {0}: Name={1}; FullName={2}; Saved={3}" -f $index, $document.Name, $document.FullName, $document.Saved)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document)
    }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}
catch {
    Write-Host ("No accessible Word ROT instance: {0}" -f $_.Exception.Message)
}

try {
    $powerPoint = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
    Write-Host ("Active PowerPoint: Presentations={0}; Windows={1}" -f $powerPoint.Presentations.Count, $powerPoint.Windows.Count)
    for ($index = 1; $index -le $powerPoint.Presentations.Count; $index++) {
        $presentation = $powerPoint.Presentations.Item($index)
        Write-Host ("PowerPoint presentation {0}: Name={1}; FullName={2}; Saved={3}" -f $index, $presentation.Name, $presentation.FullName, $presentation.Saved)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($presentation)
    }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint)
}
catch {
    Write-Host ("No accessible PowerPoint ROT instance: {0}" -f $_.Exception.Message)
}
