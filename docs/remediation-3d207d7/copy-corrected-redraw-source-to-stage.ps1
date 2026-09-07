param(
    [Parameter(Mandatory=$true)][int]$SourceWordProcessId,
    [Parameter(Mandatory=$true)][string]$SourceDocumentName,
    [Parameter(Mandatory=$true)][string]$Stage
)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$recordPath=Join-Path $PSScriptRoot "evidence\$Stage-word-process.json"
$record=Get-Content -LiteralPath $recordPath -Raw -Encoding UTF8 | ConvertFrom-Json
if(!$record.pid -or [int]$record.pid -eq $SourceWordProcessId){throw 'Source and target Word processes must be different.'}
$sourceWord=Connect-RunningWord $SourceWordProcessId
$targetWord=$null
$sourceDoc=$null
$targetDoc=$null
$targetRange=$null
try {
    $sourceDoc=$sourceWord.Documents.Item($SourceDocumentName)
    $sourceText=$sourceDoc.Content.Text
    if([string]::IsNullOrEmpty($sourceText)){throw 'Source document is empty.'}
    $firstBad='\begin{pmatrix}1&2\3&4\end{pmatrix}'
    $firstGood='\begin{pmatrix}1&2\\3&4\end{pmatrix}'
    $secondBad='\begin{pmatrix}a&b\c&d\end{pmatrix}'
    $secondGood='\begin{pmatrix}a&b\\c&d\end{pmatrix}'
    $firstMatches=([regex]::Matches($sourceText,[regex]::Escape($firstBad))).Count
    $secondMatches=([regex]::Matches($sourceText,[regex]::Escape($secondBad))).Count
    if($firstMatches -ne 1 -or $secondMatches -ne 1){throw "Expected exactly the two known malformed matrix rows; found first=$firstMatches second=$secondMatches."}
    $corrected=$sourceText.Replace($firstBad,$firstGood).Replace($secondBad,$secondGood)
    if($corrected.Length -ne $sourceText.Length + 2){throw 'Corrected source changed by more than the two missing row separators.'}
    $targetWord=Connect-RunningWord ([int]$record.pid)
    if($targetWord.Documents.Count -ne 1){throw 'Target stage must contain exactly one seed document.'}
    $targetDoc=$targetWord.ActiveDocument
    if(!$targetDoc.Saved -or $targetDoc.Path -or $targetDoc.Content.End -ne 1 -or $targetDoc.OMaths.Count -ne 0 -or $targetDoc.InlineShapes.Count -ne 0){throw 'Target stage seed is not an unchanged empty document.'}
    $payload=$corrected
    if($payload.EndsWith("`r")){$payload=$payload.Substring(0,$payload.Length-1)}
    $targetRange=$targetDoc.Range(0,0)
    $targetRange.Text=$payload
    if($targetDoc.Content.Text -ne $corrected){throw 'Target stage text does not exactly match the corrected source after copy.'}
    $targetDoc.Activate()
    $record | Add-Member -NotePropertyName document -NotePropertyValue $targetDoc.Name -Force
    $record | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'evidence\target-word.json') -Encoding UTF8
    Write-Output ("CORRECTED_SOURCE_COPIED|sourcePid=$SourceWordProcessId|source=$SourceDocumentName|targetPid=$($record.pid)|target=$($targetDoc.Name)|sourceLength=$($sourceText.Length)|correctedLength=$($corrected.Length)|corrections=2")
}
finally {
    if($targetRange){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($targetRange)}
    if($targetDoc){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($targetDoc)}
    if($targetWord){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($targetWord)}
    if($sourceDoc){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($sourceDoc)}
    if($sourceWord){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($sourceWord)}
}
