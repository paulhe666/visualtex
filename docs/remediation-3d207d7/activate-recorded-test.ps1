param([Parameter(Mandatory=$true)][string]$EvidenceName)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'word-connection.ps1')
if($EvidenceName -notmatch '^stage05-[a-z-]+-result-__\d+\.json$'){throw 'Only this stage recorded redraw documents are allowed.'}
$targetPath=Join-Path $PSScriptRoot 'evidence\target-word.json'
$target=Get-Content -LiteralPath $targetPath -Raw | ConvertFrom-Json
$stage=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'evidence\stage05-word-process.json') -Raw | ConvertFrom-Json
if($target.pid -ne $stage.pid -or $target.pid -eq 27664){throw 'The current process is not the redraw stage process.'}
$record=Get-Content -LiteralPath (Join-Path $PSScriptRoot ('evidence\'+$EvidenceName)) -Raw | ConvertFrom-Json
$word=Connect-RunningWord $target.pid
$document=$word.Documents.Item($record.name)
if($document.Content.End -ne $record.end -or $document.OMaths.Count -ne @($record.maths).Count -or $document.InlineShapes.Count -ne @($record.shapes).Count -or $document.Fields.Count -ne @($record.fields).Count){throw 'The recorded test document has changed; no activation or input was performed.'}
$document.Activate()
$target | Add-Member -NotePropertyName document -NotePropertyValue $document.Name -Force
$target | ConvertTo-Json | Set-Content -LiteralPath $targetPath -Encoding UTF8
Write-Output ('ACTIVATED_RECORDED_REDRAW|'+$target.pid+'|'+$document.Name)
