param([Parameter(Mandatory=$true)][int]$WordProcessId,[Parameter(Mandatory=$true)][string]$ExpectedName,[Parameter(Mandatory=$true)][string]$Label,[switch]$ResumeSaved)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$target=Get-Content (Join-Path $PSScriptRoot 'evidence\target-word.json') -Raw | ConvertFrom-Json
if($WordProcessId -eq 27664 -or $WordProcessId -ne $target.pid){throw 'Only the current stage-owned Word process may be saved or reopened.'}
if($Label -notmatch '^stage[0-9a-z]+-[a-z0-9-]+$'){throw 'Invalid stage evidence label.'}
$path=Join-Path (Join-Path $PSScriptRoot 'evidence') ($Label+'.docx')
if((Test-Path -LiteralPath $path) -and !$ResumeSaved){throw 'Evidence file already exists; refusing overwrite.'}
$word=Connect-RunningWord $WordProcessId
$document=$word.ActiveDocument
if($document.Name -ne $ExpectedName){throw 'Active document is not the explicitly named stage document.'}
if($ResumeSaved){if($document.FullName -ne $path -or !$document.Saved){throw 'Resume only the exact already-saved evidence file; no unsaved changes may be closed.'}}
elseif($document.Path){throw 'Active document is not a new stage document.'}
$before=@{name=$document.Name;maths=$document.OMaths.Count;shapes=$document.InlineShapes.Count;bookmarks=$document.Bookmarks.Count;fields=$document.Fields.Count;end=$document.Content.End}
if(!$ResumeSaved){$document.SaveAs2([string]$path,12)}
if(!$document.Saved -or $document.FullName -ne $path){throw 'Word did not save the owned test document at the expected path.'}
$document.Close(0)
$reopened=$word.Documents.Open([string]$path)
$reopened.Activate()
$word.Visible=$true
$target | Add-Member -NotePropertyName document -NotePropertyValue $reopened.Name -Force
$target | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $PSScriptRoot 'evidence\target-word.json') -Encoding UTF8
$record=@{time=[DateTime]::UtcNow.ToString('o');pid=$WordProcessId;path=$path;before=$before;after=@{name=$reopened.Name;saved=$reopened.Saved;maths=$reopened.OMaths.Count;shapes=$reopened.InlineShapes.Count;bookmarks=$reopened.Bookmarks.Count;fields=$reopened.Fields.Count;end=$reopened.Content.End}}
$record | ConvertTo-Json -Depth 5 | Set-Content (Join-Path $PSScriptRoot "evidence\$Label-save-reopen.json") -Encoding UTF8
$record | ConvertTo-Json -Depth 5
