param([Parameter(Mandatory=$true)][int]$WordProcessId,[Parameter(Mandatory=$true)][string]$Stage,[Parameter(Mandatory=$true)][ValidateSet(20,50,100)][int]$Size)
$ErrorActionPreference='Stop';[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$r=Get-Content (Join-Path $PSScriptRoot "evidence\$Stage-word-process.json") -Raw -Encoding UTF8|ConvertFrom-Json
if($r.pid -ne $WordProcessId -or $WordProcessId -in @(27664,107976)){throw 'Only the explicitly owned UI-test process is permitted.'}
$p=Get-Process -Id $WordProcessId
if([Math]::Abs(($p.StartTime.ToUniversalTime()-[DateTimeOffset]::Parse($r.started).UtcDateTime).TotalSeconds)-gt 10){throw 'Test PID was reused.'}
$word=Connect-RunningWord $WordProcessId;$document=$null
try {
 $document=$word.Documents.Add();$document.Activate()
 $name=$document.Name
 if($document.Content.End -ne 1 -or $document.OMaths.Count -ne 0 -or $document.InlineShapes.Count -ne 0){throw 'New document is not blank.'}
 $target=@{pid=$WordProcessId;stage=$Stage;document=$name;size=$Size;started=$r.started;source="$Stage-installed.json"}
 $json=$target|ConvertTo-Json
 [IO.File]::WriteAllText((Join-Path $PSScriptRoot "evidence\$Stage-size-$Size-target.json"),$json,[Text.UTF8Encoding]::new($false))
 $json
}finally{if($document){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($document)};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)}
