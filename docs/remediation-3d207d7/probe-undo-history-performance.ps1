param([int]$WordProcessId,[string]$DocumentName,[string]$Label='undo-probe')
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$word=Connect-RunningWord $WordProcessId
$doc=$word.ActiveDocument
if($doc.Name -ne $DocumentName){throw 'Wrong document; read stopped.'}
$rows=@();$watch=[Diagnostics.Stopwatch]::StartNew();$checkpoint=0.0
function Mark([string]$stage,$value){$elapsed=$watch.Elapsed.TotalMilliseconds;$script:rows+=@{stage=$stage;deltaMs=$elapsed-$script:checkpoint;value=$value};$script:checkpoint=$elapsed}
$bars=$word.CommandBars;Mark 'bars' ''
$standard=$bars.Item('Standard');Mark 'standard' ''
$control=$standard.FindControl([Type]::Missing,128);Mark 'find-control' ''
$enabled=$bars.GetEnabledMso('Undo');Mark 'enabled-mso' $enabled
$historyEnabled=$control.Enabled;Mark 'history-enabled' $historyEnabled
$count=$control.ListCount;Mark 'list-count' $count
$records=@()
for($i=1;$i -le $count;$i++){$records+=$control.List($i);if($i -le 4 -or $i -eq $count){Mark "list-$i" $records[-1]}}
Mark 'complete' $records.Count
$report=@{label=$Label;pid=$WordProcessId;document=$DocumentName;count=$count;rows=$rows;records=$records}
[IO.File]::WriteAllText((Join-Path $PSScriptRoot "evidence\$Label.json"),($report|ConvertTo-Json -Depth 5),[Text.UTF8Encoding]::new($false))
$rows|ConvertTo-Json -Depth 4
foreach($v in @($control,$standard,$bars,$doc,$word)){if($null -ne $v -and [Runtime.InteropServices.Marshal]::IsComObject($v)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($v)}}
