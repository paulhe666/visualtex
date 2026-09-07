param([int]$WordProcessId,[string]$Label='undo')
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$word=Connect-RunningWord $WordProcessId
$records=@()
foreach($id in @(128)) {
 $control=$word.CommandBars.Item('Standard').FindControl([Type]::Missing,$id)
 $row=[ordered]@{id=$id;found=($null -ne $control)}
 if($null -ne $control) {
  $row.caption=$control.Caption;$row.type=$control.Type;$row.enabled=$control.Enabled
  try{$row.listCount=$control.ListCount;$row.items=@();for($i=1;$i -le $control.ListCount;$i++){$row.items+=@{index=$i;text=$control.List($i)}}}catch{$row.readError=$_.Exception.Message}
 }
 $records+=$row
}
$standard=@();foreach($c in $word.CommandBars.Item('Standard').Controls){$standard+=@{id=$c.Id;type=$c.Type;caption=$c.Caption}}
$result=[ordered]@{pid=$WordProcessId;document=$word.ActiveDocument.Name;undoEnabled=$word.CommandBars.GetEnabledMso('Undo');recording=$word.UndoRecord.IsRecordingCustomRecord;level=$word.UndoRecord.CustomRecordLevel;controls=$records;standard=$standard}
$result|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $PSScriptRoot "evidence\$Label.json") -Encoding UTF8
$result|ConvertTo-Json -Depth 8
