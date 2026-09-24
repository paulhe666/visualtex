param(
  [Parameter(Mandatory=$true)][int]$WordProcessId,
  [Parameter(Mandatory=$true)][string]$DocumentName
)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
function Release($o){if($null-ne$o -and [Runtime.InteropServices.Marshal]::IsComObject($o)){try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($o)}catch{}}}
$word=Connect-RunningWord $WordProcessId
$document=$null
try{
  $document=$word.Documents.Item($DocumentName)
  $omml=$document.OMaths.Count
  $tables=$document.Tables.Count
  $fields=$document.Fields.Count
  $oneByThree=0
  $healthyCenters=0
  for($i=1;$i-le$tables;$i++){
    $table=$null;$cell=$null;$range=$null
    try{
      $table=$document.Tables.Item($i)
      if($table.Rows.Count-eq 1 -and $table.Columns.Count-eq 3){
        $oneByThree++
        $cell=$table.Cell(1,2)
        $range=$cell.Range
        if($range.OMaths.Count-eq 1){$healthyCenters++}
      }
    }finally{Release $range;Release $cell;Release $table}
  }
  $seq=0
  for($i=1;$i-le$fields;$i++){
    $field=$null;$code=$null
    try{
      $field=$document.Fields.Item($i)
      if($field.Type-eq [Microsoft.Office.Interop.Word.WdFieldType]::wdFieldEmbed){continue}
      $code=$field.Code
      $text=[string]$code.Text
      if($text-match '(?i)\bSEQ\s+VisualTeXEquation\b'){$seq++}
    }finally{Release $code;Release $field}
  }
  $vtomml=0;$vteq=0;$vteqnum=0;$vteqcap=0
  $bookmarks=$document.Bookmarks
  try{
    for($i=1;$i-le$bookmarks.Count;$i++){
      $bookmark=$null
      try{
        $bookmark=$bookmarks.Item($i)
        $name=[string]$bookmark.Name
        if($name-match '^VTOMML_'){$vtomml++}
        elseif($name-match '^VTEqNum_'){$vteqnum++}
        elseif($name-match '^VTEqCap_'){$vteqcap++}
        elseif($name-match '^VTEq_'){$vteq++}
      }finally{Release $bookmark}
    }
  }finally{Release $bookmarks}
  Write-Output ('REDRAW_SUMMARY|document='+$document.Name+'|omml='+$omml+'|tables='+$tables+'|oneByThree='+$oneByThree+'|healthyCenters='+$healthyCenters+'|fields='+$fields+'|seq='+$seq+'|vtomml='+$vtomml+'|vteq='+$vteq+'|vteqnum='+$vteqnum+'|vteqcap='+$vteqcap+'|end='+$document.Content.End)
}finally{Release $document;Release $word}
