param([Parameter(Mandatory=$true)][int]$WordProcessId)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$helper=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\docs\remediation-3d207d7\word-connection.ps1'))
. $helper
function Release($value){if($null-ne$value -and [Runtime.InteropServices.Marshal]::IsComObject($value)){try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($value)}catch{}}}
$word=Connect-RunningWord $WordProcessId
try{
  $documents=@()
  for($index=1;$index -le $word.Documents.Count;$index++){
    $document=$null;$content=$null
    try{
      $document=$word.Documents.Item($index)
      $content=$document.Content
      $prefix=[string]$content.Text
      if($prefix.Length -gt 120){$prefix=$prefix.Substring(0,120)}
      $documents+=@{
        index=$index
        name=[string]$document.Name
        fullName=[string]$document.FullName
        active=([string]$word.ActiveDocument.Name -eq [string]$document.Name)
        saved=[bool]$document.Saved
        saveFormat=[int]$document.SaveFormat
        compatibilityMode=[int]$document.CompatibilityMode
        paragraphs=[int]$document.Paragraphs.Count
        oMaths=[int]$document.OMaths.Count
        inlineShapes=[int]$document.InlineShapes.Count
        tables=[int]$document.Tables.Count
        bookmarks=[int]$document.Bookmarks.Count
        fields=[int]$document.Fields.Count
        contentEnd=[int]$content.End
        textPrefix=$prefix
      }
    }finally{Release $content;Release $document}
  }
  @{pid=$WordProcessId;documentCount=[int]$word.Documents.Count;activeDocument=[string]$word.ActiveDocument.Name;documents=$documents}|ConvertTo-Json -Depth 6 -Compress
}finally{Release $word}
