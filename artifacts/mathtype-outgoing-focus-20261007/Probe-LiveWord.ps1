param(
 [string]$OutputPath = ''
)
$ErrorActionPreference='Stop'
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
. (Join-Path $repo 'docs\remediation-3d207d7\word-connection.ps1')
$rows=@()
foreach($process in @(Get-Process WINWORD -ErrorAction SilentlyContinue)){
 $word=$null
 $document=$null
 try{
  $word=Connect-RunningWord $process.Id
  $document=$word.ActiveDocument
  $rows+=[ordered]@{
   Pid=$process.Id
   StartedUtc=$process.StartTime.ToUniversalTime().ToString('o')
   Hwnd=[long]$document.ActiveWindow.Hwnd
   Name=$document.Name
   FullName=$document.FullName
   Saved=[bool]$document.Saved
   InlineShapes=$document.InlineShapes.Count
   OMaths=$document.OMaths.Count
   Visible=[bool]$word.Visible
  }
 }catch{
  $rows+=[ordered]@{
   Pid=$process.Id
   StartedUtc=$process.StartTime.ToUniversalTime().ToString('o')
   Error=$_.Exception.Message
  }
 }finally{
  if($null -ne $document -and [Runtime.InteropServices.Marshal]::IsComObject($document)){
   [void][Runtime.InteropServices.Marshal]::ReleaseComObject($document)
  }
  if($null -ne $word -and [Runtime.InteropServices.Marshal]::IsComObject($word)){
   [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  }
 }
}
$json=$rows|ConvertTo-Json -Depth 4
if($OutputPath){[IO.File]::WriteAllText($OutputPath,$json,[Text.UTF8Encoding]::new($false))}
$json
