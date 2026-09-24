param([int]$WordProcessId=63624,[string]$Label='stage02f-view-export')
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'word-connection.ps1')
if($WordProcessId -eq 27664){throw 'Never change an original document view.'}
$word=Connect-RunningWord $WordProcessId
$doc=$word.ActiveDocument
$window=$word.ActiveWindow
$view=$window.View.Type;$screen=$word.ScreenUpdating;$pagination=$word.Options.Pagination
$results=@()
try {
 foreach($mode in @(@{name='original';view=$view;screen=$screen;pagination=$pagination},@{name='draft-suspended';view=1;screen=$false;pagination=$false},@{name='print-suspended';view=3;screen=$false;pagination=$false},@{name='original-restored';view=$view;screen=$screen;pagination=$pagination})) {
  $window.View.Type=$mode.view;$word.ScreenUpdating=$mode.screen;$word.Options.Pagination=$mode.pagination
  $range=$doc.Content
  $xml=[string]$range.WordOpenXML
  [IO.File]::WriteAllText((Join-Path $PSScriptRoot "evidence\$Label-$($mode.name).xml"),$xml)
  [xml]$package=$xml;$ns=[Xml.XmlNamespaceManager]::new($package.NameTable);$ns.AddNamespace('m','http://schemas.openxmlformats.org/officeDocument/2006/math');$ns.AddNamespace('o','urn:schemas-microsoft-com:office:office')
  $results+=@{mode=$mode.name;document=$doc.Name;start=$range.Start;end=$range.End;maths=$doc.OMaths.Count;shapes=$doc.InlineShapes.Count;xmlMaths=$package.SelectNodes('//m:oMath',$ns).Count;xmlOle=$package.SelectNodes('//o:OLEObject',$ns).Count;view=$window.View.Type}
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($range)
 }
 $results|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $PSScriptRoot "evidence\$Label.json") -Encoding UTF8
 $results|ConvertTo-Json -Depth 5
} finally {
 $window.View.Type=$view;$word.ScreenUpdating=$screen;$word.Options.Pagination=$pagination
 foreach($obj in @($window,$doc,$word)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj)}
}
