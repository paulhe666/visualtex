param([Parameter(Mandatory=$true)][int]$WordProcessId,[Parameter(Mandatory=$true)][string]$ExpectedDocument,[Parameter(Mandatory=$true)][string]$Needle,[Parameter(Mandatory=$true)][string]$Label,[int]$Zoom=120,[int]$ScrollLines=0)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$root=Join-Path $PSScriptRoot 'evidence\real-mathtype-20260926'
$a=Connect-RunningWord $WordProcessId
$d=$a.ActiveDocument
if($d.Name -ne $ExpectedDocument -or !$d.FullName.StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'Only owned evidence copies may be navigated'}
$r=$d.Content.Duplicate
$f=$r.Find
$f.ClearFormatting();$f.Text=$Needle;$f.Forward=$true;$f.Wrap=0;$f.MatchWildcards=$false
if(!$f.Execute()){throw ('Text not found: '+$Needle)}
$w=$d.ActiveWindow
$w.View.Zoom.Percentage=$Zoom
$r.Select();$w.ScrollIntoView($r,$true)
if($ScrollLines -ne 0){$w.SmallScroll($ScrollLines)}
$p=$r.Paragraphs.Item(1);$q=$p.Range.Duplicate
$xml=$q.WordOpenXML
[IO.File]::WriteAllText((Join-Path $root ($Label+'-paragraph.xml')),$xml,[Text.UTF8Encoding]::new($false))
$summary=[ordered]@{pid=$WordProcessId;name=$d.Name;found=$r.Text;start=$r.Start;end=$r.End;position=$r.Font.Position;size=$r.Font.Size;superscript=$r.Font.Superscript;paragraphText=$q.Text;paragraphAlignment=[int]$p.Alignment;baselineAlignment=$p.Format.BaseLineAlignment;lineRule=[int]$p.LineSpacingRule;line=$p.LineSpacing;shapes=@()}
foreach($s in $q.InlineShapes){
 $srange=$s.Range
 $summary.shapes+=@{start=$srange.Start;end=$srange.End;position=$srange.Font.Position;height=$s.Height;width=$s.Width;type=[int]$s.Type}
 [void][Runtime.InteropServices.Marshal]::ReleaseComObject($srange)
 [void][Runtime.InteropServices.Marshal]::ReleaseComObject($s)
}
[IO.File]::WriteAllText((Join-Path $root ($Label+'-paragraph.json')),($summary|ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($false))
[pscustomobject]$summary|Select-Object pid,name,found,start,end,position,size,superscript,baselineAlignment,lineRule,line|ConvertTo-Json
foreach($v in @($q,$p,$w,$f,$r,$d,$a)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($v)}
& (Join-Path $PSScriptRoot 'capture-real-word-window.ps1') -WordProcessId $WordProcessId -ExpectedDocument $ExpectedDocument -Label $Label
