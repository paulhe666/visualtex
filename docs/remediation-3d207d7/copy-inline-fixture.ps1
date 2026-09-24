param(
 [Parameter(Mandatory=$true)][int]$WordProcessId,
 [Parameter(Mandatory=$true)][string]$DocumentName,
 [ValidateSet('prepare','position','inspect','save-owned','reopen-owned','open-owned')][string]$Action='inspect',
 [string]$Marker='VT_SOURCE',
 [string]$Label='copy-inline',
 [string]$Path=''
)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$evidence=Join-Path $PSScriptRoot 'evidence'
$owned=@(Get-ChildItem $evidence -Filter '*-word-process.json' | ForEach-Object {try {Get-Content $_.FullName -Raw | ConvertFrom-Json}catch{}} | Where-Object {$_.pid -eq $WordProcessId})
if(!$owned.Count -or $WordProcessId -in @(27664,47836)){throw 'This helper only operates on recorded test-owned Word processes.'}
$word=Connect-RunningWord $WordProcessId
$doc=$word.ActiveDocument
if($doc.Name -ne $DocumentName){throw "Expected $DocumentName; active is $($doc.Name)."}
if($Action -eq 'prepare'){
 if($doc.Content.End -ne 1){throw 'Refusing to overwrite a nonempty document.'}
 $doc.Content.Text="VT_SOURCE:  after VT`rOMML_SOURCE:  after OMML`rMT_SOURCE:  after MT`rVT_COPY:  tail VT`rOMML_COPY:  tail OMML`rMT_COPY:  tail MT`r"
 $doc.Content.Font.Size=14
 $doc.Content.Font.Name='Arial'
 $doc.Styles.Item(-1).Font.Size=10.5
 Write-Output "PREPARED|$DocumentName|end=$($doc.Content.End)"
}
if($Action -eq 'position' -or $Action -eq 'prepare'){
 $range=$doc.Content.Duplicate
 if(!$range.Find.Execute($Marker+': ')){throw "Missing fixture marker: $Marker"}
 $word.Selection.SetRange($range.End,$range.End)
 Write-Output "POSITION|$Marker|$($range.End)"
}
if($Action -eq 'save-owned'){
 if(!$Path -or !([IO.Path]::GetFullPath($Path).StartsWith([IO.Path]::GetFullPath($evidence),[StringComparison]::OrdinalIgnoreCase))){throw 'Test save must be inside evidence.'}
 $doc.SaveAs2($Path,12)
 Write-Output "SAVED_OWNED|$($doc.Name)|$Path"
}
if($Action -eq 'open-owned'){
 if($doc.Content.End -ne 1 -or !$doc.Saved){throw 'Open-fixture requires an untouched blank test document.'}
 if(!$Path -or !([IO.Path]::GetFullPath($Path).StartsWith([IO.Path]::GetFullPath($evidence)+[IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) -or !(Test-Path -LiteralPath $Path)){throw 'Only an existing test fixture inside evidence may be opened.'}
 # Keep the saved reference fixture unchanged and avoid another owned Word's file lock.
 $fixtureCopy=Join-Path $evidence ($Label+'-'+$WordProcessId+'.docx')
 [IO.File]::Copy($Path,$fixtureCopy,$false)
 $doc=$word.Documents.Open($fixtureCopy,$false,$false,$false)
 $doc.Activate()
 Write-Output "OPENED_OWNED|$($doc.Name)"
}
if($Action -eq 'reopen-owned'){
 if(!$Path -or $doc.FullName -ne $Path -or !$doc.Saved){throw 'Only the explicitly saved owned test document can be reopened.'}
 $doc.Close(0)
 $doc=$word.Documents.Open($Path,$false,$false,$false)
 $doc.Activate()
 Write-Output "REOPENED_OWNED|$($doc.Name)"
}
if($Action -eq 'inspect'){
 $shapes=@();$maths=@();$bookmarks=@();$paras=@()
 for($i=1;$i -le $doc.InlineShapes.Count;$i++){
  $s=$doc.InlineShapes.Item($i);$r=$s.Range
  $shapes+=@{index=$i;start=$r.Start;end=$r.End;prog=$s.OLEFormat.ProgID;width=$s.Width;height=$s.Height;position=$r.Font.Position;size=$r.Font.Size;alternative=$s.AlternativeText}
 }
 for($i=1;$i -le $doc.OMaths.Count;$i++){
  $m=$doc.OMaths.Item($i);$r=$m.Range
  $maths+=@{index=$i;start=$r.Start;end=$r.End;type=[int]$m.Type;text=$r.Text;font=$r.Font.Name;size=$r.Font.Size}
 }
 for($i=1;$i -le $doc.Bookmarks.Count;$i++){
  $b=$doc.Bookmarks.Item($i);$r=$b.Range
  $bookmarks+=@{name=$b.Name;start=$r.Start;end=$r.End}
 }
 for($i=1;$i -le $doc.Paragraphs.Count;$i++){
  $r=$doc.Paragraphs.Item($i).Range
  $paras+=@{start=$r.Start;end=$r.End;text=$r.Text;size=$r.Font.Size;alignment=[int]$r.ParagraphFormat.Alignment}
 }
 $report=@{pid=$WordProcessId;name=$doc.Name;end=$doc.Content.End;selection=@($word.Selection.Start,$word.Selection.End);shapes=$shapes;maths=$maths;bookmarks=$bookmarks;paragraphs=$paras;fields=$doc.Fields.Count;tables=$doc.Tables.Count}
 [IO.File]::WriteAllText((Join-Path $evidence ($Label+'-detail.json')),($report|ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false))
 $report|ConvertTo-Json -Depth 10
}
