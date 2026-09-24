param([int]$WordProcessId,[string]$DocumentName,[int]$Index=12,[string]$Label='omml-height',[switch]$Show)
$ErrorActionPreference='Stop';[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$word=Connect-RunningWord $WordProcessId;$doc=$word.Documents.Item($DocumentName)
if($doc.Name -ne $DocumentName){throw 'Document mismatch.'}
$math=$doc.OMaths.Item($Index);$range=$math.Range;$format=$range.ParagraphFormat;$font=$range.Font
$data=@{pid=$WordProcessId;document=$doc.Name;index=$Index;mathCount=$doc.OMaths.Count;oleCount=$doc.InlineShapes.Count;type=[int]$math.Type;start=$range.Start;end=$range.End;text=$range.Text;font=$font.Name;size=$font.Size;fontPosition=$font.Position;lineRule=[int]$format.LineSpacingRule;lineSpacing=$format.LineSpacing;disableGrid=$format.DisableLineHeightGrid;spaceBefore=$format.SpaceBefore;spaceAfter=$format.SpaceAfter}
if($range.Tables.Count -gt 0){
 $table=$range.Tables.Item(1);$cell=$range.Cells.Item(1);$row=$table.Rows.Item($cell.RowIndex);$cellrange=$cell.Range;$cf=$cellrange.ParagraphFormat
 $data.table=@{rows=$table.Rows.Count;columns=$table.Columns.Count;heightRule=[int]$row.HeightRule;height=$row.Height;topPadding=$table.TopPadding;bottomPadding=$table.BottomPadding;cellTop=$cell.TopPadding;cellBottom=$cell.BottomPadding;cellVertical=[int]$cell.VerticalAlignment;cellLineRule=[int]$cf.LineSpacingRule;cellLineSpacing=$cf.LineSpacing;cellDisableGrid=$cf.DisableLineHeightGrid}
}
$window=$doc.ActiveWindow
if($Show){$doc.Activate();$window.ScrollIntoView($range,$true);Start-Sleep -Milliseconds 400}
try{$x=0;$y=0;$width=0;$height=0;$window.GetPoint([ref]$x,[ref]$y,[ref]$width,[ref]$height,$range);$data.rectangle=@{x=$x;y=$y;width=$width;height=$height;zoom=$window.View.Zoom.Percentage}}catch{$data.rectangleError=$_.Exception.Message}
$xml=$range.WordOpenXML
[IO.File]::WriteAllText((Join-Path $PSScriptRoot "evidence\$Label.xml"),$xml,[Text.UTF8Encoding]::new($false))
[IO.File]::WriteAllText((Join-Path $PSScriptRoot "evidence\$Label.json"),($data|ConvertTo-Json -Depth 5),[Text.UTF8Encoding]::new($false))
$data|ConvertTo-Json -Depth 5
foreach($obj in @($cf,$cellrange,$row,$cell,$table,$window,$font,$format,$range,$math,$doc,$word)){if($obj -and [Runtime.InteropServices.Marshal]::IsComObject($obj)){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj)}}
