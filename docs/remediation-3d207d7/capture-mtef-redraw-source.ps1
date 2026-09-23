param(
 [Parameter(Mandatory=$true)][int]$WordProcessId,
 [Parameter(Mandatory=$true)][string]$ExpectedDocument,
 [Parameter(Mandatory=$true)][string]$OutputPath
)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$word=Connect-RunningWord $WordProcessId
$document=$word.ActiveDocument
if($document.Name -ne $ExpectedDocument){throw "Unexpected active document: $($document.Name)"}
$text=[string]$document.Content.Text
$xml=[string]$document.Content.WordOpenXML
$sha=[Security.Cryptography.SHA256]::Create()
function HashText([string]$value) { ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($value)))).Replace('-','') }
$matches=[regex]::Matches($text,'(?s)\$\$(?<block>.*?)\$\$|(?<!\\)\$(?<inline>[^$]*?)\$')
$items=@();$i=0
foreach($match in $matches){
 $block=$match.Groups['block'].Success
 $latex=if($block){$match.Groups['block'].Value}else{$match.Groups['inline'].Value}
 $items += [pscustomobject]@{index=++$i;start=$match.Index;length=$match.Length;displayMode=if($block){'block'}else{'inline'};latex=$latex}
}
$record=[pscustomobject]@{
 capturedAt=[DateTimeOffset]::Now.ToString('o');pid=$WordProcessId;name=$document.Name;fullName=$document.FullName
 saved=$document.Saved;paragraphs=$document.Paragraphs.Count;omaths=$document.OMaths.Count;inlineShapes=$document.InlineShapes.Count
 selectionStart=$word.Selection.Start;selectionEnd=$word.Selection.End;textHash=(HashText $text);xmlHash=(HashText $xml)
 sourceText=$text;formulas=$items
}
$full=[IO.Path]::GetFullPath($OutputPath)
if([IO.File]::Exists($full)){throw "Refusing to overwrite evidence: $full"}
[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($full))
[IO.File]::WriteAllText($full,($record|ConvertTo-Json -Depth 8),[Text.UTF8Encoding]::new($false))
$record|Select-Object capturedAt,pid,name,saved,paragraphs,omaths,inlineShapes,textHash,xmlHash|ConvertTo-Json
Write-Output ('FORMULAS='+$items.Count+';INLINE='+@($items|Where-Object displayMode -eq 'inline').Count+';BLOCK='+@($items|Where-Object displayMode -eq 'block').Count)
$sha.Dispose()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($document)
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
