param(
 [Parameter(Mandatory=$true)][string]$CorpusPath,
 [Parameter(Mandatory=$true)][string]$OutputDirectory,
 [string]$AssemblyPath='apps/windows/src-windows/VisualTeX.WordVsto/bin/x64/Release/net472/VisualTeX.WordVsto.dll',
 [int[]]$NativePreviewIndices=@()
)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$root=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path (Join-Path $root 'results.json')){throw 'Refusing to overwrite probe evidence'}
[void][IO.Directory]::CreateDirectory($root)
$assembly=[Reflection.Assembly]::LoadFrom([IO.Path]::GetFullPath($AssemblyPath))
$flags=[Reflection.BindingFlags]'Static,Public,NonPublic'
$codec=$assembly.GetType('VisualTeX.WordVsto.MathTypeMtefCodec',$true)
$create=$codec.GetMethod('CreateEquationNativeAtFontSize',$flags)
$decode=$codec.GetMethod('ReadEquationNativeMathMl',$flags)
$signature=$codec.GetMethod('SemanticSignature',$flags)
$service=$assembly.GetType('VisualTeX.WordVsto.WordFormulaService',$true)
$prepare=$service.GetMethod('PrepareStandaloneMathTypeOleData',$flags)
$renderer=$assembly.GetType('VisualTeX.WordVsto.MathTypeNativePreviewRenderer',$true)
$render=$renderer.GetMethod('TryRender',$flags)
$corpus=Get-Content -LiteralPath $CorpusPath -Raw -Encoding UTF8 | ConvertFrom-Json
$sha=[Security.Cryptography.SHA256]::Create()
$results=@()
foreach($item in $corpus.formulas){
 $record=[ordered]@{index=$item.index;source=$item.source;latex=$item.latex;displayMode=$item.displayMode;passed=$false}
 try {
  if(!$item.mathMl){throw "Frontend export failed: $($item.exportError)"}
  $inline=$item.displayMode -eq 'inline'
  $generated=$create.Invoke($null,[object[]]@([string]$item.mathMl,[bool]$inline,[double]12))
  $native=[byte[]]$generated.EquationNative
  $actual=[string]$decode.Invoke($null,[object[]]@(,$native))
  $expectedSignature=[string]$signature.Invoke($null,[object[]]@([string]$item.mathMl))
  $actualSignature=[string]$signature.Invoke($null,[object[]]@($actual))
  $record.expectedSignature=$expectedSignature
  $record.actualSignature=$actualSignature
  $record.expectedMathMl=$item.mathMl
  $record.actualMathMl=$actual
  $record.equationNativeHash=([BitConverter]::ToString($sha.ComputeHash($native))).Replace('-','')
  $record.signatureMatches=$expectedSignature -ceq $actualSignature
  try {
   $null=$prepare.Invoke($null,[object[]]@([string]$item.mathMl,[bool]$inline,[double]12,[string]$item.latex))
   $record.passed=$true
  } catch { $record.error=$_.Exception.GetBaseException().Message }
  if($NativePreviewIndices -contains [int]$item.index){
   [IO.File]::WriteAllBytes((Join-Path $root ('formula-'+$item.index+'.mtef')),[byte[]]$generated.Mtef)
   $arguments=[object[]]@([byte[]]$generated.Mtef,$root,$null)
   $record.nativeRenderSucceeded=[bool]$render.Invoke($null,$arguments)
   if($record.nativeRenderSucceeded){
    $preview=$arguments[2]
    $record.nativeWidthPt=$preview.WidthPt;$record.nativeHeightPt=$preview.HeightPt
    Add-Type -AssemblyName System.Drawing
    $metafile=[Drawing.Imaging.Metafile]::new([string]$preview.WmfPath)
    try {
     $bitmap=[Drawing.Bitmap]::new([Math]::Max(1,[int]($preview.WidthPt*6)),[Math]::Max(1,[int]($preview.HeightPt*6)))
     $graphics=[Drawing.Graphics]::FromImage($bitmap)
     try {
      $graphics.Clear([Drawing.Color]::White)
      $graphics.DrawImage($metafile,[Drawing.Rectangle]::new(0,0,$bitmap.Width,$bitmap.Height))
      $record.nativePng=Join-Path $root ('formula-'+$item.index+'-native.png')
      $bitmap.Save($record.nativePng,[Drawing.Imaging.ImageFormat]::Png)
     } finally { $graphics.Dispose();$bitmap.Dispose() }
    } finally { $metafile.Dispose() }
   }
  }
 } catch { $record.error=$_.Exception.GetBaseException().Message }
 $results += [pscustomobject]$record
}
$summary=[pscustomobject]@{
 assemblyPath=$assembly.Location;assemblyHash=(Get-FileHash $assembly.Location -Algorithm SHA256).Hash
 document=$corpus.document;total=$results.Count;passed=@($results|Where-Object passed).Count
 documentTotal=@($results|Where-Object source -eq 'active-document').Count
 documentPassed=@($results|Where-Object {$_.source -eq 'active-document' -and $_.passed}).Count
 results=$results
}
[IO.File]::WriteAllText((Join-Path $root 'results.json'),($summary|ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false))
$summary|Select-Object assemblyPath,assemblyHash,document,total,passed,documentTotal,documentPassed|ConvertTo-Json
$results|Where-Object {!$_.passed}|Select-Object index,source,latex,error|ConvertTo-Json -Depth 3
$results|Where-Object {$_.PSObject.Properties.Name -contains 'nativeRenderSucceeded'}|Select-Object index,nativeRenderSucceeded,nativeWidthPt,nativeHeightPt,nativePng|ConvertTo-Json
$sha.Dispose()
