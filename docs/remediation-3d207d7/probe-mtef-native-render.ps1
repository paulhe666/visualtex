param(
 [string]$CorpusPath='docs/remediation-3d207d7/evidence/mtef-redraw-20260923/corpus.json',
 [Parameter(Mandatory=$true)][string]$OutputDirectory,
 [switch]$AllExportable,
 [string]$AssemblyPath='apps/windows/src-windows/VisualTeX.WordVsto/bin/x64/Release/net472/VisualTeX.WordVsto.dll'
)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$root=[IO.Path]::GetFullPath($OutputDirectory)
if(Test-Path $root){throw 'Refusing to overwrite native-render evidence'}
[void][IO.Directory]::CreateDirectory($root)
$assembly=[Reflection.Assembly]::LoadFrom([IO.Path]::GetFullPath($AssemblyPath))
$flags=[Reflection.BindingFlags]'Static,Public,NonPublic'
$codec=$assembly.GetType('VisualTeX.WordVsto.MathTypeMtefCodec',$true)
$create=$codec.GetMethod('CreateEquationNativeAtFontSize',$flags)
$interop=$assembly.GetType('VisualTeX.WordVsto.MathTypeOleInterop',$true)
$server=$interop.GetMethod('ResolveInstalledServerPath',$flags).Invoke($null,@())
$bridge=Join-Path $env:LOCALAPPDATA 'VisualTeX/visualtex-windows-office-bridge.exe'
$corpus=Get-Content -LiteralPath $CorpusPath -Raw -Encoding UTF8 | ConvertFrom-Json
$items=@()
foreach($item in $corpus.formulas | Where-Object {$_.mathMl -and ($AllExportable -or $_.index -in @(1,15,27,36,42,45,207,219,227))}){
 $generated=$create.Invoke($null,[object[]]@([string]$item.mathMl,[bool]($item.displayMode -eq 'inline'),[double]12))
 $mtef=Join-Path $root ('formula-'+$item.index+'.mtef')
 [IO.File]::WriteAllBytes($mtef,[byte[]]$generated.Mtef)
 $items+=@{Id=[string]$item.index;MtefPath=$mtef;WmfPath=(Join-Path $root ('formula-'+$item.index+'.wmf'))}
}
$responsePath=Join-Path $root 'response.json'
$manifest=Join-Path $root 'request.json'
@{ResultPath=$responsePath;MathTypeServerPath=[string]$server;Items=$items}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath $manifest -Encoding UTF8
$start=[Diagnostics.ProcessStartInfo]::new($bridge,('--mathtype-preview-manifest "'+$manifest+'"'))
$start.UseShellExecute=$false;$start.CreateNoWindow=$true
$process=[Diagnostics.Process]::Start($start)
if(!$process.WaitForExit(120000)){ $process.Kill();throw 'Native sidecar timed out; only the sidecar started by this test was stopped.' }
Write-Output ('BRIDGE_EXIT='+$process.ExitCode+';SERVER='+$server+';BRIDGE='+$bridge)
$process.Dispose()
if(!(Test-Path $responsePath)){throw 'Sidecar returned no response file'}
$response=Get-Content $responsePath -Raw -Encoding UTF8|ConvertFrom-Json
$response|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $root 'render-results.json') -Encoding UTF8
@{Requested=$items.Count;Succeeded=@($response.Items|Where-Object Success).Count;Error=$response.Error;Failed=@($response.Items|Where-Object {!$_.Success});AssemblyHash=(Get-FileHash $assembly.Location -Algorithm SHA256).Hash}|ConvertTo-Json -Depth 5
Add-Type -AssemblyName System.Drawing
foreach($item in $response.Items | Where-Object {$_.Success -and (!$AllExportable -or [int]$_.Id -le 50 -or [int]$_.Id -in @(207,219,227))}){
 $metafile=[Drawing.Imaging.Metafile]::new([string]$item.WmfPath)
 $bitmap=[Drawing.Bitmap]::new([Math]::Max(1,[int]($item.WidthPt*6)),[Math]::Max(1,[int]($item.HeightPt*6)))
 $graphics=[Drawing.Graphics]::FromImage($bitmap)
 try {
  $graphics.Clear([Drawing.Color]::White)
  $graphics.DrawImage($metafile,[Drawing.Rectangle]::new(0,0,$bitmap.Width,$bitmap.Height))
  $bitmap.Save((Join-Path $root ('formula-'+$item.Id+'.png')),[Drawing.Imaging.ImageFormat]::Png)
 } finally {$graphics.Dispose();$bitmap.Dispose();$metafile.Dispose()}
}
