param([Parameter(Mandatory=$true)][string]$OutputDirectory,[switch]$RenderExisting)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$root=[IO.Path]::GetFullPath($OutputDirectory)
[void][IO.Directory]::CreateDirectory($root)
$requestPath=Join-Path $root 'request.json'
$responsePath=Join-Path $root 'response.json'
if((Test-Path $requestPath) -and !$RenderExisting){throw 'Refusing to overwrite native probe evidence'}
$source=Join-Path $PSScriptRoot 'evidence\real-mathtype-20260926\codec-native-preview'
$items=@(Get-ChildItem $source -Filter '*.mtef' | ForEach-Object {@{Id=$_.BaseName;MtefPath=$_.FullName;WmfPath=(Join-Path $root ($_.BaseName+'.wmf'))}})
$exe=(Get-ItemProperty HKCU:\Software\VisualTeX\OfficeIntegration).ExecutablePath
$bridge=Join-Path (Split-Path $exe) 'visualtex-windows-office-bridge.exe'
if(!$RenderExisting){
 @{ResultPath=$responsePath;MathTypeServerPath='';Items=$items}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath $requestPath -Encoding UTF8
 $process=Start-Process -FilePath $bridge -ArgumentList @('--mathtype-preview-manifest',('"'+$requestPath+'"')) -PassThru -Wait
 Write-Output ('EXIT='+$process.ExitCode)
}
if(!(Test-Path $responsePath)){throw 'Native sidecar did not produce response'}
$response=[IO.File]::ReadAllText($responsePath)|ConvertFrom-Json
$response|ConvertTo-Json -Depth 5
Add-Type -AssemblyName System.Drawing
foreach($item in $response.Items){
 if(!$item.Success){continue}
 $metafile=[Drawing.Imaging.Metafile]::new([string]$item.WmfPath)
 $bmp=[Drawing.Bitmap]::new([Math]::Max(1,[int]($item.WidthPt*8)),[Math]::Max(1,[int]($item.HeightPt*8)))
 $g=[Drawing.Graphics]::FromImage($bmp)
 try {$g.Clear([Drawing.Color]::White);$g.DrawImage($metafile,[Drawing.Rectangle]::new(0,0,$bmp.Width,$bmp.Height));$bmp.Save((Join-Path $root ($item.Id+'.png')),[Drawing.Imaging.ImageFormat]::Png)}
 finally {$g.Dispose();$bmp.Dispose();$metafile.Dispose()}
}
