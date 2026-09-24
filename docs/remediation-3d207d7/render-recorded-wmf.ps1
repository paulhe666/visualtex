param([string]$XmlEvidence,[string]$WmfManifest,[Parameter(Mandatory=$true)][string]$Label)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Drawing
# Supplementary rendering of previews already captured from real Word. No Word connection.
$paths=@()
if($WmfManifest) {
 $manifest=Get-Content -LiteralPath (Join-Path $PSScriptRoot $WmfManifest) -Raw -Encoding UTF8 | ConvertFrom-Json
 $paths=@($manifest.Items | ForEach-Object {$_.WmfPath})
} else {
[xml]$xml=[IO.File]::ReadAllText((Join-Path $PSScriptRoot $XmlEvidence),[Text.Encoding]::UTF8)
$ns=[Xml.XmlNamespaceManager]::new($xml.NameTable)
$ns.AddNamespace('pkg','http://schemas.microsoft.com/office/2006/xmlPackage')
$index=0
foreach($part in $xml.SelectNodes('//pkg:part[pkg:binaryData]',$ns)) {
 $name=$part.GetAttribute('name','http://schemas.microsoft.com/office/2006/xmlPackage')
 if($name -notmatch '\.wmf$'){continue}
 $index++
 $path=Join-Path $PSScriptRoot "evidence\$Label-$index.wmf"
 [IO.File]::WriteAllBytes($path,[Convert]::FromBase64String($part.SelectSingleNode('pkg:binaryData',$ns).InnerText))
 $paths+=,$path
}
}
$index=0
foreach($path in $paths) {
 $index++
 $metafile=[Drawing.Imaging.Metafile]::new($path)
 try {
  $bitmap=[Drawing.Bitmap]::new(1500,[int][Math]::Ceiling(1500*$metafile.Height/$metafile.Width))
  try {
   $graphics=[Drawing.Graphics]::FromImage($bitmap)
   try {$graphics.Clear([Drawing.Color]::White);$graphics.DrawImage($metafile,[Drawing.Rectangle]::new(0,0,$bitmap.Width,$bitmap.Height))}
   finally {$graphics.Dispose()}
   $bitmap.Save((Join-Path $PSScriptRoot "evidence\$Label-$index.png"),[Drawing.Imaging.ImageFormat]::Png)
   Write-Output "$path $($metafile.Width)x$($metafile.Height)"
  } finally {$bitmap.Dispose()}
 } finally {$metafile.Dispose()}
}
