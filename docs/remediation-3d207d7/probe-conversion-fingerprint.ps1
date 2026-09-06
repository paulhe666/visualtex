param([string]$Label='stage02c-fingerprint-diagnosis',[switch]$Semantic)
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$bin=Join-Path $root 'apps\windows\src-windows\VisualTeX.WordVsto\bin\x64\Release\net472'
$resolver=[ResolveEventHandler]{param($sender,$eventArgs) $path=Join-Path $bin (([Reflection.AssemblyName]$eventArgs.Name).Name+'.dll');if(Test-Path -LiteralPath $path){return [Reflection.Assembly]::LoadFrom($path)};return $null}
[AppDomain]::CurrentDomain.add_AssemblyResolve($resolver)
try {
 $assembly=[Reflection.Assembly]::LoadFrom((Join-Path $bin 'VisualTeX.WordVsto.dll'))
 $flags=[Reflection.BindingFlags]'Static,NonPublic,Public'
 $converter=$assembly.GetType('VisualTeX.WordVsto.WordOmmlConverter',$true)
 $transform=$converter.GetMethod('TransformMathMlToOmml',$flags)
 $methodName=if($Semantic){'ComputeImportedOmmlContentSignature'}else{'ComputeOmmlFingerprint'}
 $fingerprint=$converter.GetMethod($methodName,$flags)
 [xml]$observed=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'evidence\stage02c-vt-tail-after-__19.xml'))
 $manager=[Xml.XmlNamespaceManager]::new($observed.NameTable);$manager.AddNamespace('m','http://schemas.openxmlformats.org/officeDocument/2006/math')
 $actuals=@($observed.SelectNodes('//m:oMath',$manager))
 $rows=@();$index=0
 foreach($id in @('84049e39-d78b-4574-a100-e9d17d275793','e32e45ba-1fd0-4460-bc0b-070a10e14fa2')){
  $session=Get-Content -LiteralPath (Join-Path $PSScriptRoot "evidence\stage02c-session-$id.json") -Raw -Encoding UTF8|ConvertFrom-Json
  $source=[string]$transform.Invoke($null,[object[]]@($session.exportResult.mathMl))
  $actual=$actuals[$index].OuterXml
  [IO.File]::WriteAllText((Join-Path $PSScriptRoot "evidence\$Label-$index-source.xml"),$source)
  [IO.File]::WriteAllText((Join-Path $PSScriptRoot "evidence\$Label-$index-actual.xml"),$actual)
  $sourceHash=[string]$fingerprint.Invoke($null,[object[]]@($source));$actualHash=[string]$fingerprint.Invoke($null,[object[]]@($actual))
  $rows+=@{session=$id;source=$sourceHash;actual=$actualHash;equal=($sourceHash -ceq $actualHash)};$index++
 }
 $rows|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $PSScriptRoot "evidence\$Label.json") -Encoding UTF8
 $rows|ConvertTo-Json -Depth 5
}finally{[AppDomain]::CurrentDomain.remove_AssemblyResolve($resolver)}
