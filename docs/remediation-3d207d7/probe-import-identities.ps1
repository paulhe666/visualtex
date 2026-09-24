param([string]$Stage='stage04o',[string]$Label='identity-diagnosis')
$ErrorActionPreference='Stop'
# Pure data diagnosis of already captured failure evidence, never a Word write.
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$bin=Join-Path $root 'apps\windows\src-windows\VisualTeX.WordVsto\bin\x64\Release\net472'
$handler=[ResolveEventHandler]{param($sender,$eventArgs) $p=Join-Path $bin (([Reflection.AssemblyName]$eventArgs.Name).Name+'.dll');if(Test-Path $p){return [Reflection.Assembly]::LoadFrom($p)};return $null}
[AppDomain]::CurrentDomain.add_AssemblyResolve($handler)
try {
 $asm=[Reflection.Assembly]::LoadFrom((Join-Path $bin 'VisualTeX.WordVsto.dll'))
 $method=$asm.GetType('VisualTeX.WordVsto.WordOmmlConverter',$true).GetMethod('ComputeImportedOmmlContentSignature',[Reflection.BindingFlags]'Static,NonPublic')
 $evidence=Join-Path $PSScriptRoot 'evidence'
 [xml]$body=Get-Content -LiteralPath (Join-Path $evidence "$Stage-word-hook.log.identity-body.xml") -Raw -Encoding UTF8
 $ns=[Xml.XmlNamespaceManager]::new($body.NameTable);$ns.AddNamespace('m','http://schemas.openxmlformats.org/officeDocument/2006/math')
 $maths=@($body.SelectNodes('//m:oMath',$ns));$actual=@($maths | ForEach-Object {$method.Invoke($null,[object[]]@($_.OuterXml))})
 $rows=@()
 foreach($file in Get-ChildItem -LiteralPath $evidence -Filter "$Stage-word-hook.log.identity-*.xml") {
  if($file.Name.EndsWith('identity-body.xml')){continue}
  $expected=[string]$method.Invoke($null,[object[]]@([IO.File]::ReadAllText($file.FullName)))
  $matches=@();for($i=0;$i -lt $actual.Count;$i++){if($actual[$i] -ceq $expected){$matches+=($i+1)}}
  $rows+=@{source=$file.Name;signature=$expected;matchingEquationIndices=$matches}
 }
 $report=@{dllSha256=(Get-FileHash (Join-Path $bin 'VisualTeX.WordVsto.dll')).Hash;scope='read-only data diagnosis, not Word acceptance';rows=$rows}
 $report|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $evidence "$Label.json") -Encoding UTF8
 $rows|ConvertTo-Json -Depth 4
}finally{[AppDomain]::CurrentDomain.remove_AssemblyResolve($handler)}
