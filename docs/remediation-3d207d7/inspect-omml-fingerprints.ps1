param([int]$WordProcessId,[string]$DocumentName,[string]$Label)
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'word-connection.ps1')
$word=Connect-RunningWord $WordProcessId
$document=$word.Documents.Item($DocumentName)
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$bin=Join-Path $root 'apps\windows\src-windows\VisualTeX.WordVsto\bin\x64\Release\net472'
$handler=[ResolveEventHandler]{param($sender,$eventArgs) $path=Join-Path $bin (([Reflection.AssemblyName]$eventArgs.Name).Name+'.dll');if(Test-Path $path){return [Reflection.Assembly]::LoadFrom($path)};return $null}
[AppDomain]::CurrentDomain.add_AssemblyResolve($handler)
try {
 $assembly=[Reflection.Assembly]::LoadFrom((Join-Path $bin 'VisualTeX.WordVsto.dll'))
 $flags=[Reflection.BindingFlags]'Static,NonPublic'
 $decode=$assembly.GetType('VisualTeX.WordVsto.WordOmmlFormulaStore').GetMethod('TryDecodePartXml',$flags)
 $fingerprint=$assembly.GetType('VisualTeX.WordVsto.WordOmmlConverter').GetMethod('ComputeOmmlFingerprint',$flags)
 $metadata=@()
 foreach($part in $document.CustomXMLParts){
  $arguments=[object[]]@([string]$part.XML,$null)
  if($decode.Invoke($null,$arguments)){$metadata+=@{id=$arguments[1].FormulaId;latex=$arguments[1].Latex;stored=$arguments[1].NativeOmmlFingerprint}}
 }
 # Pure-data functions above only decode recorded XML; no formula service is
 # constructed or invoked. This is a post-failure diagnostic, not acceptance.
 [xml]$xml=$document.Content.WordOpenXML
 $ns=[Xml.XmlNamespaceManager]::new($xml.NameTable);$ns.AddNamespace('m','http://schemas.openxmlformats.org/officeDocument/2006/math')
 $maths=@($xml.SelectNodes('//m:oMath',$ns));$actual=@()
 for($index=0;$index -lt $maths.Count;$index++){
  $hash=[string]$fingerprint.Invoke($null,[object[]]@($maths[$index].OuterXml))
  $actual+=@{index=$index+1;fingerprint=$hash;text=$document.OMaths.Item($index+1).Range.Text;matchingStored=@($metadata|Where-Object {$_.stored -ceq $hash}|ForEach-Object {$_.id})}
 }
 $report=@{scope='read-only real COM plus pure-data diagnostic; not acceptance';document=$document.Name;pid=$WordProcessId;dllSha256=(Get-FileHash (Join-Path $bin 'VisualTeX.WordVsto.dll')).Hash;metadata=$metadata;actual=$actual}
 $report|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $PSScriptRoot ('evidence\'+$Label+'.json')) -Encoding UTF8
 $report|ConvertTo-Json -Depth 8
}finally{[AppDomain]::CurrentDomain.remove_AssemblyResolve($handler)}
