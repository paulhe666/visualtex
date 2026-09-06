param([Parameter(Mandatory=$true)][string]$XmlEvidence,[Parameter(Mandatory=$true)][string]$Label,[switch]$IncludeRecords)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
# Pure data diagnosis after real Word insertion/conversion. No Word connection,
# no service instance, and no mutation of a Word document or the saved XML.
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$bin=Join-Path $root 'apps\windows\src-windows\VisualTeX.WordVsto\bin\x64\Release\net472'
$handler=[ResolveEventHandler]{param($sender,$eventArgs) $p=Join-Path $bin (([Reflection.AssemblyName]$eventArgs.Name).Name+'.dll');if(Test-Path $p){return [Reflection.Assembly]::LoadFrom($p)};return $null}
[AppDomain]::CurrentDomain.add_AssemblyResolve($handler)
try {
 $asm=[Reflection.Assembly]::LoadFrom((Join-Path $bin 'VisualTeX.WordVsto.dll'))
 $flags=[Reflection.BindingFlags]'Static,NonPublic,Public'
 $storage=$asm.GetType('VisualTeX.WordVsto.MathTypeOleStorage',$true)
 $codec=$asm.GetType('VisualTeX.WordVsto.MathTypeMtefCodec',$true)
 $read=$storage.GetMethod('ReadEquationNative',$flags)
 $options=$codec.GetMethod('FindEquationOptionsOffset',$flags)
 $rootOffset=$codec.GetMethod('FindRootStructureOffset',$flags)
 [xml]$xml=[IO.File]::ReadAllText((Join-Path $PSScriptRoot $XmlEvidence),[Text.Encoding]::UTF8)
 $ns=[Xml.XmlNamespaceManager]::new($xml.NameTable)
 $ns.AddNamespace('pkg','http://schemas.microsoft.com/office/2006/xmlPackage')
 $rows=@()
 foreach($part in $xml.SelectNodes('//pkg:part[pkg:binaryData]',$ns)) {
  $name=$part.GetAttribute('name','http://schemas.microsoft.com/office/2006/xmlPackage')
  if($name -notmatch '/embeddings/'){continue}
  $compound=[Convert]::FromBase64String($part.SelectSingleNode('pkg:binaryData',$ns).InnerText)
  $native=[byte[]]$read.Invoke($null,[object[]]@(,$compound))
  $header=[BitConverter]::ToUInt16($native,0);$length=[BitConverter]::ToUInt32($native,8)
  $mtef=[byte[]]::new($length);[Array]::Copy($native,$header,$mtef,0,$length)
  $cursor=1+[int]$options.Invoke($null,[object[]]@(,$mtef))
  $end=[int]$rootOffset.Invoke($null,[object[]]@(,$mtef));$preferences=@();$initialSizes=@()
  while($cursor -lt $end) {
   $tag=[int]$mtef[$cursor]
   if($tag -eq 18) {
    $count=[int]$mtef[$cursor+2];$hex=([BitConverter]::ToString($mtef,($cursor+3),($end-$cursor-3))).Replace('-','');$at=0;$values=@()
    for($i=0;$i -lt $count;$i++){$unit=[Convert]::ToInt32($hex.Substring($at++,1),16);$digits='';while($hex[$at] -ne 'F'){$ch=$hex[$at++];if($ch -eq 'A'){$digits+='.'}elseif($ch -eq 'B'){$digits+='-'}elseif($ch -match '[0-9]'){$digits+=$ch}else{throw 'Invalid dimension digit'}};$at++;$values+=@{unit=$unit;value=$digits}}
    $preferences+=@{offset=$cursor;sizes=$values}
   }
   if($tag -ge 9 -and $tag -le 14){$initialSizes+=@{offset=$cursor;tag=$tag;bytes=[BitConverter]::ToString($mtef,$cursor,([Math]::Min(5,$end-$cursor)))}}
   switch($tag) {
    19 {$cursor++;while($mtef[$cursor] -ne 0){$cursor++};$cursor++;break}
    17 {$cursor++;if($mtef[$cursor] -eq 255){$cursor+=3}else{$cursor++};while($mtef[$cursor] -ne 0){$cursor++};$cursor++;break}
    8 {$cursor++;if($mtef[$cursor] -eq 255){$cursor+=3}else{$cursor++};$cursor++;break}
    18 {$cursor=[int]$codec.GetMethod('SkipEquationPreferences',$flags).Invoke($null,[object[]]@($mtef,$cursor));break}
    16 {$cursor=[int]$codec.GetMethod('SkipColorDefinition',$flags).Invoke($null,[object[]]@($mtef,$cursor));break}
    15 {$cursor++;if($mtef[$cursor] -eq 255){$cursor+=3}else{$cursor++};break}
    default {if($tag -ge 9 -and $tag -le 14){$cursor=[int]$codec.GetMethod('SkipInitialSizeRecord',$flags).Invoke($null,[object[]]@($mtef,$cursor))}elseif($tag -ge 100){$cursor=[int]$codec.GetMethod('SkipFutureRecord',$flags).Invoke($null,[object[]]@($mtef,$cursor))}else{throw "Unsupported prefix record $tag"}}
   }
  }
  $nativeMathMl=[string]$codec.GetMethod('ReadEquationNativeMathMl',$flags).Invoke($null,[object[]]@(,$native))
  $row=@{part=$name;preferences=$preferences;initialSizes=$initialSizes;nativeLength=$native.Length;nativeMathMl=$nativeMathMl;semanticSignature=[string]$codec.GetMethod('SemanticSignature',$flags).Invoke($null,[object[]]@($nativeMathMl))}
  if($IncludeRecords){$row.rootOffset=$end;$row.mtefHex=[BitConverter]::ToString($mtef);$row.nativeBase64=[Convert]::ToBase64String($native)}
  $rows+=$row
 }
 $record=@{source=$XmlEvidence;assemblySha256=(Get-FileHash (Join-Path $bin 'VisualTeX.WordVsto.dll')).Hash;formatReference='https://docs.wiris.com/en_US/mathtype-mtef-v5-mathtype-40-and-later';equations=$rows}
 $json=$record|ConvertTo-Json -Depth 9
 [IO.File]::WriteAllText((Join-Path $PSScriptRoot "evidence\$Label.json"),$json,[Text.UTF8Encoding]::new($false))
 $json
} finally {[AppDomain]::CurrentDomain.remove_AssemblyResolve($handler)}
