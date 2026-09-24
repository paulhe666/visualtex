param([string]$Before,[string]$After,[string]$Label)
$ErrorActionPreference='Stop'
Add-Type -AssemblyName System.Xml.Linq
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$bin=Join-Path $root 'apps\windows\src-windows\VisualTeX.WordVsto\bin\x64\Release\net472'
$resolver=[ResolveEventHandler]{param($sender,$eventArgs) $path=Join-Path $bin (([Reflection.AssemblyName]$eventArgs.Name).Name+'.dll');if(Test-Path -LiteralPath $path){return [Reflection.Assembly]::LoadFrom($path)};return $null}
[AppDomain]::CurrentDomain.add_AssemblyResolve($resolver)
try{
 $assembly=[Reflection.Assembly]::LoadFrom((Join-Path $bin 'VisualTeX.WordVsto.dll'))
 $type=$assembly.GetType('VisualTeX.WordVsto.WordLocalEditSnapshot',$true)
 $flags=[Reflection.BindingFlags]'Static,NonPublic,Public'
 $normalize=$type.GetMethod('NormalizedBody',$flags);$signature=$type.GetMethod('Signature',$flags)
 $values=@();$index=0
 foreach($file in @($Before,$After)){
  $xml=[IO.File]::ReadAllText((Join-Path $PSScriptRoot ('evidence\'+$file)))
  $body=[string]$normalize.Invoke($null,[object[]]@($xml))
  [IO.File]::WriteAllText((Join-Path $PSScriptRoot "evidence\$Label-$index.xml"),[Xml.Linq.XElement]::Parse($body).ToString())
  $values+=[string]$signature.Invoke($null,[object[]]@($xml));$index++
 }
 @{before=$values[0];after=$values[1];equal=($values[0] -ceq $values[1]);kind='data-only comparison of actual Word evidence'}|ConvertTo-Json|Tee-Object -FilePath (Join-Path $PSScriptRoot "evidence\$Label.json")
}finally{[AppDomain]::CurrentDomain.remove_AssemblyResolve($resolver)}
