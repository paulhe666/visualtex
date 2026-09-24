param([string]$Label='mechanisms')
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$bin=Join-Path $root 'apps\windows\src-windows\VisualTeX.WordVsto\bin\x64\Release\net472'
$resolver=[ResolveEventHandler]{param($sender,$eventArgs) $path=Join-Path $bin (([Reflection.AssemblyName]$eventArgs.Name).Name+'.dll');if(Test-Path -LiteralPath $path){return [Reflection.Assembly]::LoadFrom($path)};return $null}
[AppDomain]::CurrentDomain.add_AssemblyResolve($resolver)
try {
 $assembly=[Reflection.Assembly]::LoadFrom((Join-Path $bin 'VisualTeX.WordVsto.dll'))
 $flags=[Reflection.BindingFlags]'Static,NonPublic,Public'
 $signature=$assembly.GetType('VisualTeX.WordVsto.WordLocalEditSnapshot',$true).GetMethod('Signature',$flags)
 $before=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'evidence\stage02-vt-tail-before-__15.xml'))
 $after=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'evidence\stage02-vt-tail-after-__15.xml'))
 $first=[string]$signature.Invoke($null,[object[]]@($before));$second=[string]$signature.Invoke($null,[object[]]@($after))
 $snapshot=@{original=$first;recovered=$second;equal=($first -ceq $second)}
 if(!$snapshot.equal){throw 'Recovery signature does not match actual Word undo evidence.'}

 $xmlType=$assembly.GetType('VisualTeX.WordVsto.MathTypeWordOpenXml',$true)
 $template=$xmlType.GetMethod('CreateVisualTeXNumberTemplate',$flags).Invoke($null,[object[]]@('heading2-dash'))
 $rewrite=$xmlType.GetMethod('RewriteMathTypePlaceRefFieldFlatOpc',$flags)
 [xml]$audit=[IO.File]::ReadAllText((Join-Path $root 'docs\audit-3d207d7\evidence\audit-ref-final-after-__70.xml'))
 $manager=[Xml.XmlNamespaceManager]::new($audit.NameTable)
 $manager.AddNamespace('w','http://schemas.openxmlformats.org/wordprocessingml/2006/main')
 $manager.AddNamespace('pkg','http://schemas.microsoft.com/office/2006/xmlPackage')
 $paragraph=@($audit.SelectNodes('//w:p',$manager)|Where-Object {$_.InnerText -like '*MACROBUTTON MTPlaceRef*'})[0]
 $children=@($paragraph.ChildNodes)
 $macro=@($children|Where-Object {$_.InnerText -like '*MACROBUTTON MTPlaceRef*'})[0]
 $begin=$macro.PreviousSibling
 foreach($child in $children){if($child -eq $begin){break};[void]$paragraph.RemoveChild($child)}
 $body=$audit.SelectSingleNode('//w:body',$manager)
 foreach($child in @($body.ChildNodes)){if($child -ne $paragraph){[void]$body.RemoveChild($child)}}
 $source=$audit.OuterXml
 $rewritten=[string]$rewrite.Invoke($null,[object[]]@($source,$template))
 # A second rewrite validates the exact alias span produced by the first.
 $roundTrip=[string]$rewrite.Invoke($null,[object[]]@($rewritten,$template))
 $unknown=$source.Replace('VTEqNum_bb58029dd9ca4a029571939017603eb4','UnknownReferencedBookmark')
 $rejected=$false
 try {$null=$rewrite.Invoke($null,[object[]]@($unknown,$template))}catch {$rejected=$_.Exception.GetBaseException() -is [IO.InvalidDataException];if(!$rejected){throw}}
 if(!$rejected){throw 'An unsupported reference bookmark was accepted.'}
 $result=@{label=$Label;dllSha256=(Get-FileHash (Join-Path $bin 'VisualTeX.WordVsto.dll')).Hash;kind='data-only supplementary verification, not real Word acceptance';recovery=$snapshot;aliasRewritten=($rewritten -like '*VTEqNum_bb58029dd9ca4a029571939017603eb4*');secondRewriteStable=($rewritten -ceq $roundTrip);unsupportedRejected=$rejected}
 $result|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $PSScriptRoot "evidence\$Label.json") -Encoding UTF8
 [IO.File]::WriteAllText((Join-Path $PSScriptRoot "evidence\$Label-alias.xml"),$rewritten,[Text.UTF8Encoding]::new($false))
 $result|ConvertTo-Json -Depth 5
} finally {[AppDomain]::CurrentDomain.remove_AssemblyResolve($resolver)}
