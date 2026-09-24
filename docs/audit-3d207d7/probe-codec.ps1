param([string]$SessionId='a64f98c5-4cc6-467b-a6da-3915d5d9af3b')
$ErrorActionPreference='Stop';[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'));$out=Join-Path $PSScriptRoot 'evidence'
$bin=Join-Path $root 'apps\windows\src-windows\VisualTeX.WordVsto\bin\x64\Release\net472'
$handler=[ResolveEventHandler]{param($sender,$eventArgs) $p=Join-Path $bin (([Reflection.AssemblyName]$eventArgs.Name).Name+'.dll');if(Test-Path $p){return [Reflection.Assembly]::LoadFrom($p)};return $null};[AppDomain]::CurrentDomain.add_AssemblyResolve($handler)
$asm=[Reflection.Assembly]::LoadFrom((Join-Path $bin 'VisualTeX.WordVsto.dll'));$t=$asm.GetType('VisualTeX.WordVsto.MathTypeMtefCodec',$true);$flags=[Reflection.BindingFlags]'Static,NonPublic,Public'
$create=$t.GetMethod('CreateEquationNative',$flags);$read=$t.GetMethod('ReadEquationNativeMathMl',$flags);$sig=$t.GetMethod('SemanticSignature',$flags)
$s=Get-Content (Join-Path $env:APPDATA ('com.visualtex.studio\office\sessions\'+$SessionId+'\session.json')) -Raw -Encoding UTF8|ConvertFrom-Json
$m=[string]$s.exportResult.mathMl;if(!$m){throw 'Actual renderer MathML missing'}
$cases=[ordered]@{actual=$m;minimal='<math xmlns="http://www.w3.org/1998/Math/MathML"><mrow><mo>{</mo><mtable columnalign="left left"><mtr><mtd><mi>x</mi></mtd></mtr><mtr><mtd><mi>y</mi></mtd></mtr></mtable><mo fence="true" stretchy="true"></mo></mrow></math>'}
$cases['actual_without_empty_close']=[regex]::Replace($m,'<mo[^>]*></mo>','')
$cases['actual_with_empty_second_column']=$m.Replace('</mtd>','</mtd><mtd></mtd>')
$report=@()
foreach($k in $cases.Keys){try{$inputXml=$cases[$k];$g=$create.Invoke($null,[object[]]@($inputXml,$false));$native=$g.GetType().GetProperty('EquationNative').GetValue($g,$null);$actual=[string]$read.Invoke($null,[object[]]@(,$native));$expectedSig=[string]$sig.Invoke($null,[object[]]@($inputXml));$actualSig=[string]$sig.Invoke($null,[object[]]@($actual));$r=[ordered]@{case=$k;session=$SessionId;input=$inputXml;output=$actual;expectedSignature=$expectedSig;actualSignature=$actualSig;equal=($expectedSig -ceq $actualSig)};$report+=$r;Write-Output ($k+'|equal='+$r.equal+'|EXPECTED='+$expectedSig+'|ACTUAL='+$actualSig)}catch{$report+=@{case=$k;error=$_.Exception.ToString()};Write-Output ($k+'|ERROR='+$_.Exception.ToString())}}
[IO.File]::WriteAllText((Join-Path $out 'codec-roundtrip.json'),($report|ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false))
[AppDomain]::CurrentDomain.remove_AssemblyResolve($handler)
