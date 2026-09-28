param(
 [Parameter(Mandatory=$true)][string]$AssemblyPath,
 [Parameter(Mandatory=$true)][string]$ReportPath,
 [switch]$ExpectUnsafeBaseline,
 [string]$DependencyRoot=''
)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
if(!$DependencyRoot){$DependencyRoot=Join-Path $PSScriptRoot '..\src-windows\VisualTeX.WordVsto\bin\x64\Release\net472'}
$dependencyRoot=[IO.Path]::GetFullPath($DependencyRoot)
$assemblyPath=[IO.Path]::GetFullPath($AssemblyPath)
$interop=Join-Path $dependencyRoot 'Microsoft.Office.Interop.Word.dll'
$resolver=[ResolveEventHandler]{param($sender,$eventArgs)
 $name=[Reflection.AssemblyName]::new($eventArgs.Name).Name+'.dll'
 $file=Join-Path $dependencyRoot $name
 if(Test-Path -LiteralPath $file){return [Reflection.Assembly]::LoadFrom($file)}
 return $null
}
[AppDomain]::CurrentDomain.add_AssemblyResolve($resolver)
try{
 Add-Type -Path (Join-Path $PSScriptRoot 'WordPassiveObservationChecks.cs') -ReferencedAssemblies $interop
 $results=@([WordPassiveObservationChecks]::Run($assemblyPath))
 $unmodelled=@($results | Where-Object { $_.UnexpectedCalls.Length -ne 0 })
 $violations=@($results | Where-Object { !$_.BoundarySatisfied })
 $passed=($unmodelled.Count -eq 0 -and $(if($ExpectUnsafeBaseline){$violations.Count -gt 0}else{$violations.Count -eq 0}))
 $report=[ordered]@{Assembly=$assemblyPath;Sha256=(Get-FileHash -LiteralPath $assemblyPath).Hash;UsesRealWord=$false;Expected=$(if($ExpectUnsafeBaseline){'Unsafe baseline must expose the old content read'}else{'Safe passive observations; explicit command still reaches content'});Passed=$passed;Results=$results}
 $report | ConvertTo-Json -Depth 7 | Set-Content -LiteralPath $ReportPath -Encoding UTF8
 $report | ConvertTo-Json -Depth 7
 if(!$passed){throw 'Word passive-observation boundary checks failed; inspect the report.'}
}finally{[AppDomain]::CurrentDomain.remove_AssemblyResolve($resolver)}
