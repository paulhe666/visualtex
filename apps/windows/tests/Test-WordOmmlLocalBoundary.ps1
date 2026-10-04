param([Parameter(Mandatory=$true)][string]$AssemblyPath,[Parameter(Mandatory=$true)][string]$ReportPath)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$dependencyRoot=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\src-windows\VisualTeX.WordVsto\bin\x64\Release\net472'))
$assemblyPath=[IO.Path]::GetFullPath($AssemblyPath)
if(Test-Path -LiteralPath $ReportPath){throw 'Evidence already exists; choose a new report path.'}
$resolver=[ResolveEventHandler]{param($sender,$eventArgs)
 $file=Join-Path $dependencyRoot (([Reflection.AssemblyName]::new($eventArgs.Name).Name)+'.dll')
 if(Test-Path -LiteralPath $file){return [Reflection.Assembly]::LoadFrom($file)}
 return $null
}
[AppDomain]::CurrentDomain.add_AssemblyResolve($resolver)
try{
 Add-Type -Path (Join-Path $PSScriptRoot 'WordOmmlLocalBoundaryChecks.cs') -ReferencedAssemblies @((Join-Path $dependencyRoot 'Microsoft.Office.Interop.Word.dll'),'System.dll','System.Core.dll')
 $results=@([WordOmmlLocalBoundaryChecks]::Run($assemblyPath))
 $failed=@($results | Where-Object { !$_.Passed })
 [ordered]@{Time=(Get-Date).ToString('o');Assembly=$assemblyPath;Sha256=(Get-FileHash -LiteralPath $assemblyPath).Hash;UsesRealWord=$false;Scope='Actual compiled local resolver; partial Range.OMaths model from real Word snapshot';Passed=($failed.Count -eq 0);Results=$results} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ReportPath -Encoding UTF8
 $results | Select-Object Case,Expected,Actual,Passed,GlobalReads,ContentOutsideHost,ParagraphExpansions | Format-Table -AutoSize -Wrap
 Write-Output ('Report: '+[IO.Path]::GetFullPath($ReportPath))
 if($failed.Count){throw ('OMML local boundary failures: '+(($failed | ForEach-Object Case) -join '; '))}
}finally{[AppDomain]::CurrentDomain.remove_AssemblyResolve($resolver)}
