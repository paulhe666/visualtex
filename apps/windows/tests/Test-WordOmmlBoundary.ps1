param([Parameter(Mandatory=$true)][string]$AssemblyPath,[Parameter(Mandatory=$true)][string]$ReportPath)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$assemblyPath=[IO.Path]::GetFullPath($AssemblyPath)
$dependencies=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\src-windows\VisualTeX.WordVsto\bin\x64\Release\net472'))
if(Test-Path -LiteralPath $ReportPath){throw 'Evidence exists; choose a new report path.'}
$resolver=[ResolveEventHandler]{param($sender,$eventArgs)
 $file=Join-Path $dependencies ([Reflection.AssemblyName]::new($eventArgs.Name).Name+'.dll')
 if(Test-Path -LiteralPath $file){return [Reflection.Assembly]::LoadFrom($file)}
 return $null
}
[AppDomain]::CurrentDomain.add_AssemblyResolve($resolver)
try{
 Add-Type -Path (Join-Path $PSScriptRoot 'WordOmmlBoundaryChecks.cs') -ReferencedAssemblies @((Join-Path $dependencies 'Microsoft.Office.Interop.Word.dll'),'System.dll','System.Core.dll')
 $results=@([WordOmmlBoundaryChecks]::Run($assemblyPath))
 $failed=@($results | Where-Object { !$_.Passed })
 [ordered]@{Time=(Get-Date).ToString('o');Assembly=$assemblyPath;Sha256=(Get-FileHash -LiteralPath $assemblyPath).Hash;UsesRealWord=$false;BehaviorModel='Measured range-scoped OMath clipping; ParentOMath returns physical host';Passed=($failed.Count -eq 0);Results=$results} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ReportPath -Encoding UTF8
 $results | Select-Object Case,Expected,Actual,Passed,ParentReads | Format-Table -AutoSize -Wrap
 if($failed.Count){throw ('OMML boundary checks failed: '+$failed.Count)}
}finally{[AppDomain]::CurrentDomain.remove_AssemblyResolve($resolver)}
