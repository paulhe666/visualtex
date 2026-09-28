param(
 [Parameter(Mandatory=$true)][string]$AssemblyPath,
 [Parameter(Mandatory=$true)][string]$ReportPath,
 [string]$DependencyRoot=''
)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
if([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA'){throw 'Use Windows PowerShell -STA.'}
if(!$DependencyRoot){$DependencyRoot=Join-Path $PSScriptRoot '..\src-windows\VisualTeX.WordVsto\bin\x64\Release\net472'}
$dependencyRoot=[IO.Path]::GetFullPath($DependencyRoot)
$assemblyPath=[IO.Path]::GetFullPath($AssemblyPath)
if(Test-Path -LiteralPath $ReportPath){throw 'Evidence path already exists; use a new report path.'}
$resolver=[ResolveEventHandler]{param($sender,$eventArgs)
 $name=[Reflection.AssemblyName]::new($eventArgs.Name).Name+'.dll'
 $file=Join-Path $dependencyRoot $name
 if(Test-Path -LiteralPath $file){return [Reflection.Assembly]::LoadFrom($file)}
 return $null
}
[AppDomain]::CurrentDomain.add_AssemblyResolve($resolver)
try{
 Add-Type -Path (Join-Path $PSScriptRoot 'WordRibbonObservationChecks.cs') -ReferencedAssemblies @((Join-Path $dependencyRoot 'Microsoft.Office.Interop.Word.dll'),'System.dll','System.Windows.Forms.dll','System.Drawing.dll')
 $results=@([WordRibbonObservationChecks]::Run($assemblyPath))
 $failed=@($results | Where-Object { !$_.BoundarySatisfied -or !$_.BehaviorSatisfied })
 $unmodelled=@($results | Where-Object { $_.UnexpectedCalls.Length -ne 0 })
 $report=[ordered]@{Time=(Get-Date).ToString('o');Assembly=$assemblyPath;Sha256=(Get-FileHash -LiteralPath $assemblyPath).Hash;UsesRealWord=$false;Scope='Actual compiled Ribbon callbacks, complete selection events and dispatcher work on managed Office doubles';Passed=($failed.Count -eq 0);FailedCases=@($failed | ForEach-Object Case);UnmodelledCases=$unmodelled.Count;Results=$results}
 $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $ReportPath -Encoding UTF8
 $results | Select-Object Case,BoundarySatisfied,BehaviorSatisfied,Detail | Format-Table -AutoSize -Wrap
 Write-Output ('Report: '+[IO.Path]::GetFullPath($ReportPath))
 if($failed.Count){throw ('Ribbon observation checks failed: '+($report.FailedCases -join '; '))}
}finally{[AppDomain]::CurrentDomain.remove_AssemblyResolve($resolver)}
