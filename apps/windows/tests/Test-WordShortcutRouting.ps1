param([Parameter(Mandatory=$true)][string]$AssemblyPath,[Parameter(Mandatory=$true)][string]$ReportPath)
$ErrorActionPreference='Stop'
$assemblyPath=[IO.Path]::GetFullPath($AssemblyPath)
$dependencyRoot=Split-Path $assemblyPath
$resolver=[ResolveEventHandler]{param($sender,$eventArgs)
 $file=Join-Path $dependencyRoot ([Reflection.AssemblyName]::new($eventArgs.Name).Name+'.dll')
 if(Test-Path $file){return [Reflection.Assembly]::LoadFrom($file)}
 return $null
}
[AppDomain]::CurrentDomain.add_AssemblyResolve($resolver)
try{
 Add-Type -Path (Join-Path $PSScriptRoot 'WordShortcutRoutingChecks.cs') -ReferencedAssemblies @((Join-Path $dependencyRoot 'Microsoft.Office.Interop.Word.dll'),'System.Core')
 $results=@([WordShortcutRoutingChecks]::Run($assemblyPath))
 $passed=@($results|Where-Object {!$_.Passed}).Count -eq 0
 $report=[ordered]@{Assembly=$assemblyPath;Sha256=(Get-FileHash $assemblyPath).Hash;UsesRealWord=$false;UsesPhysicalInput=$false;RegistryWrites=$false;Passed=$passed;Results=$results}
 $report|ConvertTo-Json -Depth 6|Set-Content -LiteralPath $ReportPath -Encoding UTF8
 $report|ConvertTo-Json -Depth 6
 if(!$passed){throw 'Shortcut routing checks failed'}
}finally{[AppDomain]::CurrentDomain.remove_AssemblyResolve($resolver)}
