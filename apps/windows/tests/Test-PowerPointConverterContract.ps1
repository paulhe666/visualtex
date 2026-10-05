param(
    [Parameter(Mandatory=$true)][string]$CandidateDirectory,
    [Parameter(Mandatory=$true)][string]$ArtifactDirectory
)
$ErrorActionPreference='Stop'
[Console]::OutputEncoding=[Text.UTF8Encoding]::new($false)
$directory=(Resolve-Path -LiteralPath $CandidateDirectory).Path
$artifact=[IO.Path]::GetFullPath($ArtifactDirectory)
[void][IO.Directory]::CreateDirectory($artifact)
$wordPath=Join-Path $directory 'VisualTeX.WordVsto.dll'
$pptPath=Join-Path $directory 'VisualTeX.PowerPointVsto.dll'
$result=[ordered]@{
    TimeUtc=[DateTime]::UtcNow.ToString('o')
    WordDllSha256=(Get-FileHash -LiteralPath $wordPath).Hash
    PowerPointDllSha256=(Get-FileHash -LiteralPath $pptPath).Hash
    Scope='Production converter reflection contract only; no Office application created'
    Passed=$false
    Checks=@()
    Error=$null
}
try {
    foreach($name in @('VisualTeX.WindowsOffice.Contracts.dll','Microsoft.Office.Interop.Word.dll','Microsoft.Office.Interop.PowerPoint.dll')) {
        [void][Reflection.Assembly]::LoadFrom((Join-Path $directory $name))
    }
    $word=[Reflection.Assembly]::LoadFrom($wordPath)
    $ppt=[Reflection.Assembly]::LoadFrom($pptPath)
    $flags=[Reflection.BindingFlags]'Static,NonPublic'
    $converter=$word.GetType('VisualTeX.WordVsto.WordOmmlConverter',$true)
    $replacement=@($converter.GetMethods($flags) | Where-Object Name -eq 'ReplaceWithPreparedOmmlDirect')
    $result.Checks+=@{Name='Word replacement contract';Signatures=@($replacement | ForEach-Object ToString)}
    $kernel=$ppt.GetType('VisualTeX.PowerPointVsto.PowerPointNativeOmmlKernel',$true)
    $resolver=$kernel.GetMethod('ResolveConverterMethods',$flags)
    $arguments=[object[]]@($null,$null)
    [void]$resolver.Invoke($null,$arguments)
    $transform=[Reflection.MethodInfo]$arguments[0]
    $replace=[Reflection.MethodInfo]$arguments[1]
    if($null -eq $transform -or $null -eq $replace){throw 'Production resolver returned no methods.'}
    $result.Checks+=@{Name='Production method resolution';Passed=$true}
    $parameters=$replace.GetParameters()
    if($parameters.Length -ne 6 -or $parameters[5].ParameterType -ne [bool] -or $parameters[5].DefaultValue -ne $false) {
        throw 'Native inline-promotion contract differs from the PowerPoint call.'
    }
    $result.Checks+=@{Name='Native inline-promotion remains disabled';Passed=$true}
    $omml=[string]$transform.Invoke($null,[object[]]@('<math xmlns="http://www.w3.org/1998/Math/MathML"><mfrac><mi>x</mi><mn>2</mn></mfrac></math>'))
    $xml=[xml]$omml
    if($null -eq $xml.SelectSingleNode('//*[local-name()="f"]')) {throw 'Resolved transform did not produce a fraction.'}
    $result.Checks+=@{Name='Resolved transform produces structured OMML';Passed=$true}
    $result.Passed=$true
} catch {
    $failure=$_.Exception
    while($null -ne $failure.InnerException){$failure=$failure.InnerException}
    $result.Error=$failure.ToString()
} finally {
    $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $artifact 'converter-contract.json') -Encoding UTF8
}
$result | ConvertTo-Json -Depth 8
if(-not $result.Passed){exit 1}
