[CmdletBinding()]
param([string]$Configuration='Release', [string]$RuntimeRoot='', [switch]$SkipPublish)
$ErrorActionPreference='Stop'
$windowsRoot=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
if (!$RuntimeRoot) { $RuntimeRoot=Join-Path $windowsRoot 'src-tauri\target\private-mathtype-runtime' }
$RuntimeRoot=[IO.Path]::GetFullPath($RuntimeRoot)
foreach($required in @('MathType.exe','MathPage\64\MathPage.wll','System\MathTypeLib.exe','visualtex-runtime-manifest.json')) {
    if (!(Test-Path -LiteralPath (Join-Path $RuntimeRoot $required) -PathType Leaf)) { throw "Missing private runtime input: $required" }
}
$dotnet=Join-Path $env:LOCALAPPDATA 'Microsoft\dotnet\dotnet.exe'
if (!(Test-Path -LiteralPath $dotnet)) { $dotnet=(Get-Command dotnet.exe -ErrorAction Stop).Source }
if ((& $dotnet --version) -notlike '8.*') { throw 'The native-render package requires the .NET 8 SDK.' }
$env:DOTNET_ROOT=Split-Path -Parent $dotnet
$env:DOTNET_HOST_PATH=$dotnet
$env:DOTNET_MSBUILD_SDK_RESOLVER_CLI_DIR=$env:DOTNET_ROOT
$sdk=Get-ChildItem (Join-Path $env:DOTNET_ROOT 'sdk') -Directory | Where-Object Name -like '8.*' | Sort-Object { [version]$_.Name } -Descending | Select-Object -First 1
$env:MSBuildSDKsPath=Join-Path $sdk.FullName 'Sdks'
$env:MSBuildEnableWorkloadResolver='false'
$targetRoot=Join-Path $windowsRoot 'src-tauri\target'
$stage=Join-Path $targetRoot ('native-render-package-'+[guid]::NewGuid().ToString('N'))
$payload=Join-Path $stage 'payload'
[void][IO.Directory]::CreateDirectory($payload)
$serviceProject=Join-Path $windowsRoot 'src-windows\VisualTeX.NativeRenderService\VisualTeX.NativeRenderService.csproj'
$workerProject=Join-Path $windowsRoot 'src-windows\VisualTeX.NativeRenderWorker\VisualTeX.NativeRenderWorker.csproj'
$serviceOutput=Join-Path $windowsRoot "src-windows\VisualTeX.NativeRenderService\bin\$Configuration\net472"
$workerOutput=Join-Path $stage 'worker'
if (!$SkipPublish) {
    & $dotnet restore $serviceProject -r win7-x64 -p:NuGetAudit=false --nologo -v minimal
    if($LASTEXITCODE) { throw 'Native-render service restore failed.' }
    & $dotnet build $serviceProject -c $Configuration -r win7-x64 -p:AppendRuntimeIdentifierToOutputPath=false --no-restore --nologo -v minimal
    if($LASTEXITCODE) { throw 'Native-render service build failed.' }
    & $dotnet publish $workerProject -c $Configuration -p:NuGetAudit=false --nologo -v minimal -o $workerOutput
    if($LASTEXITCODE) { throw 'Native-render worker publish failed.' }
} else {
    $workerOutput=Join-Path $windowsRoot 'src-windows\VisualTeX.NativeRenderWorker\bin\Release\net8.0-windows\win-x64\publish'
}
foreach($source in @($serviceOutput,$workerOutput)) {
    Get-ChildItem -LiteralPath $source -File | Where-Object Extension -ne '.pdb' | Copy-Item -Destination $payload
}
foreach($required in @('VisualTeX.NativeRenderService.exe','VisualTeX.NativeRenderWorker.exe')) {
    if (!(Test-Path -LiteralPath (Join-Path $payload $required))) { throw "Missing build output: $required" }
}
Copy-Item -LiteralPath $RuntimeRoot -Destination (Join-Path $payload 'runtime') -Recurse
# One component per file, with deterministic identifiers/GUIDs based on destination.
# Generated WiX is a build artifact, never a hand-maintained list of runtime DLLs.
function Identifier([string]$relative) {
    $sha=[Security.Cryptography.SHA256]::Create()
    try { return 'n'+([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($relative.ToLowerInvariant()))).Replace('-','').Substring(0,30)) }
    finally { $sha.Dispose() }
}
function Esc([string]$value) { [Security.SecurityElement]::Escape($value) }
$components=[Collections.Generic.List[string]]::new()
$fragment=[Text.StringBuilder]::new()
[void]$fragment.AppendLine('<Wix xmlns="http://wixtoolset.org/schemas/v4/wxs"><Fragment><DirectoryRef Id="NATIVERENDERFOLDER">')
function Emit-Directory([string]$directory,[string]$relative) {
    foreach($child in Get-ChildItem -LiteralPath $directory -Directory | Sort-Object Name) {
        $childRelative=if($relative) { $relative+'\'+$child.Name } else { $child.Name }
        [void]$fragment.AppendLine('<Directory Id="'+(Identifier ('dir/'+$childRelative))+'" Name="'+(Esc $child.Name)+'">')
        Emit-Directory $child.FullName $childRelative
        [void]$fragment.AppendLine('</Directory>')
    }
    foreach($file in Get-ChildItem -LiteralPath $directory -File | Sort-Object Name) {
        $fileRelative=if($relative) { $relative+'\'+$file.Name } else { $file.Name }
        if($fileRelative -eq 'VisualTeX.NativeRenderService.exe') { continue }
        $id=Identifier ('file/'+$fileRelative)
        $components.Add($id)
        [void]$fragment.AppendLine('<Component Id="'+$id+'" Guid="*" Bitness="always64"><File Id="f'+$id+'" Source="'+(Esc $file.FullName)+'" KeyPath="yes" /></Component>')
    }
}
Emit-Directory $payload ''
[void]$fragment.AppendLine('</DirectoryRef><ComponentGroup Id="NativeRenderPayload">')
foreach($id in $components) { [void]$fragment.AppendLine('<ComponentRef Id="'+$id+'" />') }
[void]$fragment.AppendLine('</ComponentGroup></Fragment></Wix>')
$fragmentPath=Join-Path $stage 'Payload.wxs'
[IO.File]::WriteAllText($fragmentPath,$fragment.ToString(),[Text.UTF8Encoding]::new($false))
$project=Join-Path $windowsRoot 'src-windows\VisualTeX.NativeRender.Installer\VisualTeX.NativeRender.Installer.wixproj'
& $dotnet build $project -c $Configuration -p:Platform=x64 "-p:PayloadRoot=$payload" "-p:PayloadFragment=$fragmentPath" -p:NuGetAudit=false --nologo -v minimal
if($LASTEXITCODE) { throw 'Native-render MSI build failed.' }
$msi=Join-Path $windowsRoot "src-windows\VisualTeX.NativeRender.Installer\bin\x64\$Configuration\VisualTeX-NativeRender-x64.msi"
$resources=Join-Path $windowsRoot 'src-tauri\resources\windows-office'
[void][IO.Directory]::CreateDirectory($resources)
Copy-Item -LiteralPath $msi -Destination $resources -Force
$files=@(Get-ChildItem -LiteralPath $payload -Recurse -File | ForEach-Object {
    [ordered]@{path=$_.FullName.Substring($payload.Length+1);sha256=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash}
})
$manifest=[ordered]@{schema=1;version='1.2.9';service='VisualTeX.NativeRender';package=[ordered]@{name=[IO.Path]::GetFileName($msi);sha256=(Get-FileHash -LiteralPath $msi -Algorithm SHA256).Hash};files=$files}
$manifestPath=Join-Path $resources 'VisualTeX-NativeRender-x64.sha256.json'
[IO.File]::WriteAllText($manifestPath,($manifest | ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($false))
Write-Output "Native-render MSI: $msi"
Write-Output "Native-render manifest: $manifestPath"
