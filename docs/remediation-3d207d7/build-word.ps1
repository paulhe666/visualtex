param([string]$Label='stage01')
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$taskDotnet=Join-Path $env:LOCALAPPDATA 'Microsoft\dotnet\dotnet.exe'
$taskSdk=Get-ChildItem (Join-Path (Split-Path -Parent $taskDotnet) 'sdk') -Directory | Sort-Object {[version]$_.Name} -Descending | Select-Object -First 1
$env:DOTNET_ROOT=Split-Path -Parent $taskDotnet
$env:DOTNET_HOST_PATH=$taskDotnet
$env:DOTNET_MSBUILD_SDK_RESOLVER_CLI_DIR=$env:DOTNET_ROOT
$env:MSBuildSDKsPath=Join-Path $taskSdk.FullName 'Sdks'
$env:MSBuildEnableWorkloadResolver='false'
$taskReferenceRoot=Join-Path $env:USERPROFILE '.nuget\packages\microsoft.netframework.referenceassemblies.net472\1.0.3\build'
$taskVswhere=Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$taskMsbuild=& $taskVswhere -latest -products '*' -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe | Select-Object -First 1
& $taskMsbuild (Join-Path $root 'apps\windows\src-windows\VisualTeX.WordVsto\VisualTeX.WordVsto.csproj') /t:Build /p:Configuration=Release /p:Platform=x64 /p:TargetFrameworkRootPath=$taskReferenceRoot /v:minimal /nologo 2>&1 | Tee-Object -FilePath (Join-Path $PSScriptRoot "evidence\$Label-build.log")
if($LASTEXITCODE -ne 0){throw "Word VSTO build failed: $LASTEXITCODE"}
