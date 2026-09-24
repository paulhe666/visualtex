param([string]$Label = 'stage04ah')
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSEdition -ne 'Desktop') { throw 'Run Office dependency verification in Windows PowerShell / .NET Framework.' }
$taskRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$taskBridgeOutput = (Resolve-Path (Join-Path $taskRoot 'apps\windows\src-windows\artifacts\windows-ole-bridge')).Path
if (!$taskBridgeOutput.StartsWith($taskRoot + '\', [StringComparison]::OrdinalIgnoreCase)) { throw 'Bridge output is outside this workspace.' }
$taskAllowed = @('D3DCompiler_47_cor3.dll', 'PenImc_cor3.dll', 'PresentationNative_cor3.dll', 'vcruntime140_cor3.dll', 'visualtex-windows-office-bridge.exe', 'wpfgfx_cor3.dll')
if (@(Get-ChildItem -LiteralPath $taskBridgeOutput | Where-Object { $_.PSIsContainer -or $taskAllowed -notcontains $_.Name }).Count) { throw 'Unknown bridge output; preserve it.' }
$taskOleOutput = [IO.Path]::GetFullPath((Join-Path $taskRoot 'apps\windows\src-windows\artifacts\formula-ole-server')) + '\'
foreach ($taskServer in Get-CimInstance Win32_Process -Filter "Name='VisualTeX.FormulaOleServer.exe'") {
    if (!$taskServer.ExecutablePath) { throw 'Unable to identify running OLE server.' }
    if ($taskServer.ExecutablePath.StartsWith($taskOleOutput, [StringComparison]::OrdinalIgnoreCase)) { throw 'A build output OLE server is still active; inspect before building.' }
}
$taskVswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'
$taskMsbuild = & $taskVswhere -latest -products '*' -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe | Select-Object -First 1
if (!$taskMsbuild) { throw 'MSBuild unavailable.' }
$env:PATH = (Join-Path $env:LOCALAPPDATA 'Microsoft\dotnet') + ';' + (Split-Path $taskMsbuild -Parent) + ';' + $env:PATH
& (Join-Path $taskRoot 'apps\windows\scripts\build_windows_office.ps1') -SkipTests *> (Join-Path $PSScriptRoot "evidence\$Label-office-build-desktop.log")
if ($LASTEXITCODE -ne 0) { throw "Office package build failed: $LASTEXITCODE" }
Get-Content (Join-Path $PSScriptRoot "evidence\$Label-office-build-desktop.log") -Tail 12
