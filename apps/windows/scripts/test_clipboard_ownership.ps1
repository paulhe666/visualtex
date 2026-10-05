param([string]$OutputRoot, [string]$Dotnet = "dotnet", [string]$TargetFramework = "net8.0")
$ErrorActionPreference = 'Stop'
if (!$OutputRoot) { $OutputRoot = Join-Path $env:TEMP ('visualtex-clipboard-ownership-' + [guid]::NewGuid().ToString('N')) }
[void][IO.Directory]::CreateDirectory($OutputRoot)
$source = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot '../src-windows/VisualTeX.WindowsOffice.VstoShared/MathTypeOleStorage.cs')
$start = $source.IndexOf('    internal sealed class ClipboardTransaction : IDisposable')
$end = $source.IndexOf('    internal sealed class StandaloneClipboardTransaction : IDisposable')
if ($start -lt 0 -or $end -le $start) { throw 'Clipboard transaction source boundaries changed.' }
$fixture = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'ownership/ClipboardFixture.cs')
$program = $fixture.Replace('/* PRODUCTION_CLIPBOARD_TRANSACTION */', $source.Substring($start, $end - $start))
[IO.File]::WriteAllText((Join-Path $OutputRoot 'Program.cs'), $program)
[IO.File]::WriteAllText((Join-Path $OutputRoot 'ClipboardOwnership.csproj'), "<Project Sdk=`"Microsoft.NET.Sdk`"><PropertyGroup><OutputType>Exe</OutputType><TargetFramework>$TargetFramework</TargetFramework><Nullable>enable</Nullable><NuGetAudit>false</NuGetAudit></PropertyGroup></Project>")
# No package dependencies, Office, COM activation or system clipboard calls.
& $Dotnet run --project (Join-Path $OutputRoot 'ClipboardOwnership.csproj') --configuration Release -p:RestoreIgnoreFailedSources=true
if ($LASTEXITCODE -ne 0) { throw "Clipboard ownership test failed: $LASTEXITCODE" }
