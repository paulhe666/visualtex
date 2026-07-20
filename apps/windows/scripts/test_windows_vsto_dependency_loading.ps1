[CmdletBinding()]
param(
    [ValidateSet("x86", "x64")]
    [string]$Platform = "x64",
    [string]$Configuration = "Release",
    [string]$ProbeAssemblyPath,
    [string]$ProbeClassName
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot

function Assert-SameDirectory([string]$ExpectedDirectory, [Reflection.Assembly]$Assembly) {
    $actualDirectory = Split-Path -Parent $Assembly.Location
    if (-not [string]::Equals(
            (Resolve-Path $actualDirectory).Path,
            (Resolve-Path $ExpectedDirectory).Path,
            [StringComparison]::OrdinalIgnoreCase)) {
        throw "$($Assembly.GetName().Name) loaded from '$actualDirectory' instead of '$ExpectedDirectory'."
    }
}

function Invoke-DependencyProbe([string]$AssemblyPath, [string]$ClassName) {
    $AssemblyPath = (Resolve-Path $AssemblyPath).Path
    $directory = Split-Path -Parent $AssemblyPath
    $requiredFiles = @(
        "VisualTeX.WindowsOffice.Contracts.dll",
        "System.Text.Json.dll",
        "System.Text.Encodings.Web.dll",
        "Microsoft.Bcl.AsyncInterfaces.dll",
        "System.Memory.dll",
        "System.Buffers.dll",
        "System.Numerics.Vectors.dll",
        "System.ValueTuple.dll",
        "System.Runtime.CompilerServices.Unsafe.dll",
        "System.Threading.Tasks.Extensions.dll"
    )
    foreach ($file in $requiredFiles) {
        $candidate = Join-Path $directory $file
        if (-not (Test-Path $candidate)) {
            throw "Required VSTO runtime dependency is missing: $candidate"
        }
    }

    $addInAssembly = [Reflection.Assembly]::LoadFile($AssemblyPath)
    $addInType = $addInAssembly.GetType($ClassName, $true)
    $instance = [Activator]::CreateInstance($addInType)
    if ($null -eq $instance) {
        throw "Unable to instantiate $ClassName."
    }

    $jsonAssembly = [Reflection.Assembly]::Load(
        "System.Text.Json, Version=8.0.0.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51")
    Assert-SameDirectory $directory $jsonAssembly
    if ($jsonAssembly.GetName().Version -lt [Version]"8.0.0.5") {
        throw "System.Text.Json did not bind to the packaged patched runtime: $($jsonAssembly.FullName)"
    }

    $vectorsAssembly = [Reflection.Assembly]::Load(
        "System.Numerics.Vectors, Version=4.1.4.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
    Assert-SameDirectory $directory $vectorsAssembly

    $valueTupleAssembly = [Reflection.Assembly]::Load(
        "System.ValueTuple, Version=4.0.3.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51")
    if ($valueTupleAssembly.GetName().Version -lt [Version]"4.0.0.0") {
        throw "System.ValueTuple binding returned an unsupported runtime: $($valueTupleAssembly.FullName)"
    }

    $contractsPath = Join-Path $directory "VisualTeX.WindowsOffice.Contracts.dll"
    $contractsAssembly = [Reflection.Assembly]::LoadFile($contractsPath)
    $metadataType = $contractsAssembly.GetType(
        "VisualTeX.WindowsOffice.Contracts.FormulaMetadata",
        $true)
    $lineType = $contractsAssembly.GetType(
        "VisualTeX.WindowsOffice.Contracts.FormulaLine",
        $true)
    $codecType = $contractsAssembly.GetType(
        "VisualTeX.WindowsOffice.Contracts.FormulaMetadataCodec",
        $true)

    $metadata = [Activator]::CreateInstance($metadataType)
    $line = [Activator]::CreateInstance($lineType)
    $metadataType.GetProperty("FormulaId").SetValue(
        $metadata,
        [Guid]::NewGuid().ToString(),
        $null)
    $metadataType.GetProperty("Latex").SetValue($metadata, "x^2", $null)
    $lineType.GetProperty("Id").SetValue($line, [Guid]::NewGuid().ToString(), $null)
    $lineType.GetProperty("Latex").SetValue($line, "x^2", $null)

    $listDefinition = [System.Collections.Generic.List[object]].GetGenericTypeDefinition()
    $listType = $listDefinition.MakeGenericType(@($lineType))
    $lines = [Activator]::CreateInstance($listType)
    [void]$listType.GetMethod("Add").Invoke($lines, @($line))
    $metadataType.GetProperty("Lines").SetValue($metadata, $lines, $null)

    try {
        $json = [string]$codecType.GetMethod("SerializeJson").Invoke(
            $null,
            @($metadata))
    } catch [Reflection.TargetInvocationException] {
        throw $_.Exception.InnerException
    }
    if ($json -notmatch '"latex":"x\^2"') {
        throw "VisualTeX metadata JSON probe returned unexpected content: $json"
    }

    Write-Host "$ClassName CodeBase dependency resolution and JSON execution passed."
}

if (-not [string]::IsNullOrWhiteSpace($ProbeAssemblyPath)) {
    if ([string]::IsNullOrWhiteSpace($ProbeClassName)) {
        throw "ProbeClassName is required when ProbeAssemblyPath is supplied."
    }
    Invoke-DependencyProbe $ProbeAssemblyPath $ProbeClassName
    exit 0
}

$probes = @(
    [ordered]@{
        Assembly = Join-Path $root "src-windows\VisualTeX.WordVsto\bin\$Platform\$Configuration\net48\VisualTeX.WordVsto.dll"
        Class = "VisualTeX.WordVsto.ThisAddIn"
    },
    [ordered]@{
        Assembly = Join-Path $root "src-windows\VisualTeX.PowerPointVsto\bin\$Platform\$Configuration\net48\VisualTeX.PowerPointVsto.dll"
        Class = "VisualTeX.PowerPointVsto.ThisAddIn"
    }
)

$requires32Bit = $Platform -eq "x86"
$currentPowerShell = (Get-Process -Id $PID).Path
$probePowerShell = if ($requires32Bit -eq (-not [Environment]::Is64BitProcess)) {
    $currentPowerShell
} elseif ($requires32Bit) {
    Join-Path $env:WINDIR "SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
} else {
    Join-Path $env:WINDIR "Sysnative\WindowsPowerShell\v1.0\powershell.exe"
}
if (-not (Test-Path $probePowerShell)) {
    throw "Matching Windows PowerShell host is missing: $probePowerShell"
}

foreach ($probe in $probes) {
    if (-not (Test-Path $probe.Assembly)) {
        throw "VSTO assembly was not built: $($probe.Assembly)"
    }
    $process = Start-Process $probePowerShell -ArgumentList @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", ('"{0}"' -f $PSCommandPath),
        "-ProbeAssemblyPath", ('"{0}"' -f $probe.Assembly),
        "-ProbeClassName", $probe.Class
    ) -Wait -PassThru -NoNewWindow
    if ($process.ExitCode -ne 0) {
        throw "$($probe.Class) dependency probe failed with exit code $($process.ExitCode)."
    }
}

Write-Host "VisualTeX VSTO dependency loading passed for $Platform."
