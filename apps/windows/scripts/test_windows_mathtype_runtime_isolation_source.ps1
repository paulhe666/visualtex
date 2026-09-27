[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$windowsRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$bridgePath = Join-Path $windowsRoot 'src-windows\VisualTeX.WindowsOleBridge\MathTypeNativePreviewCommand.cs'
$interopPath = Join-Path $windowsRoot 'src-windows\VisualTeX.WordVsto\MathTypeOleInterop.cs'
$guardPath = Join-Path $windowsRoot 'scripts\manage_private_mathtype_runtime.ps1'
$hooksPath = Join-Path $windowsRoot 'src-tauri\windows\hooks.nsh'

$bridge = Get-Content -LiteralPath $bridgePath -Raw -Encoding UTF8
$interop = Get-Content -LiteralPath $interopPath -Raw -Encoding UTF8
$guard = Get-Content -LiteralPath $guardPath -Raw -Encoding UTF8
$hooks = Get-Content -LiteralPath $hooksPath -Raw -Encoding UTF8

function Assert-Source([bool]$condition, [string]$message) {
    if (-not $condition) { throw $message }
}

# The private preview runtime must not own or temporarily replace MathType's
# public per-user COM registration. A sidecar can be terminated by native code,
# so cleanup in finally is not a sufficient safety boundary.
Assert-Source (-not $bridge.Contains('PrivateLocalServerSnapshot')) `
    'The preview bridge still snapshots the public MathType COM registration.'
Assert-Source (-not $bridge.Contains('Software\Classes\CLSID')) `
    'The preview bridge still writes MathType under the public HKCU Classes tree.'
Assert-Source ([regex]::IsMatch(
        $bridge,
        'RegOverridePredefKey\(\s*HkeyClassesRoot',
        [Text.RegularExpressions.RegexOptions]::CultureInvariant)) `
    'The preview bridge no longer scopes its private class view to the sidecar process.'
Assert-Source ($bridge.Contains('Arguments = "-mtrpc"')) `
    'The private RPC renderer is not launched explicitly.'
Assert-Source ($bridge.Contains('DeleteAbandonedPrivateVirtualRegistry')) `
    'Abandoned VisualTeX-only virtual registry sessions are not reclaimed.'

# Live Word automation must resolve only the complete installed MathType build.
Assert-Source ([regex]::IsMatch(
        $interop,
        'ResolveMathPagePath\(\)\s*=>\s*ResolveInstalledMathPagePathForArchitecture',
        [Text.RegularExpressions.RegexOptions]::CultureInvariant)) `
    'Word can still load the reduced VisualTeX MathPage WLL in-process.'
$installedFirst = $interop.IndexOf(
    'return ResolveInstalledMathPagePathForArchitecture(processArchitecture)',
    [StringComparison]::Ordinal)
$bundledFallback = $interop.IndexOf(
    '?? ResolveBundledMathPagePath(processArchitecture);',
    [StringComparison]::Ordinal)
Assert-Source ($installedFirst -ge 0 -and $bundledFallback -gt $installedFirst) `
    'The reduced runtime is not an installed-MathType fallback.'
Assert-Source ($interop.Contains('!IsBundledMathTypeRuntimePath(capabilities.ServerPath)')) `
    'A stale VisualTeX private COM path can still masquerade as installed MathType.'

# Upgrades and uninstalls must repair only legacy values that point inside a
# recognized VisualTeX runtime root.
Assert-Source ($guard.Contains("RepairRegistration")) `
    'The runtime guard has no legacy registration repair mode.'
Assert-Source ($guard.Contains('$fullServerPath.StartsWith($prefix')) `
    'Legacy registration repair is not path-scoped to VisualTeX-owned files.'
$repairCalls = [regex]::Matches(
    $hooks,
    'VisualTeXRunPrivateRuntimeGuard RepairRegistration').Count
Assert-Source ($repairCalls -eq 2) `
    "Expected repair on install and uninstall, found $repairCalls calls."

Write-Output 'MathType runtime isolation source acceptance passed.'
