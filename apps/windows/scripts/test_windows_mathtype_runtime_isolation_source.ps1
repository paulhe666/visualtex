[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$windowsRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$bridgePath = Join-Path $windowsRoot 'src-windows\VisualTeX.NativeRenderWorker\MathTypeNativePreviewCommand.cs'
$interopPath = Join-Path $windowsRoot 'src-windows\VisualTeX.WindowsOffice.VstoShared\MathTypeOleInterop.cs'
$guardPath = Join-Path $windowsRoot 'scripts\manage_private_mathtype_runtime.ps1'
$nativeManagerPath = Join-Path $windowsRoot 'scripts\manage_windows_native_render.ps1'
$hooksPath = Join-Path $windowsRoot 'src-tauri\windows\hooks.nsh'
$stagingPath = Join-Path $windowsRoot 'scripts\prepare_private_mathtype_runtime.ps1'

$bridge = Get-Content -LiteralPath $bridgePath -Raw -Encoding UTF8
$interop = Get-Content -LiteralPath $interopPath -Raw -Encoding UTF8
$guard = Get-Content -LiteralPath $guardPath -Raw -Encoding UTF8
$nativeManager = Get-Content -LiteralPath $nativeManagerPath -Raw -Encoding UTF8
$hooks = Get-Content -LiteralPath $hooksPath -Raw -Encoding UTF8
$staging = Get-Content -LiteralPath $stagingPath -Raw -Encoding UTF8
$adapter = Get-Content -LiteralPath (Join-Path $windowsRoot 'src-windows\VisualTeX.WindowsOffice.VstoShared\MathTypeNativePreviewRenderer.cs') -Raw -Encoding UTF8
$worker = Get-Content -LiteralPath (Join-Path $windowsRoot 'src-windows\VisualTeX.NativeRenderWorker\Program.cs') -Raw -Encoding UTF8
$legacyEntry = Get-Content -LiteralPath (Join-Path $windowsRoot 'src-windows\VisualTeX.WindowsOleBridge\Program.cs') -Raw -Encoding UTF8

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
Assert-Source (-not $bridge.Contains('Process.Start(')) `
    'The preview command must let MTInitAPI launch as needed, not reopen the native editor on each render.'
Assert-Source ($bridge.Contains('init(MtInitLaunchAsNeeded, 8)')) `
    'The preview command no longer uses MathPage launch-as-needed initialization.'
Assert-Source ($bridge.Contains('DeleteAbandonedPrivateVirtualRegistry')) `
    'Abandoned VisualTeX-only virtual registry sessions are not reclaimed.'
Assert-Source ($staging.Contains("'Language\MT7chs.dll'")) `
    'The private runtime can still be staged without the Chinese native language dependency.'
Assert-Source ($bridge.Contains('Path.Combine("Language", "MT7chs.dll")')) `
    'The runtime validator still accepts a package that changes Chinese MathType to English.'
Assert-Source ($bridge.Contains('the saved language was not changed.')) `
    'The runtime does not fail closed when the saved native language is not bundled.'

# Word's separate native-editor diagnostics must never use the reduced runtime.
Assert-Source ([regex]::IsMatch(
        $interop,
        'ResolveMathPagePath\(\)\s*=>\s*ResolveInstalledMathPagePathForArchitecture',
        [Text.RegularExpressions.RegexOptions]::CultureInvariant)) `
    'Word can still load the reduced VisualTeX MathPage WLL in-process.'
Assert-Source (-not $bridge.Contains('ResolveMathPagePath(')) `
    'The service engine still discovers installed MathType.'
Assert-Source (-not $legacyEntry.Contains('mathtype-preview-manifest')) `
    'The legacy interactive bridge still exposes a direct native-render entry point.'
Assert-Source ($worker.Contains('SessionId != 0')) `
    'The native worker no longer rejects interactive sessions.'
Assert-Source ($adapter.Contains('NativeRenderClient.RenderAsync')) `
    'Office does not use the data-only service client.'
foreach ($forbidden in @('Process.Start', 'GetProcessesByName', 'ManagementObjectSearcher', 'AcquireSharedSession', 'ResolveMathPagePath')) {
    Assert-Source (-not $adapter.Contains($forbidden)) "Office rendering still owns legacy native runtime behavior: $forbidden"
}
Assert-Source ($interop.Contains('!IsBundledMathTypeRuntimePath(capabilities.ServerPath)')) `
    'A stale VisualTeX private COM path can still masquerade as installed MathType.'

# Current maintenance owns exactly the dedicated SCM service. It must never
# enumerate MathType.exe/MathTypeLib.exe and then infer ownership from a PID,
# because Session 0 service descendants can intentionally hide ExecutablePath.
Assert-Source ($guard.Contains('SYSTEM\CurrentControlSet\Services\VisualTeX.NativeRender')) `
    'The maintenance guard does not prove the exact SCM service identity.'
Assert-Source ($guard.Contains("NT SERVICE\VisualTeX.NativeRender")) `
    'The maintenance guard does not verify the dedicated virtual account.'
Assert-Source ($guard.Contains('Stop-Service -Name $serviceName')) `
    'The maintenance guard does not stop the owned service through SCM.'
Assert-Source (-not $guard.Contains('Get-CimInstance -ClassName Win32_Process')) `
    'The maintenance guard still enumerates MathType processes.'
Assert-Source (-not $guard.Contains('GetProcessesByName')) `
    'The maintenance guard still owns processes by executable name.'
Assert-Source ($nativeManager.Contains('Start-Service -Name $serviceName')) `
    'Same-version renderer maintenance can leave the explicitly stopped owned service offline.'
Assert-Source ($nativeManager.Contains('The native-render MSI completed without an owned service registration.')) `
    'Renderer restart is not gated by a fresh ownership proof after MSI maintenance.'
Assert-Source ($hooks.Contains('Function VisualTeXEnsureNativeRendererStopped')) `
    'Installer maintenance does not expose the exact-service stop boundary.'
Assert-Source (-not $hooks.Contains('VisualTeXPromptPrivateMathTypeRuntimeClosure')) `
    'Installer still uses the obsolete process-based private MathType prompt.'

# Upgrades and uninstalls must still repair only legacy values that point inside
# recognized VisualTeX roots. This is static registry cleanup, not process
# ownership and therefore remains independent of the current renderer service.
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
