[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$AcceptancePath,
    [Parameter(Mandatory)][string]$OutputRoot
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$acceptance = (Resolve-Path -LiteralPath $AcceptancePath).Path
$output = [IO.Path]::GetFullPath($OutputRoot)
[void][IO.Directory]::CreateDirectory($output)

function Get-PublicClassRegistration {
    $values = @()
    foreach ($view in @([Microsoft.Win32.RegistryView]::Registry32, [Microsoft.Win32.RegistryView]::Registry64)) {
        foreach ($hive in @([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryHive]::LocalMachine)) {
            $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
            try {
                foreach ($relative in @('CLSID\{0002CE03-0000-0000-C000-000000000046}\LocalServer32', 'Equation.DSMT4\CLSID', 'Equation\CLSID')) {
                    $key = $baseKey.OpenSubKey('Software\Classes\' + $relative)
                    try {
                        $value = if ($null -eq $key) { '<absent>' } else { [string]$key.GetValue('') }
                        $values += "$hive/$view/$relative=$value"
                    } finally { if ($null -ne $key) { $key.Dispose() } }
                }
            } finally { $baseKey.Dispose() }
        }
    }
    $values -join "`n"
}

function Get-PublicMathTypeState {
    $state = @{}
    $hash = [Security.Cryptography.SHA256]::Create()
    try {
        foreach ($view in @([Microsoft.Win32.RegistryView]::Registry32, [Microsoft.Win32.RegistryView]::Registry64)) {
            foreach ($hive in @([Microsoft.Win32.RegistryHive]::CurrentUser, [Microsoft.Win32.RegistryHive]::LocalMachine)) {
                $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
                try {
                    $queue = [Collections.Generic.Queue[string]]::new()
                    $queue.Enqueue('Software\Design Science\DSMT7')
                    while ($queue.Count -gt 0) {
                        $path = $queue.Dequeue()
                        $key = $base.OpenSubKey($path, $false)
                        if ($null -eq $key) { continue }
                        try {
                            foreach ($name in $key.GetValueNames()) {
                                $value = $key.GetValue($name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
                                $serialized = [Management.Automation.PSSerializer]::Serialize(@($key.GetValueKind($name).ToString(), $value))
                                # Compare state without recording licence data or user preferences.
                                $state["$hive/$view/$path/$name"] = [Convert]::ToBase64String($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($serialized)))
                            }
                            foreach ($child in $key.GetSubKeyNames()) { $queue.Enqueue($path + '\' + $child) }
                        } finally { $key.Dispose() }
                    }
                } finally { $base.Dispose() }
            }
        }
    } finally { $hash.Dispose() }
    return $state
}

$registration = Get-PublicClassRegistration
$publicState = Get-PublicMathTypeState
# The new boundary test permits an existing editor and verifies PID + creation
# time. No windowless/-mtrpc classifier, process launch override, or process kill.
& $acceptance '--mode=mathtype-native-preview-private-boundary' $output
$testExit = $LASTEXITCODE
$stateAfter = Get-PublicMathTypeState
$changed = @($publicState.Keys | Where-Object { -not $stateAfter.ContainsKey($_) -or $publicState[$_] -cne $stateAfter[$_] })
$added = @($stateAfter.Keys | Where-Object { -not $publicState.ContainsKey($_) })
$report = @{
    TestExit = $testExit
    PublicRegistrationUnchanged = ($registration -ceq (Get-PublicClassRegistration))
    PublicSettingsCompared = $publicState.Count
    ChangedPublicSettings = $changed
    AddedPublicSettings = $added
}
$report | ConvertTo-Json -Depth 5 | Tee-Object -FilePath (Join-Path $output 'public-boundary.json')
if ($testExit -ne 0) { throw "Private render boundary failed with exit $testExit." }
if (-not $report.PublicRegistrationUnchanged -or $changed.Count -ne 0 -or $added.Count -ne 0) {
    throw 'Private rendering changed public MathType configuration.'
}
