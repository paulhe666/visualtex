param(
    [string]$InstallerPath = '',
    [string]$PayloadDir = '',
    [string]$OutputDir = ''
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.UTF8Encoding]::new($false)

$windowsRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$tauriRoot = Join-Path $windowsRoot 'src-tauri'
$targetRoot = [IO.Path]::GetFullPath((Join-Path $tauriRoot 'target'))
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = Join-Path $targetRoot 'private-mathtype-runtime'
}
$output = [IO.Path]::GetFullPath($OutputDir.Trim().Trim('"'))
if (!$output.StartsWith($targetRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw "Private MathType staging output must stay under $targetRoot"
}

function Remove-OwnedDirectory([string]$path) {
    $full = [IO.Path]::GetFullPath($path)
    if (!$full.StartsWith($targetRoot + '\', [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to remove a directory outside the Tauri target tree: $full"
    }
    if (Test-Path -LiteralPath $full) {
        Remove-Item -LiteralPath $full -Recurse -Force
    }
}

$extractRoot = $null
if ([string]::IsNullOrWhiteSpace($InstallerPath) -and [string]::IsNullOrWhiteSpace($PayloadDir)) {
    $environmentInstaller = [Environment]::GetEnvironmentVariable('VISUALTEX_MATHTYPE_INSTALLER')
    if (![string]::IsNullOrWhiteSpace($environmentInstaller)) {
        $InstallerPath = $environmentInstaller
    } else {
        $existingManifestPath = Join-Path $output 'visualtex-runtime-manifest.json'
        if (Test-Path -LiteralPath $existingManifestPath -PathType Leaf) {
            $existingManifest = Get-Content -LiteralPath $existingManifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if ([string]$existingManifest.validatedVersion -ne '7.9.1') {
                throw "Existing private MathType runtime was validated for $($existingManifest.validatedVersion), expected 7.9.1."
            }
            foreach ($relative in @(
                'MathType.exe','MT7.dsc','System\MathTypeLib.exe','System\MT6.dll','System\64\MT6.dll',
                'MathPage\64\MathPage.wll','System\rt\lib\rt.jar','System\rt\lib\charsets.jar',
                'System\rt\lib\jce.jar','System\rt\lib\jsse.jar','System\rt\lib\security\java.security'
            )) {
                if (!(Test-Path -LiteralPath (Join-Path $output $relative) -PathType Leaf)) {
                    throw "Existing staged private MathType runtime is incomplete: $relative"
                }
            }
            foreach ($relative in @('System\rt\bin','System\rt\lib\ext','System\rt\lib\security')) {
                if (!(Test-Path -LiteralPath (Join-Path $output $relative) -PathType Container)) {
                    throw "Existing staged private MathType runtime directory is incomplete: $relative"
                }
            }
            $existingManifest | ConvertTo-Json -Depth 6 -Compress
            exit
        }
        throw 'Private MathType runtime is not staged. Set VISUALTEX_MATHTYPE_INSTALLER to a legally obtained MathType 7.9.1 installer, or run this script once with -InstallerPath/-PayloadDir.'
    }
}
if (![string]::IsNullOrWhiteSpace($PayloadDir)) {
    $payload = [IO.Path]::GetFullPath($PayloadDir.Trim().Trim('"'))
    if (!(Test-Path -LiteralPath $payload -PathType Container)) {
        throw "MathType payload directory not found: $payload"
    }
} else {
    if ([string]::IsNullOrWhiteSpace($InstallerPath)) {
        throw 'Provide -InstallerPath or -PayloadDir for a legally obtained MathType 7.9.1 payload.'
    }
    $installer = [IO.Path]::GetFullPath($InstallerPath.Trim().Trim('"'))
    if (!(Test-Path -LiteralPath $installer -PathType Leaf)) {
        throw "MathType installer not found: $installer"
    }
    $sevenZip = @(
        (Get-Command 7z.exe -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -First 1),
        (Join-Path $env:ProgramFiles '7-Zip\7z.exe'),
        (Join-Path ${env:ProgramFiles(x86)} '7-Zip\7z.exe')
    ) | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -First 1
    if (!$sevenZip) { throw '7-Zip is required to stage the private MathType runtime from the installer.' }
    $extractRoot = Join-Path $targetRoot 'private-mathtype-payload'
    Remove-OwnedDirectory $extractRoot
    [void][IO.Directory]::CreateDirectory($extractRoot)
    & $sevenZip x '-y' ('-o' + $extractRoot) $installer | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "7-Zip failed to extract MathType installer: $LASTEXITCODE" }
    $matches = @(Get-ChildItem -LiteralPath $extractRoot -Recurse -File -Filter 'mathtype.ex_')
    if ($matches.Count -ne 1) {
        throw "Expected exactly one MathType payload root, found $($matches.Count)."
    }
    $payload = $matches[0].Directory.FullName
}

$setupInf = Join-Path $payload 'setup.inf'
$setupLibInf = Join-Path $payload 'setup.lib.inf'
foreach ($required in @('mathtype.ex_','mt7.dsc','MathTypeLib.exe','mt6x86.dll','mt6x64.dll','mathpagex64.wll','setup.inf','setup.lib.inf')) {
    if (!(Test-Path -LiteralPath (Join-Path $payload $required) -PathType Leaf)) {
        throw "MathType payload is missing $required"
    }
}
$setupText = Get-Content -LiteralPath $setupInf -Raw -Encoding Default
if ($setupText -notmatch '(?m)^MajVersion\s*=\s*7\s*$' -or
    $setupText -notmatch '(?m)^MinVersion\s*=\s*9\s*$' -or
    $setupText -notmatch '(?m)^BugFixVersion\s*=\s*1\s*$') {
    throw 'This staging profile is validated only for MathType 7.9.1.'
}

Remove-OwnedDirectory $output
[void][IO.Directory]::CreateDirectory($output)
function Copy-Exact([string]$sourceName,[string]$relativeDestination) {
    $source = Join-Path $payload $sourceName
    $target = Join-Path $output $relativeDestination
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $target))
    [IO.File]::Copy($source,$target,$true)
}

# Native MathType server + API bridge files proven necessary by the isolated
# MTEF -> WMF probe. MathType.exe is 32-bit, hence System\MT6.dll is x86;
# the VisualTeX preview sidecar is x64 and uses System\64\MT6.dll.
Copy-Exact 'mathtype.ex_' 'MathType.exe'
Copy-Exact 'mt7.dsc' 'MT7.dsc'
Copy-Exact 'MathTypeLib.exe' 'System\MathTypeLib.exe'
Copy-Exact 'mt6x86.dll' 'System\MT6.dll'
Copy-Exact 'mt6x64.dll' 'System\64\MT6.dll'
Copy-Exact 'mathpagex64.wll' 'MathPage\64\MathPage.wll'

# setup.lib.inf is the authoritative mapping for MathTypeLib's private Java
# runtime. Keep the validated runtime closure, but omit groups proven unnecessary
# for native MTEF -> WMF rendering: resources.jar, font/CMM assets and management/
# cursor images. Do not omit rt.jar, charsets.jar, ext, crypto or security files:
# isolated probes showed that doing so prevents the renderer service from starting.
$excludedRelative = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($relative in @(
    'System\rt\lib\resources.jar',
    'System\rt\lib\fontconfig.bfc',
    'System\rt\lib\fontconfig.properties.src',
    'System\rt\lib\psfont.properties.ja',
    'System\rt\lib\psfontj2d.properties'
)) { [void]$excludedRelative.Add($relative) }
$excludedPrefixes = @(
    'System\rt\lib\fonts\',
    'System\rt\lib\cmm\',
    'System\rt\lib\management\',
    'System\rt\lib\images\'
)
foreach ($line in Get-Content -LiteralPath $setupLibInf -Encoding Default) {
    if ($line.TrimStart().StartsWith(';')) { continue }
    if ($line -notmatch '^f_[^=]+=<[^>]+>,\s*([^,]+),\s*([^,]+)') { continue }
    $sourceName = $matches[1].Trim()
    $destination = $matches[2].Trim()
    if ($destination -notlike '*<AppSystemDir>\rt\*') { continue }
    $relative = $destination.Replace('<AppSystemDir>','System').TrimStart('\')
    if ($relative -match '<[^>]+>') { continue }
    if ($excludedRelative.Contains($relative)) { continue }
    if ($excludedPrefixes | Where-Object { $relative.StartsWith($_,[StringComparison]::OrdinalIgnoreCase) }) { continue }
    $source = Join-Path $payload $sourceName
    if (!(Test-Path -LiteralPath $source -PathType Leaf)) {
        throw "MathType runtime payload is missing mapped file $sourceName"
    }
    Copy-Exact $sourceName $relative
}

$mustExist = @(
    'MathType.exe','MT7.dsc','System\MathTypeLib.exe','System\MT6.dll','System\64\MT6.dll',
    'MathPage\64\MathPage.wll','System\rt\lib\rt.jar','System\rt\lib\charsets.jar',
    'System\rt\lib\jce.jar','System\rt\lib\jsse.jar','System\rt\lib\security\java.security'
)
foreach ($relative in $mustExist) {
    if (!(Test-Path -LiteralPath (Join-Path $output $relative) -PathType Leaf)) {
        throw "Staged MathType runtime is incomplete: $relative"
    }
}
foreach ($relative in @('System\rt\bin','System\rt\lib\ext','System\rt\lib\security')) {
    if (!(Test-Path -LiteralPath (Join-Path $output $relative) -PathType Container)) {
        throw "Staged MathType runtime directory is incomplete: $relative"
    }
}

$files = @(Get-ChildItem -LiteralPath $output -Recurse -File)
$manifest = [ordered]@{
    schema = 1
    product = 'MathType'
    validatedVersion = '7.9.1'
    purpose = 'VisualTeX isolated native MTEF-to-WMF preview runtime'
    fileCount = $files.Count
    bytes = ($files | Measure-Object Length -Sum).Sum
    core = @{}
}
foreach ($relative in @('MathType.exe','MT7.dsc','System\MathTypeLib.exe','System\MT6.dll','System\64\MT6.dll','MathPage\64\MathPage.wll')) {
    $path = Join-Path $output $relative
    $manifest.core[$relative] = @{
        bytes = (Get-Item -LiteralPath $path).Length
        sha256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    }
}
$manifestPath = Join-Path $output 'visualtex-runtime-manifest.json'
[IO.File]::WriteAllText($manifestPath,($manifest | ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($false))
$manifest.fileCount = $files.Count + 1
$manifest.bytes = ($files | Measure-Object Length -Sum).Sum + (Get-Item -LiteralPath $manifestPath).Length
[IO.File]::WriteAllText($manifestPath,($manifest | ConvertTo-Json -Depth 6),[Text.UTF8Encoding]::new($false))
$manifest | ConvertTo-Json -Depth 6 -Compress
