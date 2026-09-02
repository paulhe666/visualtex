[CmdletBinding()]
param(
    [string]$ExpectedAppVersion = "1.2.5",
    [string]$ExpectedOfficeMsiVersion = "1.0.43.0"
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$installerPath = Join-Path $root "src-tauri\target\release\bundle\nsis\VisualTeX_${ExpectedAppVersion}_x64-setup.exe"
$resourceX64 = Join-Path $root "src-tauri\resources\windows-office\VisualTeX-WindowsOffice-VSTO-x64.msi"
$resourceX86 = Join-Path $root "src-tauri\resources\windows-office\VisualTeX-WindowsOffice-VSTO-x86.msi"
$manifestX64 = Join-Path $root "src-tauri\resources\windows-office\VisualTeX-WindowsOffice-VSTO-x64.sha256.json"
$manifestX86 = Join-Path $root "src-tauri\resources\windows-office\VisualTeX-WindowsOffice-VSTO-x86.sha256.json"
$vstoRuntime = Join-Path $root "src-tauri\resources\windows-office\vstor_redist.exe"
$vstoRuntimeManifest = Join-Path $root "src-tauri\resources\windows-office\vstor_redist.sha256.json"
$buildX64 = Join-Path $root "src-windows\VisualTeX.WindowsOffice.Installer\bin\x64\Release\VisualTeX-WindowsOffice-VSTO-x64.msi"
$buildX86 = Join-Path $root "src-windows\VisualTeX.WindowsOffice.Installer\bin\x86\Release\VisualTeX-WindowsOffice-VSTO-x86.msi"
$ocrPythonRoot = Join-Path $root "src-tauri\resources\ocr-python\windows-x64"
$ocrPythonManifestPath = Join-Path $ocrPythonRoot "manifest.json"
$ocrModelRoot = Join-Path $root "src-tauri\resources\ocr-models\windows-x64"
$ocrModelCatalogPath = Join-Path $ocrModelRoot "catalog.json"
$paths = @(
    $installerPath,
    $resourceX64,
    $resourceX86,
    $manifestX64,
    $manifestX86,
    $vstoRuntime,
    $vstoRuntimeManifest,
    $buildX64,
    $buildX86,
    $ocrPythonManifestPath,
    $ocrModelCatalogPath
)

foreach ($path in $paths) {
    if (-not (Test-Path -LiteralPath $path)) { throw "Release artifact is missing: $path" }
    $item = Get-Item -LiteralPath $path
    $hash = Get-FileHash -LiteralPath $path -Algorithm SHA256
    Write-Host ("{0} | {1} bytes | SHA256 {2}" -f $item.FullName, $item.Length, $hash.Hash)
}

function Assert-ManifestFileRecord {
    param(
        [string]$Root,
        [object]$Record,
        [string]$Label
    )
    $path = Join-Path $Root ([string]$Record.name)
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "$Label is missing: $path"
    }
    $item = Get-Item -LiteralPath $path
    if ([int64]$item.Length -ne [int64]$Record.size) {
        throw "$Label has the wrong size: $path; expected $($Record.size), actual $($item.Length)."
    }
    $hash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    if (-not [string]::Equals($hash, [string]$Record.sha256, [StringComparison]::OrdinalIgnoreCase)) {
        throw "$Label has the wrong SHA256: $path; expected $($Record.sha256), actual $hash."
    }
}

$ocrPythonManifest = Get-Content -LiteralPath $ocrPythonManifestPath -Raw | ConvertFrom-Json
if ([int]$ocrPythonManifest.schemaVersion -ne 2 -or
    [string]$ocrPythonManifest.platform -ne "windows" -or
    [string]$ocrPythonManifest.architecture -ne "x64" -or
    [string]$ocrPythonManifest.pythonVersion -ne "3.12.10" -or
    [string]$ocrPythonManifest.pipVersion -ne "25.1.1") {
    throw "Unexpected private OCR Python manifest metadata: $ocrPythonManifestPath"
}
Assert-ManifestFileRecord $ocrPythonRoot $ocrPythonManifest.archive "Private Python archive"
$openMpRecord = $ocrPythonManifest.appLocalRuntime.openMp.file
if ($null -eq $openMpRecord -or [string]$openMpRecord.name -ne "vcomp140.dll" -or [int64]$openMpRecord.size -le 0) {
    throw "The private OCR Python manifest is missing the app-local Microsoft OpenMP runtime record."
}
$archivePath = Join-Path $ocrPythonRoot ([string]$ocrPythonManifest.archive.name)
Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive = [System.IO.Compression.ZipFile]::OpenRead($archivePath)
try {
    $entry = @($archive.Entries | Where-Object { [string]::Equals($_.FullName, "vcomp140.dll", [StringComparison]::OrdinalIgnoreCase) })
    if ($entry.Count -ne 1) {
        throw "The private OCR Python archive must contain exactly one app-local vcomp140.dll."
    }
    if ([int64]$entry[0].Length -ne [int64]$openMpRecord.size) {
        throw "The archived vcomp140.dll size does not match the manifest."
    }
    $sha = [Security.Cryptography.SHA256]::Create()
    $stream = $entry[0].Open()
    try {
        $digest = $sha.ComputeHash($stream)
    } finally {
        $stream.Dispose()
        $sha.Dispose()
    }
    $actualOpenMpHash = -join ($digest | ForEach-Object { $_.ToString("X2") })
    if (-not [string]::Equals($actualOpenMpHash, [string]$openMpRecord.sha256, [StringComparison]::OrdinalIgnoreCase)) {
        throw "The archived vcomp140.dll checksum does not match the manifest."
    }
} finally {
    $archive.Dispose()
}
Assert-ManifestFileRecord $ocrPythonRoot $ocrPythonManifest.wheelhouse.lock "Hash-locked OCR requirements"
$wheelhouseRoot = Join-Path $ocrPythonRoot "wheelhouse"
$manifestWheels = @($ocrPythonManifest.wheelhouse.files)
if ($manifestWheels.Count -ne 71) {
    throw "The fixed Windows OCR wheelhouse must contain exactly 71 wheels; manifest contains $($manifestWheels.Count)."
}
$manifestWheelNames = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($wheel in $manifestWheels) {
    if (-not $manifestWheelNames.Add([string]$wheel.name)) {
        throw "Duplicate wheel in OCR manifest: $($wheel.name)"
    }
    Assert-ManifestFileRecord $wheelhouseRoot $wheel "OCR wheel"
}
$actualWheels = @(Get-ChildItem -LiteralPath $wheelhouseRoot -File -Filter "*.whl")
if ($actualWheels.Count -ne $manifestWheels.Count) {
    throw "OCR wheelhouse contains unexpected or missing files. Manifest=$($manifestWheels.Count); Actual=$($actualWheels.Count)."
}
foreach ($wheel in $actualWheels) {
    if (-not $manifestWheelNames.Contains($wheel.Name)) {
        throw "Unexpected wheel outside the fixed manifest: $($wheel.FullName)"
    }
}
foreach ($requiredWheelPattern in @(
    "paddlepaddle-3.3.1-*.whl",
    "paddleocr-3.7.0-*.whl",
    "tokenizers-0.19.1-*.whl"
)) {
    if (@(Get-ChildItem -LiteralPath $wheelhouseRoot -File -Filter $requiredWheelPattern).Count -ne 1) {
        throw "The fixed OCR wheelhouse is missing pinned package $requiredWheelPattern"
    }
}
$lockText = Get-Content -LiteralPath (Join-Path $ocrPythonRoot ([string]$ocrPythonManifest.wheelhouse.lock.name)) -Raw
foreach ($requiredLockMarker in @("paddlepaddle==3.3.1", "paddleocr==3.7.0", "tokenizers==0.19.1")) {
    if (-not $lockText.Contains($requiredLockMarker)) {
        throw "OCR requirements lock is missing $requiredLockMarker"
    }
}
Write-Host "Private OCR Python 3.12.10 x64 archive, app-local Microsoft OpenMP runtime and exact 71-wheel offline closure verified."

$ocrModelCatalog = Get-Content -LiteralPath $ocrModelCatalogPath -Raw | ConvertFrom-Json
if ([int]$ocrModelCatalog.schemaVersion -ne 1 -or
    [string]$ocrModelCatalog.platform -ne "windows" -or
    [string]$ocrModelCatalog.architecture -ne "x64") {
    throw "Unexpected OCR model catalog metadata: $ocrModelCatalogPath"
}
$expectedModels = @("PP-FormulaNet_plus-S", "PP-FormulaNet_plus-M", "PP-FormulaNet_plus-L")
$expectedModelBaseUrl = "https://download.visualtex.pauljianliao.com/ppformula-model"
$expectedModelPackages = @{
    "PP-FormulaNet_plus-S" = "VisualTeX_PP-FormulaNet_plus-S_windows-x64.vtxocrmodel"
    "PP-FormulaNet_plus-M" = "VisualTeX_PP-FormulaNet_plus-M_windows-x64.vtxocrmodel"
    "PP-FormulaNet_plus-L" = "VisualTeX_PP-FormulaNet_plus-L_windows-x64.vtxocrmodel"
}
$modelEntries = @($ocrModelCatalog.entries)
if ($modelEntries.Count -ne $expectedModels.Count) {
    throw "OCR model catalog must contain S, M and L exactly once."
}
foreach ($model in $expectedModels) {
    $entry = @($modelEntries | Where-Object { [string]$_.model -eq $model })
    if ($entry.Count -ne 1 -or [int64]$entry[0].size -le 0 -or [string]$entry[0].sha256 -notmatch '^[0-9a-fA-F]{64}$') {
        throw "OCR model catalog entry is invalid or duplicated: $model"
    }
    $expectedUrl = "$expectedModelBaseUrl/$($expectedModelPackages[$model])"
    if ([string]$entry[0].url -ne $expectedUrl) {
        throw "OCR model catalog URL does not match the deployed VisualTeX web download path for ${model}: $($entry[0].url)"
    }
}
$bundledModelPacks = @(Get-ChildItem -LiteralPath (Join-Path $root "src-tauri\resources") -Recurse -File -Filter "*.vtxocrmodel")
if ($bundledModelPacks.Count -ne 0) {
    throw "OCR model packages must not be embedded in the NSIS resources: $($bundledModelPacks.FullName -join ', ')"
}
$unexpectedModelResourceFiles = @(Get-ChildItem -LiteralPath $ocrModelRoot -Recurse -File | Where-Object { $_.FullName -ne $ocrModelCatalogPath })
if ($unexpectedModelResourceFiles.Count -ne 0) {
    throw "The bundled OCR model resource directory must contain only catalog.json: $($unexpectedModelResourceFiles.FullName -join ', ')"
}
Write-Host "OCR model catalog verified; S/M/L model packages are excluded from bundled resources."

function Assert-MsiComponentBitness {
    param(
        [object]$Installer,
        [string]$Path,
        [bool]$Expected64Bit
    )
    $requiredComponents = @(
        "VstoFiles",
        "NativeOleServerFiles",
        "NativeOleRegistration",
        "WordComRegistration",
        "PowerPointComRegistration",
        "WordOfficeRegistration",
        "PowerPointOfficeRegistration",
        "ModeRegistration"
    )
    $database = $null
    $view = $null
    try {
        $database = $Installer.OpenDatabase($Path, 0)
        $view = $database.OpenView('SELECT `Component`, `Attributes` FROM `Component`')
        $view.Execute()
        $found = @{}
        while ($true) {
            $record = $view.Fetch()
            if ($null -eq $record) { break }
            try {
                $name = [string]$record.StringData(1)
                $attributes = [int]$record.IntegerData(2)
                $found[$name] = $attributes
            } finally {
                if ([Runtime.InteropServices.Marshal]::IsComObject($record)) {
                    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($record)
                }
            }
        }
        foreach ($component in $requiredComponents) {
            if (-not $found.ContainsKey($component)) {
                throw "Office MSI is missing required component '$component': $Path"
            }
            $is64Bit = ([int]$found[$component] -band 256) -ne 0
            if ($is64Bit -ne $Expected64Bit) {
                throw "Office MSI component '$component' has the wrong registry/file bitness in $Path. Expected64Bit=$Expected64Bit; Attributes=$($found[$component])."
            }
        }
        Write-Host ("{0} | Component bitness {1} verified for {2} components" -f $Path, $(if ($Expected64Bit) { "x64" } else { "x86" }), $requiredComponents.Count)
    } finally {
        if ($null -ne $view -and [Runtime.InteropServices.Marshal]::IsComObject($view)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($view)
        }
        if ($null -ne $database -and [Runtime.InteropServices.Marshal]::IsComObject($database)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($database)
        }
    }
}

$installer = New-Object -ComObject WindowsInstaller.Installer
foreach ($path in @($resourceX64, $resourceX86, $buildX64, $buildX86)) {
    $database = $installer.OpenDatabase($path, 0)
    $view = $database.OpenView("SELECT `Value` FROM `Property` WHERE `Property`='ProductVersion'")
    $view.Execute()
    $record = $view.Fetch()
    $version = $record.StringData(1)
    Write-Host ("{0} | ProductVersion {1}" -f $path, $version)
    if ($version -ne $ExpectedOfficeMsiVersion) {
        throw "Unexpected Office MSI version in $path. Expected $ExpectedOfficeMsiVersion, actual $version."
    }
    $scopeView = $database.OpenView("SELECT `Value` FROM `Property` WHERE `Property`='ALLUSERS'")
    $scopeView.Execute()
    $scopeRecord = $scopeView.Fetch()
    $allUsers = if ($null -eq $scopeRecord) { "" } else { [string]$scopeRecord.StringData(1) }
    if ($allUsers -ne "1") {
        throw "Office MSI is not machine-wide: $path; ALLUSERS='$allUsers'."
    }
    Write-Host ("{0} | ALLUSERS=1 machine-wide scope verified" -f $path)
    if ($null -ne $scopeRecord -and [Runtime.InteropServices.Marshal]::IsComObject($scopeRecord)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($scopeRecord)
    }
    if ([Runtime.InteropServices.Marshal]::IsComObject($scopeView)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($scopeView)
    }
}
Assert-MsiComponentBitness $installer $resourceX64 $true
Assert-MsiComponentBitness $installer $buildX64 $true
Assert-MsiComponentBitness $installer $resourceX86 $false
Assert-MsiComponentBitness $installer $buildX86 $false

if ((Get-FileHash $resourceX64 -Algorithm SHA256).Hash -ne (Get-FileHash $buildX64 -Algorithm SHA256).Hash) {
    throw "The x64 Office MSI bundled by Tauri is not the current x64 build."
}
if ((Get-FileHash $resourceX86 -Algorithm SHA256).Hash -ne (Get-FileHash $buildX86 -Algorithm SHA256).Hash) {
    throw "The x86 Office MSI bundled by Tauri is not the current x86 build."
}

$runtimeManifest = Get-Content -LiteralPath $vstoRuntimeManifest -Raw | ConvertFrom-Json
$runtimeHash = (Get-FileHash -LiteralPath $vstoRuntime -Algorithm SHA256).Hash
if ([string]$runtimeManifest.package.sha256 -ne $runtimeHash) {
    throw "The bundled Microsoft VSTO Runtime hash manifest does not match $vstoRuntime."
}
if ($runtimeHash -ne "CFE1A40BBE4A50022DB2164ABDB0154984E2CECB761A23CDC81CB5754F6E0A18") {
    throw "The bundled Microsoft VSTO Runtime is not the pinned 10.0.60917.00 package."
}
$runtimeSignature = Get-AuthenticodeSignature -FilePath $vstoRuntime
if ($runtimeSignature.Status -ne [System.Management.Automation.SignatureStatus]::Valid -or
    $null -eq $runtimeSignature.SignerCertificate -or
    $runtimeSignature.SignerCertificate.Subject -notmatch "Microsoft Corporation") {
    throw "The bundled Microsoft VSTO Runtime has an invalid or unexpected Authenticode signature."
}
$runtimeVersion = (Get-Item -LiteralPath $vstoRuntime).VersionInfo
if ([string]$runtimeVersion.ProductVersion -ne "10.0.60917.00" -or
    [string]$runtimeVersion.CompanyName -ne "Microsoft Corporation") {
    throw "Unexpected bundled Microsoft VSTO Runtime metadata: Version=$($runtimeVersion.ProductVersion); Company=$($runtimeVersion.CompanyName)."
}

foreach ($entry in @(
    @{ Msi = $resourceX64; Manifest = $manifestX64 },
    @{ Msi = $resourceX86; Manifest = $manifestX86 }
)) {
    $manifest = Get-Content -LiteralPath $entry.Manifest -Raw | ConvertFrom-Json
    $actualHash = (Get-FileHash -LiteralPath $entry.Msi -Algorithm SHA256).Hash
    if ($manifest.package.sha256 -ne $actualHash) {
        throw "Office MSI hash manifest does not match $($entry.Msi)."
    }
}

$customNsisTemplate = Join-Path $root "src-tauri\target\nsis-template\visualtex-installer.nsi"
if (-not (Test-Path -LiteralPath $customNsisTemplate -PathType Leaf)) {
    throw "Verified custom NSIS template is missing: $customNsisTemplate"
}
$customNsisSource = Get-Content -LiteralPath $customNsisTemplate -Raw
foreach ($requiredMarker in @(
    "Same-version maintenance defaults to the second option",
    'StrCpy $ReinstallPageCheck 2',
    'Same version: always remove the installed payload before reinstalling',
    '/VISUALTEXACCEPTANCE',
    'Page custom VisualTeXOfficePageCreate VisualTeXOfficePageLeave',
    'Page custom VisualTeXOcrPageCreate VisualTeXOcrPageLeave',
    'VisualTeXCreateBundledResourceDirectory',
    'VisualTeXInstallBundledResource'
)) {
    if (-not $customNsisSource.Contains($requiredMarker)) {
        throw "Custom NSIS template is missing verified release marker: $requiredMarker"
    }
}
$officePageMarker = 'Page custom VisualTeXOfficePageCreate VisualTeXOfficePageLeave'
$maintenancePageMarker = 'Page custom PageReinstall PageLeaveReinstall'
$directoryPageMarker = '!insertmacro MUI_PAGE_DIRECTORY'
$ocrPageMarker = 'Page custom VisualTeXOcrPageCreate VisualTeXOcrPageLeave'
$officePageIndex = $customNsisSource.IndexOf($officePageMarker, [System.StringComparison]::Ordinal)
$maintenancePageIndex = $customNsisSource.IndexOf($maintenancePageMarker, [System.StringComparison]::Ordinal)
$directoryPageIndex = $customNsisSource.IndexOf($directoryPageMarker, [System.StringComparison]::Ordinal)
$ocrPageIndex = $customNsisSource.IndexOf($ocrPageMarker, [System.StringComparison]::Ordinal)
if ($officePageIndex -lt 0 -or $maintenancePageIndex -lt 0 -or $officePageIndex -ge $maintenancePageIndex) {
    throw "The Office integration/process-guard page must run before Tauri's reinstall/uninstall maintenance page."
}
if ($directoryPageIndex -lt 0 -or $ocrPageIndex -lt 0 -or $ocrPageIndex -le $directoryPageIndex) {
    throw "The optional OCR page must remain in the normal install flow after the directory page."
}
$tauriConfig = Get-Content -LiteralPath (Join-Path $root "src-tauri\tauri.conf.json") -Raw -Encoding UTF8 | ConvertFrom-Json
if ([string]$tauriConfig.bundle.windows.nsis.template -ne "./target/nsis-template/visualtex-installer.nsi") {
    throw "Tauri is not configured to bundle with the verified custom NSIS template."
}

$nsisHooksPath = Join-Path $root "src-tauri\windows\hooks.nsh"
$nsisHooksSource = Get-Content -LiteralPath $nsisHooksPath -Raw -Encoding UTF8
if ($nsisHooksSource.Contains('$INSTDIR\VisualTeX.exe')) {
    throw "NSIS Office integration hooks must not hard-code VisualTeX.exe; use the Tauri MAINBINARYNAME macro."
}
foreach ($requiredHookMarker in @(
    'IfFileExists "$INSTDIR\${MAINBINARYNAME}.exe" 0 visualtex_main_binary_missing',
    '-VisualTeXPath "$INSTDIR\${MAINBINARYNAME}.exe"',
    'Behavior:Win32/Persistence.A!ml',
    'Remove only known legacy application payloads.',
    '%APPDATA%\VisualTeX\ocr-storage.json',
    'preserve %APPDATA%\com.visualtex.studio.',
    'Function VisualTeXOcrPageCreate',
    '${NSD_CreateCheckbox} 0 38u 100% 18u',
    'Install offline OCR resources (recommended)',
    'Function VisualTeXOcrPageLeave',
    '/VISUALTEXOCR=',
    '!macro VisualTeXCreateBundledResourceDirectory DESTINATION',
    '!macro VisualTeXInstallBundledResource DESTINATION SOURCE',
    'no ocr*, wheel or private-Python resources will be written'
)) {
    if (-not $nsisHooksSource.Contains($requiredHookMarker)) {
        throw "NSIS hooks are missing a verified executable or persistent-data marker: $requiredHookMarker"
    }
}
foreach ($forbiddenUninstallPattern in @(
    '(?im)^\s*RMDir\s+/r\s+"\$APPDATA\\VisualTeX"\s*$',
    '(?im)^\s*RMDir\s+/r\s+"\$APPDATA\\com\.visualtex\.studio"\s*$',
    '(?im)^\s*Delete\s+"\$APPDATA\\VisualTeX\\ocr-storage\.json"\s*$'
)) {
    if ($nsisHooksSource -match $forbiddenUninstallPattern) {
        throw "NSIS uninstall hooks must preserve the OCR storage pointer and legacy reusable OCR environment: $forbiddenUninstallPattern"
    }
}

$ocrStorageSource = Get-Content -LiteralPath (Join-Path $root "src-tauri\src\ocr_storage.rs") -Raw -Encoding UTF8
foreach ($requiredStorageMarker in @(
    'pub const STORAGE_DIRECTORY_NAME: &str = "VisualTeX-OCR"',
    'pub const STORAGE_CONFIG_FILE: &str = "ocr-storage.json"',
    'pub const STORAGE_MARKER_FILE: &str = ".visualtex-ocr-root.json"',
    'reset_source: source_has_payload',
    'ocr-storage.invalid-'
)) {
    if (-not $ocrStorageSource.Contains($requiredStorageMarker)) {
        throw "OCR independent-storage source is missing a release-safety marker: $requiredStorageMarker"
    }
}
$ocrRuntimeSource = Get-Content -LiteralPath (Join-Path $root "src-tauri\src\lib.rs") -Raw -Encoding UTF8
foreach ($requiredRuntimeStorageMarker in @(
    '.join("VisualTeXData")',
    '.join("ocr-runtime")',
    'storage_persistent_across_uninstall: true',
    'configure_ocr_storage_location',
    'reset_runtime_contents(&current.root)'
)) {
    if (-not $ocrRuntimeSource.Contains($requiredRuntimeStorageMarker)) {
        throw "OCR runtime path resolution is missing a persistent-storage marker: $requiredRuntimeStorageMarker"
    }
}
foreach ($forbiddenStorageMigrationMarker in @(
    'copy_tree',
    'source_to_remove',
    'remove_migrated_source',
    '.visualtex-ocr-migrating-'
)) {
    if ($ocrStorageSource.Contains($forbiddenStorageMigrationMarker) -or $ocrRuntimeSource.Contains($forbiddenStorageMigrationMarker)) {
        throw "OCR storage must reset and reinstall instead of copying the previous environment: $forbiddenStorageMigrationMarker"
    }
}

# Tauri's rendered target/release/nsis/.../installer.nsi is an ephemeral
# implementation detail and may be removed immediately after makensis exits.
# Verify the persistent template that was actually configured for bundling
# (checked above), and verify the Windows-only resource overlay that Tauri
# merges into that template instead of requiring the transient rendered file.
$windowsTauriConfigPath = Join-Path $root "src-tauri\tauri.windows.conf.json"
if (-not (Test-Path -LiteralPath $windowsTauriConfigPath -PathType Leaf)) {
    throw "Windows Tauri resource configuration is missing: $windowsTauriConfigPath"
}
$windowsTauriConfig = Get-Content -LiteralPath $windowsTauriConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$windowsResources = @($windowsTauriConfig.bundle.resources.PSObject.Properties)
foreach ($requiredResource in @(
    @{ Source = "resources/windows-office/VisualTeX-WindowsOffice-VSTO-x64.msi"; Destination = "windows-office/VisualTeX-WindowsOffice-VSTO-x64.msi" },
    @{ Source = "resources/windows-office/VisualTeX-WindowsOffice-VSTO-x64.sha256.json"; Destination = "windows-office/VisualTeX-WindowsOffice-VSTO-x64.sha256.json" },
    @{ Source = "resources/windows-office/VisualTeX-WindowsOffice-VSTO-x86.msi"; Destination = "windows-office/VisualTeX-WindowsOffice-VSTO-x86.msi" },
    @{ Source = "resources/windows-office/VisualTeX-WindowsOffice-VSTO-x86.sha256.json"; Destination = "windows-office/VisualTeX-WindowsOffice-VSTO-x86.sha256.json" },
    @{ Source = "resources/windows-office/vstor_redist.exe"; Destination = "windows-office/vstor_redist.exe" },
    @{ Source = "resources/windows-office/vstor_redist.sha256.json"; Destination = "windows-office/vstor_redist.sha256.json" }
)) {
    $record = @($windowsResources | Where-Object { $_.Name -eq $requiredResource.Source })
    if ($record.Count -ne 1 -or [string]$record[0].Value -ne [string]$requiredResource.Destination) {
        throw "Windows Tauri resource mapping is missing or invalid: $($requiredResource.Source) -> $($requiredResource.Destination)"
    }
}

& node.exe (Join-Path $root "scripts\verify_embedded_frontend_assets.mjs") `
    --exe (Join-Path $root "src-tauri\target\release\visualtex.exe")
if ($LASTEXITCODE -ne 0) {
    throw "The release VisualTeX.exe failed embedded frontend verification."
}

Write-Host "VisualTeX Windows release artifacts passed static verification, including the private offline OCR runtime, model-package exclusion, embedded main frontend and patched maintenance flow."
