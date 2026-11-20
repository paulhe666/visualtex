param(
    [switch]$Force
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$PythonVersion = "3.12.10"
$PythonArchiveName = "python-$PythonVersion-embed-amd64.zip"
$PythonArchiveUrl = "https://www.python.org/ftp/python/$PythonVersion/$PythonArchiveName"
$PythonArchiveSha256 = "4ACBED6DD1C744B0376E3B1CF57CE906F9DC9E95E68824584C8099A63025A3C3"
$PipVersion = "25.1.1"
$PipWheelName = "pip-$PipVersion-py3-none-any.whl"
$PipWheelUrl = "https://files.pythonhosted.org/packages/29/a2/d40fb2460e883eca5199c62cfc2463fd261f760556ae6290f88488c362c0/$PipWheelName"
$PipWheelSha256 = "2913A38A2ABF4EA6B64AB507BD9E967F3B53DC1EDE74B01B0931E1CE548751AF"
$PaddleVersion = "3.3.1"
$PaddleOcrVersion = "3.7.0"
$TokenizersVersion = "0.19.1"
$VcompFileName = "vcomp140.dll"

$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$RequirementsSource = Join-Path $ProjectRoot "ocr\requirements.txt"
$SiteCustomizeSource = Join-Path $ProjectRoot "src-tauri\ocr\private_sitecustomize.py"
$ResourceRoot = Join-Path $ProjectRoot "src-tauri\resources\ocr-python\windows-x64"
$OutputArchiveName = "visualtex-python-$PythonVersion-windows-x64.zip"
$OutputArchive = Join-Path $ResourceRoot $OutputArchiveName
$OutputManifest = Join-Path $ResourceRoot "manifest.json"
$OutputWheelhouse = Join-Path $ResourceRoot "wheelhouse"
$OutputLockName = "requirements-windows-x64-py312.lock"
$OutputLock = Join-Path $ResourceRoot $OutputLockName
$CacheRoot = Join-Path $env:LOCALAPPDATA "VisualTeXBuildCache\ocr-python"
$PythonArchive = Join-Path $CacheRoot $PythonArchiveName
$PipWheel = Join-Path $CacheRoot $PipWheelName

function Get-Sha256([string]$Path) {
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToUpperInvariant()
}

function Get-FileRecord([string]$Path) {
    $item = Get-Item -LiteralPath $Path
    return [ordered]@{
        name = $item.Name
        size = $item.Length
        sha256 = Get-Sha256 $item.FullName
    }
}

function Get-PeMachine([string]$Path) {
    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
    $reader = $null
    try {
        $reader = New-Object System.IO.BinaryReader($stream)
        if ($reader.ReadUInt16() -ne 0x5A4D) {
            throw "Not a valid PE file: $Path"
        }
        $stream.Position = 0x3C
        $peOffset = $reader.ReadUInt32()
        if ($peOffset -gt ($stream.Length - 6)) {
            throw "Invalid PE header offset in $Path"
        }
        $stream.Position = $peOffset
        if ($reader.ReadUInt32() -ne 0x00004550) {
            throw "Missing PE signature in $Path"
        }
        return $reader.ReadUInt16()
    } finally {
        if ($null -ne $reader) { $reader.Dispose() }
        else { $stream.Dispose() }
    }
}

function Resolve-Vcomp140Dll {
    $roots = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    if ($env:VCToolsRedistDir -and (Test-Path -LiteralPath $env:VCToolsRedistDir -PathType Container)) {
        [void]$roots.Add((Resolve-Path -LiteralPath $env:VCToolsRedistDir).Path)
    }
    foreach ($programFilesRoot in @($env:ProgramFiles, ${env:ProgramFiles(x86)})) {
        if (-not $programFilesRoot) { continue }
        $visualStudioRoot = Join-Path $programFilesRoot "Microsoft Visual Studio"
        if (-not (Test-Path -LiteralPath $visualStudioRoot -PathType Container)) { continue }
        foreach ($redistRoot in Get-ChildItem -LiteralPath $visualStudioRoot -Directory -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -like "*\VC\Redist\MSVC" }) {
            [void]$roots.Add($redistRoot.FullName)
        }
    }

    $candidates = @()
    foreach ($root in $roots) {
        foreach ($item in Get-ChildItem -LiteralPath $root -Filter $VcompFileName -File -Recurse -ErrorAction SilentlyContinue) {
            if ($item.FullName -like "*\onecore\*") { continue }
            if ($item.FullName -notmatch '\\x64\\Microsoft\.VC\d+\.OpenMP\\vcomp140\.dll$') { continue }
            if ((Get-PeMachine $item.FullName) -ne 0x8664) { continue }
            if ([string]$item.VersionInfo.CompanyName -notmatch "Microsoft") { continue }
            $signature = Get-AuthenticodeSignature -LiteralPath $item.FullName
            if ($signature.Status -ne [System.Management.Automation.SignatureStatus]::Valid) { continue }
            if ([string]$signature.SignerCertificate.Subject -notmatch "Microsoft") { continue }
            $versionText = [string]$item.VersionInfo.FileVersion
            $version = [version]::new(0, 0, 0, 0)
            [void][version]::TryParse($versionText, [ref]$version)
            $candidates += [pscustomobject]@{
                Path = $item.FullName
                Version = $version
                VersionText = $versionText
            }
        }
    }
    $selected = $candidates | Sort-Object Version -Descending | Select-Object -First 1
    if ($null -eq $selected) {
        throw "Microsoft x64 vcomp140.dll was not found in a Visual Studio VC Redist directory. Install the Visual C++ build tools with the OpenMP runtime before producing the VisualTeX installer."
    }
    Write-Host "Using Microsoft OpenMP app-local runtime: $($selected.Path) ($($selected.VersionText))"
    return [string]$selected.Path
}

function Write-Utf8NoBom([string]$Path, [string]$Content) {
    $encoding = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $encoding)
}

function Get-VerifiedDownload(
    [string]$Url,
    [string]$Path,
    [string]$ExpectedSha256
) {
    if (Test-Path -LiteralPath $Path) {
        $actual = Get-Sha256 $Path
        if ($actual -eq $ExpectedSha256) {
            return
        }
        Remove-Item -LiteralPath $Path -Force
    }

    $directory = Split-Path -Parent $Path
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    $temporary = "$Path.download"
    Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue
    Write-Host "Downloading $Url"
    Invoke-WebRequest -UseBasicParsing -Uri $Url -OutFile $temporary
    $actual = Get-Sha256 $temporary
    if ($actual -ne $ExpectedSha256) {
        Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue
        throw "SHA-256 mismatch for $Url. Expected $ExpectedSha256, actual $actual"
    }
    Move-Item -LiteralPath $temporary -Destination $Path -Force
}

function Invoke-PrivatePython(
    [string]$Python,
    [string[]]$Arguments,
    [string]$Label,
    [hashtable]$ExtraEnvironment = @{}
) {
    $saved = @{}
    $baseEnvironment = @{
        PYTHONNOUSERSITE = "1"
        PYTHONSAFEPATH = "1"
        PYTHONUTF8 = "1"
        PYTHONIOENCODING = "utf-8"
        PIP_DISABLE_PIP_VERSION_CHECK = "1"
        PIP_NO_INPUT = "1"
    }
    foreach ($entry in $ExtraEnvironment.GetEnumerator()) {
        $baseEnvironment[$entry.Key] = [string]$entry.Value
    }
    try {
        foreach ($entry in $baseEnvironment.GetEnumerator()) {
            $saved[$entry.Key] = [Environment]::GetEnvironmentVariable($entry.Key, "Process")
            [Environment]::SetEnvironmentVariable($entry.Key, $entry.Value, "Process")
        }
        Write-Host "+ $Python $($Arguments -join ' ')"
        & $Python @Arguments
        if ($LASTEXITCODE -ne 0) {
            throw "$Label failed with exit code $LASTEXITCODE"
        }
    } finally {
        foreach ($entry in $saved.GetEnumerator()) {
            [Environment]::SetEnvironmentVariable($entry.Key, $entry.Value, "Process")
        }
    }
}

function Initialize-PrivateRuntime(
    [string]$RuntimeRoot,
    [string]$WorkRoot,
    [string]$VcompSource
) {
    New-Item -ItemType Directory -Path $RuntimeRoot -Force | Out-Null
    Expand-Archive -LiteralPath $PythonArchive -DestinationPath $RuntimeRoot -Force
    $vcompTarget = Join-Path $RuntimeRoot $VcompFileName
    Copy-Item -LiteralPath $VcompSource -Destination $vcompTarget -Force
    if ((Get-PeMachine $vcompTarget) -ne 0x8664) {
        throw "Bundled Microsoft OpenMP runtime is not x64: $vcompTarget"
    }
    $vcompSignature = Get-AuthenticodeSignature -LiteralPath $vcompTarget
    if ($vcompSignature.Status -ne [System.Management.Automation.SignatureStatus]::Valid -or
        [string]$vcompSignature.SignerCertificate.Subject -notmatch "Microsoft") {
        throw "Bundled Microsoft OpenMP runtime does not have a valid Microsoft Authenticode signature: $vcompTarget"
    }

    $sitePackages = Join-Path $RuntimeRoot "Lib\site-packages"
    New-Item -ItemType Directory -Path $sitePackages -Force | Out-Null
    $pipWheelZip = Join-Path $WorkRoot "pip-wheel.zip"
    Copy-Item -LiteralPath $PipWheel -Destination $pipWheelZip -Force
    Expand-Archive -LiteralPath $pipWheelZip -DestinationPath $sitePackages -Force

    $pthPath = Join-Path $RuntimeRoot "python312._pth"
    @(
        "python312.zip",
        ".",
        "Lib\site-packages",
        "import site"
    ) | Set-Content -LiteralPath $pthPath -Encoding ASCII

    if (-not (Test-Path -LiteralPath $SiteCustomizeSource -PathType Leaf)) {
        throw "VisualTeX private Python sitecustomize source is missing: $SiteCustomizeSource"
    }
    $siteCustomize = Get-Content -LiteralPath $SiteCustomizeSource -Raw
    Write-Utf8NoBom (Join-Path $sitePackages "sitecustomize.py") $siteCustomize

    $vcompItem = Get-Item -LiteralPath $vcompTarget
    $runtimeMetadata = [ordered]@{
        schemaVersion = 2
        pythonVersion = $PythonVersion
        architecture = "x64"
        distribution = "python.org embeddable"
        pipVersion = $PipVersion
        appLocalRuntime = [ordered]@{
            openMp = [ordered]@{
                name = $VcompFileName
                version = [string]$vcompItem.VersionInfo.FileVersion
                sha256 = Get-Sha256 $vcompTarget
            }
        }
    }
    Write-Utf8NoBom (Join-Path $RuntimeRoot "visualtex-python.json") ($runtimeMetadata | ConvertTo-Json -Depth 4)
}

function Test-PreparedBaseRuntime([string]$RuntimeRoot) {
    $python = Join-Path $RuntimeRoot "python.exe"
    if (-not (Test-Path -LiteralPath $python -PathType Leaf)) {
        throw "Prepared OCR Python is missing python.exe"
    }
    $vcomp = Join-Path $RuntimeRoot $VcompFileName
    if (-not (Test-Path -LiteralPath $vcomp -PathType Leaf)) {
        throw "Prepared OCR Python is missing the app-local Microsoft OpenMP runtime: $vcomp"
    }
    if ((Get-PeMachine $vcomp) -ne 0x8664) {
        throw "Prepared OCR Python contains a non-x64 Microsoft OpenMP runtime: $vcomp"
    }
    $probeScript = Join-Path $RuntimeRoot "_visualtex_runtime_probe.py"
    @'
import ctypes
import importlib.metadata as metadata
import json
import os
import platform
import site
import struct
import sys
runtime_root = os.path.dirname(sys.executable)
vcomp_path = os.path.join(runtime_root, "vcomp140.dll")
ctypes.WinDLL(vcomp_path)
user_site = site.getusersitepackages()
user_sites = [user_site] if isinstance(user_site, str) else list(user_site)
normalized = lambda value: os.path.normcase(os.path.abspath(value))
user_site_on_path = any(
    normalized(path) == normalized(candidate)
    for path in sys.path
    for candidate in user_sites
    if candidate
)
print(json.dumps({
    "pythonVersion": platform.python_version(),
    "pipVersion": metadata.version("pip"),
    "bits": struct.calcsize("P") * 8,
    "machine": platform.machine(),
    "executable": sys.executable,
    "prefix": sys.prefix,
    "vcompPath": vcomp_path,
    "userSiteEnabled": bool(site.ENABLE_USER_SITE),
    "userSiteOnPath": user_site_on_path,
}))
'@ | Set-Content -LiteralPath $probeScript -Encoding UTF8
    try {
        $probe = & $python -I -X utf8 $probeScript
        if ($LASTEXITCODE -ne 0) {
            throw "Prepared OCR Python validation failed with exit code $LASTEXITCODE"
        }
    } finally {
        Remove-Item -LiteralPath $probeScript -Force -ErrorAction SilentlyContinue
    }
    $parsed = $probe | Select-Object -Last 1 | ConvertFrom-Json
    if ($parsed.pythonVersion -ne $PythonVersion) {
        throw "Prepared OCR Python version is $($parsed.pythonVersion), expected $PythonVersion"
    }
    if ($parsed.pipVersion -ne $PipVersion) {
        throw "Prepared OCR pip version is $($parsed.pipVersion), expected $PipVersion"
    }
    if ($parsed.bits -ne 64 -or $parsed.machine -notmatch "^(AMD64|x86_64|x64)$") {
        throw "Prepared OCR Python architecture is invalid: $($parsed.bits)-bit $($parsed.machine)"
    }
    if ($parsed.userSiteEnabled -or $parsed.userSiteOnPath) {
        throw "Prepared OCR Python is not isolated from user site-packages"
    }
    if (-not ([string]$parsed.vcompPath).StartsWith($RuntimeRoot, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Prepared OCR Python did not load the bundled app-local Microsoft OpenMP runtime: $($parsed.vcompPath)"
    }
    $pipOutput = & $python -I -X utf8 -m pip --version
    if ($LASTEXITCODE -ne 0 -or -not ($pipOutput -match [regex]::Escape($RuntimeRoot))) {
        throw "python -m pip did not resolve inside the prepared private runtime: $pipOutput"
    }
}

function New-FixedWheelhouseLock(
    [string]$Python,
    [string]$Wheelhouse,
    [string]$LockPath,
    [string]$ScriptPath
) {
    @'
from __future__ import annotations

import hashlib
import pathlib
import sys
from pip._vendor.packaging.utils import canonicalize_name, parse_wheel_filename

wheelhouse = pathlib.Path(sys.argv[1])
lock_path = pathlib.Path(sys.argv[2])
records = []
seen = set()
for wheel in sorted(wheelhouse.glob("*.whl"), key=lambda path: path.name.lower()):
    name, version, _build, _tags = parse_wheel_filename(wheel.name)
    canonical = canonicalize_name(name)
    if canonical in seen:
        raise SystemExit(f"duplicate distribution in wheelhouse: {canonical}")
    seen.add(canonical)
    digest = hashlib.sha256(wheel.read_bytes()).hexdigest()
    records.append(f"{canonical}=={version} --hash=sha256:{digest}")
if not records:
    raise SystemExit("wheelhouse is empty")
lock_path.write_text(
    "# VisualTeX Windows OCR fixed offline dependency closure\n"
    "# Generated from binary wheels for CPython 3.12 x64.\n"
    + "\n".join(records)
    + "\n",
    encoding="utf-8",
    newline="\n",
)
'@ | Set-Content -LiteralPath $ScriptPath -Encoding UTF8
    Invoke-PrivatePython $Python @("-I", "-X", "utf8", $ScriptPath, $Wheelhouse, $LockPath) "Generate fixed OCR wheelhouse lock"
}

function Test-FullOfflineRuntime(
    [string]$RuntimeRoot,
    [string]$Wheelhouse,
    [string]$LockPath,
    [string]$WorkRoot
) {
    $python = Join-Path $RuntimeRoot "python.exe"
    Invoke-PrivatePython $python @(
        "-I", "-X", "utf8", "-m", "pip", "install",
        "--isolated",
        "--no-index",
        "--find-links", $Wheelhouse,
        "--only-binary=:all:",
        "--require-hashes",
        "--disable-pip-version-check",
        "--no-input",
        "--requirement", $LockPath
    ) "Offline OCR wheelhouse installation" @{
        PIP_NO_INDEX = "1"
        PIP_FIND_LINKS = $Wheelhouse
    }
    Invoke-PrivatePython $python @("-I", "-X", "utf8", "-m", "pip", "check") "Offline OCR dependency closure check" @{
        PIP_NO_INDEX = "1"
    }

    $probeScript = Join-Path $WorkRoot "full-runtime-probe.py"
    @'
import ctypes
import importlib.metadata as metadata
import json
import os
import platform
import site
import struct
import sys
import paddle
import paddleocr
import tokenizers
import imagesize
import ftfy
import wand
from paddleocr import FormulaRecognition

kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
kernel32.GetModuleHandleW.argtypes = [ctypes.c_wchar_p]
kernel32.GetModuleHandleW.restype = ctypes.c_void_p
kernel32.GetModuleFileNameW.argtypes = [ctypes.c_void_p, ctypes.c_wchar_p, ctypes.c_uint32]
kernel32.GetModuleFileNameW.restype = ctypes.c_uint32
vcomp_handle = kernel32.GetModuleHandleW("vcomp140.dll")
if not vcomp_handle:
    raise RuntimeError("Paddle did not load vcomp140.dll")
vcomp_buffer = ctypes.create_unicode_buffer(32768)
if not kernel32.GetModuleFileNameW(vcomp_handle, vcomp_buffer, len(vcomp_buffer)):
    raise ctypes.WinError(ctypes.get_last_error())
vcomp_path = vcomp_buffer.value

user_site = site.getusersitepackages()
user_sites = [user_site] if isinstance(user_site, str) else list(user_site)
normalized = lambda value: os.path.normcase(os.path.abspath(value))
user_site_on_path = any(
    normalized(path) == normalized(candidate)
    for path in sys.path
    for candidate in user_sites
    if candidate
)
print(json.dumps({
    "pythonVersion": platform.python_version(),
    "bits": struct.calcsize("P") * 8,
    "machine": platform.machine(),
    "paddleVersion": paddle.__version__,
    "paddleocrVersion": metadata.version("paddleocr"),
    "tokenizersVersion": metadata.version("tokenizers"),
    "formulaRecognition": FormulaRecognition.__name__,
    "executable": sys.executable,
    "vcompPath": vcomp_path,
    "userSiteEnabled": bool(site.ENABLE_USER_SITE),
    "userSiteOnPath": user_site_on_path,
}))
'@ | Set-Content -LiteralPath $probeScript -Encoding UTF8
    $probe = & $python -I -X utf8 $probeScript
    if ($LASTEXITCODE -ne 0) {
        throw "Full offline OCR runtime interface validation failed with exit code $LASTEXITCODE"
    }
    $parsed = $probe | Select-Object -Last 1 | ConvertFrom-Json
    if ($parsed.pythonVersion -ne $PythonVersion -or $parsed.bits -ne 64) {
        throw "Offline OCR Python version or architecture mismatch"
    }
    if ($parsed.paddleVersion -ne $PaddleVersion) {
        throw "Offline OCR PaddlePaddle version is $($parsed.paddleVersion), expected $PaddleVersion"
    }
    if ($parsed.paddleocrVersion -ne $PaddleOcrVersion) {
        throw "Offline OCR PaddleOCR version is $($parsed.paddleocrVersion), expected $PaddleOcrVersion"
    }
    if ($parsed.tokenizersVersion -ne $TokenizersVersion) {
        throw "Offline OCR tokenizers version is $($parsed.tokenizersVersion), expected $TokenizersVersion"
    }
    if ($parsed.formulaRecognition -ne "FormulaRecognition") {
        throw "PaddleOCR FormulaRecognition interface is unavailable"
    }
    if ($parsed.userSiteEnabled -or $parsed.userSiteOnPath) {
        throw "Full offline OCR runtime used user site-packages"
    }
    if (-not ([string]$parsed.executable).StartsWith($RuntimeRoot, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Full offline OCR runtime resolved the wrong interpreter: $($parsed.executable)"
    }
    if (-not ([string]$parsed.vcompPath).StartsWith($RuntimeRoot, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Paddle loaded Microsoft OpenMP from outside the bundled OCR runtime: $($parsed.vcompPath)"
    }
}

function Install-StagedDirectory(
    [string]$Staged,
    [string]$Target,
    [string]$Backup
) {
    Remove-Item -LiteralPath $Backup -Recurse -Force -ErrorAction SilentlyContinue
    $hadTarget = Test-Path -LiteralPath $Target
    if ($hadTarget) {
        Move-Item -LiteralPath $Target -Destination $Backup -Force
    }
    try {
        Move-Item -LiteralPath $Staged -Destination $Target -Force
        Remove-Item -LiteralPath $Backup -Recurse -Force -ErrorAction SilentlyContinue
    } catch {
        Remove-Item -LiteralPath $Target -Recurse -Force -ErrorAction SilentlyContinue
        if ($hadTarget -and (Test-Path -LiteralPath $Backup)) {
            Move-Item -LiteralPath $Backup -Destination $Target -Force
        }
        throw
    }
}

if (-not (Test-Path -LiteralPath $RequirementsSource -PathType Leaf)) {
    throw "OCR requirements file is missing: $RequirementsSource"
}
New-Item -ItemType Directory -Path $ResourceRoot -Force | Out-Null
New-Item -ItemType Directory -Path $CacheRoot -Force | Out-Null
Get-VerifiedDownload $PythonArchiveUrl $PythonArchive $PythonArchiveSha256
Get-VerifiedDownload $PipWheelUrl $PipWheel $PipWheelSha256
$Vcomp140Source = Resolve-Vcomp140Dll

$workRoot = Join-Path $env:TEMP ("visualtex-ocr-python-build-" + [guid]::NewGuid().ToString("N"))
$runtimeRoot = Join-Path $workRoot "runtime"
$validationRoot = Join-Path $workRoot "validation"
$stagedWheelhouse = Join-Path $workRoot "wheelhouse"
$stagedLock = Join-Path $workRoot $OutputLockName
try {
    Initialize-PrivateRuntime $runtimeRoot $workRoot $Vcomp140Source
    Test-PreparedBaseRuntime $runtimeRoot
    $python = Join-Path $runtimeRoot "python.exe"

    New-Item -ItemType Directory -Path $stagedWheelhouse -Force | Out-Null
    Invoke-PrivatePython $python @(
        "-I", "-X", "utf8", "-m", "pip", "download",
        "--dest", $stagedWheelhouse,
        "--only-binary=:all:",
        "--disable-pip-version-check",
        "--no-input",
        "--requirement", $RequirementsSource
    ) "Download complete fixed Windows OCR wheelhouse"

    $lockScript = Join-Path $workRoot "generate-wheelhouse-lock.py"
    New-FixedWheelhouseLock $python $stagedWheelhouse $stagedLock $lockScript

    Remove-Item -LiteralPath $validationRoot -Recurse -Force -ErrorAction SilentlyContinue
    Copy-Item -LiteralPath $runtimeRoot -Destination $validationRoot -Recurse -Force
    Test-FullOfflineRuntime $validationRoot $stagedWheelhouse $stagedLock $workRoot

    # Archive only the private Python bootstrap. The fixed wheelhouse remains a
    # separate installer resource so runtime installation can use --find-links.
    $stagedArchive = Join-Path $workRoot $OutputArchiveName
    Compress-Archive -Path (Join-Path $runtimeRoot "*") -DestinationPath $stagedArchive -CompressionLevel Optimal
    $archiveValidation = Join-Path $workRoot "archive-validation"
    New-Item -ItemType Directory -Path $archiveValidation -Force | Out-Null
    Expand-Archive -LiteralPath $stagedArchive -DestinationPath $archiveValidation -Force
    Test-PreparedBaseRuntime $archiveValidation

    $archiveBackup = "$OutputArchive.backup"
    Remove-Item -LiteralPath $archiveBackup -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $OutputArchive) {
        Move-Item -LiteralPath $OutputArchive -Destination $archiveBackup -Force
    }
    try {
        Move-Item -LiteralPath $stagedArchive -Destination $OutputArchive -Force
        Remove-Item -LiteralPath $archiveBackup -Force -ErrorAction SilentlyContinue
    } catch {
        Remove-Item -LiteralPath $OutputArchive -Force -ErrorAction SilentlyContinue
        if (Test-Path -LiteralPath $archiveBackup) {
            Move-Item -LiteralPath $archiveBackup -Destination $OutputArchive -Force
        }
        throw
    }

    Install-StagedDirectory $stagedWheelhouse $OutputWheelhouse "$OutputWheelhouse.backup"
    Copy-Item -LiteralPath $stagedLock -Destination $OutputLock -Force

    $wheelRecords = @(
        Get-ChildItem -LiteralPath $OutputWheelhouse -File -Filter "*.whl" |
            Sort-Object Name |
            ForEach-Object { Get-FileRecord $_.FullName }
    )
    if ($wheelRecords.Count -eq 0) {
        throw "Bundled OCR wheelhouse is empty after activation"
    }
    $manifest = [ordered]@{
        schemaVersion = 2
        platform = "windows"
        architecture = "x64"
        pythonVersion = $PythonVersion
        pipVersion = $PipVersion
        archive = Get-FileRecord $OutputArchive
        appLocalRuntime = [ordered]@{
            openMp = [ordered]@{
                file = Get-FileRecord (Join-Path $runtimeRoot $VcompFileName)
                version = [string](Get-Item -LiteralPath (Join-Path $runtimeRoot $VcompFileName)).VersionInfo.FileVersion
                source = "Microsoft Visual C++ Redistributable app-local OpenMP runtime"
            }
        }
        wheelhouse = [ordered]@{
            lock = Get-FileRecord $OutputLock
            files = $wheelRecords
        }
        expected = [ordered]@{
            paddlepaddle = $PaddleVersion
            paddleocr = $PaddleOcrVersion
            tokenizers = $TokenizersVersion
            formulaRecognition = "paddleocr.FormulaRecognition"
        }
        upstream = [ordered]@{
            pythonArchive = $PythonArchiveName
            pythonArchiveSha256 = $PythonArchiveSha256
            pipWheel = $PipWheelName
            pipWheelSha256 = $PipWheelSha256
            requirements = (Split-Path -Leaf $RequirementsSource)
            requirementsSha256 = Get-Sha256 $RequirementsSource
        }
    }
    Write-Utf8NoBom $OutputManifest ($manifest | ConvertTo-Json -Depth 8)

    Write-Host "Prepared bundled OCR Python: $OutputArchive"
    Write-Host "Runtime SHA-256: $($manifest.archive.sha256)"
    Write-Host "App-local OpenMP runtime: $($manifest.appLocalRuntime.openMp.file.name) $($manifest.appLocalRuntime.openMp.version)"
    Write-Host "Wheelhouse files: $($wheelRecords.Count)"
    Write-Host "Wheelhouse bytes: $((Get-ChildItem -LiteralPath $OutputWheelhouse -File | Measure-Object Length -Sum).Sum)"
    Write-Host "Offline validation: Python $PythonVersion x64, app-local Microsoft OpenMP, PaddlePaddle $PaddleVersion, PaddleOCR $PaddleOcrVersion, tokenizers $TokenizersVersion, FormulaRecognition available"
} finally {
    Remove-Item -LiteralPath $workRoot -Recurse -Force -ErrorAction SilentlyContinue
}
