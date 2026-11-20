[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ExecutablePath,
    [Parameter(Mandatory = $true)][string]$IdentityRoot,
    [string]$OfficeUiRoot,
    [ValidateRange(1, 65535)][int]$Port = 43129,
    [ValidatePattern('^[A-Za-z0-9_-]+$')][string]$RunName = ("run-" + (Get-Date -Format "yyyyMMdd-HHmmss")),
    [switch]$Start
)
# Source-only acceptance launcher. No installer, registry writes, certificate
# generation/import, credential copies, user documents, or installed-process stop.
$ErrorActionPreference = "Stop"
$repo = (Resolve-Path (Join-Path $PSScriptRoot "..\..\..")).Path
$executable = (Resolve-Path -LiteralPath $ExecutablePath).Path
$identity = (Resolve-Path -LiteralPath $IdentityRoot).Path.TrimEnd('\')
if ($Port -eq 43127) { throw "The acceptance port must differ from installed Companion port 43127." }
if (-not $executable.StartsWith($repo.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw "Launch only a rebuilt source executable under this checkout."
}
$identityOffice = Join-Path $identity "office"
foreach ($name in @("install.json", "localhost-cert.pem", "localhost-key.pem", "certificate.json")) {
    if (-not (Test-Path -LiteralPath (Join-Path $identityOffice $name) -PathType Leaf)) {
        throw "The existing read-only Office identity is incomplete: $name"
    }
}
# Read only the public certificate and source metadata. Never print or copy token/key.
$certificatePath = Join-Path $identityOffice "localhost-cert.pem"
$pem = [IO.File]::ReadAllText($certificatePath)
$der = [Convert]::FromBase64String(($pem -replace '-----BEGIN CERTIFICATE-----|-----END CERTIFICATE-----|\s', ''))
$certificate = New-Object Security.Cryptography.X509Certificates.X509Certificate2 -ArgumentList @(,$der)
$thumbprint = $certificate.Thumbprint
$sourceConfig = Get-Content -LiteralPath (Join-Path $identityOffice "install.json") -Raw | ConvertFrom-Json
if ($sourceConfig.port -ne 43127 -or $sourceConfig.protocolVersion -ne 1 -or $sourceConfig.installToken.Length -ne 64) {
    throw "Existing Office identity metadata does not match installed port/protocol/token requirements."
}
$sourceConfig = $null
$uiRoot = if ($OfficeUiRoot) { (Resolve-Path -LiteralPath $OfficeUiRoot).Path } else { Join-Path $repo "apps\windows\dist-office-windows-native" }
if (-not $uiRoot.StartsWith($repo.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw "Acceptance UI must be a source build under this checkout."
}
$ui = Join-Path $uiRoot "dialog\index.html"
if (-not (Test-Path -LiteralPath $ui -PathType Leaf)) { throw "Build the source Office UI before launch." }
$runRoot = Join-Path $repo (".cache\companion-source-acceptance\" + $RunName)
if (Test-Path -LiteralPath $runRoot) { throw "Choose a fresh run name; existing acceptance artifacts are preserved." }
# Check the exact loopback endpoint before starting. Never stop an existing listener.
$probe = New-Object Net.Sockets.TcpClient
try {
    $pending = $probe.BeginConnect("127.0.0.1", $Port, $null, $null)
    if ($pending.AsyncWaitHandle.WaitOne(400)) {
        try { $probe.EndConnect($pending); throw "Acceptance port $Port is already in use." }
        catch [Net.Sockets.SocketException] { }
    }
} finally { $probe.Dispose() }
if (-not $Start) {
    [pscustomobject]@{ ReadyToLaunch = $true; Executable = $executable; Port = $Port; ArtifactRoot = $runRoot; IdentityRoot = $identity }
    return
}
New-Item -ItemType Directory -Path $runRoot | Out-Null
$appData = (New-Item -ItemType Directory -Path (Join-Path $runRoot "app-data")).FullName
$environment = [ordered]@{
    VISUALTEX_ACCEPTANCE_OFFICE_COMPANION_PORT = "$Port"
    VISUALTEX_ACCEPTANCE_OFFICE_APP_DATA_ROOT = $appData
    VISUALTEX_ACCEPTANCE_OFFICE_IDENTITY_ROOT = $identity
    VISUALTEX_ACCEPTANCE_OFFICE_IDENTITY_PORT = "43127"
    VISUALTEX_ACCEPTANCE_OFFICE_CERTIFICATE_PATH = $certificatePath
    VISUALTEX_ACCEPTANCE_OFFICE_CERTIFICATE_THUMBPRINT = $thumbprint
    VISUALTEX_ACCEPTANCE_OFFICE_PROTOCOL_VERSION = "1"
    VISUALTEX_ACCEPTANCE_OFFICE_EXECUTABLE_PATH = $executable
    VISUALTEX_ACCEPTANCE_OFFICE_UI_ROOT = $uiRoot
}
$previous = @{}
try {
    foreach ($key in $environment.Keys) {
        $previous[$key] = [Environment]::GetEnvironmentVariable($key, "Process")
        [Environment]::SetEnvironmentVariable($key, $environment[$key], "Process")
    }
    $process = Start-Process -FilePath $executable -ArgumentList "--office-background" -WorkingDirectory (Split-Path $executable) -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $runRoot "stdout.log") -RedirectStandardError (Join-Path $runRoot "stderr.log")
} finally {
    foreach ($key in $previous.Keys) { [Environment]::SetEnvironmentVariable($key, $previous[$key], "Process") }
}
$manifest = [ordered]@{
    ProcessId = $process.Id
    Executable = $executable
    ExecutableSha256 = (Get-FileHash -LiteralPath $executable -Algorithm SHA256).Hash
    StartedUtc = [DateTime]::UtcNow.ToString("o")
    Endpoint = "https://127.0.0.1:$Port"
    Environment = $environment
}
$manifestPath = Join-Path $runRoot "runtime.json"
$manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifestPath -Encoding UTF8
try {
    $deadline = [DateTime]::UtcNow.AddSeconds(40)
    $health = $null
    $lastError = "The source Companion has not started listening."
    do {
        if ($process.HasExited) { throw "Source Companion exited with code $($process.ExitCode); inspect $runRoot." }
        try {
            # Normal Windows trust validation remains enabled. A missing trusted
            # identity is a blocker, never a reason to bypass TLS or install trust.
            $request = [Net.HttpWebRequest]::Create($manifest.Endpoint + "/health")
            $request.Proxy = $null
            $request.Timeout = 2000
            $response = $request.GetResponse()
            try {
                $stream = New-Object IO.StreamReader($response.GetResponseStream())
                try { $health = $stream.ReadToEnd() | ConvertFrom-Json } finally { $stream.Dispose() }
            } finally { $response.Dispose() }
            if ($health.protocolVersion -ne 1) { throw "Companion health protocol mismatch." }
            # Verify the same endpoint's certificate with standard chain/name
            # validation plus the exact installed public-certificate pin.
            $tcp = New-Object Net.Sockets.TcpClient("127.0.0.1", $Port)
            try {
                $tls = New-Object Net.Security.SslStream($tcp.GetStream(), $false)
                try {
                    $tls.AuthenticateAsClient("127.0.0.1")
                    if ($tls.RemoteCertificate.GetCertHashString() -ne $thumbprint) { throw "Companion certificate pin mismatch." }
                } finally { $tls.Dispose() }
            } finally { $tcp.Dispose() }
            break
        } catch {
            $health = $null
            $lastError = $_.Exception.Message
            Start-Sleep -Milliseconds 250
        }
    } while ([DateTime]::UtcNow -lt $deadline)
    if ($null -eq $health) { throw "Source Companion readiness failed: $lastError" }
    [pscustomobject]@{ ProcessId = $process.Id; Endpoint = $manifest.Endpoint; Health = $health; RuntimeManifest = $manifestPath }
} catch {
    # This handle belongs only to the new process created above, never installed Companion.
    if (-not $process.HasExited) { $process.Kill(); $process.WaitForExit() }
    throw
} finally { $certificate.Dispose() }
