param([string]$OutputRoot)
$ErrorActionPreference = 'Stop'
if (!$OutputRoot) { $OutputRoot = Join-Path $env:TEMP ('visualtex-pipe-ownership-' + [guid]::NewGuid().ToString('N')) }
[void][IO.Directory]::CreateDirectory($OutputRoot)
$source = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot '../src-tauri/src/office/windows_pipe.rs')
$end = $source.IndexOf('pub fn locate_sidecar(')
if ($end -lt 0) { throw 'Sidecar source boundary changed.' }
# Copy the real client unchanged, omitting only unrelated Tauri path discovery.
$tests = Get-Content -Raw -LiteralPath (Join-Path $PSScriptRoot 'ownership/pipe_tests.rs')
[IO.File]::WriteAllText((Join-Path $OutputRoot 'lib.rs'), $source.Substring(0, $end) + $tests)
$manifest = @'
[package]
name = "visualtex-pipe-ownership"
version = "0.0.0"
edition = "2021"
[workspace]
[lib]
path = "lib.rs"
[dependencies]
getrandom = "0.3"
serde = { version = "1", features = ["derive"] }
serde_json = "1"
hex = "0.4"
uuid = { version = "1", features = ["v4"] }
'@
[IO.File]::WriteAllText((Join-Path $OutputRoot 'Cargo.toml'), $manifest)
& cargo test --offline --manifest-path (Join-Path $OutputRoot 'Cargo.toml') -- --test-threads=1
if ($LASTEXITCODE -ne 0) { throw "Pipe ownership test failed: $LASTEXITCODE" }
