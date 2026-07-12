Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-VisualTeXExecutable {
    if ($env:VISUALTEX_BIN) { return $env:VISUALTEX_BIN }
    return "visualtex"
}

function Invoke-VisualTeX {
    param([Parameter(Mandatory = $true)][string[]]$Arguments)
    $executable = Get-VisualTeXExecutable
    $previousConsoleEncoding = [Console]::OutputEncoding
    $previousOutputEncoding = $OutputEncoding
    $previousErrorActionPreference = $ErrorActionPreference
    try {
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        [Console]::OutputEncoding = $utf8
        $OutputEncoding = $utf8
        $ErrorActionPreference = "Continue"
        & $executable @Arguments
        $exitCode = $LASTEXITCODE
        if ($exitCode -ne 0) {
            throw "VisualTeX command failed with exit code $exitCode"
        }
    }
    finally {
        [Console]::OutputEncoding = $previousConsoleEncoding
        $OutputEncoding = $previousOutputEncoding
        $ErrorActionPreference = $previousErrorActionPreference
    }
}

function Invoke-VisualTeXJsonToFile {
    param(
        [Parameter(Mandatory = $true)][string[]]$Arguments,
        [Parameter(Mandatory = $true)][string]$OutputJson
    )
    $executable = Get-VisualTeXExecutable
    $temporary = $OutputJson + ".tmp." + $PID
    $previousConsoleEncoding = [Console]::OutputEncoding
    $previousOutputEncoding = $OutputEncoding
    $previousErrorActionPreference = $ErrorActionPreference
    try {
        $utf8 = New-Object System.Text.UTF8Encoding($false)
        [Console]::OutputEncoding = $utf8
        $OutputEncoding = $utf8
        $ErrorActionPreference = "Continue"
        & $executable @Arguments | Set-Content -Encoding UTF8 $temporary
        $exitCode = $LASTEXITCODE
        if ($exitCode -ne 0) {
            throw "VisualTeX command failed with exit code $exitCode"
        }
        Move-Item -Force $temporary $OutputJson
    }
    finally {
        [Console]::OutputEncoding = $previousConsoleEncoding
        $OutputEncoding = $previousOutputEncoding
        $ErrorActionPreference = $previousErrorActionPreference
        Remove-Item -Force $temporary -ErrorAction SilentlyContinue
    }
}

function Test-VisualTeXBridge {
    param(
        [Parameter(Mandatory = $true)][string]$Executable,
        [Parameter(Mandatory = $true)][string]$ProjectRoot
    )
    $previousErrorActionPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        & $Executable bridge-request $ProjectRoot initialize --params "{}" 1> $null 2> $null
        return $LASTEXITCODE -eq 0
    }
    finally {
        $ErrorActionPreference = $previousErrorActionPreference
    }
}

function Ensure-VisualTeXBridge {
    param([Parameter(Mandatory = $true)][string]$ProjectRoot)
    $executable = Get-VisualTeXExecutable
    if (Test-VisualTeXBridge -Executable $executable -ProjectRoot $ProjectRoot) { return }

    $bridgeDirectory = Join-Path $ProjectRoot ".visualtex\bridge"
    New-Item -ItemType Directory -Force -Path $bridgeDirectory | Out-Null
    $logFile = Join-Path $bridgeDirectory "texstudio-bridge.log"
    $quotedProject = '"' + $ProjectRoot.Replace('"', '\"') + '"'
    Start-Process -FilePath $executable `
        -ArgumentList @("bridge-serve", $quotedProject) `
        -RedirectStandardOutput $logFile `
        -RedirectStandardError ($logFile + ".error") `
        -WindowStyle Hidden | Out-Null

    for ($attempt = 0; $attempt -lt 50; $attempt++) {
        Start-Sleep -Milliseconds 100
        if (Test-VisualTeXBridge -Executable $executable -ProjectRoot $ProjectRoot) { return }
    }
    throw "VisualTeX bridge did not become ready. See $logFile"
}
