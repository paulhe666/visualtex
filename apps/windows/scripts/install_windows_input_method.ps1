[CmdletBinding()]
param(
    [string]$SourceDirectory,
    [switch]$Uninstall,
    [switch]$Elevated
)

# Installs VisualTeX's Windows input method (pinyin, English and formulas in any application,
# apps/windows/input-method) from the VisualTeX installation:
# - machine-wide, elevated (Windows lists only machine-wide text services): the files go to
#   %ProgramFiles%\VisualTeX Input\<build id> (a new folder per build, since applications keep
#   the previous DLL loaded) and the TIP is registered (regsvr32);
# - for the user: the input method is added to the Chinese (Simplified) language's keyboards
#   and VisualTeX Input (its RIME server) is started.
# -Uninstall reverses both. An input method that is already current is left alone (no UAC).

$ErrorActionPreference = "Stop"
$clsid = "{B6E0A6C4-6F52-4C7E-9B7A-2F1E5A3D9C41}"
$profileGuid = "{D3A1F7B2-8C4E-4E2A-A5D6-6B9E0C1F2A73}"
$tipId = "0804:$clsid$profileGuid"
$installRoot = Join-Path $env:ProgramFiles "VisualTeX Input"
$logRoot = Join-Path $env:LOCALAPPDATA "VisualTeX"
New-Item -Path $logRoot -ItemType Directory -Force | Out-Null
$logPath = Join-Path $logRoot "input-method-install.log"

function Write-Log([string]$Line) {
    $stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logPath -Value "$stamp $Line" -Encoding UTF8
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Quote-ProcessArgument([string]$Value) {
    return '"' + $Value.Replace('"', '\"') + '"'
}

function Get-RegisteredDll {
    $key = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\$clsid\InprocServer32"
    if (-not (Test-Path $key)) { return $null }
    return [string](Get-ItemProperty -Path $key).'(default)'
}

function Read-BuildId([string]$Directory) {
    $file = Join-Path $Directory "build-id"
    if (-not (Test-Path $file)) { return $null }
    return (Get-Content -Path $file -TotalCount 1).Trim()
}

function Invoke-Elevated {
    $arguments = @("-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", (Quote-ProcessArgument $PSCommandPath), "-Elevated")
    if ($Uninstall) { $arguments += "-Uninstall" }
    if ($SourceDirectory) { $arguments += @("-SourceDirectory", (Quote-ProcessArgument ([IO.Path]::GetFullPath($SourceDirectory)))) }
    $process = Start-Process -FilePath "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -ArgumentList ($arguments -join " ") -Verb RunAs -WindowStyle Hidden -PassThru
    try {
        $process.WaitForExit()
        return $process.ExitCode
    } finally {
        $process.Dispose()
    }
}

function Stop-InputServer {
    Get-Process -Name "VisualTeXInput" -ErrorAction SilentlyContinue | ForEach-Object {
        try { $_.Kill(); $_.WaitForExit(3000) | Out-Null } catch { }
    }
}

function Set-UserKeyboard([bool]$Present) {
    $list = Get-WinUserLanguageList
    $chinese = $list | Where-Object { $_.LanguageTag -eq "zh-Hans-CN" } | Select-Object -First 1
    if ($Present) {
        if ($null -eq $chinese) {
            $list.Add("zh-Hans-CN")
            $chinese = $list | Where-Object { $_.LanguageTag -eq "zh-Hans-CN" } | Select-Object -First 1
        }
        if ($chinese.InputMethodTips -notcontains $tipId) { $chinese.InputMethodTips.Add($tipId) }
    } else {
        if ($null -eq $chinese -or $chinese.InputMethodTips -notcontains $tipId) { return }
        [void]$chinese.InputMethodTips.Remove($tipId)
    }
    Set-WinUserLanguageList -LanguageList $list -Force
}

# The elevated part: files and registration.
if ($Elevated) {
    if (-not (Test-IsAdministrator)) { throw "The VisualTeX input method needs administrator rights." }
    $registered = Get-RegisteredDll
    if ($Uninstall) {
        if ($registered -and (Test-Path $registered)) {
            & "$env:WINDIR\System32\regsvr32.exe" /u /s $registered
            Write-Log "unregistered $registered (exit $LASTEXITCODE)"
        }
        Remove-Item -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\$clsid" -Recurse -ErrorAction SilentlyContinue
        # Folders still loaded by running applications go at the next install.
        Get-ChildItem -Path $installRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
        }
        exit 0
    }
    $buildId = Read-BuildId $SourceDirectory
    if (-not $buildId) { throw "The input method's build id is missing from $SourceDirectory" }
    $target = Join-Path $installRoot $buildId
    New-Item -Path $target -ItemType Directory -Force | Out-Null
    Copy-Item -Path (Join-Path $SourceDirectory "*") -Destination $target -Recurse -Force
    & "$env:WINDIR\System32\regsvr32.exe" /s (Join-Path $target "visualtex_tip.dll")
    $code = $LASTEXITCODE
    Write-Log "registered $target (exit $code)"
    if ($code -ne 0) { exit $code }
    # Earlier builds' folders, unless an application still has their DLL loaded.
    Get-ChildItem -Path $installRoot -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne $buildId } | ForEach-Object {
        Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
    }
    exit 0
}

# The user's part.
try {
    if ($Uninstall) {
        Stop-InputServer
        Set-UserKeyboard $false
        if (Get-RegisteredDll) {
            $code = Invoke-Elevated
            Write-Log "uninstall elevated exit $code"
        }
        exit 0
    }
    if (-not $SourceDirectory -or -not (Test-Path (Join-Path $SourceDirectory "visualtex_tip.dll"))) {
        throw "The input method is missing from the VisualTeX installation ($SourceDirectory)."
    }
    $buildId = Read-BuildId $SourceDirectory
    $registered = Get-RegisteredDll
    $current = $registered -and (Test-Path $registered) -and ((Read-BuildId (Split-Path -Parent $registered)) -eq $buildId)
    if (-not $current) {
        Stop-InputServer
        $code = Invoke-Elevated
        Write-Log "install elevated exit $code"
        if ($code -ne 0) { throw "Registering the input method failed or was cancelled (exit $code)." }
    } else {
        Write-Log "input method $buildId is current"
    }
    Set-UserKeyboard $true
    $registered = Get-RegisteredDll
    $server = Join-Path (Split-Path -Parent $registered) "VisualTeXInput.exe"
    if (-not (Get-Process -Name "VisualTeXInput" -ErrorAction SilentlyContinue) -and (Test-Path $server)) {
        Start-Process -FilePath $server
    }
    Write-Log "input method ready ($registered)"
    exit 0
} catch {
    Write-Log "failed: $($_.Exception.Message)"
    Write-Output $_.Exception.Message
    exit 1
}
