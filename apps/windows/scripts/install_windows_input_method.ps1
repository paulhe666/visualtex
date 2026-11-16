[CmdletBinding()]
param(
    [string]$SourceDirectory,
    [switch]$Uninstall,
    [switch]$Elevated
)

# Installs VisualTeX's Windows input method (formulas in any application,
# apps/windows/input-method) from the VisualTeX installation:
# - machine-wide, elevated (Windows loads only machine-wide text services): the files go to
#   %ProgramFiles%\VisualTeX Input\<build id> (a new folder per build, since applications keep
#   the previous DLL loaded) and the TIP is registered (regsvr32);
# - for the user: it is NOT one of the user's input methods. It is taken off the Chinese
#   (Simplified) language's keyboards (earlier builds were added there) and its profile is
#   disabled, so no input switcher offers it; VisualTeX's shortcut is the only way into it (the
#   input method enables its profile for the time of a formula). The earlier builds' pinyin
#   server is stopped.
# -Uninstall removes it, except when the VisualTeX installer is replacing this installation
# (it sets VISUALTEX_REINSTALL for the previous uninstaller): the input method stays, and the
# new installation's step keeps it or brings its new build. An input method that is already
# current is left alone (no UAC).

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

# The installer runs this script in 32-bit PowerShell (NSIS is 32-bit), where HKLM\SOFTWARE\Classes
# reads the WOW64 view; the TIP is a 64-bit COM server registered in the 64-bit view.
function Open-MachineClasses([bool]$Writable = $false) {
    $view = if ([Environment]::Is64BitOperatingSystem) { [Microsoft.Win32.RegistryView]::Registry64 } else { [Microsoft.Win32.RegistryView]::Default }
    $machine = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $view)
    return $machine.OpenSubKey("SOFTWARE\Classes\CLSID", $Writable)
}

function Get-RegisteredDll {
    $classes = Open-MachineClasses
    if ($null -eq $classes) { return $null }
    $key = $classes.OpenSubKey("$clsid\InprocServer32")
    if ($null -eq $key) { return $null }
    return [string]$key.GetValue("")
}

function Remove-Registration {
    $classes = Open-MachineClasses $true
    if ($null -ne $classes -and $null -ne $classes.OpenSubKey($clsid)) { $classes.DeleteSubKeyTree($clsid) }
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

# regsvr32 is a Windows program: "&" would not wait for it (nor give its exit code), and the
# registration read after it would still be the previous build's. The 64-bit one (Sysnative from
# the installer's 32-bit PowerShell) registers the 64-bit TIP itself.
function Invoke-Regsvr32([string[]]$Arguments) {
    $program = Join-Path $env:WINDIR "Sysnative\regsvr32.exe"
    if (-not (Test-Path $program)) { $program = Join-Path $env:WINDIR "System32\regsvr32.exe" }
    $process = Start-Process -FilePath $program -ArgumentList $Arguments -Wait -PassThru -WindowStyle Hidden
    try { return $process.ExitCode } finally { $process.Dispose() }
}

function Stop-InputServer {
    Get-Process -Name "VisualTeXInput" -ErrorAction SilentlyContinue | ForEach-Object {
        try { $_.Kill(); $_.WaitForExit(3000) | Out-Null } catch { }
    }
}

# VisualTeX is none of the user's input methods: off the language's keyboards (where earlier
# builds put it), and its profile disabled for the user (the input switcher lists enabled
# profiles; Windows keeps the user's choice under HKCU).
function Remove-UserKeyboard {
    $list = Get-WinUserLanguageList
    $chinese = $list | Where-Object { $_.LanguageTag -eq "zh-Hans-CN" } | Select-Object -First 1
    if ($null -ne $chinese -and $chinese.InputMethodTips -contains $tipId) {
        [void]$chinese.InputMethodTips.Remove($tipId)
        Set-WinUserLanguageList -LanguageList $list -Force
        Write-Log "taken off the user's keyboards"
    }
    $profileKey = "HKCU:\Software\Microsoft\CTF\TIP\$clsid\LanguageProfile\0x00000804\$profileGuid"
    if (-not (Test-Path $profileKey)) { New-Item -Path $profileKey -Force | Out-Null }
    Set-ItemProperty -Path $profileKey -Name "Enable" -Value 0 -Type DWord
}

# The elevated part: files and registration.
if ($Elevated) {
    try {
        if (-not (Test-IsAdministrator)) { throw "The VisualTeX input method needs administrator rights." }
        $registered = Get-RegisteredDll
        if ($Uninstall) {
            if ($registered -and (Test-Path $registered)) {
                $code = Invoke-Regsvr32 @("/u", "/s", (Quote-ProcessArgument $registered))
                Write-Log "unregistered $registered (exit $code)"
            }
            Remove-Registration
            # Folders still loaded by running applications go at the next install.
            Get-ChildItem -Path $installRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
            }
            exit 0
        }
        $buildId = Read-BuildId $SourceDirectory
        if (-not $buildId) { throw "The input method's build id is missing from $SourceDirectory" }
        $target = Join-Path $installRoot $buildId
        if ((Read-BuildId $target) -eq $buildId) {
            # This build's folder is already there (a reinstall of the same build): applications
            # may have its DLL loaded, so its files cannot be replaced, and need not be.
            Write-Log "files of $buildId already in place"
        } else {
            # An uninstall leaves the files applications still had loaded (this build's, so the
            # same): only the others are copied, the build id last.
            $source = (Get-Item -LiteralPath $SourceDirectory).FullName.TrimEnd('\')
            New-Item -Path $target -ItemType Directory -Force | Out-Null
            Get-ChildItem -LiteralPath $source -Recurse -File | Where-Object { $_.Name -ne "build-id" } | ForEach-Object {
                $destination = Join-Path $target $_.FullName.Substring($source.Length + 1)
                if ((Test-Path $destination) -and (Get-FileHash $destination).Hash -eq (Get-FileHash $_.FullName).Hash) { return }
                New-Item -Path (Split-Path -Parent $destination) -ItemType Directory -Force | Out-Null
                Copy-Item -Path $_.FullName -Destination $destination -Force
            }
            Copy-Item -Path (Join-Path $source "build-id") -Destination $target -Force
        }
        $code = Invoke-Regsvr32 @("/s", (Quote-ProcessArgument (Join-Path $target "visualtex_tip.dll")))
        Write-Log "registered $target (exit $code)"
        if ($code -ne 0) { exit $code }
        # Earlier builds' folders, unless an application still has their DLL loaded.
        Get-ChildItem -Path $installRoot -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne $buildId } | ForEach-Object {
            Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
        }
        exit 0
    } catch {
        Write-Log "elevated step failed at line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
        exit 1
    }
}

# The user's part.
try {
    if ($Uninstall) {
        if ($env:VISUALTEX_REINSTALL -eq "1") {
            Write-Log "uninstall for a reinstall: the input method stays"
            exit 0
        }
        Stop-InputServer
        Remove-UserKeyboard
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
    Stop-InputServer
    Remove-UserKeyboard
    $registered = Get-RegisteredDll
    Write-Log "input method ready ($registered)"
    exit 0
} catch {
    Write-Log "failed: $($_.Exception.Message)"
    Write-Output $_.Exception.Message
    exit 1
}
