[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$resources=Join-Path $root 'src-tauri\resources\windows-office'
$msi=Join-Path $resources 'VisualTeX-NativeRender-x64.msi'
$manifest=Get-Content -LiteralPath (Join-Path $resources 'VisualTeX-NativeRender-x64.sha256.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if ($manifest.schema -ne 1 -or $manifest.service -cne 'VisualTeX.NativeRender' -or
    $manifest.package.name -cne 'VisualTeX-NativeRender-x64.msi' -or
    (Get-FileHash -LiteralPath $msi -Algorithm SHA256).Hash -ine $manifest.package.sha256) {
    throw 'The isolated renderer MSI does not match its manifest.'
}
$build=Join-Path $root 'src-windows\VisualTeX.NativeRender.Installer\bin\x64\Release\VisualTeX-NativeRender-x64.msi'
if ((Get-FileHash -LiteralPath $build -Algorithm SHA256).Hash -ine $manifest.package.sha256) {
    throw 'Bundled native-render MSI is not the current build.'
}
$paths=@($manifest.files | ForEach-Object { [string]$_.path })
foreach($required in @('VisualTeX.NativeRenderService.exe','VisualTeX.NativeRenderWorker.exe','runtime\MathType.exe','runtime\Language\MT7chs.dll','runtime\MathPage\64\MathPage.wll','runtime\Fonts\TrueType\mtextra.ttf','runtime\Fonts\TrueType\eumat1.ttf','runtime\Fonts\TrueType\eumat2.ttf')) {
    if ($required -notin $paths) { throw "Native-render manifest is missing $required" }
}
$config=Get-Content -LiteralPath (Join-Path $root 'src-tauri\tauri.windows.conf.json') -Raw | ConvertFrom-Json
if ($config.bundle.resources.PSObject.Properties.Name -contains 'target/private-mathtype-runtime/') {
    throw 'The desktop bundle must not deploy another interactive MathType runtime.'
}
$installer=New-Object -ComObject WindowsInstaller.Installer
$database=$null
function Read-Rows([string]$query,[int]$columns) {
    $view=$database.OpenView($query)
    try {
        [void]$view.Execute()
        while($null -ne ($record=$view.Fetch())) {
            try {
                $values=@(for($i=1;$i -le $columns;$i++) { [string]$record.StringData($i) })
                ,$values
            } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($record) }
        }
    } finally { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($view) }
}
try {
    $database=$installer.OpenDatabase($msi,0)
    $services=@(Read-Rows 'SELECT `Name`, `StartName`, `StartType`, `ServiceType` FROM `ServiceInstall`' 4)
    if ($services.Count -ne 1 -or $services[0][0] -cne 'VisualTeX.NativeRender' -or
        $services[0][1] -cne 'NT SERVICE\VisualTeX.NativeRender' -or $services[0][2] -ne '2' -or $services[0][3] -ne '16') {
        throw 'MSI does not install the tested noninteractive virtual-account service.'
    }
    $permissions=@(Read-Rows 'SELECT `LockObject`, `SDDLText` FROM `MsiLockPermissionsEx`' 2)
    foreach($folder in @('NativeRenderData','NativeRenderJobs')) {
        $rows=@($permissions | Where-Object { $_[0] -ceq $folder })
        if ($rows.Count -ne 1 -or $rows[0][1] -cne 'D:P(A;OICI;FA;;;SY)(A;OICI;FA;;;BA)(A;OICI;FA;;;S-1-5-80-3568014781-357338252-412558032-1798669891-843860098)') {
            throw "MSI lacks the explicit private service ACL for $folder."
        }
    }
    foreach($row in @(Read-Rows 'SELECT `Key` FROM `Registry`' 1)) {
        if ($row[0] -cne 'Software\VisualTeX\NativeRender') { throw 'Renderer MSI must not change public MathType/COM or user preferences.' }
    }
    # MathType's own fonts: added only where missing, never overwritten, never
    # removed, so an installed MathType (or its later install) is not affected.
    $fonts=@(Read-Rows 'SELECT `File_` FROM `Font`' 1)
    if ($fonts.Count -ne 23) { throw "Renderer MSI must carry MathType's 23 TrueType fonts; found $($fonts.Count)." }
    $components=@(Read-Rows 'SELECT `Component`, `Attributes`, `Condition`, `Directory_` FROM `Component`' 4)
    foreach($font in $fonts) {
        $file=@(Read-Rows ('SELECT `Component_` FROM `File` WHERE `File`=''' + $font[0] + '''') 1)
        $component=@($components | Where-Object { $_[0] -ceq $file[0][0] })
        if ($component.Count -ne 1 -or $component[0][3] -cne 'FontsFolder' -or
            ([int]$component[0][1] -band 0x10) -eq 0 -or ([int]$component[0][1] -band 0x80) -eq 0 -or
            $component[0][2] -notmatch '^NOT FONT_[A-Z0-9]+$') {
            throw "Font component $($file[0][0]) must be conditional, permanent and never overwrite."
        }
    }
} finally {
    foreach($value in @($database,$installer)) { if ($null -ne $value) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value) } }
}
Write-Output 'Native-render package identity, service account and private directory ACLs verified.'
