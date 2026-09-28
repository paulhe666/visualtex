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
foreach($required in @('VisualTeX.NativeRenderService.exe','VisualTeX.NativeRenderWorker.exe','runtime\MathType.exe','runtime\Language\MT7chs.dll','runtime\MathPage\64\MathPage.wll')) {
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
} finally {
    foreach($value in @($database,$installer)) { if ($null -ne $value) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value) } }
}
Write-Output 'Native-render package identity, service account and private directory ACLs verified.'
