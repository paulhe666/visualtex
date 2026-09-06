param([string]$Label='stage04ai')
$ErrorActionPreference='Stop'
if($Label -notmatch '^[a-zA-Z0-9-]+$'){throw 'Invalid evidence label.'}
$root=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$source=Join-Path $root 'apps\windows\dist-office-windows-native'
$installedRoot=Join-Path $env:LOCALAPPDATA 'VisualTeX'
$executable=Join-Path $installedRoot 'visualtex.exe'
$destination=Join-Path $installedRoot 'office'
$registered=(Get-ItemProperty 'HKCU:\Software\VisualTeX\OfficeIntegration').ExecutablePath
if($registered -ne $executable){throw 'The current installed companion path differs.'}
$companions=@(Get-Process visualtex -ErrorAction Stop)
if($companions.Count -ne 1 -or $companions[0].Path -ne $executable){throw 'Expected one currently installed companion.'}
$visible=& (Join-Path $PSScriptRoot 'word-ui.ps1') -Action windows | ConvertFrom-Json
$editors=@($visible | Where-Object {$_.process -eq 'visualtex' -and $_.title -ne 'com.visualtex.studio-siw'})
if($editors.Count){throw 'An actual editor/main window is open; leave it untouched.'}
$backup=Join-Path $PSScriptRoot ('evidence\'+$Label+'-office-ui-backup')
if(Test-Path -LiteralPath $backup){throw 'Candidate evidence already exists; inspect instead of repeating.'}
[void][IO.Directory]::CreateDirectory($backup)
$records=@()
# Assets first, entry HTML last. Existing hashed assets are retained. Only these
# generated source files are copied; no installed directory is deleted or moved.
$files=@(Get-ChildItem -LiteralPath $source -File -Recurse | Sort-Object @{Expression={$_.Extension -eq '.html'}},FullName)
foreach($file in $files){
 $relative=$file.FullName.Substring($source.Length+1)
 if($relative -notmatch '^(assets\\[^\\]+|dialog\\index\.html)$'){throw "Unexpected Office UI resource: $relative"}
 $target=[IO.Path]::GetFullPath((Join-Path $destination $relative))
 if(-not $target.StartsWith($destination+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Target escaped the installed Office UI directory.'}
 $before=$null
 if(Test-Path -LiteralPath $target){
  $before=(Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
  $backupFile=Join-Path $backup $relative
  [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($backupFile))
  Copy-Item -LiteralPath $target -Destination $backupFile
 }
 $expected=(Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
 [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($target))
 if($before -ne $expected){Copy-Item -LiteralPath $file.FullName -Destination $target -Force}
 $after=(Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
 if($after -ne $expected){throw "Installed UI hash mismatch: $relative"}
 $records+=@{file=$relative;before=$before;after=$after}
}
$record=@{label=$Label;time=[DateTime]::UtcNow.ToString('o');source=$source;destination=$destination;companionPid=$companions[0].Id;companionSha256=(Get-FileHash -LiteralPath $executable -Algorithm SHA256).Hash;backup=$backup;files=$records;nsisInstalled=$false;wordRestarted=$false}
$record | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $PSScriptRoot ('evidence\'+$Label+'-office-ui-installed.json')) -Encoding UTF8
Write-Output ('OFFICE_UI_CANDIDATE_INSTALLED|'+$files.Count+'|'+$destination)
