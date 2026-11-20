[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Write-Warning "uninstall_windows_ole.ps1 is retired. Current VisualTeX releases do not install Office.js Trusted Catalog integration. Use uninstall_windows_vsto.ps1 to remove the native Ribbon + OLE LocalServer integration."
