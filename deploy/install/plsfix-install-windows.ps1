<#
Installs pls,fix into Excel and PowerPoint on this PC: downloads the add-in
manifest and registers it under HKCU\Software\Microsoft\Office\16.0\WEF\Developer,
the key Microsoft's own add-in tooling uses for sideloaded add-ins. No admin
rights, nothing else installed; the panes load from the pls,fix server.
Usage (PowerShell): irm https://github.com/bendiksdaniels/plsfix/releases/latest/download/plsfix-install-windows.ps1 | iex
Self-hosted:        $env:PLSFIX_MANIFEST_URL = "https://your.host/manifest.xml"  before the same line
#>
param(
  [string]$ManifestUrl = $(if ($env:PLSFIX_MANIFEST_URL) { $env:PLSFIX_MANIFEST_URL } else { "https://github.com/bendiksdaniels/plsfix/releases/latest/download/manifest.prod.xml" }),
  [switch]$DryRun
)
$ErrorActionPreference = "Stop"

$base = if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [System.IO.Path]::GetTempPath() }
$dir = Join-Path $base "plsfix"
New-Item -ItemType Directory -Force -Path $dir | Out-Null
$manifest = Join-Path $dir "manifest.xml"

Write-Host "Downloading the pls,fix manifest from $ManifestUrl"
Invoke-WebRequest -UseBasicParsing -Uri $ManifestUrl -OutFile $manifest
$xml = Get-Content -Raw -Path $manifest
if ($xml -notmatch "<OfficeApp") { throw "That file is not an Office add-in manifest: $ManifestUrl" }
if ($xml -notmatch "<Id>([0-9A-Fa-f-]{36})</Id>") { throw "The manifest carries no add-in id: $ManifestUrl" }
$id = $Matches[1]
$version = if ($xml -match "<Version>([0-9.]+)</Version>") { $Matches[1] } else { "" }

$key = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
if ($DryRun) {
  Write-Host "Dry run: would register value $id = $manifest under $key"
} else {
  New-Item -Path $key -Force | Out-Null
  New-ItemProperty -Path $key -Name $id -Value $manifest -PropertyType String -Force | Out-Null
}

Write-Host "pls,fix $version is installed for Excel and PowerPoint."
Write-Host "Close and reopen Excel and PowerPoint: the pls,fix tab appears on the ribbon, Ctrl+Shift+M opens the pane."
Write-Host "To remove it: irm https://github.com/bendiksdaniels/plsfix/releases/latest/download/plsfix-uninstall-windows.ps1 | iex"
