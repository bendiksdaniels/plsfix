<#
Removes pls,fix from Excel and PowerPoint on this PC: deletes the registration
the installer wrote under HKCU\Software\Microsoft\Office\16.0\WEF\Developer and
the downloaded manifest. Office forgets the add-in on its next launch.
Usage (PowerShell): irm https://github.com/bendiksdaniels/plsfix/releases/latest/download/plsfix-uninstall-windows.ps1 | iex
#>
param([switch]$DryRun)
$ErrorActionPreference = "Stop"

$base = if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [System.IO.Path]::GetTempPath() }
$dir = Join-Path $base "plsfix"
$manifest = Join-Path $dir "manifest.xml"
$key = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"

$removed = 0
if (-not $DryRun -and (Test-Path $key)) {
  $values = Get-ItemProperty -Path $key
  foreach ($property in $values.PSObject.Properties) {
    if ($property.Value -eq $manifest) {
      Remove-ItemProperty -Path $key -Name $property.Name
      $removed += 1
    }
  }
}
if (Test-Path $dir) {
  if ($DryRun) { Write-Host "Dry run: would delete $dir and its registration under $key" }
  else { Remove-Item -Recurse -Force -Path $dir }
}

if ($removed -eq 0 -and -not $DryRun) {
  Write-Host "pls,fix was not installed by the installer on this PC (nothing to remove)."
} else {
  Write-Host "pls,fix is removed. Close and reopen Excel and PowerPoint."
  Write-Host "If the tab lingers, clear the Office add-in cache: delete the contents of $env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\ and relaunch."
}
