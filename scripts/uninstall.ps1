[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
$AddinId = "4d5bf4c8-7851-4f6f-b595-66fc244898a4"
$InstallRoot = Join-Path $env:LOCALAPPDATA "Iconify Search for PowerPoint"
$DeveloperKey = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer"

Write-Host ""
Write-Host " Uninstall Iconify Search for PowerPoint" -ForegroundColor White
Write-Host " ---------------------------------------" -ForegroundColor DarkGray
Write-Host ""

if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
  Write-Host " Close every PowerPoint window, then press Enter." -ForegroundColor Yellow
  Read-Host | Out-Null
  if (Get-Process POWERPNT -ErrorAction SilentlyContinue) {
    throw "PowerPoint is still open. Close it and run the uninstaller again."
  }
}

if (Test-Path -LiteralPath $DeveloperKey) {
  Remove-ItemProperty -Path $DeveloperKey -Name $AddinId -ErrorAction SilentlyContinue
  New-ItemProperty -Path $DeveloperKey -Name "RefreshAddins" -Value 1 -PropertyType DWord -Force | Out-Null
}

$expectedRoot = [IO.Path]::GetFullPath((Join-Path $env:LOCALAPPDATA "Iconify Search for PowerPoint"))
$resolvedRoot = [IO.Path]::GetFullPath($InstallRoot)
if ($resolvedRoot -ne $expectedRoot -or [IO.Path]::GetFileName($resolvedRoot) -ne "Iconify Search for PowerPoint") {
  throw "Safety check failed for the installation directory."
}
if (Test-Path -LiteralPath $resolvedRoot) {
  Remove-Item -LiteralPath $resolvedRoot -Recurse -Force
}

Write-Host " Uninstallation complete." -ForegroundColor Green
Write-Host " The add-in can be installed again at any time."
Write-Host ""
