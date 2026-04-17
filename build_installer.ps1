Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

$possibleIscc = @(
  "$env:ProgramFiles(x86)\Inno Setup 6\ISCC.exe",
  "$env:ProgramFiles\Inno Setup 6\ISCC.exe"
)

$iscc = $null
foreach ($p in $possibleIscc) {
  if (Test-Path $p) {
    $iscc = $p
    break
  }
}

if (-not $iscc) {
  throw 'Inno Setup not found. Install Inno Setup 6 first, then rerun this script.'
}

if (-not (Test-Path 'dist\ELC_Packing_Tool_v1.0.exe')) {
  throw 'App build not found. Run build_windows_app.ps1 first.'
}

Write-Host 'Creating installer...'
& $iscc 'ELC_Packing_Tool_v1.0.iss'

Write-Host ''
Write-Host 'Installer created in dist_installer\ELC_Packing_Tool_v1.0_Setup.exe' -ForegroundColor Green
