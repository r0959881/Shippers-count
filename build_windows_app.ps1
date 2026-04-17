Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $scriptDir

function Invoke-WithPython {
  param(
    [Parameter(Mandatory = $true)]
    [string[]]$Args
  )

  $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
  if ($pythonCmd -and ($pythonCmd.Source -notmatch 'WindowsApps')) {
    & $pythonCmd.Source @Args
    if ($LASTEXITCODE -ne 0) { throw "Python command failed: python $($Args -join ' ')" }
    return
  }

  $pyLauncher = Get-Command py -ErrorAction SilentlyContinue
  if ($pyLauncher) {
    & $pyLauncher.Source -3 @Args
    if ($LASTEXITCODE -ne 0) { throw "Python command failed: py -3 $($Args -join ' ')" }
    return
  }

  $knownPythonPaths = @(
    'C:/Program Files/Python313/python.exe',
    "$env:LocalAppData/Programs/Python/Python313/python.exe"
  )

  foreach ($path in $knownPythonPaths) {
    if (Test-Path $path) {
      & $path @Args
      if ($LASTEXITCODE -ne 0) { throw "Python command failed: $path $($Args -join ' ')" }
      return
    }
  }

  throw 'Python was not found. Install Python 3 and rerun this script.'
}

Write-Host 'Installing/updating Python dependencies...'
Invoke-WithPython -Args @('-m', 'pip', 'install', '-r', 'requirements.txt')

Write-Host 'Cleaning old build output...'
if (Test-Path build) { Remove-Item build -Recurse -Force }
if (Test-Path dist) { Remove-Item dist -Recurse -Force }
if (Test-Path ELC_Packing_Tool_v1.0.spec) { Remove-Item ELC_Packing_Tool_v1.0.spec -Force }

Write-Host 'Building standalone Windows app...'
Invoke-WithPython -Args @(
  '-m', 'PyInstaller',
  '--noconfirm',
  '--clean',
  '--onefile',
  '--windowed',
  '--name', 'ELC_Packing_Tool_v1.0',
  '--collect-all', 'pandas',
  '--collect-all', 'openpyxl',
  'packing_tool.py'
)

Write-Host ''
Write-Host 'Build complete.' -ForegroundColor Green
Write-Host 'EXE file: dist\ELC_Packing_Tool_v1.0.exe'
Write-Host 'Run app: dist\ELC_Packing_Tool_v1.0.exe'
