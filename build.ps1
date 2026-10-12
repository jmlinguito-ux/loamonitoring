# Builds dist\LOA-Monitoring-Tool.exe from a clean virtual environment, so only
# the packages in requirements.txt end up in the exe (not torch, scipy, ... from
# the global Python install).
#
#   powershell -ExecutionPolicy Bypass -File .\build.ps1
#
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot

$venv = Join-Path $PSScriptRoot '.build-venv'
if (-not (Test-Path (Join-Path $venv 'Scripts\python.exe'))) {
    Write-Host 'Creating clean build environment (.build-venv)...'
    python -m venv $venv
}
$py = Join-Path $venv 'Scripts\python.exe'

& $py -m pip install --upgrade pip
& $py -m pip install -r requirements.txt pyinstaller

Write-Host 'Running tests...'
& $py -m unittest discover -s tests -p "test_*.py"
if ($LASTEXITCODE -ne 0) { throw 'Tests failed; the exe was not built.' }

& $py -m PyInstaller LOA-Monitoring-Tool.spec --noconfirm --clean
if ($LASTEXITCODE -ne 0) { throw 'PyInstaller failed.' }

$exe = Join-Path $PSScriptRoot 'dist\LOA-Monitoring-Tool.exe'
'{0:N1} MB  {1}' -f ((Get-Item $exe).Length / 1MB), $exe
