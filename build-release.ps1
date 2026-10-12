[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^\d+\.\d+\.\d+([-.][0-9A-Za-z.-]+)?$')]
    [string]$Version,

    [string]$DataverseConfig = '',
    [string]$ReleaseNotes = '',
    [string]$OutputDirectory = 'Releases'
)

$ErrorActionPreference = 'Stop'
$projectDirectory = $PSScriptRoot
$packDirectory = Join-Path $projectDirectory 'dist\LOA-Monitoring-Tool'
$releaseDirectory = Join-Path $projectDirectory $OutputDirectory

$env:LOA_BUILD_ONEDIR = '1'
if ($DataverseConfig) {
    $resolvedDataverseConfig = (Resolve-Path -LiteralPath $DataverseConfig).Path
    $env:LOA_DATAVERSE_CONFIG = $resolvedDataverseConfig
}
try {
    # Use python interpreter from environment
    & python -m PyInstaller --clean --noconfirm (Join-Path $projectDirectory 'LOA-Monitoring-Tool.spec')
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller failed with exit code $LASTEXITCODE." }
} finally {
    Remove-Item Env:LOA_BUILD_ONEDIR -ErrorAction SilentlyContinue
    Remove-Item Env:LOA_DATAVERSE_CONFIG -ErrorAction SilentlyContinue
}

if (-not (Test-Path -LiteralPath (Join-Path $packDirectory 'LOA-Monitoring-Tool.exe'))) {
    throw "Expected packaged application was not found in $packDirectory."
}
if (-not (Get-Command vpk -ErrorAction SilentlyContinue)) {
    throw 'Velopack CLI is missing. Install it once with: dotnet tool install -g vpk --version 1.2.0'
}

$packArgs = @(
    'pack',
    '--packId', 'LOAMonitoringTool',
    '--packVersion', $Version,
    '--packDir', $packDirectory,
    '--mainExe', 'LOA-Monitoring-Tool.exe',
    '--packTitle', 'LOA Monitoring Tool',
    '--packAuthors', 'Land Acquisition Team',
    '--channel', 'win',
    '--outputDir', $releaseDirectory,
    '--icon', (Join-Path $projectDirectory 'icon.ico')
)
if ($ReleaseNotes) {
    $notesPath = (Resolve-Path -LiteralPath $ReleaseNotes).Path
    $packArgs += @('--releaseNotes', $notesPath)
}

& vpk @packArgs
if ($LASTEXITCODE -ne 0) { throw "Velopack packaging failed with exit code $LASTEXITCODE." }

Write-Host "Release $Version created in $releaseDirectory"
Write-Host 'Distribute the generated Setup.exe for the first install, then publish every generated release file to the configured GitHub release.'
