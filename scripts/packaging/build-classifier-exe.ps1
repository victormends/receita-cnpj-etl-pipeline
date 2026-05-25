param(
    [string]$IconPath = (Join-Path $PSScriptRoot 'app.ico'),
    [string]$OutputDir = 'dist'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$entryScript = Join-Path $projectRoot 'scripts\classify-clientes-launcher.ps1'
$classifierScript = Join-Path $projectRoot 'scripts\classify-clientes.ps1'
$stagingBuilderScript = Join-Path $projectRoot 'scripts\build-client-staging.ps1'
$configTemplate = Join-Path $projectRoot 'config.example.ps1'
$preflightScript = Join-Path $projectRoot 'scripts\lib\preflight.ps1'
$resolvedIcon = if (Test-Path -LiteralPath $IconPath -PathType Leaf) { (Resolve-Path -LiteralPath $IconPath).Path } else { $null }
$resolvedOutputDir = if ([System.IO.Path]::IsPathRooted($OutputDir)) {
    [System.IO.Path]::GetFullPath($OutputDir)
}
else {
    [System.IO.Path]::GetFullPath((Join-Path $projectRoot $OutputDir))
}

if (-not (Test-Path -LiteralPath $entryScript)) {
    throw "Launcher script not found: $entryScript"
}

if (-not (Test-Path -LiteralPath $classifierScript)) {
    throw "Classifier script not found: $classifierScript"
}

if (-not (Test-Path -LiteralPath $configTemplate)) {
    throw "Config template not found: $configTemplate"
}

if (-not (Test-Path -LiteralPath $preflightScript)) {
    throw "Preflight script not found: $preflightScript"
}

if (-not (Test-Path -LiteralPath $resolvedOutputDir)) {
    New-Item -ItemType Directory -Path $resolvedOutputDir | Out-Null
}

$outputExe = Join-Path $resolvedOutputDir 'CnpjClientClassifier.exe'

$ps2exeArgs = @{
    inputFile = $entryScript
    outputFile = $outputExe
    title = 'CNPJ Client Classifier'
    description = 'Interactive Windows launcher for classifying clients by Receita Simples data.'
    company = 'OpenCode'
    product = 'CNPJ Client Classifier'
    copyright = 'Copyright (c) 2026'
    version = '1.0.0.0'
}
if ($resolvedIcon) { $ps2exeArgs.iconFile = $resolvedIcon }

Invoke-ps2exe @ps2exeArgs

Copy-Item -LiteralPath $classifierScript -Destination (Join-Path $resolvedOutputDir 'classify-clientes.ps1') -Force
if (Test-Path -LiteralPath $stagingBuilderScript -PathType Leaf) {
    Copy-Item -LiteralPath $stagingBuilderScript -Destination (Join-Path $resolvedOutputDir 'build-client-staging.ps1') -Force
}
Copy-Item -LiteralPath $configTemplate -Destination (Join-Path $resolvedOutputDir 'config.ps1') -Force

$libOutputDir = Join-Path $resolvedOutputDir 'lib'
if (-not (Test-Path -LiteralPath $libOutputDir)) {
    New-Item -ItemType Directory -Path $libOutputDir | Out-Null
}
Copy-Item -LiteralPath $preflightScript -Destination (Join-Path $libOutputDir 'preflight.ps1') -Force

Write-Host "Created executable: $outputExe"
Write-Host "Copied classifier script: $(Join-Path $resolvedOutputDir 'classify-clientes.ps1')"
Write-Host "Copied staging builder script: $(Join-Path $resolvedOutputDir 'build-client-staging.ps1')"
Write-Host "Copied config template: $(Join-Path $resolvedOutputDir 'config.example.ps1')"
Write-Host "Copied preflight script: $(Join-Path $libOutputDir 'preflight.ps1')"
