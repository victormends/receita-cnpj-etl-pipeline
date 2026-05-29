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
$configScript = Join-Path $projectRoot 'config.ps1'
$preflightScript = Join-Path $projectRoot 'scripts\lib\preflight.ps1'
$defaultEnrichmentScript = Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads\Clientes_A_limpo_cnpj_corrigido.csv'
$resolvedIcon = (Resolve-Path -LiteralPath $IconPath).Path
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

if (-not (Test-Path -LiteralPath $configScript)) {
    throw "Config script not found: $configScript"
}

if (-not (Test-Path -LiteralPath $preflightScript)) {
    throw "Preflight script not found: $preflightScript"
}

if (-not (Test-Path -LiteralPath $resolvedOutputDir)) {
    New-Item -ItemType Directory -Path $resolvedOutputDir | Out-Null
}

$outputExe = Join-Path $resolvedOutputDir 'CnpjClientClassifier.exe'

Invoke-ps2exe `
    -inputFile $entryScript `
    -outputFile $outputExe `
    -iconFile $resolvedIcon `
    -title 'CNPJ Client Classifier' `
    -description 'Interactive Windows launcher for classifying clients by Receita Simples data.' `
    -company 'OpenCode' `
    -product 'CNPJ Client Classifier' `
    -copyright 'Copyright (c) 2026' `
    -version '1.0.0.0'

$captureScript        = Join-Path $projectRoot 'scripts\capture-active-clients.ps1'
$xmlAuditorScript     = Join-Path $projectRoot 'scripts\audit-xml-regime.ps1'
$recentDiffScript     = Join-Path $projectRoot 'scripts\export-recent-xml-regime-diff.ps1'
$consultaExportScript = Join-Path $projectRoot 'scripts\export-consulta-optantes-list.ps1'

Copy-Item -LiteralPath $classifierScript -Destination (Join-Path $resolvedOutputDir 'classify-clientes.ps1') -Force
if (Test-Path -LiteralPath $stagingBuilderScript -PathType Leaf) {
    Copy-Item -LiteralPath $stagingBuilderScript -Destination (Join-Path $resolvedOutputDir 'build-client-staging.ps1') -Force
}
if (Test-Path -LiteralPath $captureScript -PathType Leaf) {
    Copy-Item -LiteralPath $captureScript -Destination (Join-Path $resolvedOutputDir 'capture-active-clients.ps1') -Force
}
Copy-Item -LiteralPath $configScript -Destination (Join-Path $resolvedOutputDir 'config.ps1') -Force

# Copy XML auditor scripts (needed for automatic recent-divergence report)
# These go alongside classify-clientes.ps1 in the dist root so $scriptRoot lookup works.
if (Test-Path -LiteralPath $xmlAuditorScript -PathType Leaf) {
    Copy-Item -LiteralPath $xmlAuditorScript -Destination (Join-Path $resolvedOutputDir 'audit-xml-regime.ps1') -Force
}
if (Test-Path -LiteralPath $recentDiffScript -PathType Leaf) {
    Copy-Item -LiteralPath $recentDiffScript -Destination (Join-Path $resolvedOutputDir 'export-recent-xml-regime-diff.ps1') -Force
}
if (Test-Path -LiteralPath $consultaExportScript -PathType Leaf) {
    Copy-Item -LiteralPath $consultaExportScript -Destination (Join-Path $resolvedOutputDir 'export-consulta-optantes-list.ps1') -Force
}

if (Test-Path -LiteralPath $defaultEnrichmentScript -PathType Leaf) {
    $dataOutputDir = Join-Path $resolvedOutputDir 'data'
    if (-not (Test-Path -LiteralPath $dataOutputDir)) {
        New-Item -ItemType Directory -Path $dataOutputDir | Out-Null
    }
    Copy-Item -LiteralPath $defaultEnrichmentScript -Destination (Join-Path $dataOutputDir 'Clientes_A_limpo_cnpj_corrigido.csv') -Force
    Write-Host "Copied default enrichment CSV: $(Join-Path $dataOutputDir 'Clientes_A_limpo_cnpj_corrigido.csv')"
}

$libOutputDir = Join-Path $resolvedOutputDir 'lib'
if (-not (Test-Path -LiteralPath $libOutputDir)) {
    New-Item -ItemType Directory -Path $libOutputDir | Out-Null
}
Copy-Item -LiteralPath $preflightScript -Destination (Join-Path $libOutputDir 'preflight.ps1') -Force

Write-Host "Created executable: $outputExe"
Write-Host "Copied classifier script: $(Join-Path $resolvedOutputDir 'classify-clientes.ps1')"
Write-Host "Copied staging builder script: $(Join-Path $resolvedOutputDir 'build-client-staging.ps1')"
Write-Host "Copied config script: $(Join-Path $resolvedOutputDir 'config.ps1')"
Write-Host "Copied preflight script: $(Join-Path $libOutputDir 'preflight.ps1')"
