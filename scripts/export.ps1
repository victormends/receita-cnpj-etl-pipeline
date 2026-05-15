# =============================================================
# scripts/export.ps1 - Export filtered data to CSV
# =============================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot 'lib\preflight.ps1')
$ScriptDir = Resolve-ScriptRoot -Invocation $MyInvocation -FallbackBaseDirectory $PSScriptRoot

$bootstrap = Import-CnpjConfig -ScriptRoot $ScriptDir
$ProjectRoot = $bootstrap.ProjectRoot
$config = $bootstrap.Config

Write-Step 'Export preflight checks'
$tools = Invoke-CnpjPreflight -Config $config -EnsureDirectories @($config.dirTemp, $config.dirOut) -MinFreeSpaceGB @{ $config.dirOut = 2 } -RequirePostgres -RequireDatabase
$psqlPath = $tools.PsqlPath

Write-Host "===================================================" -ForegroundColor Cyan
Write-Host " EXPORT - Generating final CSV file" -ForegroundColor Cyan
Write-Host "===================================================" -ForegroundColor Cyan

$outputFile = Join-Path $config.dirOut "cnpj_export_$($config.anoMes).csv"
$outputFilePsql = ConvertTo-PsqlPathLiteral -Path $outputFile

Write-Host " Destination: $outputFile" -ForegroundColor Yellow
Write-Host " Exporting..." -NoNewline

Set-PostgresPasswordEnv -Config $config

if (Test-Path -LiteralPath $outputFile) {
    Remove-CnpjArtifactSafe -Config $config -Path $outputFile -AllowedRoot $config.dirOut -AllowedPatterns @('cnpj_export_*.csv') -Reason 'replace previous export file' | Out-Null
}

$exportSelect = "SELECT c.cnpj, c.razao_social, c.nome_fantasia, c.data_inicio_atividade, c.cnae_fiscal_principal, c.cnae_fiscal_secundaria, c.tipo_logradouro, c.logradouro, c.numero, c.complemento, c.bairro, c.cep, c.uf, c.municipio_nome, c.telefone_1, c.telefone_2, c.email, c.regime_tributario FROM estabelecimentos_crm c ORDER BY c.data_inicio_atividade DESC"
$sql = "\copy ($exportSelect) TO '$outputFilePsql' WITH (FORMAT CSV, HEADER TRUE, DELIMITER ';', ENCODING 'UTF8');"

$sqlFile = Join-Path $config.dirTemp "export_query.sql"
Set-Content -Path $sqlFile -Value $sql -Encoding UTF8

try {
    $exportTimeout = if ($config.ContainsKey('sqlCommandTimeoutSeconds')) { [int]$config.sqlCommandTimeoutSeconds } else { 3600 }
    Invoke-PsqlFileChecked -PsqlPath $psqlPath -Config $config -FilePath $sqlFile -Description 'CSV export' -TimeoutSec $exportTimeout
}
catch {
    Write-Host " ERROR during export: $($_.Exception.Message)" -ForegroundColor Red
    if (Test-Path -LiteralPath $sqlFile) {
        Remove-CnpjArtifactSafe -Config $config -Path $sqlFile -AllowedRoot $config.dirTemp -AllowedPatterns @('export_query.sql') -Reason 'failed export SQL file' | Out-Null
    }
    exit 1
}

if (Test-Path -LiteralPath $sqlFile) {
    Remove-CnpjArtifactSafe -Config $config -Path $sqlFile -AllowedRoot $config.dirTemp -AllowedPatterns @('export_query.sql') -Reason 'temporary export SQL file' | Out-Null
}

if (-not (Test-TextFileReadable -Path $outputFile)) {
    throw "Export did not create a readable CSV file: $outputFile"
}

$reader = $null
try {
    $reader = [System.IO.File]::OpenText($outputFile)
    $header = $reader.ReadLine()
    if ($header -notlike 'cnpj;*') {
        throw "Export CSV header is unexpected: $header"
    }
}
finally {
    if ($reader) { $reader.Dispose() }
}

$size = (Get-Item $outputFile).Length / 1MB
Write-Host " OK ($([math]::Round($size,1)) MB)" -ForegroundColor Green
Write-Host "`n CSV generated successfully!" -ForegroundColor Green
