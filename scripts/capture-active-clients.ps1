# =============================================================
# scripts/capture-active-clients.ps1
# Populate clientes_ativos_governo from existing LIMPO_ files
# without running the full ETL import pipeline.
#
# Usage:
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\capture-active-clients.ps1
#
# Optional parameters:
#   -ActiveClientsPath  Path to CSV with cnpj_normalizado column
#   -DirTemp            Temp dir with LIMPO_*.ESTABELE / LIMPO_*.EMPRECSV
#   -DbHost / -DbPort / -DbName / -DbUser
# =============================================================

[CmdletBinding()]
param(
    [string]$ActiveClientsPath = '',
    [string]$DirTemp = '',
    [string]$DbHost = 'localhost',
    [string]$DbPort = '5253',
    [string]$DbName = 'postgres',
    [string]$DbUser = 'pedroso',
    [string]$PsqlExe = 'D:\Postgres\bin\psql.exe'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# --- Resolve config-file defaults when params are blank ---
$ScriptDir = if ($MyInvocation.MyCommand.Path) { Split-Path $MyInvocation.MyCommand.Path -Parent } else { $PSScriptRoot }
$ProjectRoot = Split-Path $ScriptDir -Parent
$configPath = Join-Path $ProjectRoot 'config.ps1'
if (Test-Path -LiteralPath $configPath -PathType Leaf) {
    $config = @{}
    . $configPath
    if ($config.ContainsKey('dbHost'))   { if (-not $DbHost -or $DbHost -eq 'localhost') { $DbHost = [string]$config.dbHost } }
    if ($config.ContainsKey('dbPort'))   { if ($DbPort -eq '5253') { $DbPort = [string]$config.dbPort } }
    if ($config.ContainsKey('dbName'))   { if (-not $DbName -or $DbName -eq 'postgres') { $DbName = [string]$config.dbName } }
    if ($config.ContainsKey('dbUser'))   { if (-not $DbUser) { $DbUser = [string]$config.dbUser } }
    if ($config.ContainsKey('dbPassword') -and -not [string]::IsNullOrWhiteSpace([string]$config.dbPassword)) {
        $env:PGPASSWORD = [string]$config.dbPassword
    }
    if (-not $DirTemp -and $config.ContainsKey('dirTemp')) { $DirTemp = [string]$config.dirTemp }
    if (-not $ActiveClientsPath -and $config.ContainsKey('activeClientsPath')) {
        $ActiveClientsPath = [string]$config.activeClientsPath
    }
}

# --- Apply env-var password if set ---
if ([string]::IsNullOrWhiteSpace($env:PGPASSWORD) -and $env:CNPJ_ETL_DB_PASSWORD) {
    $env:PGPASSWORD = $env:CNPJ_ETL_DB_PASSWORD
}

# --- Defaults ---
if (-not $DirTemp) { $DirTemp = 'C:\Postgres17\data\temp' }
if (-not $ActiveClientsPath) {
    $ActiveClientsPath = Join-Path $env:USERPROFILE 'Downloads\clientes_classificados.csv'
}
if (-not (Test-Path -LiteralPath $PsqlExe -PathType Leaf)) {
    $found = Get-Command 'psql.exe' -ErrorAction SilentlyContinue
    if ($found) { $PsqlExe = $found.Source } else { throw "psql.exe not found at $PsqlExe and not in PATH." }
}

Write-Host '======================================================' -ForegroundColor Cyan
Write-Host '  CAPTURE ACTIVE CLIENTS - populate clientes_ativos_governo' -ForegroundColor Cyan
Write-Host '======================================================' -ForegroundColor Cyan
Write-Host "  DB      : ${DbHost}:${DbPort} / $DbName / user=$DbUser" -ForegroundColor Gray
Write-Host "  Temp dir: $DirTemp" -ForegroundColor Gray
Write-Host "  Active  : $ActiveClientsPath" -ForegroundColor Gray
Write-Host "  psql    : $PsqlExe" -ForegroundColor Gray

# --- Validate inputs ---
if (-not (Test-Path -LiteralPath $ActiveClientsPath -PathType Leaf)) {
    throw "Active clients file not found: $ActiveClientsPath"
}
$estabelecimentos = @(Get-ChildItem -Path $DirTemp -Filter 'LIMPO_*.ESTABELE' | Sort-Object Name | Select-Object -ExpandProperty FullName)
$empresas         = @(Get-ChildItem -Path $DirTemp -Filter 'LIMPO_*.EMPRECSV' | Sort-Object Name | Select-Object -ExpandProperty FullName)
if ($estabelecimentos.Count -eq 0) { throw "No LIMPO_*.ESTABELE files found in $DirTemp" }
if ($empresas.Count -eq 0)         { throw "No LIMPO_*.EMPRECSV files found in $DirTemp" }
Write-Host "  ESTABELE: $($estabelecimentos.Count) files" -ForegroundColor Gray
Write-Host "  EMPRECSV: $($empresas.Count) files" -ForegroundColor Gray

# --- Helper: run psql SQL file ---
$sqlTempFile = Join-Path $DirTemp 'capture_active_clients.sql'
function Run-PsqlFile {
    param([string]$Sql, [string]$Desc)
    Write-Host " » $Desc" -ForegroundColor Gray
    Set-Content -Path $sqlTempFile -Value $Sql -Encoding UTF8
    $out = & $PsqlExe -h $DbHost -p $DbPort -d $DbName -U $DbUser -w -v ON_ERROR_STOP=1 -f $sqlTempFile 2>&1
    $err = $out | Where-Object { $_ -match '^(ERROR|FATAL|psql: error)' }
    if ($err) { throw "psql error during [$Desc]:`n$($err -join "`n")" }
    return $out
}

function Run-PsqlQuery {
    param([string]$Sql, [string]$Desc)
    Write-Host " » $Desc" -ForegroundColor Gray
    $out = & $PsqlExe -h $DbHost -p $DbPort -d $DbName -U $DbUser -w -tAc $Sql 2>&1
    $err = $out | Where-Object { $_ -match '^(ERROR|FATAL|psql: error)' }
    if ($err) { throw "psql error during [$Desc]:`n$($err -join "`n")" }
    return ($out | Where-Object { $_ -ne $null -and "$_".Trim() -ne '' } | Select-Object -First 1)
}

# --- Helper: convert path for psql COPY (forward slashes) ---
function To-PsqlPath { param([string]$P); return $P.Replace('\', '/') }

# --- Step 1: Load active client CNPJs ---
Write-Host "`n[1/4] Loading active client CNPJ targets..." -ForegroundColor Yellow
$header = Get-Content -LiteralPath $ActiveClientsPath -Encoding UTF8 -TotalCount 1
$delim = if ($header -match ';') { ';' } else { ',' }
$activeRows = @(Import-Csv -LiteralPath $ActiveClientsPath -Delimiter $delim -Encoding UTF8)

$targetCsv = Join-Path $DirTemp 'clientes_ativos_alvo_capture.csv'
$targetRows = foreach ($row in $activeRows) {
    $raw = ''
    foreach ($field in @('cnpj_normalizado','cnpj','CNPJ','cnpj_corrigido')) {
        $p = $row.PSObject.Properties[$field]
        if ($p -and -not [string]::IsNullOrWhiteSpace([string]$p.Value)) { $raw = [string]$p.Value; break }
    }
    $digits = $raw -replace '\D',''
    $cnpj = if ($digits.Length -eq 14) { $digits } elseif ($digits.Length -gt 0) { $digits.PadLeft(14,'0') } else { '' }
    if ($cnpj -match '^\d{14}$') {
        [pscustomobject]@{ cnpj = $cnpj; cnpj_basico = $cnpj.Substring(0,8) }
    }
}
$targetRows = @($targetRows | Sort-Object cnpj -Unique)
if ($targetRows.Count -eq 0) { throw "No valid 14-digit CNPJs found in $ActiveClientsPath" }
$targetRows | Export-Csv -LiteralPath $targetCsv -Delimiter ';' -Encoding UTF8 -NoTypeInformation
Write-Host "  Loaded $($targetRows.Count) unique CNPJ targets." -ForegroundColor Gray

$targetPsqlPath = To-PsqlPath $targetCsv

$setupSql = @"
DROP TABLE IF EXISTS tmp_clientes_ativos_alvo CASCADE;
CREATE UNLOGGED TABLE tmp_clientes_ativos_alvo (
    cnpj VARCHAR(14) PRIMARY KEY,
    cnpj_basico VARCHAR(8)
);
COPY tmp_clientes_ativos_alvo FROM '$targetPsqlPath'
    WITH (FORMAT CSV, HEADER TRUE, DELIMITER ';', ENCODING 'UTF8', NULL '');
CREATE INDEX IF NOT EXISTS idx_tmp_clientes_ativos_alvo_basico ON tmp_clientes_ativos_alvo(cnpj_basico);
ANALYZE tmp_clientes_ativos_alvo;
TRUNCATE TABLE clientes_ativos_governo;

DROP TABLE IF EXISTS tmp_estabelecimentos_stage CASCADE;
CREATE UNLOGGED TABLE tmp_estabelecimentos_stage (
    cnpj_basico VARCHAR(8), cnpj_ordem VARCHAR(4), cnpj_dv VARCHAR(2),
    identificador_matriz_filial CHAR(1),
    nome_fantasia TEXT, situacao_cadastral VARCHAR(2),
    data_situacao_cadastral VARCHAR(8), motivo_situacao_cadastral VARCHAR(2),
    nome_cidade_exterior VARCHAR(100), pais VARCHAR(3),
    data_inicio_atividade VARCHAR(8), cnae_fiscal_principal VARCHAR(7),
    cnae_fiscal_secundaria TEXT,
    tipo_logradouro VARCHAR(50), logradouro VARCHAR(255),
    numero VARCHAR(20), complemento TEXT,
    bairro VARCHAR(100), cep VARCHAR(8), uf CHAR(2), municipio VARCHAR(4),
    ddd_1 VARCHAR(5), telefone_1 VARCHAR(10), ddd_2 VARCHAR(5),
    telefone_2 VARCHAR(10), ddd_fax VARCHAR(5), fax VARCHAR(10),
    correio_eletronico VARCHAR(255),
    situacao_especial VARCHAR(100), data_situacao_especial VARCHAR(8)
);

DROP TABLE IF EXISTS tmp_empresas_stage CASCADE;
CREATE UNLOGGED TABLE tmp_empresas_stage (
    cnpj_basico VARCHAR(8), razao_social TEXT, natureza_juridica VARCHAR(4),
    qualificacao_responsavel VARCHAR(10), capital_social VARCHAR(20),
    porte_empresa VARCHAR(10), ente_federativo_responsavel TEXT
);
"@
Run-PsqlFile $setupSql "Setup staging tables and load targets"

# --- Step 2: Process ESTABELECIMENTOS files ---
Write-Host "`n[2/4] Capturing active-client establishments..." -ForegroundColor Yellow
$estIdx = 0
foreach ($f in $estabelecimentos) {
    $estIdx++
    $leaf = Split-Path $f -Leaf
    $fPsql = To-PsqlPath $f

    $sql = @"
TRUNCATE tmp_estabelecimentos_stage;
COPY tmp_estabelecimentos_stage
FROM '$fPsql'
WITH (FORMAT CSV, DELIMITER ';', NULL '', ENCODING 'LATIN1');

INSERT INTO clientes_ativos_governo (
    cnpj, cnpj_basico, nome_fantasia, situacao_cadastral,
    data_inicio_atividade, cnae_fiscal_principal, cnae_fiscal_secundaria,
    uf, municipio, atualizado_em
)
SELECT
    LPAD(regexp_replace(t.cnpj_basico, '\D', '', 'g'), 8, '0')
    || LPAD(regexp_replace(t.cnpj_ordem, '\D', '', 'g'), 4, '0')
    || LPAD(regexp_replace(t.cnpj_dv,   '\D', '', 'g'), 2, '0')   AS cnpj,
    LPAD(regexp_replace(t.cnpj_basico,  '\D', '', 'g'), 8, '0')   AS cnpj_basico,
    NULLIF(t.nome_fantasia, '')        AS nome_fantasia,
    t.situacao_cadastral,
    t.data_inicio_atividade,
    NULLIF(t.cnae_fiscal_principal, '') AS cnae_fiscal_principal,
    NULLIF(t.cnae_fiscal_secundaria, '') AS cnae_fiscal_secundaria,
    t.uf,
    t.municipio,
    NOW()
FROM tmp_estabelecimentos_stage t
JOIN tmp_clientes_ativos_alvo a
  ON a.cnpj = LPAD(regexp_replace(t.cnpj_basico, '\D', '', 'g'), 8, '0')
           || LPAD(regexp_replace(t.cnpj_ordem,   '\D', '', 'g'), 4, '0')
           || LPAD(regexp_replace(t.cnpj_dv,      '\D', '', 'g'), 2, '0')
WHERE t.situacao_cadastral = '02'
ON CONFLICT (cnpj) DO UPDATE SET
    cnpj_basico           = EXCLUDED.cnpj_basico,
    nome_fantasia         = EXCLUDED.nome_fantasia,
    situacao_cadastral    = EXCLUDED.situacao_cadastral,
    data_inicio_atividade = EXCLUDED.data_inicio_atividade,
    cnae_fiscal_principal  = EXCLUDED.cnae_fiscal_principal,
    cnae_fiscal_secundaria = EXCLUDED.cnae_fiscal_secundaria,
    uf                    = EXCLUDED.uf,
    municipio             = EXCLUDED.municipio,
    atualizado_em         = NOW();
"@
    Run-PsqlFile $sql "ESTABELE $estIdx/$($estabelecimentos.Count): $leaf"
}

$afterEst = Run-PsqlQuery "SELECT COUNT(*) FROM clientes_ativos_governo;" "Count after ESTABELE pass"
Write-Host "  Captured so far: $afterEst rows" -ForegroundColor Gray

# --- Step 3: Enrich with Empresas (razao_social, natureza_juridica, etc.) ---
Write-Host "`n[3/4] Enriching with Empresas data..." -ForegroundColor Yellow
$empIdx = 0
foreach ($f in $empresas) {
    $empIdx++
    $leaf = Split-Path $f -Leaf
    $fPsql = To-PsqlPath $f

    $sql = @"
TRUNCATE tmp_empresas_stage;
COPY tmp_empresas_stage
FROM '$fPsql'
WITH (FORMAT CSV, DELIMITER ';', NULL '', ENCODING 'LATIN1');

WITH normalized AS (
    SELECT
        LPAD(regexp_replace(t.cnpj_basico, '\D', '', 'g'), 8, '0') AS cnpj_basico,
        t.razao_social,
        regexp_replace(t.natureza_juridica, '\D', '', 'g') AS natureza_juridica,
        t.capital_social,
        t.porte_empresa
    FROM tmp_empresas_stage t
)
UPDATE clientes_ativos_governo c
SET razao_social       = n.razao_social,
    natureza_juridica  = n.natureza_juridica,
    capital_social     = n.capital_social,
    porte_empresa      = n.porte_empresa,
    atualizado_em      = NOW()
FROM normalized n
WHERE c.cnpj_basico = n.cnpj_basico
  AND n.cnpj_basico ~ '^\d{8}`$';
"@
    Run-PsqlFile $sql "EMPRECSV $empIdx/$($empresas.Count): $leaf"
}

# --- Step 4: Final counts ---
Write-Host "`n[4/4] Verifying results..." -ForegroundColor Yellow
$totalCount = Run-PsqlQuery "SELECT COUNT(*) FROM clientes_ativos_governo;" "Total rows"
$cnaeCount  = Run-PsqlQuery "SELECT COUNT(*) FROM clientes_ativos_governo WHERE cnae_fiscal_principal IS NOT NULL AND cnae_fiscal_principal <> '';" "Rows with CNAE"
$natCount   = Run-PsqlQuery "SELECT COUNT(*) FROM clientes_ativos_governo WHERE natureza_juridica IS NOT NULL AND natureza_juridica <> '';" "Rows with natureza_juridica"
$razaoCount = Run-PsqlQuery "SELECT COUNT(*) FROM clientes_ativos_governo WHERE razao_social IS NOT NULL AND razao_social <> '';" "Rows with razao_social"

Write-Host "`n================================================" -ForegroundColor Cyan
Write-Host "  DONE - clientes_ativos_governo populated" -ForegroundColor Green
Write-Host "  Total rows    : $totalCount" -ForegroundColor White
Write-Host "  With CNAE     : $cnaeCount" -ForegroundColor White
Write-Host "  With nat.jur. : $natCount" -ForegroundColor White
Write-Host "  With razao    : $razaoCount" -ForegroundColor White
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "`nNext step: run the classifier with -IncludeGovernmentData to fill CNAE/category columns in report.csv" -ForegroundColor Yellow

# Cleanup temp SQL file
if (Test-Path -LiteralPath $sqlTempFile) { Remove-Item -LiteralPath $sqlTempFile -Force -ErrorAction SilentlyContinue }
