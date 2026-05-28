# =============================================================
# scripts/import.ps1 - Import data into PostgreSQL and join datasets
# =============================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot 'lib\preflight.ps1')
$ScriptDir = Resolve-ScriptRoot -Invocation $MyInvocation -FallbackBaseDirectory $PSScriptRoot

$bootstrap = Import-CnpjConfig -ScriptRoot $ScriptDir
$ProjectRoot = $bootstrap.ProjectRoot
$config = $bootstrap.Config

$env:CNPJ_ETL_DB_USER = $config.dbUser
$env:CNPJ_ETL_DB_PASSWORD = $config.dbPassword
$env:CNPJ_ETL_DB_PORT = $config.dbPort

Write-Step 'Import preflight checks'
$tools = Invoke-CnpjPreflight -Config $config -EnsureDirectories @($config.dirTemp, $config.dirOut) -MinFreeSpaceGB @{ $config.dirTemp = 6; $config.dirOut = 2 } -RequirePostgres -RequireDatabase -CheckMemory
$psqlPath = $tools.PsqlPath

Write-Host "===================================================" -ForegroundColor Cyan
Write-Host " IMPORT - Fast filtered load into PostgreSQL" -ForegroundColor Cyan
Write-Host "===================================================" -ForegroundColor Cyan

if (!(Test-Path $config.dirOut)) {
    New-Item -ItemType Directory -Path $config.dirOut -Force | Out-Null
}

# --- FIND CLEANED FILES ---
$estabelecimentos = @(Get-ChildItem -Path $config.dirTemp -Filter "LIMPO_*.ESTABELE" | Sort-Object Name | Select-Object -ExpandProperty FullName)
$empresas = @(Get-ChildItem -Path $config.dirTemp -Filter "LIMPO_*.EMPRECSV" | Sort-Object Name | Select-Object -ExpandProperty FullName)
$simples = @(Get-ChildItem -Path $config.dirTemp -Filter "LIMPO_*.SIMPLES*" | Sort-Object Name | Select-Object -ExpandProperty FullName | Select-Object -First 1)
$municipios = @(Get-ChildItem -Path $config.dirTemp -Filter "LIMPO_*.MUNICCSV" | Sort-Object Name | Select-Object -ExpandProperty FullName | Select-Object -First 1)

if ($estabelecimentos.Count -eq 0 -or $empresas.Count -eq 0) {
    Write-Host "[!] Required LIMPO_ files are missing. Run clean.ps1 first." -ForegroundColor Red
    Write-Host "    Required: at least one LIMPO_*.ESTABELE and one LIMPO_*.EMPRECSV" -ForegroundColor Red
    exit 1
}

Set-PostgresPasswordEnv -Config $config
$sqlTimeout = if ($config.ContainsKey('sqlCommandTimeoutSeconds')) { [int]$config.sqlCommandTimeoutSeconds } else { 3600 }
$cleanupMode = Get-CleanupMode -Config $config
$resetFinalTablesOnImport = -not $config.ContainsKey('resetFinalTablesOnImport') -or [bool]$config.resetFinalTablesOnImport
$requireEnrichmentMatches = -not $config.ContainsKey('requireEnrichmentMatches') -or [bool]$config.requireEnrichmentMatches
$filterSimplesNacional = $config.ContainsKey('filterSimplesNacional') -and [bool]$config.filterSimplesNacional
$captureActiveClients = $config.ContainsKey('captureActiveClients') -and [bool]$config.captureActiveClients
$activeClientsPath = if ($config.ContainsKey('activeClientsPath')) { [string]$config.activeClientsPath } else { '' }
$effectiveDataCorte = if ($config.ContainsKey('dataCorte') -and -not [string]::IsNullOrWhiteSpace([string]$config.dataCorte)) {
    [string]$config.dataCorte
} else {
    (Get-Date).AddMonths(-6).ToString('yyyyMMdd')
}
$usedCleanFiles = @($estabelecimentos + $empresas + $simples + $municipios) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

Write-Host " Opening date cutoff: $effectiveDataCorte (last 6 months unless dataCorte override is set)" -ForegroundColor Gray
Write-Host " Simples Nacional only: $filterSimplesNacional; MEI rows removed" -ForegroundColor Gray
Write-Host " Capture active clients lookup: $captureActiveClients" -ForegroundColor Gray

function Run-Sql {
    param([string]$sql, [string]$desc)
    Write-Host " » $desc" -ForegroundColor Gray
    $start = Get-Date
    $sqlFile = Join-Path $config.dirTemp "temp_query.sql"
    Set-Content -Path $sqlFile -Value $sql -Encoding UTF8

    $args = @(
        '-h', $config.dbHost,
        '-p', [string]$config.dbPort,
        '-d', $config.dbName,
        '-U', $config.dbUser,
        '-w',
        '-f', $sqlFile,
        '-v', 'ON_ERROR_STOP=1'
    )

    try {
        $job = Start-Job -ScriptBlock {
            param($exe, $argList)
            & $exe $argList 2>&1
        } -ArgumentList $psqlPath, $args

        $completed = Wait-Job -Job $job -Timeout $sqlTimeout
        if (-not $completed) {
            Stop-Job -Job $job -ErrorAction SilentlyContinue
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
            throw "SQL execution timed out after $sqlTimeout seconds during: $desc"
        }

        $result = Receive-Job -Job $job -ErrorAction SilentlyContinue
        $jobOutput = $result | Out-String
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue

        $errors = $result | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] -or $_ -match '^(ERROR|FATAL)' }
        if ($errors) {
            Write-Host "`n[!] ERROR while executing: $desc" -ForegroundColor Red
            Write-Host ($errors | Out-String) -ForegroundColor Red
            throw "Process aborted during step: $desc"
        }
        $elapsed = [math]::Round(((Get-Date) - $start).TotalMinutes, 1)
        Write-Host "   completed in $elapsed min" -ForegroundColor DarkGray
        return $jobOutput
    } finally {
        if (Test-Path -LiteralPath $sqlFile) {
            Remove-CnpjArtifactSafe -Config $config -Path $sqlFile -AllowedRoot $config.dirTemp -AllowedPatterns @('temp_query.sql') -Reason 'temporary import SQL file' | Out-Null
        }
    }
}

function Run-SqlQuery {
    param([string]$sql, [string]$desc)
    Write-Host " » $desc" -ForegroundColor Gray

    $args = @(
        '-h', $config.dbHost,
        '-p', [string]$config.dbPort,
        '-d', $config.dbName,
        '-U', $config.dbUser,
        '-w',
        '-tA',
        '-c', $sql,
        '-v', 'ON_ERROR_STOP=1'
    )

    $job = Start-Job -ScriptBlock {
        param($exe, $argList)
        & $exe $argList 2>&1
    } -ArgumentList $psqlPath, $args

    $completed = Wait-Job -Job $job -Timeout $sqlTimeout
    if (-not $completed) {
        Stop-Job -Job $job -ErrorAction SilentlyContinue
        Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        throw "SQL execution timed out after $sqlTimeout seconds during: $desc"
    }

    $result = Receive-Job -Job $job -ErrorAction SilentlyContinue
    Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
    $errors = $result | Where-Object { $_ -is [System.Management.Automation.ErrorRecord] -or $_ -match '^(ERROR|FATAL)' }
    if ($errors) {
        Write-Host "`n[!] ERROR while executing: $desc" -ForegroundColor Red
        Write-Host ($errors | Out-String) -ForegroundColor Red
        throw "Process aborted during step: $desc"
    }

    return ($result | Where-Object { $_ -ne $null } | Select-Object -First 1)
}

function Join-SqlQuotedList {
    param([string[]]$Values, [string]$Name)
    if ($Values.Count -eq 0) { return $null }
    foreach ($value in $Values) {
        if ($value -notmatch '^[A-Za-z0-9_ -]+$') {
            throw "Invalid value in ${Name}: $value"
        }
    }
    return "'" + (($Values | ForEach-Object { $_.Replace("'", "''") }) -join "','") + "'"
}

function New-CopyFromFileSql {
    param(
        [Parameter(Mandatory = $true)][string]$TableName,
        [Parameter(Mandatory = $true)][string]$Path
    )

    $copyPath = ConvertTo-PsqlPathLiteral -Path $Path
    return "COPY $TableName FROM '$copyPath' WITH (DELIMITER ';', FORMAT CSV, HEADER FALSE, ENCODING 'LATIN1', NULL '')"
}

function New-CnaeKeepCondition {
    param([string]$ColumnName)
    $clauses = @()
    foreach ($c in $config.cnaeWhitelist) {
        if ($c -notmatch '^\d{1,7}$') { throw "Invalid CNAE prefix in cnaeWhitelist: $c" }
        $clauses += "$ColumnName LIKE '$c%'"
    }
    foreach ($r in $config.cnaeRanges) {
        if ($r.Count -ne 2 -or $r[0] -notmatch '^\d{2}$' -or $r[1] -notmatch '^\d{2}$') {
            throw "Invalid CNAE range in cnaeRanges: $($r -join ',')"
        }
        $clauses += "LEFT($ColumnName, 2) BETWEEN '$($r[0])' AND '$($r[1])'"
    }
    if ($clauses.Count -eq 0) { throw 'At least one CNAE whitelist prefix or range is required.' }
    return '(' + ($clauses -join " OR ") + ')'
}

function Get-CsvDelimiterLocal {
    param([string]$Path)
    $header = Get-Content -LiteralPath $Path -Encoding UTF8 -TotalCount 1
    if ($header -match ';') { return ';' }
    return ','
}

function Get-PropValueLocal {
    param($Row, [string[]]$Names)
    foreach ($name in $Names) {
        $property = $Row.PSObject.Properties[$name]
        if ($property) { return $property.Value }
    }
    return ''
}

function Normalize-CnpjLocal {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    $digits = ([string]$Value) -replace '\D', ''
    if ($digits.Length -eq 14) { return $digits }
    if ($digits.Length -gt 0 -and $digits.Length -lt 14) { return $digits.PadLeft(14, '0') }
    return ''
}

$ufsFilter = Join-SqlQuotedList -Values @($config.ufs) -Name 'ufs'
$municipioFilterSql = ''
if ($config.ContainsKey('municipiosWhitelist') -and @($config.municipiosWhitelist).Count -gt 0) {
    $municipiosFilter = Join-SqlQuotedList -Values @($config.municipiosWhitelist) -Name 'municipiosWhitelist'
    $municipioFilterSql = "AND t.municipio IN ($municipiosFilter)"
}
$cnaeKeepSql = New-CnaeKeepCondition -ColumnName 't.cnae_fiscal_principal'

# --- STEP 1: CREATE REUSABLE STAGING TABLES ---
Write-Host "`n[1/7] Preparing reusable staging tables..." -ForegroundColor Yellow

$resetSql = if ($resetFinalTablesOnImport) {
@"
TRUNCATE TABLE estabelecimentos_crm;
TRUNCATE TABLE empresas_dados;
"@
} else {
    ""
}

$sqlTemp = @"
$resetSql

DROP TABLE IF EXISTS tmp_estabelecimentos_stage CASCADE;
CREATE UNLOGGED TABLE tmp_estabelecimentos_stage (
    cnpj_basico VARCHAR(8), cnpj_ordem VARCHAR(4), cnpj_dv VARCHAR(2), identificador_matriz_filial CHAR(1),
    nome_fantasia TEXT, situacao_cadastral VARCHAR(2), data_situacao_cadastral VARCHAR(8),
    motivo_situacao_cadastral VARCHAR(2), nome_cidade_exterior VARCHAR(100), pais VARCHAR(3),
    data_inicio_atividade VARCHAR(8), cnae_fiscal_principal VARCHAR(7), cnae_fiscal_secundaria TEXT,
    tipo_logradouro VARCHAR(50), logradouro VARCHAR(255), numero VARCHAR(20), complemento TEXT,
    bairro VARCHAR(100), cep VARCHAR(8), uf CHAR(2), municipio VARCHAR(4),
    ddd_1 VARCHAR(5), telefone_1 VARCHAR(10), ddd_2 VARCHAR(5), telefone_2 VARCHAR(10),
    ddd_fax VARCHAR(5), fax VARCHAR(10), correio_eletronico VARCHAR(255),
    situacao_especial VARCHAR(100), data_situacao_especial VARCHAR(8)
);

DROP TABLE IF EXISTS tmp_empresas_stage CASCADE;
CREATE UNLOGGED TABLE tmp_empresas_stage (
    cnpj_basico VARCHAR(8), razao_social TEXT, natureza_juridica VARCHAR(4),
    qualificacao_responsavel VARCHAR(10), capital_social VARCHAR(20),
    porte_empresa VARCHAR(10), ente_federativo_responsavel TEXT
);

DROP TABLE IF EXISTS tmp_simples_stage CASCADE;
CREATE UNLOGGED TABLE tmp_simples_stage (
    cnpj_basico VARCHAR(8), opcao_pelo_simples CHAR(1), data_opcao_simples VARCHAR(8),
    data_exclusao_simples VARCHAR(8), opcao_pelo_mei CHAR(1), data_opcao_mei VARCHAR(8),
    data_exclusao_mei VARCHAR(8)
);

DROP TABLE IF EXISTS tmp_municipios CASCADE;
CREATE UNLOGGED TABLE tmp_municipios (codigo VARCHAR(4), nome TEXT);

DROP TABLE IF EXISTS tmp_cnpj_basico_alvo CASCADE;
CREATE UNLOGGED TABLE tmp_cnpj_basico_alvo (cnpj_basico VARCHAR(8) PRIMARY KEY);

DROP TABLE IF EXISTS tmp_active_client_targets CASCADE;
CREATE UNLOGGED TABLE tmp_active_client_targets (cnpj VARCHAR(14) PRIMARY KEY, cnpj_basico VARCHAR(8));
"@
$resetDescription = if ($resetFinalTablesOnImport) { 'Reset final tables and create staging tables' } else { 'Drop/Create staging tables' }
Run-Sql $sqlTemp $resetDescription

if ($captureActiveClients) {
    if ([string]::IsNullOrWhiteSpace($activeClientsPath) -or -not (Test-Path -LiteralPath $activeClientsPath -PathType Leaf)) {
        Write-Host "    [WARN] Active clients file not found; lookup capture skipped: $activeClientsPath" -ForegroundColor Yellow
        $captureActiveClients = $false
    }
    else {
        Write-Host "`n[1b/7] Loading active-client CNPJ target list..." -ForegroundColor Yellow
        $delimiter = Get-CsvDelimiterLocal -Path $activeClientsPath
        $activeRows = @(Import-Csv -LiteralPath $activeClientsPath -Delimiter $delimiter -Encoding UTF8)
        $targetPath = Join-Path $config.dirTemp 'active_client_targets.csv'
        $activeTargets = foreach ($row in $activeRows) {
            $cnpj = Normalize-CnpjLocal -Value (Get-PropValueLocal -Row $row -Names @('cnpj_normalizado', 'cnpj', 'CNPJ', 'cnpj_corrigido'))
            if ($cnpj -match '^\d{14}$') { [pscustomobject]@{ cnpj = $cnpj; cnpj_basico = $cnpj.Substring(0, 8) } }
        }
        $activeTargets = @($activeTargets | Sort-Object cnpj -Unique)
        if ($activeTargets.Count -eq 0) {
            Write-Host '    [WARN] Active clients file had no valid CNPJs; lookup capture skipped.' -ForegroundColor Yellow
            $captureActiveClients = $false
        }
        else {
            $activeTargets | Export-Csv -LiteralPath $targetPath -Delimiter ';' -Encoding UTF8 -NoTypeInformation
            Run-Sql "COPY tmp_active_client_targets FROM '$(ConvertTo-PsqlPathLiteral -Path $targetPath)' WITH (FORMAT CSV, HEADER TRUE, DELIMITER ';', ENCODING 'UTF8', NULL ''); CREATE INDEX IF NOT EXISTS idx_tmp_active_client_targets_basico ON tmp_active_client_targets(cnpj_basico); ANALYZE tmp_active_client_targets; TRUNCATE TABLE active_clients_public_enrichment;" "Load active client targets ($($activeTargets.Count) CNPJs)"
            $loadedActiveTargets = Run-SqlQuery "SELECT COUNT(*) FROM tmp_active_client_targets;" 'Verify active client target load'
            Write-Host "    Active client CNPJ targets loaded: $loadedActiveTargets" -ForegroundColor Gray
            Remove-CnpjArtifactSafe -Config $config -Path $targetPath -AllowedRoot $config.dirTemp -AllowedPatterns @('active_client_targets.csv') -Reason 'active client target load complete' | Out-Null
        }
    }
}

if ($municipios) {
    Write-Host "`n[2/7] Loading municipality lookup..." -ForegroundColor Yellow
    $path = ConvertTo-PsqlPathLiteral -Path $municipios[0]
    Run-Sql (New-CopyFromFileSql -TableName 'tmp_municipios' -Path $municipios[0]) "COPY Municipios"
    Run-Sql "CREATE INDEX IF NOT EXISTS idx_tmp_municipios_codigo ON tmp_municipios(codigo); ANALYZE tmp_municipios;" "Index municipality lookup"
} else {
    Write-Host "`n[2/7] No municipality lookup file found; export will keep municipality codes as fallback." -ForegroundColor Yellow
}

# --- STEP 3: PROCESS ESTABELECIMENTOS FILE BY FILE ---
Write-Host "`n[3/7] Importing and filtering establishment files early..." -ForegroundColor Yellow
$estIndex = 0
foreach ($f in $estabelecimentos) {
    $estIndex++
    $leaf = Split-Path $f -Leaf
    $sqlCopyEstabelecimentos = @"
TRUNCATE tmp_estabelecimentos_stage;
$(New-CopyFromFileSql -TableName 'tmp_estabelecimentos_stage' -Path $f)
"@
    Run-Sql $sqlCopyEstabelecimentos "COPY Estabelecimentos $estIndex/$($estabelecimentos.Count): $leaf"

    if ($captureActiveClients) {
        $sqlCaptureActiveEstabelecimentos = @"
INSERT INTO active_clients_public_enrichment (
    cnpj, cnpj_basico, nome_fantasia, situacao_cadastral, data_inicio_atividade,
    cnae_fiscal_principal, cnae_fiscal_secundaria, uf, municipio, atualizado_em
)
SELECT
    LPAD(regexp_replace(t.cnpj_basico, '\D', '', 'g'), 8, '0') || LPAD(regexp_replace(t.cnpj_ordem, '\D', '', 'g'), 4, '0') || LPAD(regexp_replace(t.cnpj_dv, '\D', '', 'g'), 2, '0') AS cnpj,
    LPAD(regexp_replace(t.cnpj_basico, '\D', '', 'g'), 8, '0') AS cnpj_basico,
    NULLIF(t.nome_fantasia, '') AS nome_fantasia,
    t.situacao_cadastral,
    t.data_inicio_atividade,
    t.cnae_fiscal_principal,
    t.cnae_fiscal_secundaria,
    t.uf,
    t.municipio,
    NOW()
FROM tmp_estabelecimentos_stage t
JOIN tmp_active_client_targets a ON a.cnpj = LPAD(regexp_replace(t.cnpj_basico, '\D', '', 'g'), 8, '0') || LPAD(regexp_replace(t.cnpj_ordem, '\D', '', 'g'), 4, '0') || LPAD(regexp_replace(t.cnpj_dv, '\D', '', 'g'), 2, '0')
WHERE t.situacao_cadastral = '02'
ON CONFLICT (cnpj) DO UPDATE SET
    cnpj_basico = EXCLUDED.cnpj_basico,
    nome_fantasia = EXCLUDED.nome_fantasia,
    situacao_cadastral = EXCLUDED.situacao_cadastral,
    data_inicio_atividade = EXCLUDED.data_inicio_atividade,
    cnae_fiscal_principal = EXCLUDED.cnae_fiscal_principal,
    cnae_fiscal_secundaria = EXCLUDED.cnae_fiscal_secundaria,
    uf = EXCLUDED.uf,
    municipio = EXCLUDED.municipio,
    atualizado_em = NOW();
"@
        Run-Sql $sqlCaptureActiveEstabelecimentos "Capture active-client establishments from $leaf"
    }

$sqlInsertEstabelecimentos = @"
WITH upserted AS (
INSERT INTO estabelecimentos_crm (
    cnpj, nome_fantasia, data_inicio_atividade, cnae_fiscal_principal, cnae_fiscal_secundaria,
    tipo_logradouro, logradouro, numero, complemento, bairro, cep, uf,
    municipio, municipio_nome, telefone_1, telefone_2, email
)
SELECT
    t.cnpj_basico || t.cnpj_ordem || t.cnpj_dv AS cnpj,
    NULLIF(t.nome_fantasia, '') AS nome_fantasia,
    t.data_inicio_atividade,
    t.cnae_fiscal_principal,
    t.cnae_fiscal_secundaria,
    t.tipo_logradouro,
    t.logradouro,
    t.numero,
    t.complemento,
    t.bairro,
    t.cep,
    t.uf,
    t.municipio,
    COALESCE(m.nome, t.municipio) AS municipio_nome,
    CASE WHEN t.ddd_1 IS NOT NULL AND t.ddd_1 != '' THEN '(' || t.ddd_1 || ') ' || t.telefone_1 ELSE NULLIF(t.telefone_1, '') END AS telefone_1,
    CASE WHEN t.ddd_2 IS NOT NULL AND t.ddd_2 != '' THEN '(' || t.ddd_2 || ') ' || t.telefone_2 ELSE NULLIF(t.telefone_2, '') END AS telefone_2,
    NULLIF(t.correio_eletronico, '') AS email
FROM tmp_estabelecimentos_stage t
LEFT JOIN tmp_municipios m ON t.municipio = m.codigo
WHERE t.situacao_cadastral = '02'
AND t.identificador_matriz_filial = '1'
AND t.uf IN ($ufsFilter)
AND t.data_inicio_atividade >= '$effectiveDataCorte'
$municipioFilterSql
AND $cnaeKeepSql
ON CONFLICT (cnpj) DO UPDATE SET
    nome_fantasia = EXCLUDED.nome_fantasia,
    data_inicio_atividade = EXCLUDED.data_inicio_atividade,
    cnae_fiscal_principal = EXCLUDED.cnae_fiscal_principal,
    cnae_fiscal_secundaria = EXCLUDED.cnae_fiscal_secundaria,
    tipo_logradouro = EXCLUDED.tipo_logradouro,
    logradouro = EXCLUDED.logradouro,
    numero = EXCLUDED.numero,
    complemento = EXCLUDED.complemento,
    bairro = EXCLUDED.bairro,
    cep = EXCLUDED.cep,
    uf = EXCLUDED.uf,
    municipio = EXCLUDED.municipio,
    municipio_nome = EXCLUDED.municipio_nome,
    telefone_1 = EXCLUDED.telefone_1,
    telefone_2 = EXCLUDED.telefone_2,
    email = EXCLUDED.email
RETURNING LEFT(cnpj, 8) AS cnpj_basico
)
INSERT INTO tmp_cnpj_basico_alvo (cnpj_basico)
SELECT DISTINCT LPAD(regexp_replace(cnpj_basico, '\D', '', 'g'), 8, '0') FROM upserted
ON CONFLICT (cnpj_basico) DO NOTHING;
"@
    Run-Sql $sqlInsertEstabelecimentos "Filter into CRM from $leaf"
}

# --- STEP 4: TARGET CNPJ LIST ---
Write-Host "`n[4/7] Building target CNPJ list for enrichment..." -ForegroundColor Yellow
$targetCount = [int](Run-SqlQuery "SELECT COUNT(*) FROM tmp_cnpj_basico_alvo;" 'Count target cnpj_basico list')
if ($targetCount -eq 0 -and -not ([bool]$config.allowZeroFinalRows)) {
    throw 'No establishments matched the configured filters. LIMPO_* files were kept for inspection.'
}
Run-Sql "ANALYZE tmp_cnpj_basico_alvo;" "Analyze target cnpj_basico list ($targetCount rows)"

# --- STEP 4b: DIAGNOSE CNPJ FORMAT BEFORE ENRICHMENT ---
# Load first available Empresas file to inspect raw cnpj_basico format
$diagFile = $empresas | Select-Object -First 1
if ($diagFile) {
    $diagLeaf = Split-Path $diagFile -Leaf
    $diagSql = @"
TRUNCATE tmp_empresas_stage;
$(New-CopyFromFileSql -TableName 'tmp_empresas_stage' -Path $diagFile)
"@
    Run-Sql $diagSql "DIAG: Load first Empresas file ($diagLeaf)"
    $sampleEmpresas = Run-SqlQuery "SELECT cnpj_basico, length(cnpj_basico), cnpj_basico ~ '^\d+$' AS all_digits FROM tmp_empresas_stage LIMIT 5;" 'DIAG: Sample cnpj_basico from Empresas'
    $sampleAlvo = Run-SqlQuery "SELECT cnpj_basico, length(cnpj_basico) FROM tmp_cnpj_basico_alvo LIMIT 5;" 'DIAG: Sample cnpj_basico from target list'
    Write-Host " [DIAG] Empresas cnpj_basico samples: $sampleEmpresas" -ForegroundColor Magenta
    Write-Host " [DIAG] Target cnpj_basico samples: $sampleAlvo" -ForegroundColor Magenta
    $matchTest = Run-SqlQuery @"
SELECT COUNT(*) FROM tmp_empresas_stage t
JOIN tmp_cnpj_basico_alvo a ON LPAD(regexp_replace(t.cnpj_basico, '\D', '', 'g'), 8, '0') = a.cnpj_basico;
"@ 'DIAG: Test join match count on first Empresas file'
    Write-Host " [DIAG] Rows matching via normalized join: $matchTest" -ForegroundColor Magenta
}

# --- STEP 5: PROCESS EMPRESAS FILE BY FILE, KEEPING ONLY TARGET CNPJS ---
Write-Host "`n[5/7] Enriching legal names from Empresas files..." -ForegroundColor Yellow
$empIndex = 0
foreach ($f in $empresas) {
    $empIndex++
    $leaf = Split-Path $f -Leaf
    $sqlCopyEmpresas = @"
TRUNCATE tmp_empresas_stage;
$(New-CopyFromFileSql -TableName 'tmp_empresas_stage' -Path $f)
"@
    Run-Sql $sqlCopyEmpresas "COPY Empresas $empIndex/$($empresas.Count): $leaf"

    $sqlUpsertEmpresas = @"
WITH normalized AS (
    SELECT
        LPAD(regexp_replace(t.cnpj_basico, '\D', '', 'g'), 8, '0') AS cnpj_basico,
        t.razao_social,
        regexp_replace(t.natureza_juridica, '\D', '', 'g') AS natureza_juridica,
        t.capital_social,
        t.porte_empresa
    FROM tmp_empresas_stage t
)
INSERT INTO empresas_dados (cnpj_basico, razao_social, natureza_juridica, capital_social, porte_empresa)
SELECT DISTINCT ON (n.cnpj_basico)
    n.cnpj_basico, n.razao_social, n.natureza_juridica, n.capital_social, n.porte_empresa
FROM normalized n
JOIN tmp_cnpj_basico_alvo a ON a.cnpj_basico = n.cnpj_basico
WHERE n.cnpj_basico ~ '^\d{8}$'
ORDER BY n.cnpj_basico
ON CONFLICT (cnpj_basico) DO UPDATE SET
    razao_social = EXCLUDED.razao_social,
    natureza_juridica = EXCLUDED.natureza_juridica,
    capital_social = EXCLUDED.capital_social,
    porte_empresa = EXCLUDED.porte_empresa;
"@
    Run-Sql $sqlUpsertEmpresas "Apply Empresas enrichment from $leaf"

    if ($captureActiveClients) {
        $sqlUpdateActiveClientsEmpresas = @"
WITH normalized AS (
    SELECT
        LPAD(regexp_replace(t.cnpj_basico, '\D', '', 'g'), 8, '0') AS cnpj_basico,
        t.razao_social,
        regexp_replace(t.natureza_juridica, '\D', '', 'g') AS natureza_juridica,
        t.capital_social,
        t.porte_empresa
    FROM tmp_empresas_stage t
)
UPDATE active_clients_public_enrichment c
SET razao_social = n.razao_social,
    natureza_juridica = n.natureza_juridica,
    capital_social = n.capital_social,
    porte_empresa = n.porte_empresa,
    atualizado_em = NOW()
FROM normalized n
WHERE c.cnpj_basico = n.cnpj_basico
  AND n.cnpj_basico ~ '^\d{8}$';
"@
        Run-Sql $sqlUpdateActiveClientsEmpresas "Apply Empresas enrichment to active clients from $leaf"
    }
}

$empresaMatches = [int](Run-SqlQuery "SELECT COUNT(*) FROM empresas_dados;" 'Verify Empresas enrichment matches')
if ($targetCount -gt 0 -and $empresaMatches -eq 0 -and $requireEnrichmentMatches) {
    throw 'Empresas enrichment matched zero target CNPJs. LIMPO_* files were kept for inspection; check cnpj_basico normalization and source files.'
}

$sqlApplyEmpresas = @"
UPDATE estabelecimentos_crm c
SET razao_social = e.razao_social,
    nome_fantasia = COALESCE(NULLIF(c.nome_fantasia, ''), e.razao_social)
FROM empresas_dados e
WHERE LEFT(c.cnpj, 8) = e.cnpj_basico;
"@
Run-Sql $sqlApplyEmpresas "Apply legal names to CRM rows"

Write-Host "`n[6/8] Applying Simples and exclusion rules..." -ForegroundColor Yellow
if ($simples) {
    $sqlCopySimples = @"
TRUNCATE tmp_simples_stage;
$(New-CopyFromFileSql -TableName 'tmp_simples_stage' -Path $simples[0])
"@
    Run-Sql $sqlCopySimples "COPY Simples"

    $sqlSimples = @"
CREATE INDEX IF NOT EXISTS idx_tmp_simples_stage_cnpj_basico_norm ON tmp_simples_stage((LPAD(regexp_replace(cnpj_basico, '\D', '', 'g'), 8, '0')));
ANALYZE tmp_simples_stage;

WITH normalized AS (
    SELECT
        LPAD(regexp_replace(s.cnpj_basico, '\D', '', 'g'), 8, '0') AS cnpj_basico,
        s.opcao_pelo_simples,
        s.opcao_pelo_mei
    FROM tmp_simples_stage s
)
UPDATE estabelecimentos_crm c
SET regime_tributario = CASE
    WHEN n.opcao_pelo_mei = 'S' THEN 'MEI'
    WHEN n.opcao_pelo_simples = 'S' THEN 'Simples Nacional'
    ELSE 'Normal'
END
FROM normalized n
JOIN tmp_cnpj_basico_alvo a ON a.cnpj_basico = n.cnpj_basico
WHERE LEFT(c.cnpj, 8) = n.cnpj_basico;
"@
    Run-Sql $sqlSimples "Apply Simples enrichment only to target CNPJs"

    $simplesMatches = [int](Run-SqlQuery "SELECT COUNT(*) FROM estabelecimentos_crm WHERE regime_tributario IS NOT NULL;" 'Verify Simples enrichment matches')
    if ($targetCount -gt 0 -and $simplesMatches -eq 0 -and $requireEnrichmentMatches) {
        throw 'Simples enrichment matched zero target CNPJs. LIMPO_* files were kept for inspection; check cnpj_basico normalization and source files.'
    }

    if ($filterSimplesNacional) {
        Run-Sql "DELETE FROM estabelecimentos_crm WHERE COALESCE(regime_tributario, '') <> 'Simples Nacional';" "Filter to Simples Nacional rows and remove MEI/Normal rows"
    }
} else {
    Write-Host " » No Simples file found; regime_tributario will remain unchanged." -ForegroundColor Gray
    if ($filterSimplesNacional) {
        throw 'filterSimplesNacional is enabled, but no Simples file was found. LIMPO_* files were kept for inspection.'
    }
}

$sqlDeleteExcluded = @"
DELETE FROM estabelecimentos_crm c
USING empresas_dados e
WHERE LEFT(c.cnpj, 8) = e.cnpj_basico
AND e.natureza_juridica = '2135';
"@
Run-Sql $sqlDeleteExcluded "Remove excluded MEI legal nature rows"

# --- STEP 7: APPLY CNAE CATEGORY RULES ---
Write-Host "`n[7/8] Applying CNAE category rules..." -ForegroundColor Yellow

$sqlCategories = @"
INSERT INTO cnae_categoria_map (categoria, match_type, cnae_inicio, cnae_fim, prioridade, observacao)
VALUES
('agro_produtor_rural', 'range', '01', '03', 10, 'Agricultura, pecuaria, pesca e atividades relacionadas'),
('industria', 'range', '10', '33', 50, 'Industria geral'),
('construtora', 'range', '41', '43', 15, 'Construcao civil'),
('transporte', 'range', '49', '53', 10, 'Transporte, armazenagem e correio'),
('telecom', 'range', '61', '61', 10, 'Telecomunicacoes'),
('ti_software', 'range', '62', '63', 10, 'TI, software e servicos de informacao'),
('financeiro', 'range', '64', '66', 10, 'Financeiras, seguros e servicos auxiliares'),
('saude_clinica', 'range', '86', '86', 10, 'Atividades de saude humana'),
('gas_combustivel', 'exact', '4731800', '', 5, 'Comercio varejista de combustiveis para veiculos'),
('gas_combustivel', 'exact', '4784900', '', 5, 'Comercio varejista de gas liquefeito de petroleo'),
('gas_combustivel', 'exact', '3520401', '', 5, 'Producao de gas; processamento de gas natural'),
('farmacia', 'exact', '4771701', '', 5, 'Farmacia sem manipulacao'),
('farmacia', 'exact', '4771702', '', 5, 'Farmacia com manipulacao'),
('farmacia', 'exact', '4771703', '', 5, 'Farmacia homeopatica'),
('imobiliario', 'exact', '6810201', '', 5, 'Compra e venda de imoveis proprios'),
('imobiliario', 'exact', '6810202', '', 5, 'Aluguel de imoveis proprios'),
('imobiliario', 'exact', '6821801', '', 5, 'Corretagem de imoveis'),
('frigorifico', 'exact', '1011201', '', 5, 'Frigorifico - abate de bovinos'),
('frigorifico', 'exact', '1011202', '', 5, 'Frigorifico - abate de equinos'),
('frigorifico', 'exact', '1012101', '', 5, 'Abate de aves'),
('grafica', 'prefix', '181', '', 10, 'Impressao e servicos graficos'),
('cooperativa', 'prefix', '94', '', 80, 'Atividades associativas; revisar natureza juridica para cooperativas')
ON CONFLICT (categoria, match_type, cnae_inicio, cnae_fim) DO UPDATE SET
    prioridade = EXCLUDED.prioridade,
    ativo = TRUE,
    observacao = EXCLUDED.observacao;

TRUNCATE TABLE estabelecimentos_categorias;

WITH cnaes AS (
    SELECT cnpj, regexp_replace(cnae_fiscal_principal, '\D', '', 'g') AS cnae, 'principal' AS origem
    FROM estabelecimentos_crm
    WHERE COALESCE(cnae_fiscal_principal, '') <> ''
    UNION ALL
    SELECT e.cnpj, regexp_replace(value, '\D', '', 'g') AS cnae, 'secundaria' AS origem
    FROM estabelecimentos_crm e
    CROSS JOIN LATERAL regexp_split_to_table(COALESCE(e.cnae_fiscal_secundaria, ''), '[^0-9]+') AS value
    WHERE value <> ''
), matches AS (
    SELECT DISTINCT ON (c.cnpj, m.categoria, c.cnae, c.origem)
        c.cnpj,
        m.categoria,
        c.cnae AS cnae_match,
        c.origem AS cnae_origem,
        m.match_type,
        CASE WHEN c.origem = 'principal' THEN m.prioridade ELSE m.prioridade + 20 END AS prioridade
    FROM cnaes c
    JOIN cnae_categoria_map m ON m.ativo
    WHERE c.cnae <> ''
    AND (
        (m.match_type = 'exact' AND c.cnae = m.cnae_inicio)
        OR (m.match_type = 'prefix' AND c.cnae LIKE m.cnae_inicio || '%')
        OR (m.match_type = 'range' AND LEFT(c.cnae, 2) BETWEEN m.cnae_inicio AND NULLIF(m.cnae_fim, ''))
    )
    ORDER BY c.cnpj, m.categoria, c.cnae, c.origem, prioridade
), government_matches AS (
    SELECT
        c.cnpj,
        'cooperativa' AS categoria,
        e.natureza_juridica AS cnae_match,
        'governo' AS cnae_origem,
        'exact' AS match_type,
        5 AS prioridade
    FROM estabelecimentos_crm c
    JOIN empresas_dados e ON e.cnpj_basico = LEFT(c.cnpj, 8)
    WHERE e.natureza_juridica = '2143'
), all_matches AS (
    SELECT * FROM matches
    UNION ALL
    SELECT * FROM government_matches
), principal AS (
    SELECT cnpj, categoria, cnae_match, cnae_origem,
        ROW_NUMBER() OVER (
            PARTITION BY cnpj
            ORDER BY prioridade, CASE WHEN cnae_origem = 'principal' THEN 0 ELSE 1 END, categoria
        ) AS rn
    FROM all_matches
)
INSERT INTO estabelecimentos_categorias (
    cnpj, categoria, categoria_principal, cnae_match, cnae_origem, match_type, prioridade, classificacao_fonte
)
SELECT
    m.cnpj,
    m.categoria,
    p.rn = 1 AS categoria_principal,
    m.cnae_match,
    m.cnae_origem,
    m.match_type,
    m.prioridade,
    'Receita CNAE'
FROM all_matches m
JOIN principal p ON p.cnpj = m.cnpj
    AND p.categoria = m.categoria
    AND p.cnae_match = m.cnae_match
    AND p.cnae_origem = m.cnae_origem;
"@
Run-Sql $sqlCategories "Seed CNAE category rules and generate category matches"

# --- STEP 8: LOG AND VERIFY ---
Write-Host "`n[8/8] Logging and verifying import..." -ForegroundColor Yellow

$sqlLog = @"
INSERT INTO import_log (fase, arquivo, linhas_importadas)
SELECT 'IMPORT_FULL', '$($config.anoMes)', COUNT(*) FROM estabelecimentos_crm;
"@
Run-Sql $sqlLog "Write import log"

$logCount = [int](Run-SqlQuery "SELECT COUNT(*) FROM import_log WHERE fase = 'IMPORT_FULL' AND arquivo = '$($config.anoMes)';" 'Verify import log')
if ($logCount -lt 1) {
    throw 'Import completed, but no IMPORT_FULL log entry was found. LIMPO_* files were kept for inspection.'
}

$finalRows = [int](Run-SqlQuery "SELECT COUNT(*) FROM estabelecimentos_crm;" 'Verify final table row count')
if ($finalRows -eq 0 -and -not ([bool]$config.allowZeroFinalRows)) {
    throw 'Import completed, but estabelecimentos_crm has zero rows. LIMPO_* files were kept for inspection.'
}

if ($cleanupMode -in @('aggressive', 'balanced')) {
    foreach ($file in $usedCleanFiles) {
        Remove-CnpjArtifactSafe -Config $config -Path $file -AllowedRoot $config.dirTemp -AllowedPatterns @('LIMPO_*') -Reason 'import transaction verified' | Out-Null
    }
}

Write-Host "`n===================================================" -ForegroundColor Cyan
Write-Host " Import and processing completed! Final CRM rows: $finalRows" -ForegroundColor Green
