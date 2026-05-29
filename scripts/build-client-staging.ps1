# =============================================================
# build-client-staging.ps1 - Build PostgreSQL-ready existing-client staging CSV
# =============================================================

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [string]$EnrichmentPath,

    [string]$SimplesPath,

    [string]$RegimeOverridePath,

    [string]$ClassifiedInputPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [string]$ClassifiedOutputPath = '',

    [string]$ReviewPath = '',

    [string]$SummaryPath = '',

    [string]$XmlRoot = '',

    [string]$RecentXmlDiffPath = '',

    [int]$XmlDiffMonthsBack = 2,

    [string]$Delimiter = ';',

    [ValidateSet('Database', 'File')]
    [string]$Mode = 'Database',

    [switch]$IncludeGovernmentData
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$projectRoot = Split-Path -Parent $scriptRoot
$preflightPath = Join-Path $scriptRoot 'lib\preflight.ps1'
if (Test-Path -LiteralPath $preflightPath -PathType Leaf) { . $preflightPath }

function Resolve-RequiredFileLocal {
    param([string]$Path, [string]$Label)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "$Label was not found: $Path" }
    return (Resolve-Path -LiteralPath $Path).Path
}

function Resolve-OptionalFileLocal {
    param([string]$Path, [string]$Label)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "$Label was not found: $Path" }
    return (Resolve-Path -LiteralPath $Path).Path
}

function Resolve-ConfiguredXmlRootLocal {
    param([string]$RequestedRoot)

    if (-not [string]::IsNullOrWhiteSpace($RequestedRoot)) { return $RequestedRoot }
    if (-not [string]::IsNullOrWhiteSpace($env:CNPJ_XML_ROOT)) { return $env:CNPJ_XML_ROOT }

    $configPath = Join-Path $projectRoot 'config.ps1'
    if (Test-Path -LiteralPath $configPath -PathType Leaf) {
        $config = & $configPath
        if ($config -is [hashtable]) {
            foreach ($key in @('xmlRoot', 'XmlRoot', 'xmlAuditRoot', 'XmlAuditRoot')) {
                if ($config.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$config[$key])) { return [string]$config[$key] }
            }
        }
    }
    return ''
}

function Invoke-RecentXmlDiffIfConfigured {
    param([string]$ExpectedPath)

    $resolvedXmlRoot = Resolve-ConfiguredXmlRootLocal -RequestedRoot $XmlRoot
    if ([string]::IsNullOrWhiteSpace($resolvedXmlRoot)) {
        Write-Host 'Recent XML divergence report skipped: set -XmlRoot, CNPJ_XML_ROOT, or config xmlRoot to enable it.' -ForegroundColor Yellow
        return
    }

    $diffScript = Join-Path $scriptRoot 'export-recent-xml-regime-diff.ps1'
    if (-not (Test-Path -LiteralPath $diffScript -PathType Leaf)) {
        Write-Host "Recent XML divergence report skipped: script not found: $diffScript" -ForegroundColor Yellow
        return
    }

    $targetPath = $RecentXmlDiffPath
    if ([string]::IsNullOrWhiteSpace($targetPath)) {
        $targetPath = [System.IO.Path]::ChangeExtension($OutputPath, '.xml-divergencias-recentes.csv')
    }

    & $diffScript -ExpectedCsvPath $ExpectedPath -XmlRoot $resolvedXmlRoot -OutputPath $targetPath -MonthsBack $XmlDiffMonthsBack -Delimiter $Delimiter
}

function Resolve-DefaultEnrichmentPath {
    $candidates = @(
        (Join-Path $scriptRoot 'data\Clientes_A_limpo_cnpj_corrigido.csv'),
        (Join-Path $projectRoot 'data\Clientes_A_limpo_cnpj_corrigido.csv'),
        (Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads\Clientes_A_limpo_cnpj_corrigido.csv')
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return (Resolve-Path -LiteralPath $candidate).Path }
    }

    throw 'Enrichment CSV was not provided and the default Clientes_A_limpo_cnpj_corrigido.csv was not found beside the executable, in project data, or in Downloads.'
}

function Normalize-DigitsLocal {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value) -replace '\D', ''
}

function Normalize-IdLocal {
    param([object]$Value)
    $digits = Normalize-DigitsLocal -Value $Value
    if ([string]::IsNullOrWhiteSpace($digits)) { return '' }
    $trimmed = $digits.TrimStart('0')
    if ($trimmed -eq '') { return '0' }
    return $trimmed
}

function Normalize-CnpjLocal {
    param([object]$Value)
    $digits = Normalize-DigitsLocal -Value $Value
    if ($digits.Length -eq 14) { return $digits }
    if ($digits.Length -gt 0 -and $digits.Length -lt 14) { return $digits.PadLeft(14, '0') }
    return ''
}

function Format-CnpjLocal {
    param([string]$Cnpj)
    if ($Cnpj -notmatch '^\d{14}$') { return '' }
    return '{0}.{1}.{2}/{3}-{4}' -f $Cnpj.Substring(0, 2), $Cnpj.Substring(2, 3), $Cnpj.Substring(5, 3), $Cnpj.Substring(8, 4), $Cnpj.Substring(12, 2)
}

function Get-PropValue {
    param($Row, [string[]]$Names)
    foreach ($name in $Names) {
        $property = $Row.PSObject.Properties[$name]
        if ($property) { return $property.Value }
    }
    return ''
}

function Get-ClassifierOutputValue {
    param($Row, [string]$BaseName)

    $classifierProperty = @($Row.PSObject.Properties |
        Where-Object { $_.Name -eq "${BaseName}_classificador" -or $_.Name -match "^$([regex]::Escape($BaseName))_classificador_\d+$" } |
        Sort-Object Name |
        Select-Object -Last 1)
    if ($classifierProperty.Count -gt 0) { return $classifierProperty[0].Value }

    return (Get-PropValue -Row $Row -Names @($BaseName))
}

function Convert-ToPgBoolText {
    param([object]$Value)
    $text = ([string]$Value).Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($text)) { return '' }
    if ($text -in @('sim', 's', 'yes', 'y', 'true', 't', '1')) { return 'true' }
    if ($text -in @('nao', 'não', 'n', 'no', 'false', 'f', '0')) { return 'false' }
    return ''
}

function Add-ColumnValue {
    param([System.Collections.Specialized.OrderedDictionary]$Row, [string]$Name, [object]$Value)
    $Row[$Name] = if ($null -eq $Value) { '' } else { $Value }
}

function New-EmptyGovernmentRow {
    return [pscustomobject]@{
        cnae_fiscal_principal = ''
        cnae_fiscal_secundaria = ''
        categoria_principal = ''
        categorias_detectadas = ''
        categoria_transporte = ''
        categoria_gas_combustivel = ''
        categoria_farmacia = ''
        categoria_construtora = ''
        categoria_industria = ''
        categoria_agro_produtor_rural = ''
        categoria_telecom = ''
        categoria_ti_software = ''
        categoria_saude_clinica = ''
        categoria_financeiro = ''
        categoria_cooperativa = ''
        categoria_imobiliario = ''
        categoria_grafica = ''
        categoria_frigorifico = ''
        categoria_importador_exportador_status = 'external_source_required'
        categoria_beneficio_fiscal_status = 'external_source_required'
    }
}

function Export-GovernmentCategories {
    param([string]$Path)

    $psqlCandidates = @(
        'D:\Postgres\bin\psql.exe',
        'C:\Postgres17\bin\psql.exe',
        'C:\Postgres16\bin\psql.exe',
        'psql.exe'
    )
    $psqlPath = $psqlCandidates | Where-Object { (Get-Command $_ -ErrorAction SilentlyContinue) -or (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -First 1
    if (-not $psqlPath) { throw 'psql.exe was not found. Cannot export government CNAE data.' }

    $hostCandidates = @('localhost', '127.0.0.1')
    $portCandidates = @($env:CNPJ_ETL_DB_PORT, $env:PGPORT, '5253', '5432') | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique
    $dbName = if ($env:CNPJ_ETL_DB_NAME) { $env:CNPJ_ETL_DB_NAME } else { 'postgres' }
    $dbUser = if ($env:CNPJ_ETL_DB_USER) { $env:CNPJ_ETL_DB_USER } elseif ($env:PGUSER) { $env:PGUSER } else { 'postgres' }

    # Capture password from any known env var before launching jobs
    $pgPassword = if ($env:PGPASSWORD) { $env:PGPASSWORD } elseif ($env:CNPJ_ETL_DB_PASSWORD) { $env:CNPJ_ETL_DB_PASSWORD } else { '' }

    $connection = $null
    foreach ($hostName in $hostCandidates) {
        foreach ($port in $portCandidates) {
            $job = Start-Job -ScriptBlock {
                param($Exe, $HostName, $Port, $Database, $User, $PgPass)
                if ($PgPass) { $env:PGPASSWORD = $PgPass }
                $output = & $Exe -h $HostName -p $Port -d $Database -U $User -w -tAc 'SELECT 1;' 2>&1
                [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = ($output | Out-String).Trim() }
            } -ArgumentList $psqlPath, $hostName, $port, $dbName, $dbUser, $pgPassword
            $completed = Wait-Job -Job $job -Timeout 8
            if ($completed) {
                $result = Receive-Job -Job $job -ErrorAction SilentlyContinue
                Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
                if ($result.ExitCode -eq 0) {
                    $connection = [pscustomobject]@{ Host = $hostName; Port = $port; Database = $dbName; User = $dbUser; PgPass = $pgPassword }
                    break
                }
            }
            else {
                Stop-Job -Job $job -ErrorAction SilentlyContinue
                Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
            }
        }
        if ($connection) { break }
    }

    if (-not $connection) { throw 'Could not connect to PostgreSQL without prompting for a password. Set CNPJ_ETL_DB_PORT/PGPORT and PGPASSWORD, or provide an exported CNAE file.' }

    $sql = @"
COPY (
    WITH estabelecimentos_governo AS (
        SELECT
            cnpj,
            left(cnpj, 8) AS cnpj_basico,
            cnae_fiscal_principal,
            cnae_fiscal_secundaria,
            '' AS natureza_juridica
        FROM estabelecimentos_crm
        UNION
        SELECT
            cnpj_basico || cnpj_ordem || cnpj_dv AS cnpj,
            cnpj_basico,
            cnae_fiscal_principal,
            cnae_fiscal_secundaria,
            '' AS natureza_juridica
        FROM tmp_estabelecimentos
        WHERE situacao_cadastral = '02'
        UNION
        SELECT
            cnpj,
            cnpj_basico,
            cnae_fiscal_principal,
            cnae_fiscal_secundaria,
            natureza_juridica
        FROM clientes_ativos_governo
    ),
    cnaes AS (
        SELECT cnpj, NULLIF(regexp_replace(COALESCE(cnae_fiscal_principal, ''), '[^0-9]', '', 'g'), '') AS cnae, 'principal' AS cnae_origem
        FROM estabelecimentos_governo
        UNION ALL
        SELECT e.cnpj, s.cnae, 'secundaria' AS cnae_origem
        FROM estabelecimentos_governo e
        CROSS JOIN LATERAL regexp_split_to_table(COALESCE(e.cnae_fiscal_secundaria, ''), '[^0-9]+') AS s(cnae)
        WHERE s.cnae <> ''
    ),
    cnae_matches AS (
        SELECT cnpj, cnae, cnae_origem, v.categoria, v.prioridade
        FROM cnaes
        CROSS JOIN LATERAL (VALUES
            ('gas_combustivel', cnae IN ('4731800', '4784900', '3520401'), 10),
            ('farmacia', cnae IN ('4771701', '4771702', '4771703'), 10),
            ('imobiliario', cnae IN ('6810201', '6810202', '6821801'), 10),
            ('frigorifico', cnae IN ('1011201', '1011202', '1012101'), 10),
            ('grafica', cnae LIKE '181%', 80),
            ('agro_produtor_rural', substring(cnae, 1, 2) BETWEEN '01' AND '03', 50),
            ('industria', substring(cnae, 1, 2) BETWEEN '10' AND '33', 50),
            ('construtora', substring(cnae, 1, 2) BETWEEN '41' AND '43', 50),
            ('transporte', substring(cnae, 1, 2) BETWEEN '49' AND '53', 50),
            ('telecom', substring(cnae, 1, 2) = '61', 50),
            ('ti_software', substring(cnae, 1, 2) BETWEEN '62' AND '63', 50),
            ('financeiro', substring(cnae, 1, 2) BETWEEN '64' AND '66', 50),
            ('saude_clinica', substring(cnae, 1, 2) = '86', 50)
        ) AS v(categoria, matched, prioridade)
        WHERE cnae IS NOT NULL AND v.matched
    ),
    government_matches AS (
        SELECT DISTINCT e.cnpj, '' AS cnae, 'governo' AS cnae_origem, 'cooperativa' AS categoria, 5 AS prioridade
        FROM estabelecimentos_governo e
        LEFT JOIN (
            SELECT cnpj_basico, natureza_juridica FROM empresas_dados
            UNION
            SELECT cnpj_basico, natureza_juridica FROM tmp_empresas
        ) d ON d.cnpj_basico = e.cnpj_basico
        WHERE COALESCE(NULLIF(e.natureza_juridica, ''), d.natureza_juridica) = '2143'
    ),
    all_matches AS (
        SELECT * FROM cnae_matches
        UNION ALL
        SELECT * FROM government_matches
    ),
    principal_category AS (
        SELECT DISTINCT ON (cnpj) cnpj, categoria
        FROM all_matches
        ORDER BY cnpj, prioridade, CASE cnae_origem WHEN 'governo' THEN 0 WHEN 'principal' THEN 1 ELSE 2 END, categoria
    ),
    category_flags AS (
        SELECT
            cnpj,
            string_agg(DISTINCT categoria, ',' ORDER BY categoria) AS categorias_detectadas,
            max(categoria) FILTER (WHERE categoria = (SELECT pc.categoria FROM principal_category pc WHERE pc.cnpj = all_matches.cnpj)) AS categoria_principal,
            bool_or(categoria = 'transporte') AS categoria_transporte,
            bool_or(categoria = 'gas_combustivel') AS categoria_gas_combustivel,
            bool_or(categoria = 'farmacia') AS categoria_farmacia,
            bool_or(categoria = 'construtora') AS categoria_construtora,
            bool_or(categoria = 'industria') AS categoria_industria,
            bool_or(categoria = 'agro_produtor_rural') AS categoria_agro_produtor_rural,
            bool_or(categoria = 'telecom') AS categoria_telecom,
            bool_or(categoria = 'ti_software') AS categoria_ti_software,
            bool_or(categoria = 'saude_clinica') AS categoria_saude_clinica,
            bool_or(categoria = 'financeiro') AS categoria_financeiro,
            bool_or(categoria = 'cooperativa') AS categoria_cooperativa,
            bool_or(categoria = 'imobiliario') AS categoria_imobiliario,
            bool_or(categoria = 'grafica') AS categoria_grafica,
            bool_or(categoria = 'frigorifico') AS categoria_frigorifico
        FROM all_matches
        GROUP BY cnpj
    )
    SELECT
        e.cnpj,
        e.cnae_fiscal_principal,
        e.cnae_fiscal_secundaria,
        COALESCE(f.categoria_principal, '') AS categoria_principal,
        COALESCE(f.categorias_detectadas, '') AS categorias_detectadas,
        COALESCE(f.categoria_transporte, false) AS categoria_transporte,
        COALESCE(f.categoria_gas_combustivel, false) AS categoria_gas_combustivel,
        COALESCE(f.categoria_farmacia, false) AS categoria_farmacia,
        COALESCE(f.categoria_construtora, false) AS categoria_construtora,
        COALESCE(f.categoria_industria, false) AS categoria_industria,
        COALESCE(f.categoria_agro_produtor_rural, false) AS categoria_agro_produtor_rural,
        COALESCE(f.categoria_telecom, false) AS categoria_telecom,
        COALESCE(f.categoria_ti_software, false) AS categoria_ti_software,
        COALESCE(f.categoria_saude_clinica, false) AS categoria_saude_clinica,
        COALESCE(f.categoria_financeiro, false) AS categoria_financeiro,
        COALESCE(f.categoria_cooperativa, false) AS categoria_cooperativa,
        COALESCE(f.categoria_imobiliario, false) AS categoria_imobiliario,
        COALESCE(f.categoria_grafica, false) AS categoria_grafica,
        COALESCE(f.categoria_frigorifico, false) AS categoria_frigorifico,
        'external_source_required' AS categoria_importador_exportador_status,
        'external_source_required' AS categoria_beneficio_fiscal_status
    FROM estabelecimentos_governo e
    LEFT JOIN category_flags f ON f.cnpj = e.cnpj
    ORDER BY e.cnpj
) TO STDOUT
WITH (FORMAT CSV, HEADER TRUE, DELIMITER ';', ENCODING 'UTF8', NULL '');
"@

    $sqlFile = Join-Path ([System.IO.Path]::GetTempPath()) ("client-government-export-{0}.sql" -f ([guid]::NewGuid().ToString('N')))
    [System.IO.File]::WriteAllText($sqlFile, $sql, [System.Text.Encoding]::UTF8)
    try {
        $job = Start-Job -ScriptBlock {
            param($Exe, $Connection, $SqlFile, $OutPath)
            if ($Connection.PgPass) { $env:PGPASSWORD = $Connection.PgPass }
            $output = & $Exe -h $Connection.Host -p $Connection.Port -d $Connection.Database -U $Connection.User -w -v ON_ERROR_STOP=1 -f $SqlFile 2>&1
            if ($LASTEXITCODE -eq 0) { [System.IO.File]::WriteAllLines($OutPath, [string[]]$output, [System.Text.Encoding]::UTF8) }
            [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = ($output | Out-String).Trim() }
        } -ArgumentList $psqlPath, $connection, $sqlFile, $Path
        $completed = Wait-Job -Job $job -Timeout 120
        if (-not $completed) {
            Stop-Job -Job $job -ErrorAction SilentlyContinue
            throw 'Government CNAE export timed out after 120s.'
        }
        $result = Receive-Job -Job $job -ErrorAction SilentlyContinue
        if ($result.ExitCode -ne 0) { throw "Government CNAE export failed: $($result.Output)" }
    }
    finally { Remove-Item -LiteralPath $sqlFile -Force -ErrorAction SilentlyContinue }
}

if ($Delimiter.Length -ne 1) { throw 'Delimiter must be a single character.' }
$InputPath = Resolve-RequiredFileLocal -Path $InputPath -Label 'Input file'
if ([string]::IsNullOrWhiteSpace($EnrichmentPath)) { $EnrichmentPath = Resolve-DefaultEnrichmentPath }
else { $EnrichmentPath = Resolve-RequiredFileLocal -Path $EnrichmentPath -Label 'Enrichment CSV' }
$SimplesPath = Resolve-OptionalFileLocal -Path $SimplesPath -Label 'Simples CSV'
$RegimeOverridePath = Resolve-OptionalFileLocal -Path $RegimeOverridePath -Label 'Regime override CSV'
$ClassifiedInputPath = Resolve-OptionalFileLocal -Path $ClassifiedInputPath -Label 'Existing classified CSV'

$outputParent = Split-Path -Parent $OutputPath
if (-not [string]::IsNullOrWhiteSpace($outputParent) -and -not (Test-Path -LiteralPath $outputParent)) { New-Item -ItemType Directory -Path $outputParent -Force | Out-Null }
if ([string]::IsNullOrWhiteSpace($ClassifiedOutputPath)) { $ClassifiedOutputPath = Join-Path $outputParent 'clientes_base_classificado.csv' }
if ([string]::IsNullOrWhiteSpace($ReviewPath)) { $ReviewPath = [System.IO.Path]::ChangeExtension($OutputPath, '.revisar.csv') }
if ([string]::IsNullOrWhiteSpace($SummaryPath)) { $SummaryPath = [System.IO.Path]::ChangeExtension($OutputPath, '.summary.json') }

$classifier = Join-Path $scriptRoot 'classify-clientes.ps1'
if (-not (Test-Path -LiteralPath $classifier -PathType Leaf)) { throw "Classifier script was not found: $classifier" }

Write-Host '===================================================' -ForegroundColor Cyan
Write-Host ' CLIENT STAGING BUILDER' -ForegroundColor Cyan
Write-Host '===================================================' -ForegroundColor Cyan
Write-Host " Base input:  $InputPath" -ForegroundColor Yellow
Write-Host " Enrichment:  $EnrichmentPath" -ForegroundColor Yellow
Write-Host " Classified:  $(if ($ClassifiedInputPath) { $ClassifiedInputPath } else { 'generated from base input' })" -ForegroundColor Yellow
Write-Host " Simples:     $(if ($SimplesPath) { $SimplesPath } elseif ($ClassifiedInputPath) { 'already supplied by classified input when present' } else { 'not provided; regime fields will stay unclassified' })" -ForegroundColor Yellow
Write-Host " Override:    $(if ($RegimeOverridePath) { $RegimeOverridePath } else { 'not provided' })" -ForegroundColor Yellow
Write-Host " Output:      $OutputPath" -ForegroundColor Yellow
Write-Host '-------------------------------------------------' -ForegroundColor Gray

if ($ClassifiedInputPath) {
    $ClassifiedOutputPath = $ClassifiedInputPath
}
else {
    $classifierArgs = @{
        InputPath = $InputPath
        OutputPath = $ClassifiedOutputPath
        Delimiter = $Delimiter
        Mode = $Mode
    }
    if (-not [string]::IsNullOrWhiteSpace($SimplesPath)) { $classifierArgs.SimplesPath = $SimplesPath }
    if (-not [string]::IsNullOrWhiteSpace($RegimeOverridePath)) { $classifierArgs.RegimeOverridePath = $RegimeOverridePath }
    & $classifier @classifierArgs
}

$classifiedRows = @(Import-Csv -LiteralPath $ClassifiedOutputPath -Delimiter $Delimiter -Encoding UTF8)
$enrichmentRows = @(Import-Csv -LiteralPath $EnrichmentPath -Delimiter $Delimiter -Encoding UTF8)
if ($classifiedRows.Count -eq 0) { throw 'Classifier output has no rows.' }

$enrichmentGroups = @{}
$line = 1
foreach ($row in $enrichmentRows) {
    $line++
    $id = Normalize-IdLocal -Value (Get-PropValue -Row $row -Names @('id_normalizado', 'ID'))
    if ([string]::IsNullOrWhiteSpace($id)) { continue }
    if (-not $enrichmentGroups.ContainsKey($id)) { $enrichmentGroups[$id] = New-Object System.Collections.ArrayList }
    $row | Add-Member -NotePropertyName '__linha_origem' -NotePropertyValue $line -Force
    [void]$enrichmentGroups[$id].Add($row)
}

$enrichmentLookup = @{}
$duplicateEnrichmentIds = 0
foreach ($id in $enrichmentGroups.Keys) {
    $group = @($enrichmentGroups[$id])
    if ($group.Count -gt 1) { $duplicateEnrichmentIds++ }
    $chosen = $group | Sort-Object `
        @{ Expression = { if ((Normalize-CnpjLocal -Value $_.cnpj_corrigido) -match '^\d{14}$') { 0 } else { 1 } } }, `
        @{ Expression = { -(@($_.PSObject.Properties | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Value) }).Count) } }, `
        @{ Expression = { [int]$_.__linha_origem } } | Select-Object -First 1
    $enrichmentLookup[$id] = [pscustomobject]@{ Row = $chosen; IsDuplicate = ($group.Count -gt 1) }
}

$governmentLookup = @{}
if ($IncludeGovernmentData) {
    $govPath = Join-Path ([System.IO.Path]::GetTempPath()) ("client-government-categories-{0}.csv" -f ([guid]::NewGuid().ToString('N')))
    try {
        Export-GovernmentCategories -Path $govPath
        if (Test-Path -LiteralPath $govPath -PathType Leaf) {
            foreach ($row in @(Import-Csv -LiteralPath $govPath -Delimiter ';' -Encoding UTF8)) {
                $cnpj = Normalize-CnpjLocal -Value $row.cnpj
                if ($cnpj) { $governmentLookup[$cnpj] = $row }
            }
        }
    }
    finally { Remove-Item -LiteralPath $govPath -Force -ErrorAction SilentlyContinue }
}

$baseIds = @{}
$duplicateBaseIds = New-Object System.Collections.ArrayList
$results = New-Object System.Collections.ArrayList
$review = New-Object System.Collections.ArrayList
$matchedEnrichmentRows = 0
$unmatchedEnrichmentRows = 0
$safeCnpjCount = 0
$missingCnpjCount = 0
$cnpjDivergenceCount = 0

foreach ($base in $classifiedRows) {
    $idOriginal = Get-PropValue -Row $base -Names @('ID', 'id_original')
    $id = Normalize-IdLocal -Value $idOriginal
    if ($id) {
        if ($baseIds.ContainsKey($id)) { [void]$duplicateBaseIds.Add($id) } else { $baseIds[$id] = $true }
    }

    $enrichmentMatch = $null
    $enriched = $null
    $matchStatus = 'sem_match_enriquecimento'
    if ($id -and $enrichmentLookup.ContainsKey($id)) {
        $enrichmentMatch = $enrichmentLookup[$id]
        $enriched = $enrichmentMatch.Row
        $matchStatus = if ($enrichmentMatch.IsDuplicate) { 'duplicate_resolved' } else { 'matched' }
    }
    if ($matchStatus -eq 'sem_match_enriquecimento') { $unmatchedEnrichmentRows++ } else { $matchedEnrichmentRows++ }

    $classifierCnpj = Normalize-CnpjLocal -Value (Get-PropValue -Row $base -Names @('cnpj_normalizado'))
    $baseCnpj = Normalize-CnpjLocal -Value (Get-PropValue -Row $base -Names @('CNPJ', 'cnpj'))
    $enrichedCnpj = if ($enriched) { Normalize-CnpjLocal -Value (Get-PropValue -Row $enriched -Names @('cnpj_corrigido')) } else { '' }
    $cnpj = ''
    $cnpjFonte = 'sem_cnpj_seguro'
    if ($classifierCnpj -match '^\d{14}$') { $cnpj = $classifierCnpj; $cnpjFonte = 'classificador' }
    elseif ($baseCnpj -match '^\d{14}$') { $cnpj = $baseCnpj; $cnpjFonte = 'base_xlsx' }
    elseif ($enrichedCnpj -match '^\d{14}$') { $cnpj = $enrichedCnpj; $cnpjFonte = 'enriquecimento_fallback' }

    $diverge = ''
    if ($cnpj -and $enrichedCnpj) { $diverge = if ($cnpj -eq $enrichedCnpj) { 'false' } else { 'true' } }
    if ($cnpj -match '^\d{14}$') { $safeCnpjCount++ } else { $missingCnpjCount++ }
    if ($diverge -eq 'true') { $cnpjDivergenceCount++ }

    $gov = if ($cnpj -and $governmentLookup.ContainsKey($cnpj)) { $governmentLookup[$cnpj] } else { New-EmptyGovernmentRow }
    $auditRow = [ordered]@{}
    Add-ColumnValue $auditRow 'id_normalizado' $id
    Add-ColumnValue $auditRow 'nome' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('nome_normalizado') } else { Get-PropValue -Row $base -Names @('Nome', 'nome') })
    Add-ColumnValue $auditRow 'razao_social' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('razao_social_normalizada') } else { Get-PropValue -Row $base -Names @('Razão Social', 'Razao Social', 'razao_social') })
    Add-ColumnValue $auditRow 'cidade' (Get-PropValue -Row $base -Names @('Cidade', 'cidade'))
    Add-ColumnValue $auditRow 'estado' (Get-PropValue -Row $base -Names @('Estado', 'estado'))
    Add-ColumnValue $auditRow 'cnpj' $cnpj
    Add-ColumnValue $auditRow 'cnpj_normalizado' $cnpj
    Add-ColumnValue $auditRow 'cnpj_basico' $(if ($cnpj -match '^\d{14}$') { $cnpj.Substring(0, 8) } else { '' })
    Add-ColumnValue $auditRow 'regime_tributario' (Get-ClassifierOutputValue -Row $base -BaseName 'regime_tributario')
    Add-ColumnValue $auditRow 'postgres_versao' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('postgres_versao_normalizada') } else { '' })
    Add-ColumnValue $auditRow 'arquitetura_os' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('arquitetura_os_normalizada') } else { '' })
    Add-ColumnValue $auditRow 'provedor' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('provedor_normalizado') } else { '' })
    Add-ColumnValue $auditRow 'cpf_cnpj_internet' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('cpf_cnpj_internet_normalizado') } else { '' })
    Add-ColumnValue $auditRow 'nome_internet' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('nome_internet_normalizado') } else { '' })
    Add-ColumnValue $auditRow 'atualizando' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('atualizando_normalizado') } else { '' })
    Add-ColumnValue $auditRow 'flag_postgres_17_ou_nuvem' $(if ($enriched) { Convert-ToPgBoolText (Get-PropValue -Row $enriched -Names @('Postgres 17 ou banco nuvem')) } else { '' })
    Add-ColumnValue $auditRow 'flag_postgres_desatualizado_32_bits' $(if ($enriched) { Convert-ToPgBoolText (Get-PropValue -Row $enriched -Names @('Postgres desatualizado 32 bits')) } else { '' })
    Add-ColumnValue $auditRow 'flag_postgres_atualizado_sem_replicacao' $(if ($enriched) { Convert-ToPgBoolText (Get-PropValue -Row $enriched -Names @('Postgres atualizado sem replicação')) } else { '' })
    Add-ColumnValue $auditRow 'flag_multiplos_computadores' $(if ($enriched) { Convert-ToPgBoolText (Get-PropValue -Row $enriched -Names @('Múltiplos computadores')) } else { '' })
    Add-ColumnValue $auditRow 'flag_sistema_web' $(if ($enriched) { Convert-ToPgBoolText (Get-PropValue -Row $enriched -Names @('Sistema web')) } else { '' })
    Add-ColumnValue $auditRow 'flag_sistema_deluito_contador' $(if ($enriched) { Convert-ToPgBoolText (Get-PropValue -Row $enriched -Names @('Sistema Deluito contador')) } else { '' })
    Add-ColumnValue $auditRow 'cnae_fiscal_principal' (Get-PropValue -Row $gov -Names @('cnae_fiscal_principal'))
    Add-ColumnValue $auditRow 'cnae_fiscal_secundaria' (Get-PropValue -Row $gov -Names @('cnae_fiscal_secundaria'))
    Add-ColumnValue $auditRow 'categoria_principal' (Get-PropValue -Row $gov -Names @('categoria_principal'))
    Add-ColumnValue $auditRow 'categorias_detectadas' (Get-PropValue -Row $gov -Names @('categorias_detectadas'))

    $object = [pscustomobject]$auditRow
    [void]$results.Add($object)

    $reviewRow = [ordered]@{}
    Add-ColumnValue $reviewRow 'id_original' $idOriginal
    Add-ColumnValue $reviewRow 'id_normalizado' $id
    Add-ColumnValue $reviewRow 'cnpj' $cnpj
    Add-ColumnValue $reviewRow 'cnpj_fonte' $cnpjFonte
    Add-ColumnValue $reviewRow 'cnpj_enriquecimento' $enrichedCnpj
    Add-ColumnValue $reviewRow 'cnpj_diverge_enriquecimento' $diverge
    Add-ColumnValue $reviewRow 'enrichment_match_status' $matchStatus
    Add-ColumnValue $reviewRow 'linha_origem_enriquecimento' $(if ($enriched) { $enriched.__linha_origem } else { '' })
    Add-ColumnValue $reviewRow 'classificacao_fonte' (Get-ClassifierOutputValue -Row $base -BaseName 'classificacao_fonte')
    Add-ColumnValue $reviewRow 'classificacao_observacao' (Get-ClassifierOutputValue -Row $base -BaseName 'classificacao_observacao')
    Add-ColumnValue $reviewRow 'cnae_fiscal_principal' (Get-PropValue -Row $gov -Names @('cnae_fiscal_principal'))

    if ([string]::IsNullOrWhiteSpace($id) -or $matchStatus -ne 'matched' -or [string]::IsNullOrWhiteSpace($cnpj) -or $diverge -eq 'true') {
        [void]$review.Add([pscustomobject]$reviewRow)
    }
}

if ($duplicateBaseIds.Count -gt 0) { throw "Duplicate base IDs after normalization: $((@($duplicateBaseIds) | Select-Object -Unique) -join ', ')" }

$results | Export-Csv -LiteralPath $OutputPath -Delimiter $Delimiter -Encoding UTF8 -NoTypeInformation
$review | Export-Csv -LiteralPath $ReviewPath -Delimiter $Delimiter -Encoding UTF8 -NoTypeInformation
Invoke-RecentXmlDiffIfConfigured -ExpectedPath $OutputPath

$summary = [ordered]@{
    total_base_rows = $classifiedRows.Count
    total_final_rows = $results.Count
    matched_enrichment_rows = $matchedEnrichmentRows
    unmatched_enrichment_rows = $unmatchedEnrichmentRows
    duplicate_enrichment_ids = $duplicateEnrichmentIds
    safe_cnpj_count = $safeCnpjCount
    missing_cnpj_count = $missingCnpjCount
    cnpj_divergence_count = $cnpjDivergenceCount
    review_rows = $review.Count
    output_path = $OutputPath
    output_columns = if ($results.Count -gt 0) { @($results[0].PSObject.Properties).Count } else { 0 }
    review_path = $ReviewPath
    classified_output_path = $ClassifiedOutputPath
}
$summary | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $SummaryPath -Encoding UTF8

Write-Host '-------------------------------------------------' -ForegroundColor Gray
Write-Host "Final rows: $($results.Count)" -ForegroundColor Gray
Write-Host "Output columns: $($summary.output_columns)" -ForegroundColor Gray
Write-Host "Matched enrichment: $($summary.matched_enrichment_rows)" -ForegroundColor Gray
Write-Host "Review rows: $($review.Count)" -ForegroundColor Gray
Write-Host "Done: $OutputPath" -ForegroundColor Green
