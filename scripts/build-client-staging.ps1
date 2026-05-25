# =============================================================
# build-client-staging.ps1 - Build PostgreSQL-ready existing-client staging CSV
# =============================================================

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [string]$EnrichmentPath,

    [string]$SimplesPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [string]$ClassifiedOutputPath = '',

    [string]$ReviewPath = '',

    [string]$SummaryPath = '',

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

function Resolve-DefaultEnrichmentPath {
    $candidates = @(
        (Join-Path $scriptRoot 'data\operational-enrichment.csv'),
        (Join-Path $projectRoot 'data\operational-enrichment.csv'),
        (Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads\operational-enrichment.csv')
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) { return (Resolve-Path -LiteralPath $candidate).Path }
    }

    throw 'Enrichment CSV was not provided and the default operational-enrichment.csv was not found beside the executable, in project data, or in Downloads.'
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

function Convert-ToPgBoolText {
    param([object]$Value)
    $text = ([string]$Value).Trim().ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($text)) { return '' }
    if ($text -in @('sim', 's', 'yes', 'y', 'true', '1')) { return 'true' }
    if ($text -in @('nao', 'não', 'n', 'no', 'false', '0')) { return 'false' }
    return ''
}

function Add-ColumnValue {
    param([System.Collections.Specialized.OrderedDictionary]$Row, [string]$Name, [object]$Value)
    $Row[$Name] = if ($null -eq $Value) { '' } else { $Value }
}

function New-EmptyGovernmentRow {
    return [pscustomobject]@{
        cnae_fiscal_principal = ''; cnae_fiscal_secundaria = ''; categoria_principal = ''; categorias_detectadas = ''
        categoria_transporte = ''; categoria_gas_combustivel = ''; categoria_farmacia = ''; categoria_construtora = ''
        categoria_industria = ''; categoria_agro_produtor_rural = ''; categoria_telecom = ''; categoria_ti_software = ''
        categoria_saude_clinica = ''; categoria_financeiro = ''; categoria_cooperativa = ''; categoria_imobiliario = ''
        categoria_grafica = ''; categoria_frigorifico = ''; categoria_importador_exportador_status = 'external_source_required'
        categoria_beneficio_fiscal_status = 'external_source_required'
    }
}

function Export-GovernmentCategories {
    param([string]$Path)

    if (-not (Get-Command Invoke-CnpjPreflight -ErrorAction SilentlyContinue)) {
        throw 'Government data export requires scripts\lib\preflight.ps1.'
    }

    $bootstrap = Import-CnpjConfig -ScriptRoot $scriptRoot
    $config = $bootstrap.Config
    $timeout = if ($config.ContainsKey('sqlCommandTimeoutSeconds')) { [int]$config.sqlCommandTimeoutSeconds } else { 3600 }
    $tools = Invoke-CnpjPreflight -Config $config -EnsureDirectories @($config.dirTemp, $config.dirOut) -MinFreeSpaceGB @{ $config.dirTemp = 1; $config.dirOut = 1 } -RequirePostgres -RequireDatabase
    Set-PostgresPasswordEnv -Config $config

    $copyPath = ConvertTo-PsqlPathLiteral -Path $Path
    $sql = @"
COPY (
    WITH category_flags AS (
        SELECT
            cnpj,
            string_agg(DISTINCT categoria, ',' ORDER BY categoria) AS categorias_detectadas,
            max(categoria) FILTER (WHERE categoria_principal) AS categoria_principal,
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
        FROM estabelecimentos_categorias
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
    FROM estabelecimentos_crm e
    LEFT JOIN category_flags f ON f.cnpj = e.cnpj
    ORDER BY e.cnpj
) TO '$copyPath'
WITH (FORMAT CSV, HEADER TRUE, DELIMITER ';', ENCODING 'UTF8', NULL '');
"@

    $sqlFile = Join-Path ([System.IO.Path]::GetTempPath()) ("client-government-export-{0}.sql" -f ([guid]::NewGuid().ToString('N')))
    [System.IO.File]::WriteAllText($sqlFile, $sql, [System.Text.Encoding]::UTF8)
    try { Invoke-PsqlFileChecked -PsqlPath $tools.PsqlPath -Config $config -FilePath $sqlFile -Description 'export government CNAE categories' -TimeoutSec $timeout }
    finally { Remove-Item -LiteralPath $sqlFile -Force -ErrorAction SilentlyContinue }
}

if ($Delimiter.Length -ne 1) { throw 'Delimiter must be a single character.' }
$InputPath = Resolve-RequiredFileLocal -Path $InputPath -Label 'Input file'
if ([string]::IsNullOrWhiteSpace($EnrichmentPath)) { $EnrichmentPath = Resolve-DefaultEnrichmentPath }
else { $EnrichmentPath = Resolve-RequiredFileLocal -Path $EnrichmentPath -Label 'Enrichment CSV' }
$SimplesPath = Resolve-OptionalFileLocal -Path $SimplesPath -Label 'Simples CSV'

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
Write-Host " Simples:     $(if ($SimplesPath) { $SimplesPath } else { 'not provided; regime fields will stay unclassified' })" -ForegroundColor Yellow
Write-Host " Output:      $OutputPath" -ForegroundColor Yellow
Write-Host '-------------------------------------------------' -ForegroundColor Gray

$classifierArgs = @{ InputPath = $InputPath; OutputPath = $ClassifiedOutputPath; Delimiter = $Delimiter; Mode = $Mode }
if (-not [string]::IsNullOrWhiteSpace($SimplesPath)) { $classifierArgs.SimplesPath = $SimplesPath }
& $classifier @classifierArgs

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
        @{ Expression = { -($_.PSObject.Properties | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Value) }).Count } }, `
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

    $gov = if ($cnpj -and $governmentLookup.ContainsKey($cnpj)) { $governmentLookup[$cnpj] } else { New-EmptyGovernmentRow }
    $row = [ordered]@{}
    Add-ColumnValue $row 'id_original' $idOriginal
    Add-ColumnValue $row 'id_normalizado' $id
    Add-ColumnValue $row 'nome' (Get-PropValue -Row $base -Names @('Nome', 'nome'))
    Add-ColumnValue $row 'nome_normalizado' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('nome_normalizado') } else { '' })
    Add-ColumnValue $row 'razao_social' (Get-PropValue -Row $base -Names @('Razão Social', 'Razao Social', 'razao_social'))
    Add-ColumnValue $row 'razao_social_normalizada' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('razao_social_normalizada') } else { '' })
    Add-ColumnValue $row 'validade_senha' (Get-PropValue -Row $base -Names @('Validade da Senha', 'validade_senha'))
    Add-ColumnValue $row 'cidade' (Get-PropValue -Row $base -Names @('Cidade', 'cidade'))
    Add-ColumnValue $row 'estado' (Get-PropValue -Row $base -Names @('Estado', 'estado'))
    Add-ColumnValue $row 'cnpj' $cnpj
    Add-ColumnValue $row 'cnpj_formatado' (Format-CnpjLocal -Cnpj $cnpj)
    Add-ColumnValue $row 'cnpj_basico' $(if ($cnpj -match '^\d{14}$') { $cnpj.Substring(0, 8) } else { '' })
    Add-ColumnValue $row 'cnpj_fonte' $cnpjFonte
    Add-ColumnValue $row 'cnpj_enriquecimento' $enrichedCnpj
    Add-ColumnValue $row 'cnpj_fonte_enriquecimento' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('cnpj_fonte_correcao') } else { '' })
    Add-ColumnValue $row 'cnpj_diverge_enriquecimento' $diverge
    Add-ColumnValue $row 'regime_tributario' (Get-PropValue -Row $base -Names @('regime_tributario'))
    Add-ColumnValue $row 'classificacao_fonte' (Get-PropValue -Row $base -Names @('classificacao_fonte'))
    Add-ColumnValue $row 'classificacao_observacao' (Get-PropValue -Row $base -Names @('classificacao_observacao'))
    Add-ColumnValue $row 'postgres_versao' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('Versão Postgres') } else { '' })
    Add-ColumnValue $row 'postgres_versao_normalizada' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('postgres_versao_normalizada') } else { '' })
    Add-ColumnValue $row 'arquitetura_os' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('32/64') } else { '' })
    Add-ColumnValue $row 'arquitetura_os_normalizada' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('arquitetura_os_normalizada') } else { '' })
    Add-ColumnValue $row 'provedor' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('Provedor') } else { '' })
    Add-ColumnValue $row 'provedor_normalizado' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('provedor_normalizado') } else { '' })
    Add-ColumnValue $row 'cpf_cnpj_internet' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('CPF/CNPJ Internet') } else { '' })
    Add-ColumnValue $row 'cpf_cnpj_internet_normalizado' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('cpf_cnpj_internet_normalizado') } else { '' })
    Add-ColumnValue $row 'nome_internet' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('Nome Internet') } else { '' })
    Add-ColumnValue $row 'nome_internet_normalizado' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('nome_internet_normalizado') } else { '' })
    Add-ColumnValue $row 'atualizando' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('Atualizando') } else { '' })
    Add-ColumnValue $row 'atualizando_normalizado' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('atualizando_normalizado') } else { '' })
    Add-ColumnValue $row 'flag_postgres_17_ou_nuvem' $(if ($enriched) { Convert-ToPgBoolText (Get-PropValue -Row $enriched -Names @('Postgres 17 ou banco nuvem')) } else { '' })
    Add-ColumnValue $row 'flag_postgres_desatualizado_32_bits' $(if ($enriched) { Convert-ToPgBoolText (Get-PropValue -Row $enriched -Names @('Postgres desatualizado 32 bits')) } else { '' })
    Add-ColumnValue $row 'flag_postgres_atualizado_sem_replicacao' $(if ($enriched) { Convert-ToPgBoolText (Get-PropValue -Row $enriched -Names @('Postgres atualizado sem replicação')) } else { '' })
    Add-ColumnValue $row 'flag_multiplos_computadores' $(if ($enriched) { Convert-ToPgBoolText (Get-PropValue -Row $enriched -Names @('Múltiplos computadores')) } else { '' })
    Add-ColumnValue $row 'flag_sistema_web' $(if ($enriched) { Convert-ToPgBoolText (Get-PropValue -Row $enriched -Names @('Sistema web')) } else { '' })
    Add-ColumnValue $row 'flag_sistema_deluito_contador' $(if ($enriched) { Convert-ToPgBoolText (Get-PropValue -Row $enriched -Names @('Sistema Deluito contador')) } else { '' })
    Add-ColumnValue $row 'classificacao_por_cor' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('Classificação por cor') } else { '' })
    Add-ColumnValue $row 'status_linha_pela_cor' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('Status da linha pela cor') } else { '' })
    Add-ColumnValue $row 'cor_predominante_linha' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('Cor predominante da linha') } else { '' })
    Add-ColumnValue $row 'cor_celula_id' $(if ($enriched) { Get-PropValue -Row $enriched -Names @('Cor da célula ID') } else { '' })
    Add-ColumnValue $row 'enrichment_match_status' $matchStatus
    Add-ColumnValue $row 'linha_origem_enriquecimento' $(if ($enriched) { $enriched.__linha_origem } else { '' })
    foreach ($name in @('cnae_fiscal_principal','cnae_fiscal_secundaria','categoria_principal','categorias_detectadas','categoria_transporte','categoria_gas_combustivel','categoria_farmacia','categoria_construtora','categoria_industria','categoria_agro_produtor_rural','categoria_telecom','categoria_ti_software','categoria_saude_clinica','categoria_financeiro','categoria_cooperativa','categoria_imobiliario','categoria_grafica','categoria_frigorifico','categoria_importador_exportador_status','categoria_beneficio_fiscal_status')) {
        Add-ColumnValue $row $name (Get-PropValue -Row $gov -Names @($name))
    }

    $object = [pscustomobject]$row
    [void]$results.Add($object)
    if ([string]::IsNullOrWhiteSpace($id) -or $matchStatus -ne 'matched' -or [string]::IsNullOrWhiteSpace($cnpj) -or $diverge -eq 'true' -or ($IncludeGovernmentData -and [string]::IsNullOrWhiteSpace([string]$gov.cnae_fiscal_principal))) {
        [void]$review.Add($object)
    }
}

if ($duplicateBaseIds.Count -gt 0) { throw "Duplicate base IDs after normalization: $((@($duplicateBaseIds) | Select-Object -Unique) -join ', ')" }

$results | Export-Csv -LiteralPath $OutputPath -Delimiter $Delimiter -Encoding UTF8 -NoTypeInformation
$review | Export-Csv -LiteralPath $ReviewPath -Delimiter $Delimiter -Encoding UTF8 -NoTypeInformation

$summary = [ordered]@{
    total_base_rows = $classifiedRows.Count
    total_final_rows = $results.Count
    matched_enrichment_rows = @($results | Where-Object { $_.enrichment_match_status -eq 'matched' }).Count
    unmatched_enrichment_rows = @($results | Where-Object { $_.enrichment_match_status -eq 'sem_match_enriquecimento' }).Count
    duplicate_enrichment_ids = $duplicateEnrichmentIds
    safe_cnpj_count = @($results | Where-Object { $_.cnpj -match '^\d{14}$' }).Count
    missing_cnpj_count = @($results | Where-Object { $_.cnpj -notmatch '^\d{14}$' }).Count
    cnpj_divergence_count = @($results | Where-Object { $_.cnpj_diverge_enriquecimento -eq 'true' }).Count
    review_rows = $review.Count
    output_path = $OutputPath
    review_path = $ReviewPath
    classified_output_path = $ClassifiedOutputPath
}
$summary | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $SummaryPath -Encoding UTF8

Write-Host '-------------------------------------------------' -ForegroundColor Gray
Write-Host "Final rows: $($results.Count)" -ForegroundColor Gray
Write-Host "Matched enrichment: $($summary.matched_enrichment_rows)" -ForegroundColor Gray
Write-Host "Review rows: $($review.Count)" -ForegroundColor Gray
Write-Host "Done: $OutputPath" -ForegroundColor Green
