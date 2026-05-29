# =============================================================
# export-recent-xml-regime-diff.ps1 - Build a recent XML regime divergence report
# =============================================================

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ExpectedCsvPath,

    [string]$XmlRoot = '',

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [string]$AuditOutputPath = '',

    [string[]]$IncludeTypes = @('NF-e', 'NFC-e', 'NFS-e', 'CT-e'),

    [int]$MonthsBack = 2,

    [int]$MaxFilesPerCompanyType = 5,

    [int]$ThrottleLimit = 0,

    [char]$Delimiter = ';'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$projectRoot = Split-Path -Parent $scriptRoot

function Resolve-ConfiguredXmlRoot {
    param([string]$RequestedRoot)

    if (-not [string]::IsNullOrWhiteSpace($RequestedRoot)) { return $RequestedRoot }
    if (-not [string]::IsNullOrWhiteSpace($env:CNPJ_XML_ROOT)) { return $env:CNPJ_XML_ROOT }

    $configPath = Join-Path $projectRoot 'config.ps1'
    if (Test-Path -LiteralPath $configPath -PathType Leaf) {
        $config = & $configPath
        if ($config -is [hashtable]) {
            foreach ($key in @('xmlRoot', 'XmlRoot', 'xmlAuditRoot', 'XmlAuditRoot')) {
                if ($config.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$config[$key])) {
                    return [string]$config[$key]
                }
            }
        }
    }

    return ''
}

function Normalize-DigitsLocal {
    param([object]$Value)
    if ($null -eq $Value) { return '' }
    return ([string]$Value) -replace '\D', ''
}

function Normalize-CnpjLocal {
    param([object]$Value)
    $digits = Normalize-DigitsLocal -Value $Value
    if ([string]::IsNullOrWhiteSpace($digits)) { return '' }
    if ($digits.Length -gt 14) { return $digits }
    return $digits.PadLeft(14, '0')
}

function Get-FirstExistingPropertyNameLocal {
    param($Row, [string[]]$Candidates)
    $names = @($Row.PSObject.Properties.Name)
    foreach ($candidate in $Candidates) {
        if ($names -contains $candidate) { return $candidate }
    }
    return ''
}

function Get-FirstValueLocal {
    param($Row, [string[]]$Names)
    foreach ($name in $Names) {
        if ($Row.PSObject.Properties.Name -contains $name) {
            $value = $Row.$name
            if (-not [string]::IsNullOrWhiteSpace([string]$value)) { return [string]$value }
        }
    }
    return ''
}

function Import-ExpectedMetadata {
    param([string]$Path, [char]$DelimiterChar)

    $rows = @(Import-Csv -LiteralPath $Path -Delimiter $DelimiterChar -Encoding UTF8)
    if ($rows.Count -gt 0 -and @($rows[0].PSObject.Properties.Name).Count -eq 1 -and $DelimiterChar -ne ',') {
        $commaRows = @(Import-Csv -LiteralPath $Path -Delimiter ',' -Encoding UTF8)
        if ($commaRows.Count -gt 0 -and @($commaRows[0].PSObject.Properties.Name).Count -gt 1) { $rows = $commaRows }
    }
    if ($rows.Count -eq 0) { throw 'Expected CSV has no data rows.' }

    $cnpjColumn = Get-FirstExistingPropertyNameLocal -Row $rows[0] -Candidates @('cnpj', 'CNPJ', 'cnpj_normalizado', 'documento')
    if ([string]::IsNullOrWhiteSpace($cnpjColumn)) { throw 'Expected CSV must contain a CNPJ column.' }

    $metadata = @{}
    foreach ($row in $rows) {
        $cnpj = Normalize-CnpjLocal -Value $row.$cnpjColumn
        if ($cnpj.Length -ne 14 -or $metadata.ContainsKey($cnpj)) { continue }
        $metadata[$cnpj] = [pscustomobject]@{
            Nome = Get-FirstValueLocal -Row $row -Names @('nome', 'Nome', 'nome_fantasia', 'Nome Fantasia')
            RazaoSocial = Get-FirstValueLocal -Row $row -Names @('razao_social', 'Razão Social', 'Razao Social')
            Cidade = Get-FirstValueLocal -Row $row -Names @('cidade', 'Cidade', 'municipio_nome', 'municipio')
            Estado = Get-FirstValueLocal -Row $row -Names @('estado', 'Estado', 'uf', 'UF')
        }
    }
    return $metadata
}

function Get-ProblemGroupPtBrLocal {
    param($Row)
    switch ($Row.reason) {
        'ExpectedSimplesXmlNormal' { return 'Simples com XML Normal' }
        'ExpectedNormalXmlSimples' { return 'Normal com XML Simples' }
        'Sublimite' { return 'XML Simples com Excesso de Sublimite' }
        'XmlInternalConflict' { return 'XML com conflito interno' }
        'UnknownXmlRegime' { return 'Regime do XML indefinido' }
        'UnknownExpectedRegime' { return 'Regime esperado indefinido' }
        'IssuerMismatch' { return 'CNPJ do emitente diferente da pasta' }
        'NoRecentXml' { return 'Sem XML recente nos tipos auditados' }
        'NoCompanyXmlFolder' { return 'Sem pasta de XML para o CNPJ' }
        default { return $Row.reason }
    }
}

if ($MonthsBack -lt 1) { throw 'MonthsBack must be at least 1.' }

$ExpectedCsvPath = (Resolve-Path -LiteralPath $ExpectedCsvPath).Path
$resolvedXmlRoot = Resolve-ConfiguredXmlRoot -RequestedRoot $XmlRoot
if ([string]::IsNullOrWhiteSpace($resolvedXmlRoot)) {
    Write-Host 'Recent XML divergence report skipped: XmlRoot was not provided and CNPJ_XML_ROOT/config xmlRoot is empty.' -ForegroundColor Yellow
    exit 0
}
if (-not (Test-Path -LiteralPath $resolvedXmlRoot -PathType Container)) { throw "XmlRoot was not found: $resolvedXmlRoot" }

$outputParent = Split-Path -Parent $OutputPath
if (-not [string]::IsNullOrWhiteSpace($outputParent) -and -not (Test-Path -LiteralPath $outputParent)) { New-Item -ItemType Directory -Path $outputParent -Force | Out-Null }
if ([string]::IsNullOrWhiteSpace($AuditOutputPath)) { $AuditOutputPath = [System.IO.Path]::ChangeExtension($OutputPath, '.audit.csv') }

$cutoffMonth = (Get-Date -Day 1).AddMonths(-1 * ($MonthsBack - 1)).ToString('yyyyMM')
$auditorPath = Join-Path $scriptRoot 'audit-xml-regime.ps1'
if (-not (Test-Path -LiteralPath $auditorPath -PathType Leaf)) { throw "XML auditor script was not found: $auditorPath" }

Write-Host '===================================================' -ForegroundColor Cyan
Write-Host ' RECENT XML REGIME DIVERGENCE REPORT' -ForegroundColor Cyan
Write-Host '===================================================' -ForegroundColor Cyan
Write-Host " Expected: $ExpectedCsvPath" -ForegroundColor Yellow
Write-Host " XmlRoot:  $resolvedXmlRoot" -ForegroundColor Yellow
Write-Host " Months:   >= $cutoffMonth ($MonthsBack month window)" -ForegroundColor Yellow
Write-Host " Output:   $OutputPath" -ForegroundColor Yellow
Write-Host '-------------------------------------------------' -ForegroundColor Gray

& $auditorPath `
    -XmlRoot $resolvedXmlRoot `
    -ExpectedCsvPath $ExpectedCsvPath `
    -OutputPath $AuditOutputPath `
    -IncludeTypes $IncludeTypes `
    -LatestMonthsToTry $MonthsBack `
    -MaxFilesPerCompanyType $MaxFilesPerCompanyType `
    -ThrottleLimit $ThrottleLimit `
    -Delimiter $Delimiter | Out-Host

$auditRows = @(Import-Csv -LiteralPath $AuditOutputPath -Delimiter $Delimiter -Encoding UTF8)
$metadata = Import-ExpectedMetadata -Path $ExpectedCsvPath -DelimiterChar $Delimiter
$problems = @($auditRows | Where-Object {
        $_.invoice_month -ge $cutoffMonth -and $_.status -in @('Mismatch', 'Review')
    })

$report = foreach ($row in $problems) {
    $meta = if ($metadata.ContainsKey($row.cnpj)) { $metadata[$row.cnpj] } else { $null }
    [pscustomobject][ordered]@{
        grupo_divergencia = Get-ProblemGroupPtBrLocal -Row $row
        prioridade = if ($row.status -eq 'Mismatch') { 'Alta' } else { 'Revisar XML' }
        cnpj = $row.cnpj
        nome = if ($meta) { $meta.Nome } else { '' }
        razao_social = if ($meta) { $meta.RazaoSocial } else { '' }
        cidade = if ($meta) { $meta.Cidade } else { '' }
        estado = if ($meta) { $meta.Estado } else { '' }
        regime_classificado = $row.expected_regime_raw
        familia_classificada = $row.expected_regime_family
        regime_xml = $row.xml_regime_family
        crt_xml = $row.xml_crt
        crt_descricao = $row.xml_crt_label
        tipo_xml = $row.invoice_type
        competencia_xml = $row.invoice_month
        emissao_xml = $row.dh_emi
        modelo_xml = $row.xml_modelo
        numero_xml = $row.xml_numero
        arquivo_xml = $row.xml_file
        motivo = $row.reason
        consulta_status = ''
        consulta_regime_tributario = ''
        consulta_observacao = ''
    }
}

@($report) | Export-Csv -LiteralPath $OutputPath -Delimiter ';' -Encoding UTF8 -NoTypeInformation

Write-Host '-------------------------------------------------' -ForegroundColor Gray
Write-Host "Recent divergence rows: $(@($report).Count)" -ForegroundColor Gray
Write-Host "Done: $OutputPath" -ForegroundColor Green
