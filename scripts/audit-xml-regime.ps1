# =============================================================
# audit-xml-regime.ps1 - Audit expected regimes against NF-e/NFC-e XML evidence
# =============================================================

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$XmlRoot,

    [Parameter(Mandatory = $true)]
    [string]$ExpectedCsvPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [string[]]$IncludeTypes = @('NF-e', 'NFC-e'),

    [int]$LatestMonthsToTry = 3,

    [int]$MaxFilesPerCompanyType = 3,

    [switch]$OnlyProblems,

    [string]$ProblemReportPath = '',

    [string]$GroupedReportPath = '',

    [int]$ThrottleLimit = 0,

    [char]$Delimiter = ';'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Resolve-RequiredPath {
    param([string]$Path, [string]$Label, [switch]$Directory)

    $pathType = if ($Directory) { 'Container' } else { 'Leaf' }
    if (-not (Test-Path -LiteralPath $Path -PathType $pathType)) {
        throw "$Label was not found: $Path"
    }

    return (Resolve-Path -LiteralPath $Path).Path
}

function Normalize-Digits {
    param([object]$Value)

    if ($null -eq $Value) { return '' }
    return ([string]$Value) -replace '\D', ''
}

function Normalize-Cnpj {
    param([object]$Value)

    $digits = Normalize-Digits -Value $Value
    if ([string]::IsNullOrWhiteSpace($digits)) { return '' }
    if ($digits.Length -gt 14) { return $digits }
    return $digits.PadLeft(14, '0')
}

function Get-FirstExistingPropertyName {
    param($Row, [string[]]$Candidates)

    $names = @($Row.PSObject.Properties.Name)
    foreach ($candidate in $Candidates) {
        if ($names -contains $candidate) { return $candidate }
    }

    return ''
}

function Normalize-ExpectedRegime {
    param([object]$Value)

    $raw = if ($null -eq $Value) { '' } else { ([string]$Value).Trim() }
    $normalized = ($raw.ToLowerInvariant() -replace '[áàãâä]', 'a' -replace '[éèêë]', 'e' -replace '[íìîï]', 'i' -replace '[óòõôö]', 'o' -replace '[úùûü]', 'u' -replace 'ç', 'c')

    if ([string]::IsNullOrWhiteSpace($normalized)) { return 'Unknown' }
    if ($normalized -in @('mei', 'simples', 'simples nacional')) { return 'Simples' }
    if ($normalized -in @('normal', 'regime normal', 'lucro presumido', 'lucro real')) { return 'Normal' }

    return 'Unknown'
}

function Get-CrtLabel {
    param([string]$Crt)

    switch ($Crt) {
        '1' { return 'Simples Nacional' }
        '2' { return 'Simples Nacional - Excesso de Sublimite' }
        '3' { return 'Regime Normal' }
        default { return 'Unknown' }
    }
}

function Import-ExpectedRegimes {
    param([string]$Path, [char]$DelimiterChar)

    $rows = @(Import-Csv -LiteralPath $Path -Delimiter $DelimiterChar -Encoding UTF8)
    if ($rows.Count -gt 0 -and @($rows[0].PSObject.Properties.Name).Count -eq 1 -and $DelimiterChar -ne ',') {
        $commaRows = @(Import-Csv -LiteralPath $Path -Delimiter ',' -Encoding UTF8)
        if ($commaRows.Count -gt 0 -and @($commaRows[0].PSObject.Properties.Name).Count -gt 1) {
            $rows = $commaRows
            $script:DetectedInputDelimiter = ','
        }
    }
    if ($rows.Count -eq 0) { throw 'Expected regime CSV has no data rows.' }

    $cnpjColumn = Get-FirstExistingPropertyName -Row $rows[0] -Candidates @('cnpj', 'CNPJ', 'cnpj_normalizado', 'documento')
    $regimeColumn = Get-FirstExistingPropertyName -Row $rows[0] -Candidates @('expected_regime', 'regime_tributario', 'regime', 'tipo_regime')

    if ([string]::IsNullOrWhiteSpace($cnpjColumn)) {
        throw 'Expected regime CSV must contain a CNPJ column. Supported names: cnpj, CNPJ, cnpj_normalizado, documento.'
    }
    if ([string]::IsNullOrWhiteSpace($regimeColumn)) {
        throw 'Expected regime CSV must contain a regime column. Supported names: expected_regime, regime_tributario, regime, tipo_regime.'
    }

    $expected = @{}
    foreach ($row in $rows) {
        $cnpj = Normalize-Cnpj -Value $row.$cnpjColumn
        $rawRegime = if ($null -eq $row.$regimeColumn) { '' } else { ([string]$row.$regimeColumn).Trim() }
        if ($cnpj.Length -ne 14) { continue }

        if (-not $expected.ContainsKey($cnpj)) {
            $expected[$cnpj] = [pscustomobject]@{
                Cnpj = $cnpj
                ExpectedRegimeRaw = $rawRegime
                ExpectedRegimeFamily = Normalize-ExpectedRegime -Value $rawRegime
            }
        }
    }

    if ($expected.Count -eq 0) { throw 'Expected regime CSV did not contain any valid 14-digit CNPJ.' }
    return $expected
}

function Get-ProblemGroupPtBr {
    param($Row)

    switch ($Row.reason) {
        'ExpectedSimplesXmlNormal' { return 'Simples com XML Normal' }
        'ExpectedNormalXmlSimples' { return 'Normal com XML Simples' }
        'Sublimite' { return 'XML Simples com Excesso de Sublimite' }
        'XmlInternalConflict' { return 'XML com conflito interno' }
        'UnknownXmlRegime' { return 'Regime do XML indefinido' }
        'UnknownExpectedRegime' { return 'Regime esperado indefinido' }
        'IssuerMismatch' { return 'CNPJ do emitente diferente da pasta' }
        'NoCompanyXmlFolder' { return 'Sem pasta de XML para o CNPJ' }
        'NoRecentXml' { return 'Sem XML recente nos tipos auditados' }
        'ParseError' { return 'Erro ao ler XML' }
        default {
            if ($Row.status -eq 'Mismatch') { return 'Divergencia de regime' }
            if ($Row.status -eq 'Review') { return 'Revisao manual' }
            if ($Row.status -eq 'NoEvidence') { return 'Sem evidencia XML' }
            if ($Row.status -eq 'Error') { return 'Erro' }
            return $Row.status
        }
    }
}

function ConvertTo-ProblemReportRow {
    param($Row)

    return [pscustomobject][ordered]@{
        grupo_divergencia = Get-ProblemGroupPtBr -Row $Row
        tipo_xml = $Row.invoice_type
        status = $Row.status
        motivo = $Row.reason
        cnpj = $Row.cnpj
        regime_esperado = $Row.expected_regime_raw
        familia_esperada = $Row.expected_regime_family
        familia_xml = $Row.xml_regime_family
        crt_xml = $Row.xml_crt
        crt_descricao = $Row.xml_crt_label
        competencia_xml = $Row.invoice_month
        emissao_xml = $Row.dh_emi
        modelo_xml = $Row.xml_modelo
        numero_xml = $Row.xml_numero
        emitente_xml = $Row.xml_emit_cnpj
        arquivo_xml = $Row.xml_file
        observacao = $Row.notes
    }
}

function Write-OptionalCsv {
    param($Rows, [string]$Path, [char]$DelimiterChar)

    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    $parent = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($parent) -and -not (Test-Path -LiteralPath $parent)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    @($Rows) | Export-Csv -LiteralPath $Path -Delimiter $DelimiterChar -Encoding UTF8 -NoTypeInformation
}

function Get-RecentMonthDirectories {
    param([string]$TypeDirectory, [int]$Limit)

    if (-not (Test-Path -LiteralPath $TypeDirectory -PathType Container)) { return @() }

    return @(Get-ChildItem -LiteralPath $TypeDirectory -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^\d{6}$' } |
        Sort-Object Name -Descending |
        Select-Object -First $Limit)
}

function Get-CandidateXmlFiles {
    param(
        [string]$CompanyDirectory,
        [string[]]$Types,
        [int]$LatestMonths,
        [int]$MaxFilesPerType
    )

    $candidates = New-Object System.Collections.ArrayList

    foreach ($type in $Types) {
        $typeCandidates = New-Object System.Collections.ArrayList
        foreach ($typeDirectoryInfo in @(Get-InvoiceTypeDirectories -CompanyDirectory $CompanyDirectory -Type $type)) {
            if ($typeCandidates.Count -ge $MaxFilesPerType) { break }
            $typeDirectory = $typeDirectoryInfo.Directory
            $canonicalType = $typeDirectoryInfo.CanonicalType
            foreach ($monthDirectory in @(Get-RecentMonthDirectories -TypeDirectory $typeDirectory -Limit $LatestMonths)) {
                if ($typeCandidates.Count -ge $MaxFilesPerType) { break }
                $remaining = $MaxFilesPerType - $typeCandidates.Count
                $files = @(Get-ChildItem -LiteralPath $monthDirectory.FullName -Filter '*.xml' -File -Recurse -ErrorAction SilentlyContinue |
                    Sort-Object LastWriteTime -Descending |
                    Select-Object -First $remaining)

                foreach ($file in $files) {
                    [void]$typeCandidates.Add([pscustomobject]@{
                        Path = $file.FullName
                        InvoiceType = $canonicalType
                        InvoiceMonth = $monthDirectory.Name
                        LastWriteTime = $file.LastWriteTime
                    })
                }
            }
        }
        foreach ($candidate in $typeCandidates) { [void]$candidates.Add($candidate) }
    }

    return @($candidates)
}

function Get-InvoiceTypeDirectories {
    param([string]$CompanyDirectory, [string]$Type)

    $canonical = Get-CanonicalInvoiceType -Type $Type
    $aliases = switch ($canonical) {
        'NF-e' { @('NF-e', 'NFE') }
        'NFC-e' { @('NFC-e', 'NFCE') }
        'NFS-e' { @('NFS-e', 'NFSE', 'NFS') }
        'CT-e' { @('CT-e', 'CTE') }
        default { @($Type) }
    }

    $result = New-Object System.Collections.ArrayList
    foreach ($alias in $aliases) {
        $directory = Join-Path $CompanyDirectory $alias
        if (Test-Path -LiteralPath $directory -PathType Container) {
            [void]$result.Add([pscustomobject]@{ Directory = $directory; CanonicalType = $canonical })
        }
    }
    return @($result)
}

function Get-CanonicalInvoiceType {
    param([string]$Type)

    $normalized = (([string]$Type).Trim().ToUpperInvariant() -replace '[^A-Z0-9]', '')
    switch ($normalized) {
        'NFE' { return 'NF-e' }
        'NFCE' { return 'NFC-e' }
        'CTE' { return 'CT-e' }
        'NFSE' { return 'NFS-e' }
        'NFS' { return 'NFS-e' }
        default { return ([string]$Type).Trim() }
    }
}

function Test-StackContains {
    param([System.Collections.ArrayList]$Stack, [string]$Name)

    foreach ($item in $Stack) {
        if ($item -eq $Name) { return $true }
    }
    return $false
}

function Get-XmlRegimeEvidence {
    param(
        [string]$Path,
        [string]$ExpectedFolderCnpj,
        [string]$InvoiceType,
        [string]$InvoiceMonth,
        [datetime]$LastWriteTime
    )

    $settings = New-Object System.Xml.XmlReaderSettings
    $settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
    $settings.IgnoreComments = $true
    $settings.IgnoreProcessingInstructions = $true

    $emitCnpj = ''
    $crt = ''
    $dhEmi = ''
    $modelo = ''
    $numero = ''
    $opSimpNac = ''
    $regEspTrib = ''
    $icmsGroup = ''
    $hasCsosn = $false
    $hasIcmsCst = $false
    $hasSimplesIcmsGroup = $false
    $hasNormalIcmsGroup = $false
    $currentElement = ''
    $stack = New-Object System.Collections.ArrayList
    $reader = $null

    try {
        $reader = [System.Xml.XmlReader]::Create($Path, $settings)
        while ($reader.Read()) {
            if ($reader.NodeType -eq [System.Xml.XmlNodeType]::Element) {
                $localName = $reader.LocalName
                [void]$stack.Add($localName)
                $currentElement = $localName

                if ((Test-StackContains -Stack $stack -Name 'ICMS') -and $localName -ne 'ICMS') {
                    if ([string]::IsNullOrWhiteSpace($icmsGroup) -and $localName -match '^ICMS') { $icmsGroup = $localName }
                    if ($localName -match '^ICMSSN$|^ICMSSN\d+') { $hasSimplesIcmsGroup = $true }
                    elseif ($localName -match '^ICMS\d+') { $hasNormalIcmsGroup = $true }
                }

                if ($reader.IsEmptyElement) {
                    [void]$stack.RemoveAt($stack.Count - 1)
                    $currentElement = if ($stack.Count -gt 0) { [string]$stack[$stack.Count - 1] } else { '' }
                }
                continue
            }

            if ($reader.NodeType -eq [System.Xml.XmlNodeType]::Text -or $reader.NodeType -eq [System.Xml.XmlNodeType]::CDATA) {
                $value = $reader.Value.Trim()
                if ([string]::IsNullOrWhiteSpace($value)) { continue }

                $insideEmit = Test-StackContains -Stack $stack -Name 'emit'
                $insideIde = Test-StackContains -Stack $stack -Name 'ide'
                $insideIcms = Test-StackContains -Stack $stack -Name 'ICMS'

                if ($insideEmit -and $currentElement -eq 'CNPJ' -and [string]::IsNullOrWhiteSpace($emitCnpj)) { $emitCnpj = Normalize-Cnpj -Value $value }
                elseif ($insideEmit -and $currentElement -eq 'CRT' -and [string]::IsNullOrWhiteSpace($crt)) { $crt = $value }
                elseif ($insideIde -and $currentElement -eq 'dhEmi' -and [string]::IsNullOrWhiteSpace($dhEmi)) { $dhEmi = $value }
                elseif ($insideIde -and $currentElement -eq 'mod' -and [string]::IsNullOrWhiteSpace($modelo)) { $modelo = $value }
                elseif ($insideIde -and $currentElement -eq 'nNF' -and [string]::IsNullOrWhiteSpace($numero)) { $numero = $value }
                elseif ($insideIde -and $currentElement -eq 'nCT' -and [string]::IsNullOrWhiteSpace($numero)) { $numero = $value }
                elseif ($currentElement -eq 'nNFSe' -and [string]::IsNullOrWhiteSpace($numero)) { $numero = $value }
                elseif ($currentElement -eq 'dhProc' -and [string]::IsNullOrWhiteSpace($dhEmi)) { $dhEmi = $value }
                elseif ($currentElement -eq 'opSimpNac' -and [string]::IsNullOrWhiteSpace($opSimpNac)) { $opSimpNac = $value }
                elseif ($currentElement -eq 'regEspTrib' -and [string]::IsNullOrWhiteSpace($regEspTrib)) { $regEspTrib = $value }
                elseif ($insideIcms -and $currentElement -eq 'CSOSN') { $hasCsosn = $true }
                elseif ($insideIcms -and $currentElement -eq 'CST') { $hasIcmsCst = $true }
                elseif ($insideIcms -and $currentElement -eq 'indSN' -and $value -eq '1') { $hasSimplesIcmsGroup = $true }
                continue
            }

            if ($reader.NodeType -eq [System.Xml.XmlNodeType]::EndElement) {
                if ($stack.Count -gt 0) { [void]$stack.RemoveAt($stack.Count - 1) }
                $currentElement = if ($stack.Count -gt 0) { [string]$stack[$stack.Count - 1] } else { '' }
            }
        }
    }
    finally {
        if ($reader) { $reader.Close() }
    }

    $regimeFamily = 'Unknown'
    $confidence = 'Low'
    $reason = ''

    switch ($crt) {
        '1' { $regimeFamily = 'Simples'; $confidence = 'High' }
        '2' { $regimeFamily = 'SimplesExcessoSublimite'; $confidence = 'High' }
        '3' { $regimeFamily = 'Normal'; $confidence = 'High' }
        default {
            if ($InvoiceType -eq 'NFS-e' -and $opSimpNac -in @('2', '3')) { $regimeFamily = 'Simples'; $confidence = 'Medium' }
            elseif ($InvoiceType -eq 'NFS-e' -and $opSimpNac -eq '1') { $regimeFamily = 'Normal'; $confidence = 'Medium' }
            elseif ($hasSimplesIcmsGroup -or $hasCsosn) { $regimeFamily = 'Simples'; $confidence = 'Medium' }
            elseif ($hasNormalIcmsGroup -or $hasIcmsCst) { $regimeFamily = 'Normal'; $confidence = 'Medium' }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($crt)) {
        if (($crt -in @('1', '2')) -and ($hasNormalIcmsGroup -or $hasIcmsCst) -and -not ($hasSimplesIcmsGroup -or $hasCsosn)) { $reason = 'XmlInternalConflict' }
        elseif ($crt -eq '3' -and ($hasSimplesIcmsGroup -or $hasCsosn)) { $reason = 'XmlInternalConflict' }
    }

    if ([string]::IsNullOrWhiteSpace($emitCnpj) -or $emitCnpj -ne $ExpectedFolderCnpj) {
        $reason = 'IssuerMismatch'
    }

    $parsedDate = $null
    $parsedDateValue = [datetime]::MinValue
    if (-not [string]::IsNullOrWhiteSpace($dhEmi) -and [datetime]::TryParse($dhEmi, [ref]$parsedDateValue)) {
        $parsedDate = $parsedDateValue
    }

    return [pscustomobject]@{
        XmlFile = $Path
        InvoiceType = $InvoiceType
        InvoiceMonth = $InvoiceMonth
        LastWriteTime = $LastWriteTime
        EffectiveDate = if ($parsedDate) { $parsedDate } else { $LastWriteTime }
        DhEmi = $dhEmi
        EmitCnpj = $emitCnpj
        Modelo = $modelo
        Numero = $numero
        Crt = $crt
        CrtLabel = if ($InvoiceType -eq 'NFS-e' -and -not [string]::IsNullOrWhiteSpace($opSimpNac)) { "NFS-e opSimpNac=$opSimpNac regEspTrib=$regEspTrib" } else { Get-CrtLabel -Crt $crt }
        XmlRegimeFamily = $regimeFamily
        Confidence = $confidence
        IcmsGroupSample = $icmsGroup
        HasCsosn = $hasCsosn
        HasIcmsCst = $hasIcmsCst
        EvidenceReason = $reason
    }
}

function Normalize-IncludeTypes {
    param([string[]]$Types)

    $normalized = New-Object System.Collections.ArrayList
    foreach ($type in $Types) {
        foreach ($part in (([string]$type) -split ',')) {
            $trimmed = $part.Trim()
            if (-not [string]::IsNullOrWhiteSpace($trimmed)) {
                $canonical = Get-CanonicalInvoiceType -Type $trimmed
                if (-not $normalized.Contains($canonical)) { [void]$normalized.Add($canonical) }
            }
        }
    }
    return @($normalized)
}

function Select-BestEvidence {
    param($EvidenceRows)

    $valid = @($EvidenceRows | Where-Object { $_.EvidenceReason -ne 'IssuerMismatch' })
    if ($valid.Count -eq 0) { return $null }

    return @($valid | Sort-Object @{ Expression = 'EffectiveDate'; Descending = $true }, @{ Expression = { if ([string]::IsNullOrWhiteSpace($_.Crt)) { 0 } else { 1 } }; Descending = $true }, @{ Expression = 'LastWriteTime'; Descending = $true } | Select-Object -First 1)[0]
}

function New-EmptyAuditRow {
    param($Expected, [string]$Status, [string]$Reason, [string]$Notes)

    return [pscustomobject][ordered]@{
        cnpj = $Expected.Cnpj
        expected_regime_raw = $Expected.ExpectedRegimeRaw
        expected_regime_family = $Expected.ExpectedRegimeFamily
        xml_regime_family = ''
        xml_crt = ''
        xml_crt_label = ''
        status = $Status
        reason = $Reason
        confidence = ''
        invoice_type = ''
        invoice_month = ''
        dh_emi = ''
        xml_emit_cnpj = ''
        xml_modelo = ''
        xml_numero = ''
        xml_file = ''
        icms_group_sample = ''
        has_csosn = ''
        has_icms_cst = ''
        notes = $Notes
    }
}

function Resolve-RegimeComparison {
    param($Expected, $Evidence)

    if ($Expected.ExpectedRegimeFamily -eq 'Unknown') {
        $status = 'Review'; $reason = 'UnknownExpectedRegime'
    }
    elseif ($Evidence.EvidenceReason -eq 'XmlInternalConflict') {
        $status = 'Review'; $reason = 'XmlInternalConflict'
    }
    elseif ($Evidence.XmlRegimeFamily -eq 'Unknown') {
        $status = 'Review'; $reason = 'UnknownXmlRegime'
    }
    elseif ($Evidence.XmlRegimeFamily -eq 'SimplesExcessoSublimite') {
        $status = 'Review'; $reason = 'Sublimite'
    }
    elseif ($Expected.ExpectedRegimeFamily -eq $Evidence.XmlRegimeFamily) {
        $status = 'OK'; $reason = 'MatchedExpectedRegime'
    }
    elseif ($Expected.ExpectedRegimeFamily -eq 'Simples' -and $Evidence.XmlRegimeFamily -eq 'Normal') {
        $status = 'Mismatch'; $reason = 'ExpectedSimplesXmlNormal'
    }
    elseif ($Expected.ExpectedRegimeFamily -eq 'Normal' -and $Evidence.XmlRegimeFamily -eq 'Simples') {
        $status = 'Mismatch'; $reason = 'ExpectedNormalXmlSimples'
    }
    else {
        $status = 'Review'; $reason = 'UnhandledComparison'
    }

    return [pscustomobject][ordered]@{
        cnpj = $Expected.Cnpj
        expected_regime_raw = $Expected.ExpectedRegimeRaw
        expected_regime_family = $Expected.ExpectedRegimeFamily
        xml_regime_family = $Evidence.XmlRegimeFamily
        xml_crt = $Evidence.Crt
        xml_crt_label = $Evidence.CrtLabel
        status = $status
        reason = $reason
        confidence = $Evidence.Confidence
        invoice_type = $Evidence.InvoiceType
        invoice_month = $Evidence.InvoiceMonth
        dh_emi = $Evidence.DhEmi
        xml_emit_cnpj = $Evidence.EmitCnpj
        xml_modelo = $Evidence.Modelo
        xml_numero = $Evidence.Numero
        xml_file = $Evidence.XmlFile
        icms_group_sample = $Evidence.IcmsGroupSample
        has_csosn = $Evidence.HasCsosn
        has_icms_cst = $Evidence.HasIcmsCst
        notes = if ($status -eq 'Mismatch') { 'Expected regime differs from emitted XML regime evidence.' } else { '' }
    }
}

function Invoke-CompanyAudit {
    param(
        $Expected,
        [string]$CompanyDirectory,
        [string[]]$Types,
        [int]$LatestMonths,
        [int]$MaxFilesPerType
    )

    $parseErrors = 0
    if ([string]::IsNullOrWhiteSpace($CompanyDirectory)) {
        return [pscustomobject]@{
            Row = (New-EmptyAuditRow -Expected $Expected -Status 'NoEvidence' -Reason 'NoCompanyXmlFolder' -Notes 'No matching CNPJ folder was found under XmlRoot.')
            ParseErrors = $parseErrors
        }
    }

    $candidates = @(Get-CandidateXmlFiles -CompanyDirectory $CompanyDirectory -Types $Types -LatestMonths $LatestMonths -MaxFilesPerType $MaxFilesPerType)
    if ($candidates.Count -eq 0) {
        return [pscustomobject]@{
            Row = (New-EmptyAuditRow -Expected $Expected -Status 'NoEvidence' -Reason 'NoRecentXml' -Notes 'No XML files were found in the selected invoice types and recent month window.')
            ParseErrors = $parseErrors
        }
    }

    $evidenceRows = New-Object System.Collections.ArrayList
    $lastParseError = ''
    foreach ($candidate in $candidates) {
        try {
            [void]$evidenceRows.Add((Get-XmlRegimeEvidence -Path $candidate.Path -ExpectedFolderCnpj $Expected.Cnpj -InvoiceType $candidate.InvoiceType -InvoiceMonth $candidate.InvoiceMonth -LastWriteTime $candidate.LastWriteTime))
        }
        catch {
            $parseErrors++
            $lastParseError = $_.Exception.Message
        }
    }

    if ($evidenceRows.Count -eq 0) {
        return [pscustomobject]@{
            Row = (New-EmptyAuditRow -Expected $Expected -Status 'Error' -Reason 'ParseError' -Notes $lastParseError)
            ParseErrors = $parseErrors
        }
    }

    $issuerMismatch = @($evidenceRows | Where-Object { $_.EvidenceReason -eq 'IssuerMismatch' })
    $bestEvidence = Select-BestEvidence -EvidenceRows $evidenceRows
    if ($null -eq $bestEvidence) {
        $row = New-EmptyAuditRow -Expected $Expected -Status 'Review' -Reason 'IssuerMismatch' -Notes 'Recent XML evidence exists, but emitter CNPJ does not match the company folder.'
        if ($issuerMismatch.Count -gt 0) {
            $row.xml_emit_cnpj = $issuerMismatch[0].EmitCnpj
            $row.xml_file = $issuerMismatch[0].XmlFile
        }
        return [pscustomobject]@{ Row = $row; ParseErrors = $parseErrors }
    }

    return [pscustomobject]@{
        Row = (Resolve-RegimeComparison -Expected $Expected -Evidence $bestEvidence)
        ParseErrors = $parseErrors
    }
}

function Split-IntoChunks {
    param($Items, [int]$ChunkCount)

    $chunks = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $ChunkCount; $i++) { [void]$chunks.Add((New-Object System.Collections.ArrayList)) }
    for ($i = 0; $i -lt $Items.Count; $i++) {
        [void]$chunks[$i % $ChunkCount].Add($Items[$i])
    }
    return @($chunks | Where-Object { $_.Count -gt 0 })
}

if ($LatestMonthsToTry -lt 1) { throw 'LatestMonthsToTry must be at least 1.' }
if ($MaxFilesPerCompanyType -lt 1) { throw 'MaxFilesPerCompanyType must be at least 1.' }
if ($IncludeTypes.Count -eq 0) { throw 'At least one invoice type must be included.' }
if ($ThrottleLimit -lt 0) { throw 'ThrottleLimit must be zero or greater.' }
$IncludeTypes = @(Normalize-IncludeTypes -Types $IncludeTypes)
if ($IncludeTypes.Count -eq 0) { throw 'At least one invoice type must be included.' }
if ($ThrottleLimit -eq 0) { $ThrottleLimit = [Math]::Min([Math]::Max([Environment]::ProcessorCount, 2), 8) }

$XmlRoot = Resolve-RequiredPath -Path $XmlRoot -Label 'XML root' -Directory
$ExpectedCsvPath = Resolve-RequiredPath -Path $ExpectedCsvPath -Label 'Expected regime CSV'
$script:DetectedInputDelimiter = $Delimiter

$outputParent = Split-Path -Parent $OutputPath
if (-not [string]::IsNullOrWhiteSpace($outputParent) -and -not (Test-Path -LiteralPath $outputParent)) {
    New-Item -ItemType Directory -Path $outputParent -Force | Out-Null
}

Write-Host '===================================================' -ForegroundColor Cyan
Write-Host " XML REGIME AUDITOR - $($IncludeTypes -join '/')" -ForegroundColor Cyan
Write-Host '===================================================' -ForegroundColor Cyan
Write-Host " XML root:    $XmlRoot" -ForegroundColor Yellow
Write-Host " Expected:    $ExpectedCsvPath" -ForegroundColor Yellow
Write-Host " Output:      $OutputPath" -ForegroundColor Yellow
Write-Host " Types:       $($IncludeTypes -join ', ')" -ForegroundColor Yellow
Write-Host " Workers:     $ThrottleLimit" -ForegroundColor Yellow
Write-Host '-------------------------------------------------' -ForegroundColor Gray

$expectedByCnpj = Import-ExpectedRegimes -Path $ExpectedCsvPath -DelimiterChar $Delimiter
if ($script:DetectedInputDelimiter -ne $Delimiter) {
    Write-Host " Input delimiter detected: $script:DetectedInputDelimiter" -ForegroundColor Yellow
}
$companyDirectories = @(Get-ChildItem -LiteralPath $XmlRoot -Directory -ErrorAction SilentlyContinue)
$companyDirectoryByCnpj = @{}
foreach ($companyDirectory in $companyDirectories) {
    $folderCnpj = Normalize-Cnpj -Value $companyDirectory.Name
    if ($folderCnpj.Length -eq 14 -and -not $companyDirectoryByCnpj.ContainsKey($folderCnpj)) {
        $companyDirectoryByCnpj[$folderCnpj] = $companyDirectory.FullName
    }
}

$results = New-Object System.Collections.ArrayList
$parseErrorCount = 0
$expectedItems = @($expectedByCnpj.Keys | Sort-Object | ForEach-Object { $expectedByCnpj[$_] })

if ($ThrottleLimit -le 1 -or $expectedItems.Count -le 1) {
    foreach ($expected in $expectedItems) {
        $companyDirectory = if ($companyDirectoryByCnpj.ContainsKey($expected.Cnpj)) { $companyDirectoryByCnpj[$expected.Cnpj] } else { '' }
        $companyResult = Invoke-CompanyAudit -Expected $expected -CompanyDirectory $companyDirectory -Types $IncludeTypes -LatestMonths $LatestMonthsToTry -MaxFilesPerType $MaxFilesPerCompanyType
        [void]$results.Add($companyResult.Row)
        $parseErrorCount += $companyResult.ParseErrors
    }
}
else {
    $workerCount = [Math]::Min($ThrottleLimit, $expectedItems.Count)
    $chunks = @(Split-IntoChunks -Items $expectedItems -ChunkCount $workerCount)
    $functionNames = @(
        'Normalize-Digits', 'Normalize-Cnpj', 'Get-CrtLabel', 'Get-RecentMonthDirectories',
        'Get-CandidateXmlFiles', 'Get-InvoiceTypeDirectories', 'Get-CanonicalInvoiceType',
        'Test-StackContains', 'Get-XmlRegimeEvidence', 'Select-BestEvidence',
        'New-EmptyAuditRow', 'Resolve-RegimeComparison', 'Invoke-CompanyAudit'
    )
    $functionText = ($functionNames | ForEach-Object { "function $_ {`n$((Get-Content function:\$_) -join "`n")`n}" }) -join "`n`n"
    $jobs = New-Object System.Collections.ArrayList

    foreach ($chunk in $chunks) {
        $job = Start-Job -ScriptBlock {
            param($Functions, $ExpectedChunk, $CompanyMap, $Types, $LatestMonths, $MaxFiles)
            Set-StrictMode -Version Latest
            $ErrorActionPreference = 'Stop'
            Invoke-Expression $Functions
            $rows = New-Object System.Collections.ArrayList
            $errors = 0
            foreach ($expected in $ExpectedChunk) {
                $directory = if ($CompanyMap.ContainsKey($expected.Cnpj)) { $CompanyMap[$expected.Cnpj] } else { '' }
                $result = Invoke-CompanyAudit -Expected $expected -CompanyDirectory $directory -Types $Types -LatestMonths $LatestMonths -MaxFilesPerType $MaxFiles
                [void]$rows.Add($result.Row)
                $errors += $result.ParseErrors
            }
            [pscustomobject]@{ Rows = @($rows); ParseErrors = $errors }
        } -ArgumentList $functionText, @($chunk), $companyDirectoryByCnpj, $IncludeTypes, $LatestMonthsToTry, $MaxFilesPerCompanyType
        [void]$jobs.Add($job)
    }

    $completedJobs = 0
    try {
        while ($jobs.Count -gt 0) {
            $finished = Wait-Job -Job @($jobs) -Any
            $jobResult = Receive-Job -Job $finished -ErrorAction Stop
            foreach ($row in @($jobResult.Rows)) { [void]$results.Add($row) }
            $parseErrorCount += $jobResult.ParseErrors
            $completedJobs++
            Write-Host ("  Worker chunk {0}/{1} completed ({2} rows)." -f $completedJobs, $chunks.Count, @($jobResult.Rows).Count) -ForegroundColor Gray
            Remove-Job -Job $finished -Force -ErrorAction SilentlyContinue
            $jobs = @($jobs | Where-Object { $_.Id -ne $finished.Id })
        }
    }
    finally {
        foreach ($job in $jobs) { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue }
    }
}

$sortedResults = New-Object System.Collections.ArrayList
foreach ($row in @($results | Sort-Object cnpj)) { [void]$sortedResults.Add($row) }
$results = $sortedResults

$finalResults = @($results)
if ($OnlyProblems) {
    $finalResults = @($finalResults | Where-Object { $_.status -ne 'OK' })
}

$finalResults | Export-Csv -LiteralPath $OutputPath -Delimiter $Delimiter -Encoding UTF8 -NoTypeInformation

$problemRows = @($results | Where-Object { $_.status -ne 'OK' } | ForEach-Object { ConvertTo-ProblemReportRow -Row $_ } | Sort-Object grupo_divergencia, tipo_xml, cnpj)
Write-OptionalCsv -Rows $problemRows -Path $ProblemReportPath -DelimiterChar $Delimiter

$groupedRows = @($problemRows |
    Group-Object grupo_divergencia, tipo_xml |
    ForEach-Object {
        $first = $_.Group[0]
        [pscustomobject][ordered]@{
            grupo_divergencia = $first.grupo_divergencia
            tipo_xml = $first.tipo_xml
            quantidade = $_.Count
        }
    } |
    Sort-Object grupo_divergencia, tipo_xml)
Write-OptionalCsv -Rows $groupedRows -Path $GroupedReportPath -DelimiterChar $Delimiter

Write-Host '-------------------------------------------------' -ForegroundColor Gray
foreach ($status in @($results | Group-Object status | Sort-Object Name)) {
    Write-Host ("{0}: {1}" -f $status.Name, $status.Count) -ForegroundColor Gray
}
if ($parseErrorCount -gt 0) { Write-Host "Parse errors while scanning candidate XMLs: $parseErrorCount" -ForegroundColor Yellow }
Write-Host "Done: $OutputPath" -ForegroundColor Green
if (-not [string]::IsNullOrWhiteSpace($ProblemReportPath)) { Write-Host "Problems report: $ProblemReportPath" -ForegroundColor Green }
if (-not [string]::IsNullOrWhiteSpace($GroupedReportPath)) { Write-Host "Grouped report: $GroupedReportPath" -ForegroundColor Green }
