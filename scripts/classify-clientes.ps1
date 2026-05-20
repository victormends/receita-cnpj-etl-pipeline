# =============================================================
# classify-clientes.ps1 - Classify existing clients by Receita Simples data
# =============================================================

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [Parameter(Mandatory = $true)]
    [string]$SimplesPath,

    [string]$Delimiter = ';',

    [ValidateSet('UTF8', 'Default', 'Unicode', 'BigEndianUnicode', 'UTF7', 'UTF32', 'ASCII', 'OEM')]
    [string]$Encoding = 'UTF8',

    [string]$CnpjColumn = 'CNPJ',

    [ValidateSet('Database', 'File')]
    [string]$Mode = 'Database'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$projectRoot = Split-Path -Parent $scriptRoot
$preflightPath = Join-Path $scriptRoot 'lib\preflight.ps1'
if (-not (Test-Path -LiteralPath $preflightPath -PathType Leaf)) {
    $preflightPath = Join-Path $projectRoot 'scripts\lib\preflight.ps1'
}
if (Test-Path -LiteralPath $preflightPath -PathType Leaf) {
    . $preflightPath
}

function Resolve-RequiredFile {
    param([string]$Path, [string]$Label)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
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

function Normalize-CnpjBasico {
    param([object]$Value)

    $digits = Normalize-Digits -Value $Value
    if ([string]::IsNullOrWhiteSpace($digits)) { return '' }
    if ($digits.Length -gt 8) { $digits = $digits.Substring(0, 8) }
    return $digits.PadLeft(8, '0')
}

function Get-RegimeRank {
    param([string]$Regime)

    switch ($Regime) {
        'MEI' { return 3 }
        'Simples Nacional' { return 2 }
        default { return 1 }
    }
}

function Get-RegimeFromFlags {
    param([string]$OpcaoSimples, [string]$OpcaoMei)

    $simples = ([string]$OpcaoSimples).Trim().ToUpperInvariant()
    $mei = ([string]$OpcaoMei).Trim().ToUpperInvariant()

    if ($mei -eq 'S') { return 'MEI' }
    if ($simples -eq 'S') { return 'Simples Nacional' }
    return 'Normal'
}

function Add-ResultProperty {
    param(
        [Parameter(Mandatory = $true)][System.Collections.Specialized.OrderedDictionary]$Target,
        [Parameter(Mandatory = $true)][string]$Name,
        [object]$Value
    )

    $finalName = $Name
    if ($Target.Contains($finalName)) {
        $suffix = 2
        do {
            $finalName = "${Name}_classificador_$suffix"
            $suffix++
        } while ($Target.Contains($finalName))
    }

    $Target[$finalName] = $Value
}

function Get-ZipEntryText {
    param(
        [Parameter(Mandatory = $true)][System.IO.Compression.ZipArchive]$Archive,
        [Parameter(Mandatory = $true)][string]$EntryName
    )

    $entry = $Archive.GetEntry($EntryName)
    if ($null -eq $entry) {
        $normalizedName = $EntryName -replace '\\', '/'
        foreach ($candidate in $Archive.Entries) {
            if (($candidate.FullName -replace '\\', '/') -eq $normalizedName) {
                $entry = $candidate
                break
            }
        }
    }
    if ($null -eq $entry) { return $null }

    $stream = $entry.Open()
    try {
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        try { return $reader.ReadToEnd() }
        finally { $reader.Dispose() }
    }
    finally {
        $stream.Dispose()
    }
}

function Get-ExcelColumnIndex {
    param([string]$Reference)

    if ($Reference -notmatch '^([A-Z]+)') { return -1 }

    $letters = $matches[1]
    $index = 0
    foreach ($char in $letters.ToCharArray()) {
        $index = ($index * 26) + ([int][char]$char - [int][char]'A' + 1)
    }

    return ($index - 1)
}

function Get-FirstChildByLocalName {
    param([xml.XmlNode]$Node, [string]$LocalName)

    foreach ($child in $Node.ChildNodes) {
        if ($child.LocalName -eq $LocalName) { return $child }
    }

    return $null
}

function Get-XlsxCellValue {
    param([xml.XmlNode]$Cell, [string[]]$SharedStrings)

    $type = $Cell.GetAttribute('t')
    if ($type -eq 'inlineStr') {
        $inline = Get-FirstChildByLocalName -Node $Cell -LocalName 'is'
        if ($null -eq $inline) { return '' }
        $text = Get-FirstChildByLocalName -Node $inline -LocalName 't'
        if ($null -eq $text) { return '' }
        return $text.InnerText
    }

    $value = Get-FirstChildByLocalName -Node $Cell -LocalName 'v'
    if ($null -eq $value) { return '' }

    if ($type -eq 's') {
        $sharedIndex = 0
        if ([int]::TryParse($value.InnerText, [ref]$sharedIndex) -and $sharedIndex -ge 0 -and $sharedIndex -lt $SharedStrings.Count) {
            return $SharedStrings[$sharedIndex]
        }
    }

    return $value.InnerText
}

function Import-XlsxRows {
    param([string]$Path)

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $fileStream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $archive = New-Object System.IO.Compression.ZipArchive($fileStream, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        $sharedStrings = @()
        $sharedXmlText = Get-ZipEntryText -Archive $archive -EntryName 'xl/sharedStrings.xml'
        if ($sharedXmlText) {
            [xml]$sharedXml = $sharedXmlText
            foreach ($item in $sharedXml.GetElementsByTagName('si')) {
                $parts = New-Object System.Collections.ArrayList
                foreach ($textNode in $item.GetElementsByTagName('t')) {
                    [void]$parts.Add($textNode.InnerText)
                }
                $sharedStrings += ($parts -join '')
            }
        }

        $workbookText = Get-ZipEntryText -Archive $archive -EntryName 'xl/workbook.xml'
        if ([string]::IsNullOrWhiteSpace($workbookText)) { throw 'XLSX workbook.xml was not found.' }
        [xml]$workbook = $workbookText

        $relationshipsText = Get-ZipEntryText -Archive $archive -EntryName 'xl/_rels/workbook.xml.rels'
        if ([string]::IsNullOrWhiteSpace($relationshipsText)) { throw 'XLSX workbook relationships were not found.' }
        [xml]$relationships = $relationshipsText
        $firstSheet = $workbook.GetElementsByTagName('sheet') | Select-Object -First 1
        if ($null -eq $firstSheet) { throw 'XLSX workbook has no sheets.' }

        $relationshipId = $firstSheet.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        if ([string]::IsNullOrWhiteSpace($relationshipId)) { $relationshipId = $firstSheet.GetAttribute('r:id') }

        $target = $null
        foreach ($rel in $relationships.GetElementsByTagName('Relationship')) {
            if ($rel.GetAttribute('Id') -eq $relationshipId) {
                $target = $rel.GetAttribute('Target')
                break
            }
        }

        if ([string]::IsNullOrWhiteSpace($target)) { throw 'Could not resolve first XLSX worksheet.' }
        $sheetEntry = if ($target -like '/*') { $target.TrimStart('/') } else { 'xl/' + $target.TrimStart('./') }
        $sheetText = Get-ZipEntryText -Archive $archive -EntryName $sheetEntry
        if ([string]::IsNullOrWhiteSpace($sheetText)) { throw "XLSX worksheet was not found: $sheetEntry" }
        [xml]$sheetXml = $sheetText

        $headers = $null
        $rows = New-Object System.Collections.ArrayList
        foreach ($rowNode in $sheetXml.GetElementsByTagName('row')) {
            $cellsByIndex = @{}
            $maxIndex = -1
            foreach ($cell in $rowNode.GetElementsByTagName('c')) {
                $cellRef = $cell.GetAttribute('r')
                $index = Get-ExcelColumnIndex -Reference $cellRef
                if ($index -lt 0) { continue }
                $cellsByIndex[$index] = Get-XlsxCellValue -Cell $cell -SharedStrings $sharedStrings
                if ($index -gt $maxIndex) { $maxIndex = $index }
            }

            if ($maxIndex -lt 0) { continue }
            $values = for ($i = 0; $i -le $maxIndex; $i++) {
                if ($cellsByIndex.ContainsKey($i)) { $cellsByIndex[$i] } else { '' }
            }

            if ($null -eq $headers) {
                $headers = @($values | ForEach-Object { ([string]$_).Trim() })
                continue
            }

            $record = [ordered]@{}
            for ($i = 0; $i -lt $headers.Count; $i++) {
                $name = $headers[$i]
                if ([string]::IsNullOrWhiteSpace($name)) { $name = "Column$($i + 1)" }
                $record[$name] = if ($i -lt $values.Count) { $values[$i] } else { '' }
            }
            [void]$rows.Add([pscustomobject]$record)
        }

        return @($rows)
    }
    finally {
        $archive.Dispose()
        $fileStream.Dispose()
    }
}

function Import-ClientRows {
    param([string]$Path, [string]$Delimiter, [string]$Encoding)

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    switch ($extension) {
        '.csv' { return @(Import-Csv -LiteralPath $Path -Delimiter $Delimiter -Encoding $Encoding) }
        '.xlsx' { return @(Import-XlsxRows -Path $Path) }
        default { throw "Unsupported input format '$extension'. Use .csv or .xlsx." }
    }
}

function New-CsvValue {
    param([object]$Value, [char]$DelimiterChar)

    $text = if ($null -eq $Value) { '' } else { [string]$Value }
    if ($text.IndexOfAny(@([char]'"', [char]"`r", [char]"`n", $DelimiterChar)) -ge 0) {
        return '"' + ($text -replace '"', '""') + '"'
    }

    return $text
}

function Export-PreparedClientesForPostgres {
    param(
        [Parameter(Mandatory = $true)]$PreparedClientes,
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][char]$DelimiterChar
    )

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    $writer = New-Object System.IO.StreamWriter($Path, $false, $utf8NoBom)
    try {
        foreach ($prepared in $PreparedClientes) {
            if ([string]::IsNullOrWhiteSpace($prepared.CnpjBasico)) { continue }
            $line = @(
                (New-CsvValue -Value $prepared.RowId -DelimiterChar $DelimiterChar),
                (New-CsvValue -Value $prepared.CnpjBasico -DelimiterChar $DelimiterChar)
            ) -join $DelimiterChar
            $writer.WriteLine($line)
        }
    }
    finally {
        $writer.Dispose()
    }
}

function Test-PathInsideDirectoryLocal {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Directory
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $fullDirectory = [System.IO.Path]::GetFullPath($Directory).TrimEnd('\') + '\'
    return $fullPath.StartsWith($fullDirectory, [System.StringComparison]::OrdinalIgnoreCase)
}

function Copy-FileForServerCopy {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$TempDirectory,
        [Parameter(Mandatory = $true)][string]$TempId
    )

    if (Test-PathInsideDirectoryLocal -Path $SourcePath -Directory $TempDirectory) {
        return [pscustomobject]@{ Path = $SourcePath; IsTemporary = $false }
    }

    $extension = [System.IO.Path]::GetExtension($SourcePath)
    if ([string]::IsNullOrWhiteSpace($extension)) { $extension = '.csv' }
    $destination = Join-Path $TempDirectory "simples_classificador_$TempId$extension"
    Write-Host "Copying Simples CSV to PostgreSQL temp folder for server-side COPY..." -ForegroundColor Gray
    Copy-Item -LiteralPath $SourcePath -Destination $destination -Force
    return [pscustomobject]@{ Path = $destination; IsTemporary = $true }
}

function New-ClassifierRuntimeConfig {
    if (-not (Get-Command Import-CnpjConfig -ErrorAction SilentlyContinue)) {
        throw 'PostgreSQL mode requires scripts\lib\preflight.ps1.'
    }

    $candidateConfigPaths = @(
        (Join-Path $scriptRoot 'config.ps1'),
        (Join-Path $projectRoot 'config.ps1')
    ) | Select-Object -Unique

    foreach ($localConfigPath in $candidateConfigPaths) {
        if (-not (Test-Path -LiteralPath $localConfigPath -PathType Leaf)) { continue }

        $config = & $localConfigPath
        if (-not ($config -is [hashtable])) {
            throw "config.ps1 did not return a valid hashtable: $localConfigPath"
        }
        $config = Resolve-DatabaseConfig -Config $config
        return (Resolve-PortableWorkingDirectories -Config $config)
    }

    $bootstrap = Import-CnpjConfig -ScriptRoot $scriptRoot
    return $bootstrap.Config
}

function Invoke-ClassifierSqlFile {
    param(
        [Parameter(Mandatory = $true)][string]$PsqlPath,
        [Parameter(Mandatory = $true)][hashtable]$Config,
        [Parameter(Mandatory = $true)][string]$Sql,
        [Parameter(Mandatory = $true)][string]$Description,
        [Parameter(Mandatory = $true)][int]$TimeoutSec
    )

    $sqlFile = Join-Path ([System.IO.Path]::GetTempPath()) ("cnpj-classifier-{0}.sql" -f ([guid]::NewGuid().ToString('N')))
    [System.IO.File]::WriteAllText($sqlFile, $Sql, [System.Text.Encoding]::UTF8)
    try {
        Write-Host "PostgreSQL: $Description" -ForegroundColor Gray
        Invoke-PsqlFileChecked -PsqlPath $PsqlPath -Config $Config -FilePath $sqlFile -Description $Description -TimeoutSec $TimeoutSec
    }
    finally {
        Remove-Item -LiteralPath $sqlFile -Force -ErrorAction SilentlyContinue
    }
}

function Import-SimplesMatchesWithPostgres {
    param(
        [Parameter(Mandatory = $true)]$PreparedClientes,
        [Parameter(Mandatory = $true)][string]$SimplesPath,
        [Parameter(Mandatory = $true)][char]$DelimiterChar
    )

    if (-not (Get-Command Invoke-CnpjPreflight -ErrorAction SilentlyContinue)) {
        throw 'PostgreSQL mode requires the preflight helper functions.'
    }

    $config = New-ClassifierRuntimeConfig
    $timeout = if ($config.ContainsKey('sqlCommandTimeoutSeconds')) { [int]$config.sqlCommandTimeoutSeconds } else { 14400 }
    $tools = Invoke-CnpjPreflight -Config $config -EnsureDirectories @($config.dirTemp, $config.dirOut) -MinFreeSpaceGB @{ $config.dirTemp = 6; $config.dirOut = 1 } -RequirePostgres -RequireDatabase -CheckMemory
    Set-PostgresPasswordEnv -Config $config

    $tempId = [guid]::NewGuid().ToString('N')
    $targetsPath = Join-Path $config.dirTemp "clientes_alvo_$tempId.csv"
    $matchesPath = Join-Path $config.dirTemp "clientes_matches_$tempId.csv"
    Export-PreparedClientesForPostgres -PreparedClientes $PreparedClientes -Path $targetsPath -DelimiterChar $DelimiterChar
    $serverSimples = Copy-FileForServerCopy -SourcePath $SimplesPath -TempDirectory $config.dirTemp -TempId $tempId

    $targetCopyPath = ConvertTo-PsqlPathLiteral -Path $targetsPath
    $simplesCopyPath = ConvertTo-PsqlPathLiteral -Path $serverSimples.Path
    $matchesCopyPath = ConvertTo-PsqlPathLiteral -Path $matchesPath
    $delimiterSql = $DelimiterChar.ToString().Replace("'", "''")

    $sql = @"
DROP TABLE IF EXISTS tmp_clientes_classificador_resultado;
DROP TABLE IF EXISTS tmp_clientes_classificador_alvo;
DROP TABLE IF EXISTS tmp_simples_classificador;

CREATE UNLOGGED TABLE tmp_clientes_classificador_alvo (
    row_id integer NOT NULL,
    cnpj_basico varchar(8) NOT NULL
);

COPY tmp_clientes_classificador_alvo (row_id, cnpj_basico)
FROM '$targetCopyPath'
WITH (DELIMITER '$delimiterSql', FORMAT CSV, HEADER FALSE, ENCODING 'UTF8', NULL '');

CREATE INDEX tmp_clientes_classificador_alvo_idx ON tmp_clientes_classificador_alvo (cnpj_basico);

CREATE UNLOGGED TABLE tmp_simples_classificador (
    cnpj_basico text,
    opcao_pelo_simples text,
    data_opcao_simples text,
    data_exclusao_simples text,
    opcao_pelo_mei text,
    data_opcao_mei text,
    data_exclusao_mei text
);

COPY tmp_simples_classificador
FROM '$simplesCopyPath'
WITH (DELIMITER '$delimiterSql', FORMAT CSV, HEADER FALSE, ENCODING 'LATIN1', NULL '');

CREATE INDEX tmp_simples_classificador_basico_idx
ON tmp_simples_classificador ((LPAD(regexp_replace(cnpj_basico, '\D', '', 'g'), 8, '0')));

CREATE UNLOGGED TABLE tmp_clientes_classificador_resultado AS
WITH simples_normalizado AS (
    SELECT
        LPAD(regexp_replace(cnpj_basico, '\D', '', 'g'), 8, '0') AS cnpj_basico,
        TRIM(BOTH '"' FROM TRIM(COALESCE(opcao_pelo_simples, ''))) AS opcao_pelo_simples,
        TRIM(BOTH '"' FROM TRIM(COALESCE(opcao_pelo_mei, ''))) AS opcao_pelo_mei
    FROM tmp_simples_classificador
), classificados AS (
    SELECT
        a.row_id,
        CASE
            WHEN UPPER(s.opcao_pelo_mei) = 'S' THEN 'MEI'
            WHEN UPPER(s.opcao_pelo_simples) = 'S' THEN 'Simples Nacional'
            ELSE 'Normal'
        END AS regime,
        s.opcao_pelo_simples,
        s.opcao_pelo_mei,
        CASE
            WHEN UPPER(s.opcao_pelo_mei) = 'S' THEN 3
            WHEN UPPER(s.opcao_pelo_simples) = 'S' THEN 2
            ELSE 1
        END AS regime_rank
    FROM tmp_clientes_classificador_alvo a
    JOIN simples_normalizado s ON s.cnpj_basico = a.cnpj_basico
), ranked AS (
    SELECT *, ROW_NUMBER() OVER (PARTITION BY row_id ORDER BY regime_rank DESC) AS rn
    FROM classificados
)
SELECT row_id, regime, opcao_pelo_simples, opcao_pelo_mei
FROM ranked
WHERE rn = 1;

COPY (
    SELECT row_id, regime, opcao_pelo_simples, opcao_pelo_mei
    FROM tmp_clientes_classificador_resultado
    ORDER BY row_id
) TO '$matchesCopyPath'
WITH (DELIMITER '$delimiterSql', FORMAT CSV, HEADER TRUE, ENCODING 'UTF8');
"@

    try {
        Invoke-ClassifierSqlFile -PsqlPath $tools.PsqlPath -Config $config -Sql $sql -Description 'load Simples file and classify client CNPJs' -TimeoutSec $timeout

        $matches = @{}
        if (Test-Path -LiteralPath $matchesPath -PathType Leaf) {
            foreach ($row in @(Import-Csv -LiteralPath $matchesPath -Delimiter $DelimiterChar -Encoding UTF8)) {
                $rowId = [int]$row.row_id
                $matches[$rowId] = [pscustomobject]@{
                    Regime = $row.regime
                    OpcaoPeloSimples = $row.opcao_pelo_simples
                    OpcaoPeloMei = $row.opcao_pelo_mei
                }
            }
        }

        return $matches
    }
    finally {
        Remove-Item -LiteralPath $targetsPath -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $matchesPath -Force -ErrorAction SilentlyContinue
        if ($serverSimples -and $serverSimples.IsTemporary) {
            Remove-Item -LiteralPath $serverSimples.Path -Force -ErrorAction SilentlyContinue
        }
    }
}

function Import-SimplesMatchesFromFile {
    param(
        [Parameter(Mandatory = $true)][hashtable]$TargetBasicos,
        [Parameter(Mandatory = $true)][string]$SimplesPath,
        [Parameter(Mandatory = $true)][char]$DelimiterChar
    )

    $simplesMatches = @{}
    $reader = [System.IO.File]::OpenText($SimplesPath)
    $lineNumber = 0
    try {
        while ($null -ne ($line = $reader.ReadLine())) {
            $lineNumber++
            if ([string]::IsNullOrWhiteSpace($line)) { continue }

            $parts = $line.Split($DelimiterChar)
            if ($parts.Count -lt 5) { continue }

            $cnpjBasico = Normalize-CnpjBasico -Value $parts[0]
            if ($lineNumber -eq 1 -and $cnpjBasico -eq '00000000' -and $parts[0] -match 'cnpj') { continue }
            if (-not $TargetBasicos.ContainsKey($cnpjBasico)) { continue }

            $regime = Get-RegimeFromFlags -OpcaoSimples $parts[1] -OpcaoMei $parts[4]
            if (-not $simplesMatches.ContainsKey($cnpjBasico) -or (Get-RegimeRank $regime) -gt (Get-RegimeRank $simplesMatches[$cnpjBasico].Regime)) {
                $simplesMatches[$cnpjBasico] = [pscustomobject]@{
                    Regime = $regime
                    OpcaoPeloSimples = ([string]$parts[1]).Trim()
                    OpcaoPeloMei = ([string]$parts[4]).Trim()
                }
            }
        }
    }
    finally {
        $reader.Close()
    }

    return $simplesMatches
}

$InputPath = Resolve-RequiredFile -Path $InputPath -Label 'Input file'
$SimplesPath = Resolve-RequiredFile -Path $SimplesPath -Label 'Simples CSV'

if ($Delimiter.Length -ne 1) {
    throw 'Delimiter must be a single character.'
}

$outputParent = Split-Path -Parent $OutputPath
if (-not [string]::IsNullOrWhiteSpace($outputParent) -and -not (Test-Path -LiteralPath $outputParent)) {
    New-Item -ItemType Directory -Path $outputParent -Force | Out-Null
}

Write-Host '===================================================' -ForegroundColor Cyan
Write-Host ' CLIENT CLASSIFIER - Receita Simples' -ForegroundColor Cyan
Write-Host '===================================================' -ForegroundColor Cyan
Write-Host " Input:   $InputPath" -ForegroundColor Yellow
Write-Host " Simples: $SimplesPath" -ForegroundColor Yellow
Write-Host " Output:  $OutputPath" -ForegroundColor Yellow
Write-Host '-------------------------------------------------' -ForegroundColor Gray

$clientes = @(Import-ClientRows -Path $InputPath -Delimiter $Delimiter -Encoding $Encoding)
if ($clientes.Count -eq 0) {
    throw 'Input file has no data rows.'
}

$columns = @($clientes[0].PSObject.Properties.Name)
if ($columns -notcontains $CnpjColumn) {
    throw "Input file must contain a '$CnpjColumn' column. Found: $($columns -join ', ')"
}

$targetBasicos = @{}
$preparedClientes = New-Object System.Collections.ArrayList
$rowId = 0
foreach ($cliente in $clientes) {
    $rowId++
    $cnpjNormalizado = Normalize-Cnpj -Value $cliente.$CnpjColumn
    $cnpjBasico = ''
    $status = ''

    if ([string]::IsNullOrWhiteSpace($cnpjNormalizado)) {
        $status = 'Sem CNPJ'
    }
    elseif ($cnpjNormalizado.Length -ne 14) {
        $status = 'CNPJ invalido'
    }
    else {
        $cnpjBasico = $cnpjNormalizado.Substring(0, 8)
        $targetBasicos[$cnpjBasico] = $true
    }

    [void]$preparedClientes.Add([pscustomobject]@{
        RowId = $rowId
        Row = $cliente
        CnpjNormalizado = $cnpjNormalizado
        CnpjBasico = $cnpjBasico
        Status = $status
    })
}

Write-Host "Client rows: $($preparedClientes.Count)" -ForegroundColor Gray
Write-Host "Target CNPJ basicos: $($targetBasicos.Count)" -ForegroundColor Gray

$simplesMatches = @{}
if ($targetBasicos.Count -gt 0) {
    if ($Mode -eq 'Database') {
        $simplesMatches = Import-SimplesMatchesWithPostgres -PreparedClientes $preparedClientes -SimplesPath $SimplesPath -DelimiterChar $Delimiter[0]
    }
    else {
        $matchesByBasico = Import-SimplesMatchesFromFile -TargetBasicos $targetBasicos -SimplesPath $SimplesPath -DelimiterChar $Delimiter[0]
        foreach ($prepared in $preparedClientes) {
            if ($matchesByBasico.ContainsKey($prepared.CnpjBasico)) {
                $simplesMatches[$prepared.RowId] = $matchesByBasico[$prepared.CnpjBasico]
            }
        }
    }
}

Write-Host "Matched CNPJ basicos in Simples: $($simplesMatches.Count)" -ForegroundColor Gray

$results = New-Object System.Collections.ArrayList
$counts = @{}

foreach ($prepared in $preparedClientes) {
    $regime = 'Normal'
    $fonte = 'Receita Simples'
    $observacao = 'Sem opcao por MEI ou Simples Nacional no arquivo informado.'

    if ($prepared.Status) {
        $regime = $prepared.Status
        $fonte = 'Entrada'
        $observacao = 'CNPJ ausente ou fora do formato esperado.'
    }
    elseif ($simplesMatches.ContainsKey($prepared.RowId)) {
        $match = $simplesMatches[$prepared.RowId]
        $regime = $match.Regime
        $observacao = "opcao_pelo_simples=$($match.OpcaoPeloSimples); opcao_pelo_mei=$($match.OpcaoPeloMei)"
    }

    if (-not $counts.ContainsKey($regime)) { $counts[$regime] = 0 }
    $counts[$regime]++

    $row = [ordered]@{}
    foreach ($property in $prepared.Row.PSObject.Properties) {
        $row[$property.Name] = $property.Value
    }

    Add-ResultProperty -Target $row -Name 'cnpj_normalizado' -Value $prepared.CnpjNormalizado
    Add-ResultProperty -Target $row -Name 'cnpj_basico' -Value $prepared.CnpjBasico
    Add-ResultProperty -Target $row -Name 'regime_tributario' -Value $regime
    Add-ResultProperty -Target $row -Name 'classificacao_fonte' -Value $fonte
    Add-ResultProperty -Target $row -Name 'classificacao_observacao' -Value $observacao

    [void]$results.Add([pscustomobject]$row)
}

$results | Export-Csv -LiteralPath $OutputPath -Delimiter $Delimiter -Encoding UTF8 -NoTypeInformation

Write-Host '-------------------------------------------------' -ForegroundColor Gray
foreach ($key in ($counts.Keys | Sort-Object)) {
    Write-Host ("{0}: {1}" -f $key, $counts[$key]) -ForegroundColor Gray
}
Write-Host "Done: $OutputPath" -ForegroundColor Green
