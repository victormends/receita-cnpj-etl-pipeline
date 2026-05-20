Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$ProjectRoot = Split-Path -Parent $PSScriptRoot
$PreflightPath = Join-Path $ProjectRoot 'scripts\lib\preflight.ps1'
. $PreflightPath

$script:Passed = 0
$script:Failed = 0

function Write-TestPass {
    param([string]$Name)
    $script:Passed++
    Write-Host "[PASS] $Name" -ForegroundColor Green
}

function Write-TestFail {
    param([string]$Name, [string]$Message)
    $script:Failed++
    Write-Host "[FAIL] $Name" -ForegroundColor Red
    Write-Host "       $Message" -ForegroundColor Red
}

function Invoke-Test {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][scriptblock]$Body
    )

    try {
        & $Body
        Write-TestPass -Name $Name
    }
    catch {
        Write-TestFail -Name $Name -Message $_.Exception.Message
    }
}

function Assert-True {
    param([bool]$Condition, [string]$Message = 'Expected condition to be true.')
    if (-not $Condition) { throw $Message }
}

function Assert-False {
    param([bool]$Condition, [string]$Message = 'Expected condition to be false.')
    if ($Condition) { throw $Message }
}

function Assert-Equal {
    param($Expected, $Actual, [string]$Message = 'Values are not equal.')
    if ($Expected -ne $Actual) {
        throw "$Message Expected: <$Expected> Actual: <$Actual>"
    }
}

function Assert-MatchText {
    param([string]$Text, [string]$Pattern, [string]$Message)
    if ($Text -notmatch $Pattern) { throw $Message }
}

function Assert-Throws {
    param([scriptblock]$Body, [string]$Pattern = $null)

    $threw = $false
    try {
        & $Body
    }
    catch {
        $threw = $true
        if ($Pattern -and $_.Exception.Message -notmatch $Pattern) {
            throw "Exception did not match '$Pattern'. Actual: $($_.Exception.Message)"
        }
    }

    if (-not $threw) { throw 'Expected command to throw.' }
}

function New-TestRoot {
    $root = Join-Path ([System.IO.Path]::GetTempPath()) ("cnpj-etl-tests-{0}" -f ([guid]::NewGuid().ToString('N')))
    New-Item -ItemType Directory -Path $root -Force | Out-Null
    return $root
}

function New-ZipFile {
    param(
        [Parameter(Mandatory = $true)][string]$ZipPath,
        [Parameter(Mandatory = $true)][hashtable]$Entries
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $staging = Join-Path ([System.IO.Path]::GetTempPath()) ("cnpj-etl-zip-{0}" -f ([guid]::NewGuid().ToString('N')))
    New-Item -ItemType Directory -Path $staging -Force | Out-Null
    try {
        foreach ($entry in $Entries.GetEnumerator()) {
            $entryPath = Join-Path $staging $entry.Key
            $parent = Split-Path -Parent $entryPath
            if (-not (Test-Path -LiteralPath $parent)) {
                New-Item -ItemType Directory -Path $parent -Force | Out-Null
            }
            Set-Content -Path $entryPath -Value $entry.Value -NoNewline -Encoding UTF8
        }
        [System.IO.Compression.ZipFile]::CreateFromDirectory($staging, $ZipPath)
    }
    finally {
        Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function New-TestXlsxFile {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string[][]]$Rows
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $staging = Join-Path ([System.IO.Path]::GetTempPath()) ("cnpj-etl-xlsx-{0}" -f ([guid]::NewGuid().ToString('N')))
    New-Item -ItemType Directory -Path (Join-Path $staging '_rels') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $staging 'xl\_rels') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $staging 'xl\worksheets') -Force | Out-Null

    try {
        [System.IO.File]::WriteAllText((Join-Path $staging '[Content_Types].xml'), '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>', [System.Text.Encoding]::UTF8)
        [System.IO.File]::WriteAllText((Join-Path $staging '_rels\.rels'), '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>', [System.Text.Encoding]::UTF8)
        [System.IO.File]::WriteAllText((Join-Path $staging 'xl\workbook.xml'), '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>', [System.Text.Encoding]::UTF8)
        [System.IO.File]::WriteAllText((Join-Path $staging 'xl\_rels\workbook.xml.rels'), '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>', [System.Text.Encoding]::UTF8)

        $rowXml = New-Object System.Collections.ArrayList
        for ($r = 0; $r -lt $Rows.Count; $r++) {
            $cells = New-Object System.Collections.ArrayList
            for ($c = 0; $c -lt $Rows[$r].Count; $c++) {
                $col = [char]([int][char]'A' + $c)
                $ref = "{0}{1}" -f $col, ($r + 1)
                $value = [System.Security.SecurityElement]::Escape($Rows[$r][$c])
                [void]$cells.Add("<c r=`"$ref`" t=`"inlineStr`"><is><t>$value</t></is></c>")
            }
            [void]$rowXml.Add("<row r=`"$($r + 1)`">$($cells -join '')</row>")
        }

        $sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' + ($rowXml -join '') + '</sheetData></worksheet>'
        [System.IO.File]::WriteAllText((Join-Path $staging 'xl\worksheets\sheet1.xml'), $sheet, [System.Text.Encoding]::UTF8)

        if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
        [System.IO.Compression.ZipFile]::CreateFromDirectory($staging, $Path)
    }
    finally {
        Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Get-ScriptText {
    param([string]$RelativePath)
    return [System.IO.File]::ReadAllText((Join-Path $ProjectRoot $RelativePath))
}

$tempRoots = @()
try {
    Invoke-Test 'preflight accepts valid cleanup modes and rejects invalid mode' {
        $config = @{
            anoMes = '2026-04-12'; baseUrl = 'https://example.test'; dirDownload = 'C:\tmp\download'; dirTemp = 'C:\tmp\temp'; dirOut = 'C:\tmp\out'
            dataCorte = ''; ufs = @('PR'); cnaeWhitelist = @('62'); cnaeRanges = @(@('01', '03')); dbHost = 'localhost'; dbName = 'cnpj'
            cleanupMode = 'balanced'
        }
        Assert-ConfigValid -Config $config
        $config.cleanupMode = 'unsafe'
        Assert-Throws { Assert-ConfigValid -Config $config } 'Invalid cleanupMode'
    }

    Invoke-Test 'preflight safe path helper blocks directory escape' {
        $root = New-TestRoot; $tempRoots += $root
        $safePath = Join-Path $root 'file.txt'
        Assert-Equal ([System.IO.Path]::GetFullPath($safePath)) (Assert-PathInsideDirectory -Path $safePath -AllowedRoot $root)
        Assert-Throws { Assert-PathInsideDirectory -Path (Join-Path $root '..\outside.txt') -AllowedRoot $root } 'outside allowed root'
    }

    Invoke-Test 'preflight validates ZIP archives and rejects invalid files' {
        $root = New-TestRoot; $tempRoots += $root
        $zipPath = Join-Path $root 'valid.zip'
        $badZipPath = Join-Path $root 'bad.zip'
        New-ZipFile -ZipPath $zipPath -Entries @{ 'K3241.TEST.ESTABELE' = '1;2;3' }
        Set-Content -Path $badZipPath -Value 'not a zip' -NoNewline
        Assert-True (Test-ZipValid -Path $zipPath) 'Expected valid test ZIP to pass.'
        Assert-False (Test-ZipValid -Path $badZipPath) 'Expected invalid ZIP to fail.'
    }

    Invoke-Test 'preflight safe delete honors root, allowlist, and dry run' {
        $root = New-TestRoot; $tempRoots += $root
        $keepPath = Join-Path $root 'file.zip'
        Set-Content -Path $keepPath -Value 'x'
        $dryConfig = @{ cleanupDryRun = $true; cleanupMode = 'aggressive' }
        $dryResult = Remove-CnpjArtifactSafe -Config $dryConfig -Path $keepPath -AllowedRoot $root -AllowedPatterns @('*.zip') -Reason 'test dry run'
        Assert-True $dryResult.DryRun 'Expected dry-run result.'
        Assert-True (Test-Path -LiteralPath $keepPath) 'Dry run should not delete file.'

        $deleteResult = Remove-CnpjArtifactSafe -Config @{ cleanupDryRun = $false; cleanupMode = 'aggressive' } -Path $keepPath -AllowedRoot $root -AllowedPatterns @('*.zip') -Reason 'test delete'
        Assert-True $deleteResult.Deleted 'Expected actual deletion.'
        Assert-False (Test-Path -LiteralPath $keepPath) 'Expected file to be deleted.'

        $blockedPath = Join-Path $root 'file.txt'
        Set-Content -Path $blockedPath -Value 'x'
        Assert-Throws { Remove-CnpjArtifactSafe -Config @{} -Path $blockedPath -AllowedRoot $root -AllowedPatterns @('*.zip') -Reason 'blocked' } 'disallowed name'
    }

    Invoke-Test 'preflight escapes PostgreSQL path literals' {
        Assert-Equal "C:/temp/O''Hara/file.csv" (ConvertTo-PsqlPathLiteral -Path "C:\temp\O'Hara\file.csv")
    }

    Invoke-Test 'download stage validates existing and downloaded ZIP files' {
        $text = Get-ScriptText 'scripts\download.ps1'
        Assert-MatchText $text 'Test-ZipValid -Path \$destino' 'download.ps1 must validate existing ZIP files.'
        Assert-MatchText $text 'Test-ZipValid -Path \$destinoTmp' 'download.ps1 must validate downloaded partial ZIP files before moving them.'
        Assert-MatchText $text "Remove-CnpjArtifactSafe[\s\S]*'\*\.partial'" 'download.ps1 must safe-delete stale/failed partial downloads.'
        Assert-MatchText $text 'skip valid ZIPs and retry failed files' 'download.ps1 retry guidance should not claim byte-range resume.'
    }

    Invoke-Test 'extract stage blocks zip-slip and validates exact extracted files' {
        $text = Get-ScriptText 'scripts\extract.ps1'
        Assert-MatchText $text 'function Invoke-ZipExtraction' 'extract.ps1 must expose Invoke-ZipExtraction.'
        Assert-MatchText $text 'Assert-PathInsideDirectory' 'extract.ps1 must guard extraction destinations.'
        Assert-MatchText $text 'Duplicate extraction destination' 'extract.ps1 must detect duplicate destination collisions.'
        Assert-MatchText $text 'Test-FileNonEmpty -Path \$path' 'extract.ps1 must validate extracted files before cleanup.'
        Assert-MatchText $text 'function Close-ZipArchive' 'extract.ps1 must explicitly close ZIP handles before cleanup.'
        Assert-MatchText $text 'function Remove-ZipAfterClose' 'extract.ps1 must clean ZIPs only after closing the archive.'
        Assert-MatchText $text 'foreach \(\$attempt in 1\.\.5\)' 'extract.ps1 must retry ZIP cleanup after transient file locks.'
        Assert-MatchText $text 'Reason ''extraction validated''' 'extract.ps1 must only remove ZIPs after extraction validation.'
    }

    Invoke-Test 'clean stage validates cleaned outputs before deleting raw files' {
        $text = Get-ScriptText 'scripts\clean.ps1'
        Assert-MatchText $text 'function Test-CleanOutputValid' 'clean.ps1 must validate cleaned output files.'
        Assert-MatchText $text 'Line-count mismatch after cleaning' 'clean.ps1 must compare input/output line counts.'
        Assert-MatchText $text 'cleaningMaxWorkers' 'clean.ps1 must honor bounded worker configuration.'
        Assert-MatchText $text 'Reason ''cleaned file validated''' 'clean.ps1 must delete raw files only after validation.'
        Assert-MatchText $text 'Reason ''failed cleaning output''' 'clean.ps1 must remove failed partial LIMPO outputs.'
    }

    Invoke-Test 'import stage verifies database result before deleting LIMPO files' {
        $text = Get-ScriptText 'scripts\import.ps1'
        Assert-MatchText $text 'sqlCommandTimeoutSeconds' 'import.ps1 must use configurable SQL timeout.'
        Assert-MatchText $text 'ConvertTo-PsqlPathLiteral' 'import.ps1 must escape file paths for psql.'
        Assert-MatchText $text 'function New-CopyFromFileSql' 'import.ps1 must use server-side COPY for large local source files.'
        Assert-MatchText $text 'function Run-SqlQuery' 'import.ps1 must provide query verification helper.'
        Assert-MatchText $text 'tmp_estabelecimentos_stage' 'import.ps1 must use reusable establishment staging instead of one huge temp table.'
        Assert-MatchText $text 'TRUNCATE tmp_estabelecimentos_stage' 'import.ps1 must process establishment files one at a time.'
        Assert-MatchText $text 'tmp_cnpj_basico_alvo' 'import.ps1 must build a target CNPJ list before enrichment.'
        Assert-MatchText $text 'JOIN tmp_cnpj_basico_alvo' 'import.ps1 must restrict Empresas/Simples enrichment to target CNPJs.'
        Assert-MatchText $text "LPAD\(regexp_replace\(t\.cnpj_basico, '\\D', '', 'g'\), 8, '0'\)" 'import.ps1 must normalize Empresas cnpj_basico before joining.'
        Assert-MatchText $text "LPAD\(regexp_replace\(s\.cnpj_basico, '\\D', '', 'g'\), 8, '0'\)" 'import.ps1 must normalize Simples cnpj_basico before joining.'
        Assert-MatchText $text 'municipiosWhitelist' 'import.ps1 must support configured municipality filtering.'
        Assert-MatchText $text 'resetFinalTablesOnImport' 'import.ps1 must support replacing stale final rows instead of accumulating runs.'
        Assert-MatchText $text 'TRUNCATE TABLE estabelecimentos_crm' 'import.ps1 must reset CRM rows by default before a full import.'
        Assert-MatchText $text 'requireEnrichmentMatches' 'import.ps1 must fail suspicious zero-match enrichment before deleting LIMPO files.'
        Assert-MatchText $text 'filterSimplesNacional' 'import.ps1 must support optional Simples-only filtering.'
        Assert-MatchText $text "AddMonths\(-6\)\.ToString\('yyyyMMdd'\)" 'import.ps1 must calculate the default cutoff from the execution date.'
        Assert-MatchText $text "COALESCE\(regime_tributario, ''\) <> 'Simples Nacional'" 'import.ps1 must remove Normal, MEI, and unmatched rows when Simples filter is enabled.'
        Assert-MatchText $text 'no Simples file was found' 'import.ps1 must fail if Simples-only filtering is enabled without Simples data.'
        Assert-MatchText $text "IMPORT_FULL" 'import.ps1 must verify import log entry.'
        Assert-MatchText $text 'allowZeroFinalRows' 'import.ps1 must protect against unexpected zero-row imports.'
        Assert-MatchText $text 'Reason ''import transaction verified''' 'import.ps1 must delete LIMPO files only after verification.'
    }

    Invoke-Test 'export stage escapes output path and validates CSV header' {
        $text = Get-ScriptText 'scripts\export.ps1'
        Assert-MatchText $text 'ConvertTo-PsqlPathLiteral -Path \$outputFile' 'export.ps1 must escape output path for psql.'
        Assert-MatchText $text '\\copy \(\$exportSelect\)' 'export.ps1 must keep psql copy meta-command on a single line.'
        Assert-MatchText $text 'replace previous export file' 'export.ps1 must replace stale CSV files before exporting.'
        Assert-MatchText $text 'TimeoutSec \$exportTimeout' 'export.ps1 must use long SQL timeout for exports.'
        Assert-MatchText $text 'Test-TextFileReadable -Path \$outputFile' 'export.ps1 must verify exported CSV is readable.'
        Assert-MatchText $text '\$header -notlike ''cnpj;\*''' 'export.ps1 must validate CSV header.'
        Assert-MatchText $text "Remove-CnpjArtifactSafe[\s\S]*'export_query\.sql'" 'export.ps1 must safe-delete temporary SQL file.'
    }

    Invoke-Test 'psql file helper fails on nonzero psql exit code' {
        $text = Get-ScriptText 'scripts\lib\preflight.ps1'
        Assert-MatchText $text 'ExitCode = \$LASTEXITCODE' 'Invoke-PsqlFileChecked must capture psql exit code.'
        Assert-MatchText $text 'psql exited with code' 'Invoke-PsqlFileChecked must throw when psql fails.'
        Assert-MatchText $text 'Output:' 'Invoke-PsqlFileChecked must include psql output in failures.'
        Assert-MatchText $text '\$output = \(\[string\]\$result\.Output\)\.Trim\(\)' 'Test-DatabaseConnection must validate normalized psql output.'
        Assert-MatchText $text 'PostgreSQL validation query failed with exit code' 'Test-DatabaseConnection must fail on nonzero psql exit code.'
    }

    Invoke-Test 'orchestrator skips destructive final cleanup and skipped-import verification' {
        $text = Get-ScriptText 'run-pipeline.ps1'
        Assert-MatchText $text 'function Test-DownloadComplete' 'run-pipeline.ps1 must check download completeness.'
        Assert-MatchText $text 'Test-ZipValid -Path \$path' 'run-pipeline.ps1 must validate ZIP files for download completion.'
        Assert-MatchText $text 'Downstream artifacts already exist - skipping download' 'run-pipeline.ps1 must skip download when aggressive cleanup removed ZIPs.'
        Assert-MatchText $text 'rawEstabelecimentos' 'run-pipeline.ps1 must not trust extraction markers when raw establishment files are gone.'
        Assert-MatchText $text 'rawEmpresas' 'run-pipeline.ps1 must not trust extraction markers when raw company files are gone.'
        Assert-MatchText $text 'Get-ChildItem -Path \$Config\.dirTemp -Filter ''LIMPO_\*\.ESTABELE''' 'run-pipeline.ps1 must require cleaned establishment files before declaring cleaning complete.'
        Assert-MatchText $text 'Get-ChildItem -Path \$Config\.dirTemp -Filter ''LIMPO_\*\.EMPRECSV''' 'run-pipeline.ps1 must require cleaned company files before declaring cleaning complete.'
        Assert-MatchText $text 'function Invoke-PipelineScript' 'run-pipeline.ps1 must wrap child stage execution.'
        Assert-MatchText $text 'failed with exit code \$global:LASTEXITCODE' 'run-pipeline.ps1 must stop when a child stage exits nonzero.'
        Assert-MatchText $text 'Downstream extracted/cleaned artifacts already exist - skipping extraction' 'run-pipeline.ps1 must skip extraction if cleaned artifacts already exist after ZIP cleanup.'
        Assert-MatchText $text 'if \(\$didImport\)' 'run-pipeline.ps1 must verify import only when import ran.'
        Assert-MatchText $text "Filter '\*\.partial'" 'final cleanup must be limited to transient partial files.'
        Assert-MatchText $text "Filter '\*\.sql'" 'final cleanup must be limited to transient SQL files.'
        $cleanupBody = [regex]::Match($text, 'function Invoke-PipelineCleanup[\s\S]*?^}', 'Multiline').Value
        if ($cleanupBody -match "Filter '\*\.zip'" -or $cleanupBody -match "Filter 'LIMPO_\*'" -or $cleanupBody -match "Filter '\*\.ESTABELE'") {
            throw 'run-pipeline.ps1 final cleanup must not broadly delete ZIP, raw, or LIMPO data files.'
        }
    }

    Invoke-Test 'client classifier exposes CSV and Simples parameters' {
        $text = Get-ScriptText 'scripts\classify-clientes.ps1'
        Assert-MatchText $text '\$InputPath' 'classifier must accept an input CSV path.'
        Assert-MatchText $text '\$OutputPath' 'classifier must accept an output CSV path.'
        Assert-MatchText $text '\$SimplesPath' 'classifier must accept a Simples CSV path.'
        Assert-MatchText $text '\$CnpjColumn' 'classifier must allow the CNPJ column name to be configured.'
        Assert-MatchText $text "\[ValidateSet\('Database', 'File'\)\]" 'classifier must expose database and file modes.'
        Assert-MatchText $text 'Invoke-CnpjPreflight' 'classifier database mode must reuse PostgreSQL preflight.'
        Assert-MatchText $text 'Copy-FileForServerCopy' 'classifier database mode must stage local Simples files for server-side COPY.'
        Assert-MatchText $text 'COPY tmp_simples_classificador' 'classifier database mode must use PostgreSQL COPY for large Simples files.'
        Assert-MatchText $text 'opcao_pelo_mei' 'classifier output notes must reference the MEI source flag.'
        Assert-MatchText $text 'Simples Nacional' 'classifier must classify Simples Nacional rows.'
        Assert-MatchText $text 'Import-XlsxRows' 'classifier must support XLSX input files.'
    }

    Invoke-Test 'client classifier packaging script builds dedicated executable' {
        $text = Get-ScriptText 'scripts\packaging\build-classifier-exe.ps1'
        Assert-MatchText $text 'classify-clientes-launcher\.ps1' 'classifier packaging must use the interactive launcher.'
        Assert-MatchText $text 'CnpjClientClassifier\.exe' 'classifier packaging must create the dedicated executable.'
        Assert-MatchText $text 'classify-clientes\.ps1' 'classifier packaging must copy the classifier script beside the executable.'
        Assert-MatchText $text 'config\.example\.ps1' 'classifier packaging must copy the public config template for PostgreSQL mode.'
        Assert-MatchText $text 'preflight\.ps1' 'classifier packaging must copy PostgreSQL preflight helpers.'
    }

    Invoke-Test 'client classifier sample run classifies synthetic rows' {
        $root = New-TestRoot; $tempRoots += $root
        $outputPath = Join-Path $root 'clientes_classificados.csv'
        $scriptPath = Join-Path $ProjectRoot 'scripts\classify-clientes.ps1'
        $inputPath = Join-Path $ProjectRoot 'examples\clientes.sample.csv'
        $simplesPath = Join-Path $ProjectRoot 'examples\simples.sample.csv'

        & $scriptPath -InputPath $inputPath -SimplesPath $simplesPath -OutputPath $outputPath -Delimiter ';' -Mode File | Out-Null

        Assert-True (Test-Path -LiteralPath $outputPath) 'classifier should create the output CSV.'
        $rows = @(Import-Csv -LiteralPath $outputPath -Delimiter ';')
        Assert-Equal 4 $rows.Count 'classifier should preserve all sample client rows.'
        Assert-Equal 'MEI' $rows[0].regime_tributario 'first sample row should be MEI.'
        Assert-Equal 'Simples Nacional' $rows[1].regime_tributario 'second sample row should be Simples Nacional.'
        Assert-Equal 'Normal' $rows[2].regime_tributario 'third sample row should be Normal.'
        Assert-Equal 'Sem CNPJ' $rows[3].regime_tributario 'fourth sample row should be marked as missing CNPJ.'
    }

    Invoke-Test 'client classifier reads XLSX client export' {
        $root = New-TestRoot; $tempRoots += $root
        $xlsxPath = Join-Path $root 'clientes.xlsx'
        $outputPath = Join-Path $root 'clientes_classificados.csv'
        $scriptPath = Join-Path $ProjectRoot 'scripts\classify-clientes.ps1'
        $simplesPath = Join-Path $ProjectRoot 'examples\simples.sample.csv'
        New-TestXlsxFile -Path $xlsxPath -Rows @(
            @('ID', 'Nome', 'Razao Social', 'Cidade', 'Estado', 'CNPJ'),
            @('1', 'Cliente Exemplo MEI', 'Empresa Exemplo MEI LTDA', 'Curitiba', 'PR', '01234567000189'),
            @('2', 'Cliente Exemplo Simples', 'Empresa Exemplo Simples LTDA', 'Joinville', 'SC', '11222333000181')
        )

        & $scriptPath -InputPath $xlsxPath -SimplesPath $simplesPath -OutputPath $outputPath -Delimiter ';' -Mode File | Out-Null

        $rows = @(Import-Csv -LiteralPath $outputPath -Delimiter ';')
        Assert-Equal 2 $rows.Count 'classifier should read XLSX client rows.'
        Assert-Equal 'MEI' $rows[0].regime_tributario 'first XLSX row should be MEI.'
        Assert-Equal 'Simples Nacional' $rows[1].regime_tributario 'second XLSX row should be Simples Nacional.'
    }

    Invoke-Test 'client classifier docs and ignore rules protect private data' {
        $gitignore = Get-ScriptText '.gitignore'
        $docs = Get-ScriptText 'docs\client-classifier.md'
        Assert-MatchText $gitignore '\*\.xlsx' '.gitignore must ignore Excel workbooks.'
        Assert-MatchText $gitignore '\*\.xls' '.gitignore must ignore Excel workbooks.'
        Assert-MatchText $gitignore 'output/' '.gitignore must ignore generated outputs.'
        Assert-MatchText $gitignore 'downloads/' '.gitignore must ignore downloaded Receita files.'
        Assert-MatchText $gitignore '!examples/\*\.sample\.csv' '.gitignore must allow synthetic sample CSV files.'
        Assert-MatchText $docs '\.xlsx` or `\.csv' 'client classifier docs must explain XLSX and CSV support.'
        Assert-MatchText $docs '-Mode Database' 'client classifier docs must explain PostgreSQL mode for large Simples files.'
        Assert-MatchText $docs '`Normal` does not mean Lucro Real' 'docs must explain the Normal limitation.'
        Assert-MatchText $docs 'Do not infer MEI from `natureza_juridica = 2135`' 'docs must reject natureza_juridica as the MEI source.'
    }
}
finally {
    foreach ($root in $tempRoots) {
        Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host "`nTest results: $script:Passed passed, $script:Failed failed."
if ($script:Failed -gt 0) { exit 1 }
