# =============================================================
# run-pipeline.ps1 - Main orchestrator
# =============================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Test-DownloadComplete {
    param([hashtable]$Config)

    $requiredFiles = @(
        'Cnaes.zip', 'Municipios.zip', 'Naturezas.zip', 'Simples.zip',
        'Estabelecimentos0.zip', 'Estabelecimentos1.zip', 'Estabelecimentos2.zip',
        'Estabelecimentos3.zip', 'Estabelecimentos4.zip', 'Estabelecimentos5.zip',
        'Estabelecimentos6.zip', 'Estabelecimentos7.zip', 'Estabelecimentos8.zip',
        'Estabelecimentos9.zip',
        'Empresas0.zip', 'Empresas1.zip', 'Empresas2.zip', 'Empresas3.zip',
        'Empresas4.zip', 'Empresas5.zip', 'Empresas6.zip', 'Empresas7.zip',
        'Empresas8.zip', 'Empresas9.zip'
    )

    foreach ($fileName in $requiredFiles) {
        $path = Join-Path $Config.dirDownload $fileName
        if (-not (Test-Path -LiteralPath $path)) {
            return $false
        }

        if (-not (Test-ZipValid -Path $path)) {
            return $false
        }
    }

    return $true
}

function Test-ExtractionComplete {
    param([hashtable]$Config)

    $markers = @(
        'Cnaes_extracted.txt', 'Municipios_extracted.txt', 'Naturezas_extracted.txt', 'Simples_extracted.txt',
        'Estabelecimentos0_extracted.txt', 'Estabelecimentos1_extracted.txt', 'Estabelecimentos2_extracted.txt',
        'Estabelecimentos3_extracted.txt', 'Estabelecimentos4_extracted.txt', 'Estabelecimentos5_extracted.txt',
        'Estabelecimentos6_extracted.txt', 'Estabelecimentos7_extracted.txt', 'Estabelecimentos8_extracted.txt',
        'Estabelecimentos9_extracted.txt',
        'Empresas0_extracted.txt', 'Empresas1_extracted.txt', 'Empresas2_extracted.txt', 'Empresas3_extracted.txt',
        'Empresas4_extracted.txt', 'Empresas5_extracted.txt', 'Empresas6_extracted.txt', 'Empresas7_extracted.txt',
        'Empresas8_extracted.txt', 'Empresas9_extracted.txt'
    )

    foreach ($marker in $markers) {
        if (-not (Test-Path (Join-Path $Config.dirTemp $marker))) {
            return $false
        }
    }

    $rawEstabelecimentos = @(Get-ChildItem -Path $Config.dirTemp -Filter '*.ESTABELE' -ErrorAction SilentlyContinue | Where-Object { $_.Name -notmatch '^LIMPO_' })
    $rawEmpresas = @(Get-ChildItem -Path $Config.dirTemp -Filter '*.EMPRECSV' -ErrorAction SilentlyContinue | Where-Object { $_.Name -notmatch '^LIMPO_' })
    if ($rawEstabelecimentos.Count -eq 0 -or $rawEmpresas.Count -eq 0) {
        return $false
    }

    return $true
}

function Test-CleanComplete {
    param([hashtable]$Config)

    $cleanEstabelecimentos = @(Get-ChildItem -Path $Config.dirTemp -Filter 'LIMPO_*.ESTABELE' -ErrorAction SilentlyContinue)
    $cleanEmpresas = @(Get-ChildItem -Path $Config.dirTemp -Filter 'LIMPO_*.EMPRECSV' -ErrorAction SilentlyContinue)
    if ($cleanEstabelecimentos.Count -eq 0 -or $cleanEmpresas.Count -eq 0) {
        return $false
    }

    $rawPatterns = @('*.ESTABELE', '*.EMPRECSV', '*.SIMPLES*', '*.CNAECSV', '*.MUNICCSV', '*.NATJUCSV')
    foreach ($pattern in $rawPatterns) {
        $rawFiles = @(Get-ChildItem -Path $Config.dirTemp -Filter $pattern -ErrorAction SilentlyContinue | Where-Object { $_.Name -notmatch '^LIMPO_' })
        foreach ($rawFile in $rawFiles) {
            $cleanPath = Join-Path $Config.dirTemp ("LIMPO_{0}" -f $rawFile.Name)
            if (-not (Test-Path $cleanPath)) {
                return $false
            }
        }
    }

    return $true
}

function Invoke-PipelineScript {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Description
    )

    $global:LASTEXITCODE = 0
    & $Path
    if ($global:LASTEXITCODE -ne 0) {
        throw "$Description failed with exit code $global:LASTEXITCODE."
    }
}

function Invoke-PipelineCleanup {
    param([hashtable]$Config)

    Write-Host "`n[7/7] Cleaning transient helper files..." -ForegroundColor Green

    $targets = @()
    $targets += @(Get-ChildItem -Path $Config.dirDownload -Filter '*.partial' -File -ErrorAction SilentlyContinue)
    $targets += @(Get-ChildItem -Path $Config.dirTemp -Filter '*.sql' -File -ErrorAction SilentlyContinue)
    $targets = @($targets | Sort-Object -Property FullName -Unique)
    if ($targets.Count -eq 0) {
        Write-Host '      No transient helper files were found to delete' -ForegroundColor Yellow
        return
    }

    $deletedCount = 0
    foreach ($target in $targets) {
        if ($target.DirectoryName -ieq (Get-FullPathSafe -Path $Config.dirDownload).TrimEnd('\')) {
            $result = Remove-CnpjArtifactSafe -Config $Config -Path $target.FullName -AllowedRoot $Config.dirDownload -AllowedPatterns @('*.partial') -Reason 'pipeline transient cleanup'
        }
        else {
            $result = Remove-CnpjArtifactSafe -Config $Config -Path $target.FullName -AllowedRoot $Config.dirTemp -AllowedPatterns @('*.sql') -Reason 'pipeline transient cleanup'
        }
        if ($result.Deleted) { $deletedCount++ }
    }

    Write-Host "      Removed $deletedCount transient helper file(s)" -ForegroundColor Gray
}

try {
    $BootstrapRoot = if ($PSScriptRoot) { $PSScriptRoot } else { [System.AppDomain]::CurrentDomain.BaseDirectory.TrimEnd('\\') }
    . (Join-Path $BootstrapRoot 'scripts\lib\preflight.ps1')
    $ScriptDir = Resolve-ScriptRoot -Invocation $MyInvocation -FallbackBaseDirectory $BootstrapRoot

    Write-Host "`n=======================================================================" -ForegroundColor Cyan
    Write-Host " CNPJ ETL PIPELINE" -ForegroundColor Cyan
    Write-Host "=======================================================================`n" -ForegroundColor Cyan

    Write-Host "[STEP 0] Detecting latest dataset..." -ForegroundColor Yellow
    & "$ScriptDir\scripts\update-date.ps1"

    $bootstrap = Import-CnpjConfig -ScriptRoot (Join-Path $ScriptDir 'scripts')
    $config = $bootstrap.Config

    Write-Step 'Pipeline preflight checks'
    $tools = Invoke-CnpjPreflight -Config $config -EnsureDirectories @($config.dirDownload, $config.dirTemp, $config.dirOut) -MinFreeSpaceGB @{ $config.dirDownload = 8; $config.dirTemp = 12; $config.dirOut = 2 } -RequirePostgres -RequireDatabase -CheckMemory
    $psqlPath = $tools.PsqlPath

    Write-Host " Config:" -ForegroundColor Yellow
    Write-Host "   Dataset: $($config.anoMes)" -ForegroundColor Gray
    $displayCutoff = if ($config.ContainsKey('dataCorte') -and -not [string]::IsNullOrWhiteSpace([string]$config.dataCorte)) { $config.dataCorte } else { (Get-Date).AddMonths(-6).ToString('yyyyMMdd') }
    Write-Host "   Opening Date Cutoff: $displayCutoff (last 6 months unless overridden)" -ForegroundColor Gray
    Write-Host "   DB User: $($config.dbUser)" -ForegroundColor Gray
    Write-Host "   DB: $($config.dbHost):$($config.dbPort)/$($config.dbName)`n" -ForegroundColor Gray

    # ─── STEP 1: SETUP ───
    Write-Host "[0/6] Verifying database schema..." -ForegroundColor Green

    # Release any idle/stale connections that may hold locks on pipeline tables
    Write-Host "    Releasing stale database connections..." -ForegroundColor Gray
    try {
        Set-PostgresPasswordEnv -Config $config
        $killArgs = @('-h', $config.dbHost, '-p', [string]$config.dbPort, '-d', $config.dbName, '-U', $config.dbUser, '-w', '-tA', '-c',
            "SELECT COUNT(*) FROM (SELECT pg_terminate_backend(pid) FROM pg_stat_activity WHERE datname = '$($config.dbName)' AND pid <> pg_backend_pid() AND state IN ('idle','idle in transaction','idle in transaction (aborted)')) t;")
        $killed = (& $psqlPath $killArgs 2>&1) | Where-Object { $_ -match '^\d+$' } | Select-Object -First 1
        if ($killed -and [int]$killed -gt 0) {
            Write-Host "    [WARN] Terminated $killed stale connection(s) holding locks." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "    [WARN] Could not release stale connections (non-fatal): $_" -ForegroundColor Yellow
    }

    $schemaTimeout = if ($config.ContainsKey('sqlCommandTimeoutSeconds')) { [int]$config.sqlCommandTimeoutSeconds } else { 600 }
    Invoke-PsqlFileChecked -PsqlPath $psqlPath -Config $config -FilePath "$ScriptDir\sql\schema.sql" -Description 'project schema' -TimeoutSec $schemaTimeout
    Write-Host "      Schema OK" -ForegroundColor Gray

    # ─── STEP 2: DOWNLOAD ───
    Write-Host "`n[1/6] Downloading files..." -ForegroundColor Green
    if ((Test-ExtractionComplete -Config $config) -or (Test-CleanComplete -Config $config)) {
        Write-Host '      Downstream artifacts already exist - skipping download' -ForegroundColor Yellow
    }
    elseif (Test-DownloadComplete -Config $config) {
        Write-Host '      Latest ZIP set already present - skipping download' -ForegroundColor Yellow
    }
    else {
        Invoke-PipelineScript -Path "$ScriptDir\scripts\download.ps1" -Description 'Download stage'
    }

    # ─── STEP 3: EXTRACT ───
    Write-Host "`n[2/6] Extracting files..." -ForegroundColor Green
    if ((Test-ExtractionComplete -Config $config) -or (Test-CleanComplete -Config $config)) {
        Write-Host '      Downstream extracted/cleaned artifacts already exist - skipping extraction' -ForegroundColor Yellow
    }
    else {
        Invoke-PipelineScript -Path "$ScriptDir\scripts\extract.ps1" -Description 'Extraction stage'
    }

    # ─── STEP 4: CLEAN ───
    Write-Host "`n[3/6] Cleaning null bytes..." -ForegroundColor Green
    if (Test-CleanComplete -Config $config) {
        Write-Host '      Cleaned files already available - skipping cleaning' -ForegroundColor Yellow
    }
    else {
        Invoke-PipelineScript -Path "$ScriptDir\scripts\clean.ps1" -Description 'Cleaning stage'
    }

    # ─── STEP 5: IMPORT ───
    Write-Host "`n[4/6] Importing into PostgreSQL..." -ForegroundColor Green
    $confirm = Read-Host "Do you want to import into the database? (s/n)"
    $didImport = $false
    if ($confirm -eq 's') {
        Invoke-PipelineScript -Path "$ScriptDir\scripts\import.ps1" -Description 'Import stage'
        $didImport = $true
    }

    # ─── VERIFICATION ───
    if ($didImport) {
        Write-Host "`n[5/6] Verifying import results..." -ForegroundColor Green
        Set-PostgresPasswordEnv -Config $config
        & $psqlPath -h $config.dbHost -p $config.dbPort -d $config.dbName -U $config.dbUser -c "SELECT 'Total CRM' as check, COUNT(*) FROM estabelecimentos_crm UNION ALL SELECT 'Total Empresas', COUNT(*) FROM empresas_dados UNION ALL SELECT 'Log imports', COUNT(*) FROM import_log;" -t
    }
    else {
        Write-Host "`n[5/6] Import verification skipped because import was skipped." -ForegroundColor Yellow
    }

    # ─── STEP 6: EXPORT ───
    Write-Host "`n[6/6] Exporting final CSV..." -ForegroundColor Green
    $confirm = Read-Host "Do you want to generate the final CSV? (s/n)"
    if ($confirm -eq 's') { Invoke-PipelineScript -Path "$ScriptDir\scripts\export.ps1" -Description 'Export stage' }

    Invoke-PipelineCleanup -Config $config

    Write-Host "`n===============================================================================" -ForegroundColor Cyan
    Write-Host " PIPELINE COMPLETED!" -ForegroundColor Green
    Write-Host "===============================================================================" -ForegroundColor Cyan
}
catch {
    Write-Host "`n[ERROR] $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.Exception.ToString() -ForegroundColor DarkRed
    if ($_.ScriptStackTrace) {
        Write-Host "`nStack:" -ForegroundColor Yellow
        Write-Host $_.ScriptStackTrace -ForegroundColor DarkYellow
    }
    Write-Host "Press Enter to close..." -ForegroundColor Yellow
    Read-Host | Out-Null
    exit 1
}
