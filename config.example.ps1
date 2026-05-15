# =============================================================
# EXAMPLE SETTINGS AND FILTERS - COPY TO config.ps1 BEFORE RUNNING
# =============================================================

@{
    # ---------------------------------------------------------
    # MONTHLY PARAMETERS (update as needed)
    # ---------------------------------------------------------
    anoMes       = "2026-04-12"       # Dataset folder published by Receita Federal (YYYY-MM or YYYY-MM-DD)
    prefixo      = "D61110"           # Monthly file prefix (for example: D61110, D51108). Check the extracted filenames.
    
    # ---------------------------------------------------------
    # DIRECTORIES AND URLS
    # ---------------------------------------------------------
    baseUrl      = "https://dados-abertos-rf-cnpj.casadosdados.com.br/arquivos"
    dirDownload  = "$env:USERPROFILE\CnpjEtl\downloads"  # Where ZIP files are downloaded
    dirTemp      = "$env:USERPROFILE\CnpjEtl\temp"       # Working directory for extraction and cleaning
    dirOut       = "$env:USERPROFILE\CnpjEtl\output"     # Final CSV export directory

    # ---------------------------------------------------------
    # CLEANUP AND RUNTIME SAFETY
    # cleanupMode:
    #   aggressive = delete ZIPs after extraction, raw files after cleaning, LIMPO_* after import
    #   balanced   = keep ZIPs, delete raw files after cleaning, delete LIMPO_* after import
    #   debug      = keep major artifacts for troubleshooting
    # cleanupDryRun prints what would be deleted without removing files.
    # ---------------------------------------------------------
    cleanupMode              = "aggressive"
    cleanupDryRun            = $false
    cleaningMaxWorkers       = 3
    sqlCommandTimeoutSeconds = 14400
    allowZeroFinalRows       = $false
    resetFinalTablesOnImport = $true   # Replace current filtered/company dataset on each import run.
    requireEnrichmentMatches = $true   # Fail import if Empresas/Simples files exist but match zero target CNPJs.
    filterSimplesNacional    = $true   # Keep only rows enriched as Simples Nacional. MEI rows are removed.
    
    # ---------------------------------------------------------
    # IMPORT FILTERS
    # ---------------------------------------------------------
    dataCorte    = ""                 # Optional override. Empty = opened in the last 6 months from execution date.
    ufs          = @("SP")             # Example state filter. Customize for your use case.
    municipiosWhitelist = @()          # Receita municipality codes to keep. Empty array keeps all municipalities.
    
    # CNAE filter (main code prefix)
    # Establishments starting with any of these prefixes will be kept.
    cnaeWhitelist = @("62")
    
    # Establishments whose CNAEs fall within these ranges will also be kept.
    cnaeRanges    = @(
        @("10", "33")   # Example broad industry range
    )
    
    # ---------------------------------------------------------
    # POSTGRESQL DATABASE
    # dbPort, dbUser, and dbPassword are intentionally not stored here.
    # Provide them through environment variables instead:
    #   CNPJ_ETL_DB_HOST or PGHOST
    #   CNPJ_ETL_DB_PORT or PGPORT
    #   CNPJ_ETL_DB_USER or PGUSER
    #   CNPJ_ETL_DB_PASSWORD or PGPASSWORD
    # ---------------------------------------------------------
    dbHost       = "localhost"
    dbName       = "postgres"
}

