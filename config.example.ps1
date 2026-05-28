# =============================================================
# MONTHLY SETTINGS AND FILTERS - EDIT HERE BEFORE RUNNING
# =============================================================

@{
    # ---------------------------------------------------------
    # MONTHLY PARAMETERS (update as needed)
    # ---------------------------------------------------------
    anoMes       = "2026-05-10"          # Dataset folder published by Receita Federal (YYYY-MM or YYYY-MM-DD)
    prefixo      = "D61110"           # Monthly file prefix (for example: D61110, D51108). Check the extracted filenames.
    
    # ---------------------------------------------------------
    # DIRECTORIES AND URLS
    # ---------------------------------------------------------
    baseUrl      = "https://dados-abertos-rf-cnpj.casadosdados.com.br/arquivos"
    dirDownload  = "$env:USERPROFILE\Downloads\cnpj"  # Where ZIP files are downloaded
    dirTemp      = "$env:USERPROFILE\cnpj-etl\temp"   # Working directory for extraction and cleaning
    dirOut       = "$env:USERPROFILE\cnpj-etl\output" # Final CSV export directory

    # ---------------------------------------------------------
    # CLEANUP AND RUNTIME SAFETY
    # cleanupMode:
    #   aggressive = delete ZIPs after extraction, raw files after cleaning, LIMPO_* after import
    #   balanced   = keep ZIPs, delete raw files after cleaning, delete LIMPO_* after import
    #   debug      = keep major artifacts for troubleshooting
    # cleanupDryRun prints what would be deleted without removing files.
    # ---------------------------------------------------------
    cleanupMode              = "debug"
    cleanupDryRun            = $false
    cleaningMaxWorkers       = 3
    sqlCommandTimeoutSeconds = 14400
    allowZeroFinalRows       = $false
    resetFinalTablesOnImport = $true   # Replace current CRM/Empresas dataset on each import run.
    requireEnrichmentMatches = $true   # Fail import if Empresas/Simples files exist but match zero target CNPJs.
    filterSimplesNacional    = $true   # Keep only rows enriched as Simples Nacional. MEI rows are removed.
    captureActiveClients     = $true   # Also save government CNAE data for the existing active-client list below.
    activeClientsPath        = ".\examples\active-clients.sample.csv" # CSV with CNPJ/cnpj_normalizado for sample active clients.
    
    # ---------------------------------------------------------
    # IMPORT FILTERS
    # ---------------------------------------------------------
    dataCorte    = ""                 # Optional override. Empty = opened in the last 6 months from execution date.
    ufs          = @("PR", "SC")      # Target states (example: "PR", "SC", "SP")
    municipiosWhitelist = @()          # Receita municipality codes to keep. Empty array keeps all municipalities.
    
    # CNAE filter (main code prefix)
    # Companies starting with any of these prefixes will be KEPT:
    cnaeWhitelist = @(
        "47",   # Retail, including fuel/gas/pharmacy exact categories
        "45",   # Vehicles and parts
        "01",   # Agro / rural production
        "56",   # Food service
        "462",  # Agricultural wholesale
        "463",  # Food wholesale
        "681",  # Real estate - own properties
        "682",  # Real estate administration / brokerage
        "692"   # Accounting services
    )
    
    # Companies whose CNAEs fall within these ranges will also be KEPT:
    cnaeRanges    = @(
        @("10", "33"),  # Industry
        @("41", "43"),  # Construction
        @("49", "53"),  # Transport and logistics
        @("61", "61"),  # Telecom
        @("62", "63"),  # IT, software, information services
        @("64", "66"),  # Finance, insurance, auxiliary services
        @("86", "86")   # Health clinics and services
    )
    
    # ---------------------------------------------------------
    # POSTGRESQL DATABASE
    # dbPort, dbUser, and dbPassword are intentionally not stored here.
    # Provide them through environment variables instead:
    #   CNPJ_ETL_DB_PORT or PGPORT
    #   CNPJ_ETL_DB_USER or PGUSER
    #   CNPJ_ETL_DB_PASSWORD or PGPASSWORD
    # ---------------------------------------------------------
    dbHost       = "localhost"
    dbName       = "postgres"
}



