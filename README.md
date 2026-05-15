# Receita CNPJ ETL Pipeline

Windows-first PowerShell ETL pipeline for Receita Federal CNPJ open data using PostgreSQL.

The project downloads public CNPJ datasets, prepares the raw files for import, loads them into PostgreSQL, applies configurable filters, enriches the result with supporting tables, and exports a filtered CSV for downstream analysis or operational workflows.

## What It Does

- Detects the latest available Receita Federal CNPJ dataset folder.
- Downloads the required ZIP archives with retry-safe validation.
- Extracts and validates the raw files.
- Cleans null bytes and malformed raw text before import.
- Loads staged data into PostgreSQL using server-side `COPY`.
- Applies configurable filters for state, municipality, opening date, CNAE, and Simples Nacional status.
- Exports the filtered establishments to CSV.
- Includes tests for pipeline safety checks and script behavior.

## Pipeline Architecture

The main orchestrator is `run-pipeline.ps1`.

Pipeline stages:

1. `scripts/update-date.ps1` detects the current dataset folder.
2. `scripts/download.ps1` downloads Receita Federal ZIP archives.
3. `scripts/extract.ps1` extracts and validates source files.
4. `scripts/clean.ps1` creates cleaned `LIMPO_*` files.
5. `scripts/import.ps1` loads and filters data in PostgreSQL.
6. `scripts/export.ps1` writes the final CSV export.

Shared preflight, config, PostgreSQL, cleanup, and path-safety helpers live in `scripts/lib/preflight.ps1`.

See `docs/ARCHITECTURE.md` for a short stage overview.

## Requirements

- Windows PowerShell 5.1 or later
- PostgreSQL with `psql.exe` available through `PATH`, the registry, or a common local install path
- Enough disk space for Receita Federal CNPJ archives and extracted working files
- Network access to the Receita Federal open data mirror configured in `config.ps1`

## Quick Start

Clone the repository, then create your local config file:

```powershell
Copy-Item .\config.example.ps1 .\config.ps1
```

Edit `config.ps1` for your local directories, PostgreSQL database, and filters.

Set database credentials through environment variables instead of storing them in the config file:

```powershell
$env:CNPJ_ETL_DB_PORT = "5432"
$env:CNPJ_ETL_DB_USER = "postgres"
$env:CNPJ_ETL_DB_PASSWORD = "your-password"
```

Run the pipeline:

```powershell
powershell -ExecutionPolicy Bypass -File .\run-pipeline.ps1
```

The pipeline prompts before importing into PostgreSQL and before generating the final CSV export.

## Configuration

`config.example.ps1` is a safe template. Copy it to `config.ps1` before running the project.

`config.ps1` is local-only and ignored by Git. Do not commit it.

Important settings:

- `anoMes`: Receita Federal dataset folder, such as `2026-04-12`.
- `prefixo`: monthly file prefix used in extracted filenames.
- `dirDownload`: ZIP download directory.
- `dirTemp`: extraction and cleaning working directory.
- `dirOut`: CSV export directory.
- `ufs`: state filters.
- `municipiosWhitelist`: optional Receita municipality-code allowlist.
- `cnaeWhitelist`: CNAE prefixes to keep.
- `cnaeRanges`: inclusive CNAE prefix ranges to keep.
- `filterSimplesNacional`: whether to keep only rows enriched as Simples Nacional.

The example uses neutral filters and user-relative directories. Customize them for your local workload before running the ETL.

## Database Credentials

The pipeline resolves database credentials from these environment variables:

- `CNPJ_ETL_DB_HOST` or `PGHOST`
- `CNPJ_ETL_DB_PORT` or `PGPORT`
- `CNPJ_ETL_DB_USER` or `PGUSER`
- `CNPJ_ETL_DB_PASSWORD` or `PGPASSWORD`

`config.example.ps1` intentionally does not store database passwords.

## Running The Pipeline

```powershell
powershell -ExecutionPolicy Bypass -File .\run-pipeline.ps1
```

If `config.ps1` is missing, the pipeline exits with instructions to copy `config.example.ps1` first.

The pipeline can skip stages when validated downstream artifacts already exist. Cleanup behavior is controlled by `cleanupMode` and `cleanupDryRun` in `config.ps1`.

## Running Tests

Tests do not require a committed `config.ps1` and do not run the full ETL.

```powershell
powershell -ExecutionPolicy Bypass -File .\tests\run-tests.ps1
```

## Data Safety

This repository does not include generated data.

Generated exports may contain public business contact fields from Receita Federal CNPJ open data. Treat generated files as operational data and keep them out of version control.

Do not commit:

- `config.ps1`
- `.env` files
- Generated CSV exports
- Raw Receita Federal ZIP archives or extracted files
- PostgreSQL dumps or backups
- Certificates, private keys, PFX/P12 files, or PEM files
- Logs with connection strings, credentials, local paths, or operational data

See `SECURITY.md` for the public release policy.

## What Is Not Included

- No generated CSV output
- No raw Receita Federal data archives
- No extracted or cleaned working files
- No database dumps
- No executable builds
- No private local configuration
- No private repository history

## Limitations

- The project is Windows-first and tested with PowerShell.
- Full ETL runs can require substantial disk space and database time.
- Receita Federal dataset structure, filenames, and availability can change over time.
- Filtering choices are configuration examples, not recommendations for a specific business use case.
