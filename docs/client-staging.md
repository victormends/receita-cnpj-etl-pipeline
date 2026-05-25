# Existing Client Staging

`scripts/build-client-staging.ps1` builds a PostgreSQL-ready CSV for existing clients and is also used by the interactive `CnpjClientClassifier.exe` launcher when present beside the executable.

The script reuses the current classifier, then enriches its output with operational columns from a second CSV and optional government/CNAE category data from PostgreSQL.

## Inputs

- Base client XLSX/CSV with `ID` and `CNPJ`, for example `clientes (shared) Table.xlsx`.
- Optional Receita Simples CSV, same as the existing classifier uses.
- Optional enrichment CSV, for example `Clientes_A_limpo_cnpj_corrigido.csv`.

The base file is authoritative for the final row set and CNPJ. The enrichment file is supplemental.

If `-EnrichmentPath` is omitted, the script searches for `data\Clientes_A_limpo_cnpj_corrigido.csv` beside the executable/project and then in Downloads. If `-SimplesPath` is omitted, the output is still generated and regime fields are marked as not classified.

## CNPJ Safety

The enrichment CSV can still contain a broken raw `CNPJ` column in scientific notation. The staging builder does not use that field as canonical CNPJ.

Final CNPJ precedence is:

1. classifier `cnpj_normalizado`, if 14 digits
2. base file `CNPJ`, if 14 digits
3. enrichment `cnpj_corrigido`, if 14 digits
4. blank

Audit columns include `cnpj_fonte`, `cnpj_enriquecimento`, `cnpj_fonte_enriquecimento`, and `cnpj_diverge_enriquecimento`.

## ID Matching

The join key is always `id_normalizado`:

```text
remove every non-digit character, then trim leading zeroes
```

This makes `1.003`, `1003`, and `0001003` match as `1003`.

## Government Categories

When `-IncludeGovernmentData` is set, the script exports category columns from PostgreSQL tables populated by the ETL import:

- `clientes_ativos_governo`
- `estabelecimentos_crm`
- `estabelecimentos_categorias`
- `empresas_dados`

`clientes_ativos_governo` is populated by `import.ps1` when `captureActiveClients = $true` and `activeClientsPath` points to the active-client CSV. It stores CNAE/company data for the active CNPJs before cleanup, without applying the CRM lead filters.

The category layer uses only Receita/government data. It does not infer internet provider, system version, or operational state from government data.

Categories currently generated from CNAE/legal nature include:

- `transporte`
- `gas_combustivel`
- `farmacia`
- `construtora`
- `industria`
- `agro_produtor_rural`
- `telecom`
- `ti_software`
- `saude_clinica`
- `financeiro`
- `cooperativa`
- `imobiliario`
- `grafica`
- `frigorifico`

The output marks import/export and special fiscal benefit categories as `external_source_required` because CNAE alone cannot prove them safely.

## Usage

```powershell
.\scripts\build-client-staging.ps1 `
  -InputPath '.\data\clientes.xlsx' `
  -EnrichmentPath '.\data\Clientes_A_limpo_cnpj_corrigido.csv' `
  -SimplesPath '.\data\F.K03200$W.SIMPLES.CSV' `
  -OutputPath '.\output\clientes_postgres_staging.csv' `
  -IncludeGovernmentData
```

Generated files:

- normalized staging CSV at `-OutputPath`
- intermediate classifier CSV beside the output by default
- review CSV with `.revisar.csv` extension
- summary JSON with `.summary.json` extension

## PostgreSQL Import

Create staging tables with:

```sql
\i sql/client_staging.sql
```

Then import from the local machine with `psql`:

```sql
\copy clientes_staging FROM 'C:/path/clientes_postgres_staging.csv' WITH CSV HEADER DELIMITER ';' NULL '' ENCODING 'UTF8';
```

Use client-side `\copy` for local Windows files. Server-side `COPY` only works when the PostgreSQL server can read the file path.
