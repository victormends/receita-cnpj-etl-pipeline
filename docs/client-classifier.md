# Client Classifier

`scripts/classify-clientes.ps1` classifies an existing client list as `MEI`, `Simples Nacional`, or `Normal` using the Receita Federal Simples dataset.

This tool does not use the CRM lead export because that export is filtered for lead generation and can exclude MEI or Normal companies.

## Input

Use an `.xlsx` or `.csv` file with a `CNPJ` column. The original system Excel export can be selected directly.

Example:

```powershell
.\scripts\classify-clientes.ps1 `
  -InputPath .\examples\clientes.sample.csv `
  -SimplesPath .\examples\simples.sample.csv `
  -OutputPath .\output\clientes_classificados.csv `
  -Delimiter ';' `
  -Mode File
```

For real Receita Simples files, use the default PostgreSQL mode:

```powershell
.\scripts\classify-clientes.ps1 `
  -InputPath 'C:\path\clientes.xlsx' `
  -SimplesPath 'C:\path\F.K03200$W.SIMPLES.CSV' `
  -OutputPath .\output\clientes_classificados.csv
```

`-Mode Database` is the default because Receita Simples CSV files can be several GB. It reuses the same PostgreSQL discovery and `COPY` workflow as the main ETL pipeline. Use `-Mode File` only for small sample files or tests.

If a documented official consultation contradicts the local Simples snapshot, pass an auditable override CSV with `-RegimeOverridePath`. Supported columns are `cnpj` or `cnpj_basico`, `regime_tributario`, and optional `fonte`/`observacao`:

```csv
cnpj;cnpj_basico;regime_tributario;fonte;observacao
09372387000181;09372387;Simples Nacional;Consulta Optantes Simples Nacional;Optante desde 19/02/2008
```

Overrides are applied after the Simples file and should only be used when the source and observation identify the external authority.

At the end of a classifier or staging-builder run, the tool can also generate a recent XML divergence report for manual Consulta Optantes review. Set `-XmlRoot`, the `CNPJ_XML_ROOT` environment variable, or `xmlRoot` in `config.ps1`. The report defaults to the current and previous month (`-XmlDiffMonthsBack 2`) and writes `*.xml-divergencias-recentes.csv` beside the classifier/staging output:

```powershell
$env:CNPJ_XML_ROOT = '\\your-server\XML'
.
\scripts\build-client-staging.ps1 `
  -InputPath 'C:\path\clientes.csv' `
  -ClassifiedInputPath 'C:\path\clientes_classificados.csv' `
  -OutputPath 'C:\path\clientes_postgres_normalizado.csv'
```

The report includes only recent `Mismatch` and `Review` rows from `NF-e`, `NFC-e`, `NFS-e`, and `CT-e`, with blank `consulta_*` columns so you can confirm the company status on Receita before contacting the company or accountant. `MDF-e` is intentionally excluded because the real files inspected do not contain direct tax-regime evidence.

The official Consulta Optantes page (`https://consopt.www8.receita.fazenda.gov.br/consultaoptantes`) uses hCaptcha, so bulk portal automation is not included in this repository. Consult each CNPJ manually, then copy confirmed differences into the override CSV used by `-RegimeOverridePath`.

The Simples file can be a cleaned or extracted Receita file in the standard order:

```text
cnpj_basico;opcao_pelo_simples;data_opcao_simples;data_exclusao_simples;opcao_pelo_mei;data_opcao_mei;data_exclusao_mei
```

## Output

The script preserves the original client columns and appends:

- `cnpj_normalizado`
- `cnpj_basico`
- `regime_tributario`
- `classificacao_fonte`
- `classificacao_observacao`

## Classification Rules

- `MEI`: `opcao_pelo_mei = S`
- `Simples Nacional`: not MEI and `opcao_pelo_simples = S`
- `Normal`: no MEI or Simples option found in the informed Simples file
- override CSV row: replaces the Simples snapshot result for that `cnpj_basico`

`Normal` does not mean Lucro Real or Lucro Presumido. It only means the CNPJ was not marked as MEI or Simples Nacional in the Receita Simples snapshot used.

Do not infer MEI from `natureza_juridica = 2135`; the authoritative field for this classifier is `opcao_pelo_mei` from the Simples dataset.

## Privacy

Do not commit real client spreadsheets, client CSV files, Receita ZIPs, cleaned Receita files, or generated classifier outputs. Use only synthetic sample files under `examples/`.
