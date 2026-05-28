# Sample Client Classifier

`scripts/classify-clientes.ps1` classifies a sample client list as `MEI`, `Simples Nacional`, `Normal`, `Sem CNPJ`, or `CNPJ invalido` using the Receita Federal Simples dataset.

This tool is separate from the lead export. The lead export is filtered for prospecting and can intentionally remove MEI or Normal companies, so it is not a reliable source for classifying a separate client list.

## When To Use It

Use this classifier when you already have a sample client export with CNPJs and only need the MEI/Simples/Normal status.

You only need the Receita `SIMPLES.CSV` file for this classification. You do not need the full Estabelecimentos dataset unless you also need address, CNAE, active/inactive status, or lead-generation filters.

## Input

Use an `.xlsx` or `.csv` file with a `CNPJ` column. The column name can be changed with `-CnpjColumn`, and the original system Excel export can be selected directly when it is kept outside the repository.

Small sample run:

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
  -InputPath 'C:\Work\clientes.xlsx' `
  -SimplesPath 'C:\Receita\F.K03200W.SIMPLES.CSV' `
  -OutputPath .\output\clientes_classificados.csv
```

`-Mode Database` is the default because Receita Simples CSV files can be several GB. It reuses the same PostgreSQL discovery, preflight validation, and server-side `COPY` workflow as the main ETL pipeline. Use `-Mode File` only for small sample files or tests.

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

`Normal` does not mean Lucro Real or Lucro Presumido. It only means the CNPJ was not marked as MEI or Simples Nacional in the Receita Simples snapshot used.

Do not infer MEI from `natureza_juridica = 2135`; the authoritative field for this classifier is `opcao_pelo_mei` from the Simples dataset.

## PostgreSQL Mode

Database mode requires:

- Windows PowerShell 5.1 or later
- PostgreSQL running locally or on a reachable host
- `psql.exe` available through `PATH`, the registry, or a common local install path
- a local `config.ps1` copied from `config.example.ps1`
- database credentials through `CNPJ_ETL_DB_*` or `PG*` environment variables
- write access to `dirTemp` and `dirOut`
- enough free disk space for temporary staging of the Simples file and small result files

The script stages target client CNPJs in PostgreSQL, loads the Simples file with `COPY`, joins by normalized `cnpj_basico`, writes a small match file, and then merges the result back into the original client rows.

If PostgreSQL cannot read the selected `SIMPLES.CSV` directly, the classifier copies it to `dirTemp` first so server-side `COPY` can read it. For a multi-GB file, make sure `dirTemp` has enough free space.

## File Mode

File mode scans the Simples file directly in PowerShell:

```powershell
.\scripts\classify-clientes.ps1 `
  -InputPath .\examples\clientes.sample.csv `
  -SimplesPath .\examples\simples.sample.csv `
  -OutputPath .\output\clientes_classificados.csv `
  -Mode File
```

Use this mode for examples, tests, and small files only. It is not recommended for official multi-GB Receita files.

## Interactive Launcher

Build the optional Windows launcher with:

```powershell
.\scripts\packaging\build-classifier-exe.ps1
```

The build output should be distributed as a folder or ZIP with this layout:

```text
CnpjClientClassifier/
  CnpjClientClassifier.exe
  classify-clientes.ps1
  config.example.ps1
  lib/
    preflight.ps1
```

The executable is a launcher. Keep the companion files beside it. Before running database mode from a packaged folder, copy `config.example.ps1` to `config.ps1` in that same folder and adjust local paths as needed.

## Troubleshooting

- `config.ps1 not found`: copy `config.example.ps1` to `config.ps1` and adjust local directories.
- `psql.exe` not found: install PostgreSQL client tools or add the PostgreSQL `bin` folder to `PATH`.
- database connection failed: check `CNPJ_ETL_DB_HOST`, `CNPJ_ETL_DB_PORT`, `CNPJ_ETL_DB_USER`, and `CNPJ_ETL_DB_PASSWORD`.
- permission denied on temp files: make sure PostgreSQL and the current Windows user can read/write the configured temp folder.
- huge `SIMPLES.CSV` appears slow: use `-Mode Database`, keep files on a local SSD when possible, and allow time for PostgreSQL `COPY`.
- missing `CNPJ` column: pass the correct column name with `-CnpjColumn`.
- XLSX read errors: close the workbook in Excel and retry, or export it to CSV.

## Privacy

Do not commit real client spreadsheets, client CSV files, Receita ZIPs, cleaned Receita files, or generated classifier outputs. Use only synthetic sample files under `examples/`.
