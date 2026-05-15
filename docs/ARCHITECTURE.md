# Architecture

This project is a Windows-first PowerShell ETL pipeline for Receita Federal CNPJ open data.

The pipeline stages are:

1. Detect the current Receita Federal dataset folder.
2. Download the required CNPJ ZIP archives.
3. Extract source files into a working directory.
4. Clean null bytes and malformed raw text before database loading.
5. Load staged data into PostgreSQL.
6. Apply configurable filters for state, municipality, activity date, CNAE, and Simples Nacional status.
7. Enrich filtered establishments with company, municipality, CNAE, and tax-regime data.
8. Export the filtered result as CSV for downstream analysis or operational workflows.

Generated data, raw archives, database dumps, and local configuration are intentionally excluded from this repository.
