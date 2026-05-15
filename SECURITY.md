# Security Policy

This repository is intended to contain source code, documentation, and safe example configuration only.

Do not commit:

- `config.ps1` or other local configuration files
- `.env` files or shell profiles with database credentials
- Generated CSV exports
- Raw Receita Federal archive or extracted data files
- PostgreSQL dumps, backups, or local database files
- Certificates, private keys, PFX/P12 files, PEM files, or keystores
- Logs that include connection strings, credentials, local paths, or operational data

Exports can include public business contact fields from Receita Federal CNPJ open data. Treat generated exports as operational data and keep them out of version control.

For future public updates, do not merge or rebase private repository history into this public repository. Create a fresh private-source snapshot, copy only the approved source manifest, run sanitization checks, run tests, and commit only the reviewed public tree.
