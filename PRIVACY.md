# Privacy and data-handling policy

This repository is for workflow automation only. It must not contain business PDFs, Excel workbooks, filled output files, settlement data, audit CSVs, exchange-rate cache files, company names, store names, or real transaction data.

## Stored in GitHub

Only these categories should be committed:

- Codex Skill instructions
- Python automation scripts
- Field mapping documentation
- Placeholder example configuration files
- Usage documentation

## Not stored in GitHub

Do not commit:

- Amazon settlement PDF files
- Source Excel templates that contain company/store/business data
- Filled Excel output files
- Audit CSV files
- Exchange-rate cache files
- Store/company mapping with real names
- Screenshots, logs, or exported results containing business numbers

## Local processing

The script processes files locally in the user's project folder. It does not intentionally upload PDF, Excel, audit, or exchange-rate data to GitHub.

## Required operating rule

Before pushing changes, run:

```bash
git status --short
```

Only code, documentation, and placeholder examples should appear.
