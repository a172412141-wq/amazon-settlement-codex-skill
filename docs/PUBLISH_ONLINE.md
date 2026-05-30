# Publish for teammates

Use this repository as a private internal tool. Do not make it public unless it contains only code, documentation, and placeholder examples.

## Recommended setup

1. Keep the repository private.
2. Give normal users read access.
3. Give maintainers write access only when they need to change scripts or documentation.
4. Keep all PDF, Excel, CSV, cache, log, and local mapping files outside Git.

## Teammate setup

```bash
git clone https://github.com/<owner>/amazon-settlement-codex-skill.git
cd amazon-settlement-codex-skill
python -m pip install -r .agents/skills/amazon-settlement-xlsx/scripts/requirements.txt
```

Create local folders and files:

```text
pdfs/
template.xlsx
store_mapping.local.csv
```

Run dry-run first:

```bash
./run_dry_run.sh ./template.xlsx ./pdfs ./template_filled.xlsx
```

After checking the audit CSV, run write mode:

```bash
./run_write.sh ./template.xlsx ./pdfs ./template_filled.xlsx
```

## Release check

Before pushing, run:

```bash
git status --short
```

Only commit code, documentation, and placeholder examples. Do not commit local business files or generated outputs.
