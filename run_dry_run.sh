#!/usr/bin/env bash
set -euo pipefail
python .agents/skills/amazon-settlement-xlsx/scripts/run.py \
  --workbook "${1:-./template.xlsx}" \
  --pdf-dir "${2:-./pdfs}" \
  --output "${3:-./template_filled.xlsx}" \
  --mapping .agents/skills/amazon-settlement-xlsx/assets/store_mapping.csv \
  --dry-run
