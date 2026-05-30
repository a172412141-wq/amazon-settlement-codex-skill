#!/usr/bin/env bash
set -euo pipefail
INPUT="${1:-./pdfs}"
OUTPUT="${2:-./amazon_settlement_table.xlsx}"
python .agents/skills/amazon-settlement-xlsx/scripts/pdf_to_table.py "$INPUT" --output "$OUTPUT"
