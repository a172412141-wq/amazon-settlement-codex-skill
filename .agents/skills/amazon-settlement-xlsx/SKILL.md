---
name: amazon-settlement-xlsx
summary: Extract Amazon settlement PDF summary amounts and populate an existing accounting XLSX without changing workbook formatting.
description: Use this skill when the user provides Amazon settlement PDF files and an existing accounting Excel workbook, and asks to fill income, expenses, transfer-to-bank-account, currency, exchange rate, and RMB formulas while preserving the workbook's existing formatting and formulas.
---

# Amazon Settlement PDF → Accounting XLSX Workflow

## Scope

Use this skill only for Amazon settlement reports such as `*_Standard-YYYY-MM.pdf` and an existing accounting workbook. The task is to extract report-level summary amounts and fill the workbook. Do not redesign the workbook. Do not add charts, sheets, comments, or formatting unless the user explicitly asks.

## Inputs

- One or more Amazon settlement PDF files.
- One existing `.xlsx` workbook used as the accounting template.
- Optional `assets/store_mapping.csv` with shop/site names and company-name overrides.
- Internet access if exchange rates must be pulled from ChinaMoney.

## Required output

- A new `.xlsx` copy of the workbook with only the required cells populated.
- A CSV audit file listing every PDF, extracted raw values, written values, exchange-rate source date, and verification status.
- If any extraction is ambiguous, stop and ask the user to confirm; do not guess.

## Workbook field mapping

Use header names rather than fixed letters whenever possible. In the provided template, the expected mapping is:

| Workbook field | Meaning | Default column in sample |
|---|---|---|
| 时间 | report month first day | A |
| 平台 | fixed value `亚马逊` | B |
| 店铺(按站点) | shop/site code from mapping or filename | C |
| 注册主体(公司名字) | company name from mapping or PDF | D |
| 亚马逊结算报告 | source PDF filename | E |
| 店铺销售额（原币） | Amazon summary `income` / 收入 | G |
| 销售税 Tax（原币） | Amazon summary `tax` / 税, if present | H |
| 店铺销售额（原币 包含TAX） | formula `=G+H` copied from template row | I |
| 销售费用（原币） | absolute value of Amazon summary `expenses` / 支出 | J |
| 账单回款额（原币） | absolute value of Amazon `transfer to bank account` / bank transfer | K |
| 币种 | currency code such as EUR, USD, MXN | L |
| 汇率 | CNY rate for that currency | M |
| 人民币 columns | formulas copied from template row | N:P, R, T, V, AA etc. |

## Sign convention

Amazon settlement PDFs usually show expenses and transfers as negative amounts from Amazon's perspective. The accounting workbook treats expenses and received bank transfers as positive business amounts. Therefore:

- Preserve raw PDF signs in the audit CSV.
- Write `income` as shown if positive.
- Write `expenses` to the workbook as `abs(raw_expenses)`.
- Write `transfer to bank account` to the workbook as `abs(raw_transfer)`.
- Write `tax` as shown. If the report has a negative tax subtotal, preserve that sign unless the user provides a different accounting rule.

If the workbook formula structure clearly uses the opposite sign convention, pause and ask before writing.

## Exchange-rate rule

Use the China Foreign Exchange Trade System / ChinaMoney RMB central parity rate. The target date is the first calendar day of the report month. If that date has no rate, use the last available trading day before it.

- Main historical API pattern used by the helper script:
  `https://www.chinamoney.com.cn/ags/ms/cm-u-bk-ccpr/CcprHisNew?startDate=YYYY-MM-DD&endDate=YYYY-MM-DD&currency=EUR/CNY&pageNum=1&pageSize=300`
- Query a small date window ending on the first day of the month, then select the latest returned record whose date is `<= month_start`.
- For JPY, ChinaMoney quotes `100JPY/CNY`; divide by 100 before multiplying workbook amounts.
- Cache rates in a local CSV so re-runs are deterministic.

## Extraction and verification steps

1. Read the PDF visually and textually. Prefer coordinate-based row extraction from the top `Summaries / Zusammenfassungen / Resumen / Riepilogo / Synthèse` table.
2. Extract:
   - report month
   - display/store name if available
   - registered company name if available
   - currency
   - income
   - expenses
   - tax if present
   - transfer-to-bank-account / transfers
3. Cross-check each target amount using two independent methods when possible:
   - top summary row coordinate extraction
   - text/regex fallback around target labels
4. Verify that the detailed bank-transfer row and the summary transfer row match when both are present.
5. Open the workbook and find columns by header text. Do not assume columns if the headers moved.
6. If appending a new row, copy the prior/template row style, data validation, merged-cell behavior if relevant, and formulas before writing values.
7. Write only the source/input cells and exchange-rate cell. Keep all other formulas and formats intact.
8. Recalculate formulas if Excel is available; otherwise mark the workbook for automatic recalculation on open.
9. Create an external audit CSV. Do not add an audit sheet unless the user asks.
10. Reopen or re-read the output workbook and verify the written cells against the audit CSV before returning it.

## Stop conditions

Stop and ask for confirmation if:

- Any required amount cannot be found.
- Two extraction methods produce different amounts beyond 0.01.
- The workbook headers cannot be matched.
- The ChinaMoney rate is unavailable and no cached/manual rate exists.
- The store/site or company name cannot be inferred from mapping, filename, or PDF and the workbook requires it.

## Recommended command

From the project folder containing the skill:

```bash
python .agents/skills/amazon-settlement-xlsx/scripts/run.py \
  --workbook ./template.xlsx \
  --pdf-dir ./pdfs \
  --output ./template_filled.xlsx \
  --mapping .agents/skills/amazon-settlement-xlsx/assets/store_mapping.csv
```

Use `--dry-run` first for new site/language combinations.
