# Amazon Settlement PDF → Excel 自动填表 Codex Skill

这个仓库只保存**整理流程和本地自动化脚本**，用于把 Amazon Settlement PDF 结算报告中的核心字段填入既有 Excel 表格。

## 数据边界

仓库不保存业务信息。不要提交：

- Amazon PDF 原始文件
- Excel 模板或已填 Excel
- 审计 CSV
- 汇率缓存或导出结果
- 真实店铺名、公司主体、银行信息、税务信息

`.gitignore` 已默认排除这些文件。脚本只在本地处理文件。

## 处理字段

- income / Einnahmen / Ingresos / Entrate / Recettes
- expenses / Ausgaben / Gastos / Spese / Dépenses
- transfer to bank account / Übertragungen / Transferencias / Trasferimenti / Virements
- tax / Steuer / Impuestos / Imposte / Taxes
- 币种
- 月初人民币汇率中间价
- 人民币金额

费用和回款在 Amazon PDF 中经常显示为负数；写入 Excel 时默认按业务口径写成正数，同时在审计 CSV 中保留 PDF 原始符号，方便复核。

## 目录结构

```text
project/
  .agents/skills/amazon-settlement-xlsx/
  pdfs/                 # 本地放 PDF，不提交
  template.xlsx          # 本地放 Excel，不提交
  store_mapping.local.csv # 本地映射，不提交，可选
```

## 安装

```bash
python -m pip install -r .agents/skills/amazon-settlement-xlsx/scripts/requirements.txt
```

## 第一步：先 dry-run

```bash
python .agents/skills/amazon-settlement-xlsx/scripts/run.py \
  --workbook ./template.xlsx \
  --pdf-dir ./pdfs \
  --output ./template_filled.xlsx \
  --mapping ./store_mapping.local.csv \
  --dry-run
```

检查生成的 `template_filled.audit.csv`，确认 PDF 文件名、月份、币种、income、expenses、transfer、tax、汇率日期和写入行都正确。

## 第二步：正式写入

```bash
python .agents/skills/amazon-settlement-xlsx/scripts/run.py \
  --workbook ./template.xlsx \
  --pdf-dir ./pdfs \
  --output ./template_filled.xlsx \
  --mapping ./store_mapping.local.csv
```

默认不会覆盖原始 Excel。不要使用 `--in-place`，除非已经备份。

## 本地店铺映射文件

按示例复制一份到本地：

```bash
cp examples/store_mapping.example.csv store_mapping.local.csv
```

格式：

```csv
store_site,registered_company,display_name,currency,notes
DE-SHOP,Company Placeholder,Shop Placeholder,EUR,example only
```

真实映射只保存在本地，不要提交。

## Codex 使用提示

在 Codex 中打开该项目目录后，可以输入：

```text
使用 amazon-settlement-xlsx skill：读取 pdfs 文件夹内所有 Amazon 结算 PDF，按现有 Excel 模板填写 income、expenses、transfer to bank account、tax、币种、月初汇率和人民币金额。先 dry-run，生成 audit CSV 让我复核，不要改任何格式，不要提交任何业务文件。
```

## 复核要求

每次正式写入前必须核对：

1. PDF summary 区金额与 audit CSV 完全一致。
2. expenses 和 transfer 的 PDF 原始负数已按表格业务口径写成正数。
3. Excel 原有公式、格式、列宽、样式没有被改动。
4. 汇率日期按“每月 1 号；无交易数据则回退到上月最后一个交易日”处理。
5. 每个 PDF 只写入一行，没有重复写入。
