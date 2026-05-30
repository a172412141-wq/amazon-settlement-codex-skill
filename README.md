# Amazon Settlement PDF → 数据表自动生成

这个仓库只保存**整理流程和本地自动化脚本**。现在默认流程已经简化为：

```text
输入 Amazon Settlement PDF → 自动生成 Excel 数据表
```

不再强制需要你提供原 Excel 模板。

## 数据边界

仓库不保存业务信息。不要提交：

- Amazon PDF 原始文件
- 生成的 Excel / CSV
- 审计结果
- 汇率缓存
- 真实店铺名、公司主体、银行信息、税务信息

`.gitignore` 已默认排除这些文件。脚本只在本地处理文件。

## 生成字段

自动生成的数据表包含：

- 月份
- 平台
- 店铺(按站点)
- 注册主体(公司名字)
- 亚马逊结算报告文件名
- 币种
- 汇率日期
- 汇率
- 店铺销售额（原币）
- 销售税 Tax（原币）
- 店铺销售额（原币 包含 TAX）
- 销售费用（原币）
- 账单回款额（原币）
- 店铺销售额（人民币）
- 销售税 Tax（人民币）
- 店铺销售额（人民币 包含 TAX）
- 销售费用（人民币）
- 账单回款额（人民币）
- PDF 原始 expenses
- PDF 原始 transfer
- 状态
- 备注

费用和回款在 Amazon PDF 中经常显示为负数；输出表按业务口径写成正数，同时保留 PDF 原始 expenses / transfer 方便复核。

## 最简单使用方式

### 1. 安装依赖

```bash
python -m pip install -r .agents/skills/amazon-settlement-xlsx/scripts/requirements.txt
```

### 2. 把 PDF 放进 `pdfs/` 文件夹

```text
pdfs/
  xxx_Standard-2026-04.pdf
  yyy_Standard-2026-04.pdf
```

### 3. 一键生成 Excel 数据表

Mac / Linux：

```bash
./generate_table.sh
```

Windows：

```bat
generate_table.bat
```

默认输出：

```text
amazon_settlement_table.xlsx
```

## 指定 PDF 或输出文件

读取某个文件夹：

```bash
python .agents/skills/amazon-settlement-xlsx/scripts/pdf_to_table.py ./pdfs --output ./result.xlsx
```

读取单个 PDF：

```bash
python .agents/skills/amazon-settlement-xlsx/scripts/pdf_to_table.py ./pdfs/report.pdf --output ./result.xlsx
```

同时生成 CSV：

```bash
python .agents/skills/amazon-settlement-xlsx/scripts/pdf_to_table.py ./pdfs --output ./result.xlsx --csv ./result.csv
```

## 本地店铺映射，可选

不想让脚本只从文件名推断店铺/主体时，可以在本地创建：

```text
store_mapping.local.csv
```

格式：

```csv
store_site,registered_company,display_name,currency,notes
DE-SHOP,Company Placeholder,Shop Placeholder,EUR,local only
```

真实映射只保存在本地，不要提交。

## 汇率规则

脚本按“人民币汇率中间价”处理：

- 目标日期：结算月份的每月 1 号。
- 如果 1 号没有交易数据，则使用之前最近一个交易日。
- 汇率缓存保存在本地 `.cache/`，不会提交到 GitHub。

## Codex 使用提示

在 Codex 中打开该项目目录后，可以输入：

```text
使用 amazon-settlement-xlsx skill：读取 pdfs 文件夹内所有 Amazon 结算 PDF，直接生成数据表，不需要 Excel 模板。字段包括 income、expenses、transfer to bank account、tax、币种、月初汇率和人民币金额。不要提交任何 PDF、Excel、CSV 或真实业务信息。
```

## 原模板填表模式

如果以后仍然需要写入已有 Excel 模板，可以继续使用旧脚本：

```bash
python .agents/skills/amazon-settlement-xlsx/scripts/run.py \
  --workbook ./template.xlsx \
  --pdf-dir ./pdfs \
  --output ./template_filled.xlsx
```

默认推荐使用新的 PDF-only 方式。