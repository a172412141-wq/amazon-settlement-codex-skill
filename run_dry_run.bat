@echo off
set WORKBOOK=%1
set PDFDIR=%2
set OUTPUT=%3
if "%WORKBOOK%"=="" set WORKBOOK=.\template.xlsx
if "%PDFDIR%"=="" set PDFDIR=.\pdfs
if "%OUTPUT%"=="" set OUTPUT=.\template_filled.xlsx
python .agents\skills\amazon-settlement-xlsx\scripts\run.py --workbook "%WORKBOOK%" --pdf-dir "%PDFDIR%" --output "%OUTPUT%" --mapping .agents\skills\amazon-settlement-xlsx\assets\store_mapping.csv --dry-run
