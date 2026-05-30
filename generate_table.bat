@echo off
set INPUT=%1
set OUTPUT=%2
if "%INPUT%"=="" set INPUT=.\pdfs
if "%OUTPUT%"=="" set OUTPUT=.\amazon_settlement_table.xlsx
python .agents\skills\amazon-settlement-xlsx\scripts\pdf_to_table.py "%INPUT%" --output "%OUTPUT%"
