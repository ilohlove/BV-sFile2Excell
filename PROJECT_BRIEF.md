# PROJECT_BRIEF.md

## Project Name

sFile2Excell

## Project Slug

sFile2Excell

## Purpose

Desktop tool for scanning a folder of PDF document files, extracting document number, signing date, and Vietnamese summary from filenames, then appending normalized records to an Excel management workbook.

## Main Features

- Select a folder containing PDF files.
- Select or create an Excel workbook output.
- Extract document number, signing date, and summary from PDF filenames.
- Skip duplicate rows already present in the workbook.
- Add an English summary column using `translate-rule.xlsx` rules and Google Translate fallback.
- Keep update check actions from the desktop template.

## GUI Requirements

- Main screen with app name and version.
- Folder picker for the PDF source directory.
- Excel save picker for the output workbook.
- Run button with background processing.
- Check Update button and update confirmation dialog.
- Short log area with processing summary.

## Business Rules

- Keep app-specific logic in `app/services/business_logic.py`.
- Keep framework files stable unless a real project needs a framework change.
- Never release while required metadata is incomplete.

## Input Data

- A local folder containing `.pdf` files.
- Optional `translate-rule.xlsx` with Vietnamese phrases in column A and English replacements in column B.
- Existing or new `.xlsx` output workbook.

## Output Data

- Excel workbook named `quan_ly_van_ban.xlsx` by default.
- Worksheet `Quản lý văn bản`.
- Columns: `Số`, `Ngày ký`, `Trích yếu nội dung`, `Trích yếu tiếng Anh`, `Người xin số`, `CBTT`, `BBH Liên quan`.

## Project Notes

PDF filenames should end with a valid signing date in `ddmmyyyy`, `dd.mm.yyyy`, `dd-mm-yyyy`, or `dd/mm/yyyy` style. Document numbers are parsed from the beginning of the filename when they contain a separator such as `/`, `-`, or `.`.
