# BV-sFile2Excell

CustomTkinter desktop app for converting PDF filenames into an Excel document management workbook.

## Purpose

BV-sFile2Excell scans a folder of PDF files, parses document information from each filename, and appends valid records to an Excel workbook named `quan_ly_van_ban.xlsx` by default.

## Features

- Folder picker for the PDF source directory.
- Excel output picker for creating or updating a workbook.
- Filename parsing for document number, signing date, and Vietnamese summary.
- Duplicate detection against existing workbook rows.
- Excel formatting for the `Quản lý văn bản` sheet.
- English summary generation using `translate-rule.xlsx` rules with Google Translate fallback.
- Manual Check Update action using the remote raw `latest.json` URL from `version.json`.
- Automatic update check shortly after the GUI opens.

## Run From Source

Install dependencies first:

```bat
pip install -r requirements.txt
```

Run the app:

```bat
python -m app.main
```

## Filename Format

The app expects PDF filenames to include:

- A document code at the beginning when available, such as `123/ABC`.
- A text summary after the document code.
- A signing date at the end in `ddmmyyyy`, `dd.mm.yyyy`, `dd-mm-yyyy`, or `dd/mm/yyyy` style.

Example:

```text
123-ABC Thong bao hop dai hoi 07062026.pdf
```

Optional translation rules can be placed in `translate-rule.xlsx`. Use Vietnamese phrases in column A and English replacements in column B.

## Build

First release builds the app and updater:

```bat
build.bat first
```

Update releases build only the app:

```bat
build.bat release
```

## Release

Before any release, read `RELEASE_WORKFLOW.md`.

When ready, ask Codex:

```text
Hoan thanh phien ban moi
```

Codex will review changes, bump metadata, build, self-test, scan for secrets, commit, push, create the tag and GitHub Release, upload the executable, update `latest.json`, and verify the raw `latest.json` URL.

## Security

Do not commit `.env`, credentials, tokens, logs, backups, temporary files, `build/`, `dist/`, or PyInstaller specs. The required ignore rules are in `.gitignore`.
