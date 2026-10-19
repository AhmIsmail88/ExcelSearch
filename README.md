# ExcelSearch

Local Arabic RTL application for searching large Excel files and reviewing comparisons between two datasets. Built with Python, Flask, SQLite and openpyxl.

## Features

- Index Excel workbooks into searchable local databases, with incremental synchronization.
- Search and filter across files; exclude rows containing selected words.
- Match rows by a single key or a composite key of up to six columns.
- Suggest description matches with similarity scores and manual acceptance or rejection.
- Compare yearly quantities, selected totals or custom column pairs independently of display columns.
- Review rows side by side with highlighted text differences and review notes.
- Save comparison snapshots and resume decisions, notes and the current results view after restarting.
- Export formatted Excel comparison reports, including selected descriptive columns and review notes.

## Run

Requires Python 3.10 or newer.

```sh
python -m venv .venv
# Windows PowerShell:
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

Open http://127.0.0.1:5000 and enter the path of your own workbook. The application runs locally.

## Comparison workflow

1. Select two indexed files and apply the required filters before matching.
2. Choose key matching (optionally add composite key columns) or description similarity.
3. Choose quantity column pairs and the descriptive columns to display in the report.
4. Review highlighted differences. Similarity matches are suggestions until manually accepted.
5. Save notes with the note button or a review decision, then export the report.

Comparison snapshots are immutable: run a new comparison to include subsequent source changes. Closing the result window preserves the saved session; permanent deletion is a separate action. Blank quantities remain distinct from zero. Equal totals with different annual distributions are identified separately.

## Local data and privacy

`config.json`, `databases/` and `comparisons/` contain local paths or user data and are intentionally excluded from this repository, along with source spreadsheets, reports, screenshots, logs and builds. Back up these local folders privately if you need to preserve your indexed files and review sessions. Never commit secrets or personal data. Git ignore rules are a safeguard, not a replacement for reviewing staged files.

## Tests

```sh
python -m unittest discover -p "test_*.py"
node --check static/app.js
node --check static/quantity.js
node --check static/review.js
```

Tests create synthetic workbooks and temporary databases; no personal datasets are provided.

## Build a Windows executable

```sh
pyinstaller --noconfirm ExcelSearch.spec
```

The executable is produced at `dist/ExcelSearch.exe`. The specification bundles the templates and static assets. Preserve your private configuration and data folders separately when updating.
