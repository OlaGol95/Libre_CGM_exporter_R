# LIBRE CGM Export Pipeline

Portable tool for converting raw Libre CSV exports into visit-level CGM Excel
workbooks with glucose summaries and optional R `iglu` metrics.

## Description

The program reads raw Libre CSV exports from patient folders and assigns glucose
measurements to visit windows stored in:

```text
data_visits_protected.zip
```

The main output is:

```text
<patient_export_code>_CGM.xlsx
```

Each workbook contains:

- `v1`, `v3`, `v4`, `v6` sheets with raw glucose rows,
- a `summary` sheet with Python glucose metrics,
- optional R `iglu` metrics when R and `iglu` are installed.

## Repository Files

Required files:

```text
program_LIBRE_CGM.py
data_visits_protected.zip
requirements_libre_cgm.txt
install_iglu.R
README.md
LIBRE_CSV_GITHUB_INSTRUKCJA.md
SQL_DATA_DICTIONARY.md
.gitignore
```

## Installation

Install Python dependencies:

```powershell
pip install -r requirements_libre_cgm.txt
```

Optional: install R package `iglu` for extended CGM metrics:

```powershell
Rscript install_iglu.R
```

## Configuration

Optional environment variables:

```powershell
$env:LIBRE_DIR = "<path-to>\LIBRE"
$env:LIBRE_CODES_XLSX = "<path-to>\kody.xlsx"
```

If `LIBRE_DIR` is not set, the repository directory is used as the LIBRE
workspace.

## Usage

1. Clone or download the repository.
2. Install Python dependencies.
3. Put raw Libre CSV files in patient folders:

```text
<repo>\
  G0001\
    libre_export.csv
```

4. Run:

```powershell
python program_LIBRE_CGM.py
```

5. The program writes CGM Excel workbooks to:

```text
<repo>\CGM_EXPORT\
```

## Data Dictionary

See `SQL_DATA_DICTIONARY.md` for SQL-facing definitions of glucose rows, visit
summary fields, and `iglu` columns.

## Detailed Documentation

See `LIBRE_CSV_GITHUB_INSTRUKCJA.md`.
