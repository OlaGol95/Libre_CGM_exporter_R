# LIBRE CGM and `iglu` Export Pipeline

## Description

This repository contains a portable LIBRE-only CGM export tool. The application
reads raw Libre CSV exports, maps glucose measurements to predefined visit
windows, and creates per-patient Excel workbooks with Python summary metrics and
optional R `iglu` metrics.

## Repository Contents

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

Python:

```powershell
pip install -r requirements_libre_cgm.txt
```

Optional R dependency for extended CGM metrics:

```powershell
Rscript install_iglu.R
```

When R or `iglu` is unavailable, the exporter still writes Python-derived CGM
summary metrics and records the `iglu` status in the output.

## Configuration

Optional local configuration:

```powershell
$env:LIBRE_DIR = "<path-to>\LIBRE"
$env:LIBRE_CODES_XLSX = "<path-to>\kody.xlsx"
```

If `LIBRE_DIR` is not set, the repository directory is used as the LIBRE
workspace.

## Input Layout

Portable repository layout:

```text
<repo>\
  program_LIBRE_CGM.py
  data_visits_protected.zip
  requirements_libre_cgm.txt
  install_iglu.R
  SQL_DATA_DICTIONARY.md
  G0001\
    *.csv
  G0002\
    *.csv
```

Patient folder IDs use the `G0001`, `G0002`, ... format. Raw Libre CSV exports
are placed directly inside the matching patient folder.

Supported visit labels:

```text
V1, V3, V4, V6
```

## Usage

Run the exporter:

```powershell
python program_LIBRE_CGM.py
```

For a selected workspace:

```powershell
python program_LIBRE_CGM.py --libre-dir "<path-to>\LIBRE"
```

For selected patients:

```powershell
python program_LIBRE_CGM.py --patient G0001 --patient G0002
```

The program writes CGM Excel workbooks to:

```text
<repo>\CGM_EXPORT\
```

## CSV Parsing

Libre CSV parsing attempts:

- encodings: `utf-8-sig`, `utf-8`, `cp1250`, `latin1`
- separators: autodetected delimiter, then `,`, `;`, tab
- header rows: `2`, `0`, `1`, `3`

Timestamp detection prefers Libre timestamp columns equivalent to
`Znacznik czasu w urzadzeniu`. If no named timestamp column is detected, the
third column is used as a fallback.

Glucose detection prefers `mg/dL` columns in this order:

1. historical glucose
2. scan glucose
3. strip/manual glucose
4. last numeric column fallback if no glucose header is available

## Visit Window Extraction

Visit dates are read from the protected visit-date source using `openpyxl`.

Supported date cell forms:

- native Excel dates
- ISO-like date strings
- day-first date strings

End dates that contain only a date are interpreted as the full day ending at
`23:59:59`.

## CGM Workbook Contract

Output workbook path:

```text
<repo>\CGM_EXPORT\<patient_export_code>\<patient_export_code>_CGM.xlsx
```

Worksheets:

- `v1`
- `v3`
- `v4`
- `v6`
- `summary`

Visit worksheet columns:

```text
patient_id
visit
timestamp
glucose
source_file
```

Rows are selected by visit window from all available raw Libre CSV files. The
pipeline sorts rows by timestamp and removes duplicates by timestamp, glucose,
and source file.

## Python CGM Summary Metrics

The `summary` worksheet contains per-visit Python metrics:

- `n`
- `start`
- `end`
- `mean_glucose`
- `median_glucose`
- `sd_glucose`
- `min_glucose`
- `max_glucose`
- `pct_below_54`
- `pct_below_70`
- `pct_70_180`
- `pct_above_180`
- `pct_above_250`

## Extended `iglu` Metrics

When R and the `iglu` package are available, the program additionally exports
selected `iglu` metric groups, including:

- `Active_percent`
- `ADRR`
- `AUC`
- `COGI`
- `CONGA`
- `CV`
- `eA1c`
- `GMI`
- `GRADE`
- `GVP`
- `HBGI`
- `LBGI`
- `MAGE`
- `MODD`
- `Percent_Above`
- `Percent_Below`
- `Percent_In_Range`
- `SD`

Detailed SQL column definitions are maintained in `SQL_DATA_DICTIONARY.md`.

## Files To Upload To GitHub

Upload these files:

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
