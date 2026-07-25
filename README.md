# Excel Report Generator

[![CI](https://github.com/Synth88Labs/excel-tool-reportgen/actions/workflows/ci.yml/badge.svg)](https://github.com/Synth88Labs/excel-tool-reportgen/actions/workflows/ci.yml)
[![Python 3.9+](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

Turn a **raw CSV or Excel export into a polished, formatted report** — in one command.
Styled header, frozen header row, auto-sized columns, a totals row with live `SUM`
formulas, conditional formatting, and an optional chart. Same result every time, no
manual formatting.

Built for the report you run every week or month: *"take this raw data and make it
presentable."*

> 📖 **New here?** Read the full step-by-step guide:
> [How to Auto-Generate Formatted Excel Reports From Raw Data](https://excelguru.io/tutorials/auto-generate-formatted-excel-reports-from-raw-data/) on ExcelGuru.io.

## Features

- 🎨 **Styled header** — green fill, bold white text, centered
- 📌 **Frozen header row** so it stays visible while scrolling
- ➕ **Totals row** with real `=SUM()` formulas for every numeric column
- 🔴 **Conditional formatting** — negative values in red, high values in green
- 📊 **Optional bar chart** built natively into the sheet
- 📏 **Auto-sized columns** that fit the content
- 🧩 Reads **`.csv`, `.xlsx` and `.xls`**, writes a clean `.xlsx`

## Installation

```bash
git clone https://github.com/Synth88Labs/excel-tool-reportgen.git
cd excel-tool-reportgen
pip install -r requirements.txt
```

Requires Python 3.9+.

## Usage

```bash
python make_report.py <input> [-o report.xlsx] [--chart COLUMN] [--title TEXT]
```

### Quick start (try it on the included sample data)

```bash
python make_report.py sample_data/sales_raw.csv -o march_report.xlsx --chart amount --title "March 2026 Sales Report"
```

### Options

| Option | Description |
|---|---|
| `-o`, `--output` | Output `.xlsx` path. Default: `report.xlsx` |
| `--chart COLUMN` | Numeric column to render as a bar chart |
| `--title TEXT` | Title shown at the top of the report |
| `--highlight-min N` | Highlight numeric values ≥ N in green (default: `2000`) |
| `--sheet NAME` | For Excel input, read a specific sheet (default: first) |

### What you get

From a flat file like this:

```
region,salesperson,product,units,amount
West,Alice,Widget A,12,2400
East,Bob,Widget C,-2,-440
```

…you get a formatted workbook with a styled header, a `TOTAL` row of live formulas,
Bob's negative refund row highlighted red, high-value sales highlighted green, and a
bar chart of `amount`.

## Pairs perfectly with Excel/CSV Merger

Have a folder of files first? Combine them, then report on the result:

```bash
# 1) Merge many files into one  (see the excel-tool-merger repo)
python merge_excel.py ./monthly_exports -o combined.csv

# 2) Turn the combined data into a formatted report
python make_report.py combined.csv -o report.xlsx --chart amount --title "Annual Sales"
```

→ [Excel/CSV Merger on GitHub](https://github.com/Synth88Labs/excel-tool-merger)

## Running the tests

```bash
pip install pytest
python -m pytest
```

## 📚 Learn More — Free Excel Tutorials

Level up your Excel and automation skills with free, practical guides at
**[ExcelGuru.io](https://excelguru.io/category/tutorials/)** — formulas, automation,
VBA, and more.

👉 **[Browse all free Excel tutorials on ExcelGuru.io »](https://excelguru.io/category/tutorials/)**

## License

MIT — see [LICENSE](LICENSE).
