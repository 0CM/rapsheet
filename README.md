# README.md

## Rapsheet

**Rapsheet** is a command-line tool that converts CSV files into a formatted Excel report. It supports importing individual files or directories, applying a template, updating existing Excel files, and fixes encoding issues automatically.

---

## Features

- Import one or many CSV files
- Automatically apply BOM fix for UTF-8 compatibility
- Generate Excel sheets per CSV with proper naming
- Freeze headers and apply text wrapping for long cells
- Set appropriate column widths with a max cap
- Use a custom Excel file as a template
- Avoid re-importing data when updating existing reports

---

## Usage

```bash
rapsheet -f file1.csv file2.csv -o ./output
rapsheet -d ./csv_dir -o ./output
rapsheet -f data.csv -t template.xlsx -o ./output
rapsheet -d ./csv_dir -u -o ./output
```

---

## Command-Line Arguments

| Flag       | Description                                                                 |
|------------|-----------------------------------------------------------------------------|
| `-f`       | List of CSV files to process                                                |
| `-d`       | Directory containing CSV files                                              |
| `-o`       | Output directory where the Excel report will be saved (**required**)       |
| `-t`       | (Optional) Path to a template `.xlsx` file to use as a base                |
| `-u`       | (Optional) Update mode: auto-detect `.xlsx` in directory or file list      |

---

## BOM Fix

Many CSV exports lack a UTF-8 BOM, which can cause encoding issues in Excel.  
RapsheetCLI automatically detects and prepends a BOM to ensure proper encoding.

---

## Text Wrapping & Column Widths

- Cells with content longer than **80 characters** are wrapped.
- Column widths are automatically adjusted based on content.
- Empty columns are minimized in width to keep layout clean.

---

## Tips

- Use `-u` when reprocessing a folder that already has an Excel file in it.
- Template support makes it easy to maintain branding or formatting.
- BOM handling ensures compatibility across Excel versions and OSes.

---

## Requirements

- Python 3.6+
- pandas
- openpyxl

Install with:

```bash
pip install -r requirements.txt
```

---

##  License

MIT
""
