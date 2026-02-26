# AegisFlow

**Transform, Validate, Deliver from a Single TOC**

## Overview

This project turns a single TOC-style source file into a full clinical output pipeline, including conversion, template filling, status merge, and XML delivery.

The toolkit contains five main Python scripts:

1. **mosaic_convert.py** - MOSAIC data conversion script (completed, ready to use)
2. **extract_programs.py** - SAS program list extraction script (new, extracts program list from Excel)
3. **fill_tlf_template.py** - TLF template filling script (new, integrates MOSAIC output with personnel data)
4. **fill_tlf_status.py** - TLF status filling script (new, merges QC status data)
5. **generate_batch_xml.py** - Batch List XML generation script (new, generates Adobe PDF Builder format XML from Excel)

## System Requirements

- Python 3.13 (recommended)
- Virtual environment configured (.venv)
- Required Python packages: pandas, openpyxl, tkinter (usually included by default)

### Python 3.13 Environment Setup (Windows)

If old Python versions were removed, recreate `.venv` with Python 3.13:

```bash
py -3.13 -m venv .venv
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install pandas openpyxl
```

## File Locations

All working files should be placed in the following directory:
```
c:\Users\kplp794\OneDrive - AZCollaboration\Desktop\roooooot\00-工具开发\credit_latest\03_mastertoc\
```

Output files are stored in:
```
02_output\2026-02-09\
```

## Script Descriptions

### 1. mosaic_convert.py - MOSAIC Data Conversion

**Features:**
- Convert Clinical Study Report_TiFo.csv to Excel format
- Execute data preprocessing (clean quotes, formatting)
- Add Excel formatting (fonts, colors, rich text)
- Support deduplication by tocnumber for unique outputs
- Correctly handle different tocnumbers with same program+suffix combination
- Build OUTFILE column from PROGRAM and SUFFIX (format: PROGRAM_SUFFIX)

**Usage:**
```bash
.venv\Scripts\python.exe mosaic_convert.py
```

**Input Files:**
- Clinical Study Report_TiFo.csv

**Output Files:**
- DXXXXXXXXXX_TiFo_MOSAIC_CONVERT.xlsx

**Related Documentation:**
- README_MOSAIC_CONVERT.md
- MOSAIC_CONVERT_FIX.md

---

### 2. extract_programs.py - SAS Program List Extraction (New)

**Features:**
- Extract PROGRAM list from MOSAIC_CONVERT generated Excel file
- Count tocnumber quantity for each program
- Order by first appearance in Excel
- Comments unified at the top of file
- Generate standard %runpgm format SAS script

**Usage:**
```bash
run_extract_programs.bat
```
or
```bash
.venv\Scripts\python.exe extract_programs.py
```

**Input Files:**
- MOSAIC_CONVERT generated Excel file (.xlsx)

**Output Files:**
- SAS script file (.txt, %runpgm format)

**Output Format Example:**
```sas
/* Generated SAS Program Execution Script */
/* Programs ordered by first appearance in Excel file */

/* Program Statistics: */
/*   t_ds: 3 table(s) */
/*   t_dm: 2 table(s) */

/*