```markdown
# AegisFlow - Generate Batch XML

**Transform, Validate, Deliver from a Single TOC**

## Feature Description / Overview

Generate batch list XML files for Adobe PDF Builder from Excel/CSV input automatically that comply with Adobe PDF Builder format.

Generate batch list XML files for Adobe PDF Builder from Excel/CSV input.

---

## Main Features / Features

1. ✅ Read data from Excel or CSV
2. ✅ **Graphical File Selection** - No need to manually enter file paths, use popup window to select files
3. ✅ Automatically group and sort by section (supports correct numeric sorting, e.g., 14.2.10 comes after 14.2.2)
4. ✅ Generate standard XML structure (all Adobe PDF Builder required elements **fully retained**)
5. ✅ **Preserve UTF-8 encoding** and all fixed content (ruleset, page, font, etc.)
6. ✅ Verify latin1 character compatibility and provide warnings
7. ✅ Support custom header, file location, output path
8. ✅ Interactive user interface, automatically locates to related folders

**Detailed Information**:
- XML structure rules: [XML_STRUCTURE_RULES.md](XML_STRUCTURE_RULES.md)
- Quick start: [QUICK_START_GENERATE_XML.md](QUICK_START_GENERATE_XML.md)

---

## Usage / Usage

### Method 1: Double-click batch file (Recommended)

1. Double-click `run_generate_batch_xml.bat`
2. Follow prompts to enter information
3. Wait for generation to complete

### Method 2: Run from command line

```bash
python generate_batch_xml.py
```

---

## Input File Requirements / Input Requirements

### Required Columns / Required Columns

Excel/CSV files must contain the following columns:

| Column Name | Description | Example |
|------|------|------|
| `sect_num` | Section number | 14.1, 14.2.1, 14.2.10 |
| `sect_ttl` | Section title | Study Population |
| `OUTFILE` | Output filename (without extension) | t_ds_comb |
| `Output Type (Table, Listing, Figure)` | Output type | Table, Figure, Listing |
| `tocnumber` | TOC number | 14.1.1, 14.2.1.1 |
| `Title` | Title | Disposition |

### File Format Support

- ✅ Excel files: `.xlsx`, `.xls`
- ✅ CSV files: `.csv` (supports multiple encodings: UTF-8, GBK, GB2312, latin1)

---

## Interactive Input Prompts / Interactive Prompts

After running the tool, you will be prompted sequentially:

### 1. Select Input File
```
Step 1/6: Select Input File
Please select Excel or CSV file in the popup window...
```
- File selection dialog will appear
- Default opens `02_output` directory
- Supports filtering for .xlsx, .xls, .csv files

### 2. Header Text (Header Text)
```
Step 2/6: Set Header Text
This will be used for: <header text="..."> and <document-heading text="...">
Enter header text [Default: AZD0901 CSR DR2 Batch 1 Listings]:
```
- Press Enter to use default value
- Or enter custom text, e.g.: `My Study CSR Batch 1 Tables`
- **Used for**: XML header display and document heading

### 3. Output Filename (Output Filename) - ⭐ NEW
```
Step 3/6: Set Output PDF Filename
IMPORTANT: Filename cannot contain spaces (use '_' instead)
This will be used for: <output-pdf> and <output-audit>
Enter output filename [Default: AZD0901_CSR_DR2_Batch_1_Listings]:
```
- **⚠️ IMPORTANT**: No spaces allowed in filename
- Use underscore `_` instead of spaces
- Press Enter to use default (auto-converts spaces to `_`)
- Or enter custom filename, e.g.: `Study_DR2_Tables_Batch1`
- **Used for**: PDF and audit file names in XML
  - `<output-pdf filename="YOUR_FILENAME.pdf">`
  - `<output-audit filename="YOUR_FILENAME_audit.pdf">`

**Validation**:
- ❌ `My File Name` → Error (contains spaces)
- ✅ `My_File_Name` → Valid
- ✅ `Study-DR2-Batch1` → Valid (hyphens OK)
- ✅ `AZD0901_CSR_DR2` → Valid

### 4. File Location (File Location)
```
Step 4/6: Set File Location
Enter file location [Default: root/cdar/d980/d9802c00001/ar/dr2/tlf/dev/output/]:
```
- Press Enter to use default value
- Or enter custom path

### 5. Output XML Path (Output XML Path)
```
Step 5/6: Select Output Location
Please select save location and filename in the popup window...
```
- File save dialog will appear
- Default opens `03_xml` directory
- Default filename with timestamp: `_batch_list_20260211_143025.xml`
- Can modify filename and save location

### 6. Starting Page Number (Starting Page Number)
```
Step 6/6: Set Starting Page Number
Enter starting page number [Default: 2]:
```
- Press Enter to use default value 2
- Or enter other number

---

## Output Format / Output Format

Example of generated XML file format:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<pdf-builder-metadata>
<!-- input files total to less than 100MB -->
    <ruleset>
        <headers>
            <header text="AZD0901 CSR DR2 Batch 1 Listings" startNumber="2" />
        </headers>
        <page
            orientation="landscape"
            size="letter"
            measurementUnit="in"
            marginTop="           0"
            marginLeft="           0"
            marginRight="           0"
            marginBottom="           0" />
        <font fontName="CourierNew" style="normal" size="9" />
        <!-- <character-encoding type="ascii" /> -->
        <document-heading text="AZD0901 CSR DR2 Batch 1 Listings" fontName="Times New Roman" />
    </ruleset>
    <sectionset>
        <section name="14.1 Study Population">
            <source-file filename="t_ds_comb.rtf" fileLocation="root/cdar/d980/d9802c00001/ar/dr2/tlf/dev/output/" number="Table 14.1.1" title="Disposition" />
            <source-file filename="t_aztoncsp16_itt.rtf" fileLocation="root/cdar/d980/d9802c00001/ar/dr2/tlf/dev/output/" number="Table 14.1.2" title="Recruitment per region, country/area and site (ITT analysis set)" />
        </section>
        <section name="14.2.1 Primary Endpoint - Progression Free Survival in ITT">
            ...
        </section>
    </sectionset>
    <output-pdf filename="AZD0901_CSR_DR2_Batch_1_Listings.pdf">
        <pdf-import path="root/cdar/d980/d9802c00001/ar/dr2/tlf/doc/" />
    </output-pdf>
    <output-audit filename="AZD0901_CSR_DR2_Batch_1_Listings_audit.pdf">
        <audit-import path="root/cdar/d980/d9802c00001/ar/dr2/tlf/doc/" />
    </output-audit>
</pdf-builder-metadata>
```

**Note**: Except for the following user-defined content, all other content (such as ruleset structure, page, font, etc.) is fixed:
- `header text` and `document-heading text`: User-defined (Step 2)
- `header startNumber`: User-defined (Step 6, default 2)
- `output-pdf filename`: User-defined (Step 3) - **NEW**
- `output-audit filename`: User-defined (Step 3) + "_audit" - **NEW**
- `pdf-import` and `audit-import` path: Automatically extracted from file location
- `section name`: Concatenated from Excel's sect_num and sect_ttl columns
- `source-file filename`: From Excel's OUTFILE column + ".rtf"
- `source-file fileLocation`: User-defined (Step 4)
- `source-file number`: Concatenated from Excel's "Output Type" and "tocnumber" columns
- `source-file title`: From Excel's Title column

---

## Special Features / Special Features

### 1. Intelligent Numeric Sorting

The tool correctly handles numeric sorting of section numbers:
- ✅ 14.1, 14.2, 14.10, 14.20 (Correct)
- ❌ Will not happen: 14.1, 14.10, 14.2, 14.20 (Incorrect string sorting)

### 2. Latin1 Character Check

Before generating XML, the tool checks whether all text contains only latin1 characters:

- ✅ If all compatible, generate directly
- ⚠ If non-latin1 characters found:
  - List all problem locations
  - Show specific problem characters
  - Ask whether to continue

**Common non-latin1 characters**:
- Chinese characters: such as "分析"
- Special symbols: ≥, ≤, —, ±, °, μ, etc.
- Non-ASCII quotes: " " ' '

**Solutions**:
- Replace Chinese with English
- Replace special symbols with ASCII equivalents:
  - `≥` → `>=`
  - `≤` → `<=`
  - `—` → `-`
  - `±` → `+/-`

### 3. Data Validation

The tool automatically validates:
- ✅ Whether required columns exist
- ✅ Whether file is readable
- ✅ Data integrity
- ✅ Character encoding compatibility

---

## FAQ / FAQ

### Q0: The bat file flashes and closes, what should I do?

**A**: This usually indicates that some dependent packages are not installed. Please diagnose following these steps:

1. **Run diagnostic tool**: Double-click `test_environment.bat`
2. **Check which modules are missing**
3. **Install missing packages**: 
   ```bash
   pip install pandas openpyxl
   ```
4. **Re-run** `run_generate_batch_xml.bat`

**Note**: The enhanced bat file now displays error messages and keeps the window open.

### Q1: Error says "missing required columns", what should I do?

**A**: Check whether the Excel file's column names exactly match (including case and spaces):
- `sect_num`
- `sect_ttl`
- `OUTFILE`
- `Output Type (Table, Listing, Figure)`
- `tocnumber`
- `Title`

### Q2: CSV file read failed, what should I do?

**A**: The tool will automatically try multiple encodings. If it still fails, please:
1. Open CSV in Excel
2. Save as CSV with UTF-8 encoding
3. Re-run the tool

### Q3: Found non-latin1 character warning, how to handle?

**A**: Three options:
1. **Cancel generation**, return to Excel and modify content
2. **Continue generation**, but XML may display abnormally on some systems
3. **Modify then regenerate** (recommended)

### Q4: Section order is incorrect, what should I do?

**A**: Check the `sect_num` column in Excel:
- Ensure format is consistent (all numbers, like 14.1)
- Ensure no extra spaces
- The tool will automatically perform numeric sorting

### Q5: Want to modify XML template, what should I do?

**A**: Edit the `generate_xml()` method in the `generate_batch_xml.py` file.

### Q6: Output filename contains spaces, what should I do? ⭐ NEW

**A**: The tool will automatically validate and reject filenames with spaces:
- ❌ Input contains spaces → Display error and request re-entry
- ✅ Use underscore `_` instead of spaces
- Example:
  ```
  Wrong: My Study DR2
  Correct: My_Study_DR2
  ```
- If you directly press Enter to use the default value, the tool will automatically convert spaces to underscores in the header text

---

## File Structure / File Structure

```
03_mastertoc/
├── generate_batch_xml.py          # Main program
├── run_generate_batch_xml.bat     # Windows batch file
├── test_environment.bat            # Environment test script (for diagnosis)
├── test_imports.py                 # Import test script (for diagnosis)
├── README_GENERATE_BATCH_XML.md   # This document
│
├── 02_output/                     # Input folder (store Excel/CSV)
│   └── 2026-02-11/
│       └── Clinical Study Report_TiFo.csv
│
└── 03_xml/                        # Output folder (store generated XML)
    └── _batch_list_YYYYMMDD_HHMMSS.xml
```

### Diagnostic Tools

**test_environment.bat** - Environment diagnostic tool
- Check Python version
- Check whether all required modules are installed
- If you encounter problems, run this tool first

---

## Technical Details / Technical Details

### Dependencies / Dependencies

```
pandas
openpyxl (for Excel files)
tkinter (file selection dialog, usually comes with Python)
```

### Install Dependencies

```bash
pip install pandas openpyxl
```

**Note**: `tkinter` usually comes with Python. If you encounter `ModuleNotFoundError: No module named 'tkinter'`, please refer to the Python official documentation for installation.

### Python Version Requirements

- Python 3.6+

---

## Changelog / Changelog

### Version 1.1.0 (2026-02-12) ⭐ NEW
- ✅ **New**: Separate Output PDF Filename input step
- ✅ **New**: Filename space validation (spaces not allowed, must use underscores)
- ✅ **Improvement**: Expanded from 5-step process to 6-step process for better clarity
- ✅ **Improvement**: Separated header text and output filename for greater flexibility
- ✅ **Improvement**: Automatically extract base path from file location for pdf-import/audit-import
- ✅ **Improvement**: Full English interface and error messages

### Version 1.0.0 (2026-02-11)
- ✅ Initial version released
- ✅ Support for Excel and CSV input
- ✅ Automatic numeric sorting
- ✅ Latin1 character validation
- ✅ Interactive interface

---

## Author / Author

Generated by GitHub Copilot Assistant

---

## License / License

Internal use only tool / Internal use only

```