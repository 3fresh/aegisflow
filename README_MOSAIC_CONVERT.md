# AegisFlow - MOSAIC Convert

**Transform, Validate, Deliver from a Single TOC**

## Feature Description

Convert VBA macro `MOSAIC_CONVERT` to Python script for processing TiFo CSV files and outputting as Excel format (.xlsx).

## Main Features

This script implements all functionality of the original VBA macro and adds multiple enhancements:

1. **tocnumber Uniqueness Check**: Verify no two tables use the same toc number. If duplicates detected, report error and stop
2. **seq Sequence Generation**: Automatically generate seq from CSV by order of appearance of outfile rows (starting from 1, increment by 1 with each outfile), used for transposition
3. **Transposition by seq**: Convert each seq (corresponding to multiple rows of one chart) to a single row of data
4. **Duplicate program+suffix Detection**: After transposition, check if multiple tables use the same program+suffix combination. If detected, warn and highlight in yellow in output
5. **Output Type Recognition**: Automatically identify Table, Figure, or Listing type
6. **Title Extraction**: Extract clean title text from title5 field
7. **CSV Preprocessing**: Only replace first and last single quotes with double quotes in footnote values (keep title unchanged)
8. **Excel Formatting**: 
   - All fonts using "等线" series
   - Create Index and Original worksheets
   - Header row in bold
   - Columns H, I, J with yellow background
   - H2 cell freeze panes
   - Non-latin1 characters in red font + green background (do not convert original text content)
   - Duplicate program+suffix PROGRAM/SUFFIX columns in yellow
   - Correct numeric sorting (14.1.2 before 14.1.10)

## Key Improvements

### v3.1 - OUTFILE Correction and Chinese Path Support
- ✅ **OUTFILE Data Source Correction**: Directly use `value` from CSV rows with `parm='outfile'`, no longer concatenate from PROGRAM+SUFFIX
- ✅ **Chinese Path Support**: Completely resolve path issues with Chinese characters
  - `.last_output.txt` uses UTF-8 encoding
  - `validate_output.py` automatically reads path file
  - `run_mosaic_convert.bat` simplifies validation workflow
- ✅ **Business Logic Faithfulness**: Preserve actual business logic from CSV (OUTFILE may equal PROGRAM only even when SUFFIX is not empty)

### v3.0 - Transposition by seq and Complete Check Mechanism
- ✅ New tocnumber uniqueness check (stop if duplicates found)
- ✅ Use seq sequence for transposition (based on number of outfile rows)
- ✅ Detect duplicate program+suffix after transposition and warn
- ✅ Highlight duplicate program+suffix rows in yellow in output
- ✅ Non-latin1 characters colored red + marked green, original text content not converted
- ✅ Data row count of 249 (unique row count from seq transposition)

### v2.5 - Numeric Sorting and Rich Text Formatting
- ✅ Fix tocnumber sorting (change from string sort to numeric sort)
- ✅ Implement rich text formatting for non-latin1 characters (color only problem characters red)
- ✅ Green background marking for cells with non-latin1 characters
- ✅ Preserve yellow background for duplicate program+suffix marking

### v2.4 - Complete Data Recovery and Shared Marking
- ✅ Recover from 210 rows to 245 rows (all tables included)
- ✅ Use tocnumber as unique identifier
- ✅ Automatically filter template footnote placeholders
- ✅ Mark 31 shared program+suffix combinations (affecting 66 tocnumber values)

### v2.3 - Footnote Preprocessing
- ✅ Only perform quote replacement on footnote fields
- ✅ Keep quotes in title fields unchanged

### v2.2 - File Selection Dialog
- ✅ Dynamic input file selection
- ✅ Dynamic output file save location

## Usage

### Method 1: Using Batch File (Recommended)

```bash
run_mosaic_convert.bat
```

Batch file will:
1. Automatically open file selection dialog to choose input CSV file
2. Automatically generate output XLSX file
3. Automatically validate output file (249 rows of data)
4. Fully support paths with Chinese characters

### Method 2: Run Python Script Directly

```bash
python mosaic_convert.py
```

Script opens graphical file selection dialog:
1. Select input CSV file
2. Select output XLSX file location and name
3. Automatically process and generate result

### Method 3: Call in Code

```python
from mosaic_convert import mosaic_convert

# Specify input and output files
input_file = "path/to/your/input.csv"
output_file = "path/to/your/output.xlsx"

mosaic_convert(input_file, output_file)
```

### Method 4: Edit Default Path in Script

Edit the end of `mosaic_convert.py` file:

```python
input_file = os.path.join(script_dir, "02_output", "2026-02-09", "Clinical Study Report_TiFo.csv")
```

## Dependencies

```bash
pip install pandas openpyxl
```

Or using virtual environment:
```bash
py -3.13 -m venv .venv
.venv\Scripts\activate
pip install pandas openpyxl
```

## Input File Format

CSV file should contain the following 6 columns:
- `sect_num`: Section number
- `sect_ttl`: Section title
- `program`: Program name
- `suffix`: Suffix
- `parm`: Parameter name (such as outfile, outtype, title1, footnote1, etc.)
- `value`: Parameter value

## Output File Structure

Generated Excel file contains two worksheets:

### Index Worksheet (Main Result)
Contains the following columns:
- sect_num, sect_ttl, outtype, azsolid, Core, tocnumber
- Output Type (Table, Listing, Figure), Title
- PROGRAM, SUFFIX, OUTFILE
- title1-7, footnote1-9

**Formatting Features**:
- **All Content**: Using "等线" font
- **Header Row**: Bold display
- **Columns H-J**: Yellow background
- **Non-latin1 Characters**: Red font + green background
- **Shared program+suffix**: tocnumber/PROGRAM/SUFFIX columns yellow background
- **Sorting**: Numeric sort by tocnumber (14.1.2 before 14.1.10)

### Original Worksheet
Preserve original CSV data's 6-column structure

## Data Characteristics Description

### Transposition Logic
- Based on 'outfile' rows in CSV to define seq (increment seq by 1 at each outfile)
- Transpose by seq to generate final data framework
- Total of **249 unique tables** (based on unique row count from seq)
- 8 rows with duplicate program+suffix combinations (highlighted yellow in output as warning)

### Example
| Combination | Tocnumber Quantity | Example Tocnumber |
|-----|-------------|-------------|
| t_dm + dischar_itt | 3 | 14.1.9.1, 14.1.9.2, 14.1.9.3 |
| f_km + pfs_itt | 4 | 14.2.1.3.1-14.2.1.3.4 |

### Non-latin1 Character Handling
- Automatically detect cells containing non-ASCII characters (such as "–" dash)
- Only color problem characters red (preserve original text content)
- Mark entire cell with green background for easy identification
- Characters colored red but will not be converted or deleted

## Sorting Algorithm

Use numeric sorting instead of string sorting:
```
Correct order: 14.1.1 → 14.1.2 → 14.1.10 → 14.2
String order: 14.1.1 → 14.1.10 → 14.1.2 → 14.2  (Incorrect)
```

Implementation logic:
1. Split each tocnumber by "."
2. Convert each part to integer
3. Sort by tuple numeric comparison

## Output File Naming

- If output filename not specified, default generation in same directory as input file
- File naming format: `<original_filename>_MOSAIC_CONVERT.xlsx`

Example:
- Input: `Clinical Study Report_TiFo.csv`
- Output: `Clinical Study Report_TiFo_MOSAIC_CONVERT.xlsx`

## Comparison with VBA Macro

| Feature | VBA Macro | Python Script | Description |
|------|-------|-----------|------|
| Input Format | CSV (needs manual import to Excel) | CSV | Read CSV directly |
| Output Format | xlsm (macro file) | xlsx | Standard Excel file |
| Execution Speed | Slower | Fast | Python processing more efficient |
| Dependencies | Excel application | Python + pandas | No Excel installation required |
| Automation | Requires manual clicking | Command line execution | Easy to automate |

## Precautions

1. Ensure CSV file encoding is UTF-8 (if contains Chinese characters)
2. Output directory must have write permissions
3. If output file already exists, it will be overwritten
4. Script will output processed row count and generated index row count

## Error Handling

If script runs into error, please check:
- [ ] CSV file path is correct
- [ ] CSV file format meets requirements (6 columns)
- [ ] Required Python packages are installed
- [ ] Output directory exists and is writable

## Example Output

```
✓ Conversion completed! Output file: Clinical Study Report_TiFo_MOSAIC_CONVERT.xlsx
  - Processed 3211 rows of original data
  - Generated 210 rows of index data

Success! Please see output file: Clinical Study Report_TiFo_MOSAIC_CONVERT.xlsx
```

## Technical Details

- Use pandas for data transformation and grouping operations
- Use openpyxl for Excel formatting
- Preserved all logic steps of VBA macro
- Optimized data search and matching algorithms

## Error Handling

### Stop Conditions

1. **tocnumber Duplicate Error**
   ```
   ERROR: Input data contains duplicate toc numbers
   ```
   - Program stops immediately, no output file generated
   - Need to fix duplicate tocnumber in CSV

2. **seq Sequence Generation Failure**
   ```
   ERROR: Missing required 'outfile' rows to build seq
   ```
   - CSV missing outfile rows
   - Need to check CSV structure

### Warnings

**program+suffix Duplicate Warning**
```
ERROR: Input data contains duplicate program+suffix, please update MOSAIC
```
- Program continues running, generates output file
- Duplicate rows marked in yellow in PROGRAM and SUFFIX columns
- Need to update program or suffix definitions in MOSAIC macro

## Related Tools

- **extract_programs.py**: Extract PROGRAM list from generated Excel file, generate SAS script (see README_EXTRACT_PROGRAMS.md)
- **fill_tlf_template.py**: Merge MOSAIC output with personnel data (see README.md)
- **validate_output.py**: Validate Excel file structure
- **run_mosaic_convert.bat**: Run MOSAIC_CONVERT conversion

## Version Information

- Python Version: 3.7+
- pandas Version: 1.0+
- openpyxl Version: 3.0+

---

**Author**: AI optimization from VBA macro  
**Date**: 2026-02-11  
**Version**: v3.0 (seq transposition + complete check)  
**Purpose**: TiFo data processing and MOSAIC format conversion