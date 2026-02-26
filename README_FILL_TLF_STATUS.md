# AegisFlow - Fill TLF Status

**Transform, Validate, Deliver from a Single TOC**

## 📋 Overview

`fill_tlf_status.py` is an automation tool designed to merge Comparison Status data from TLF Status files into the QC Status column of People Management files.

## 🎯 Key Features

### Core Features

1. **File Selection**
   - Select modified people_management.xlsx file
   - Select tfl_status.xlsx file

2. **Status Preprocessing**
   - Automatically convert "Match" to "Pass"
   - Automatically convert "Mismatch" to "Fail"

3. **Precise Matching and Merging**
   - Perform exact matching based on `Dataset` (tfl_status) and `Output Name` (people_management)
   - Only merge data when exact matches are found
   - Leave QC Status column empty for unmatched rows

4. **Structure Preservation**
   - Retain all sheets in people_management
   - Preserve all original columns
   - Update only the QC Status column; leave all other columns unchanged

5. **Statistical Report**
   - Total number of TLFs
   - Number of Status "Pass"
   - Number of Status "Fail"
   - Number of empty Status values
   - Matching rate percentage

## 📁 File Structure

### Input Files

1. **people_management.xlsx**
   - Required sheet: `TLF`
   - Required column: `Output Name`
   - Target column: `QC Status (Not Started, Ongoing, QC Pending, Fail, Pass)`

2. **tfl_status.xlsx**
   - Required sheet: `Overview`
   - Required columns: `Dataset`, `Comparison Status`

### Output Files

- Default filename: `people_management_with_status.xlsx`
- Users can customize the filename and save path

## 🚀 Usage

### Method 1: Using Batch File (Recommended)

```bash
run_fill_tlf_status.bat
```

### Method 2: Run Python Script Directly

```bash
py -3.13 fill_tlf_status.py
```

## 📊 Workflow

```
Input 1: people_management.xlsx (TLF sheet)
Input 2: tfl_status.xlsx (Overview sheet)
    ↓
[Step 1] Read people_management file
[Step 2] Read tfl_status file
[Step 3] Preprocess Comparison Status
         - Match → Pass
         - Mismatch → Fail
    ↓
[Step 4] Perform exact matching based on Dataset and Output Name
         - Match successful: Populate QC Status
         - Match failed: Set QC Status to empty
    ↓
[Step 5] Update Excel file (QC Status column only)
[Step 6] User selects output path and filename
    ↓
Output: people_management_with_status.xlsx
Statistics: Total/Pass count/Fail count/Empty count/Match rate
```

## 📝 Column Mapping

| Source File | Source Column | Target File | Target Column | Operation |
|-------------|---------------|-------------|---------------|-----------|
| tfl_status.xlsx | Dataset | people_management.xlsx | Output Name | Matching key |
| tfl_status.xlsx | Comparison Status | people_management.xlsx | QC Status | Merged value (post-processing) |

### Value Conversion Rules

| Original Value (tfl_status) | Converted Value (people_management) |
|-----------------------------|-------------------------------------|
| Match | Pass |
| Mismatch | Fail |
| (Other values) | (Keep as is) |

## ⚠️ Important Notes

### File Requirements

1. **people_management.xlsx**
   - Must contain `TLF` sheet
   - First row of TLF sheet is title, second row is column names
   - Must contain `Output Name` column
   - If `QC Status (Not Started, Ongoing, QC Pending, Fail, Pass)` column does not exist, the script will create it automatically

2. **tfl_status.xlsx**
   - Must contain `Overview` sheet
   - Must contain `Dataset` and `Comparison Status` columns

### Pre-Execution Checklist

- [ ] Ensure input files are not open in Excel
- [ ] Confirm file paths are correct
- [ ] Confirm Python 3.13 is available (or `.venv` is properly created)
- [ ] Ensure sufficient disk space

### Common Errors

#### Error 1: Permission denied
```
❌ Unable to read people_management file (file may be open in Excel)
```
**Solution**: Close the file in Excel and run the script again

#### Error 2: Sheet not found
```
❌ Error: 'TLF' sheet not found in people_management file
```
**Solution**: Verify that people_management.xlsx contains a TLF sheet

#### Error 3: Column not found
```
❌ Error: 'Dataset' column not found in Overview sheet of tfl_status file
```
**Solution**: Verify that tfl_status.xlsx Overview sheet contains the required columns

## 📈 Output Example

After successful execution, the following statistics will be displayed:

```
================================================================================
✓✓✓ Fill Complete!
================================================================================
Output file: C:\path\to\people_management_with_status.xlsx

Statistics:
  - Total TLF count: 249
  - Status 'Pass' count: 230
  - Status 'Fail' count: 15
  - Empty Status count: 4
  - Match rate: 245/249 (98.4%)

Note: You can open the Excel file directly to view results
      All other columns and sheets have been preserved without modification
```

## 🔧 Technical Details

### File Locking Handling

- Automatic retry mechanism (3 attempts, 1-second interval between attempts)
- User-friendly error messages

### Data Integrity

- Use pandas for data reading and processing
- Use openpyxl to preserve Excel format and structure
- Update only the target column; all other data is fully preserved

### Performance Metrics

- Read people_management: < 2 seconds
- Read tfl_status: < 1 second
- Data processing and merging: < 2 seconds
- File save: < 2 seconds

## 🆘 Support and Maintenance

If you encounter issues, please check:

1. Python version ≥ 3.7
   - Recommended: Python 3.13
2. Required packages are installed (pandas, openpyxl)
3. File format is correct
4. Files are not open in other programs
5. Virtual environment is activated

## 📚 Related Documentation

- [README.md](README.md) - Project Overview
- [QUICK_START.md](QUICK_START.md) - Quick Start Guide
- [fill_tlf_template.py](fill_tlf_template.py) - TLF Template Fill Script

---

**Created**: February 11, 2026  
**Version**: 1.0  
**Status**: ✅ Production Ready
**状态**: ✅ 生产就绪
