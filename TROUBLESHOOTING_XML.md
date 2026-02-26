```markdown
# AegisFlow - XML Troubleshooting Guide

**Transform, Validate, Deliver from a Single TOC**

## 🔧 Issue: Bat file flashes and closes

### Step 1: Run diagnostic tool

Double-click to run `test_environment.bat`

This will check:
- ✅ Whether Python is installed
- ✅ Python version
- ✅ pandas module
- ✅ openpyxl module
- ✅ tkinter module
- ✅ xml module

### Step 2: Fix based on diagnostic results

**If pandas is shown as not installed:**
```bash
pip install pandas
```

**If openpyxl is shown as not installed:**
```bash
pip install openpyxl
```

**If tkinter is shown as not installed:**
- Windows users: Reinstall Python and make sure to check "tcl/tk and IDLE" option
- Or: The program will automatically switch to command line mode (manually enter path)

**Install all dependencies at once:**
```bash
pip install pandas openpyxl
```

### Step 3: Re-run

After fixing, double-click `run_generate_batch_xml.bat` again

---

## 🔧 Issue: Error "No file selected"

**Cause**: Clicked "Cancel" in the file selection window

**Solution**: Re-run the program and select file in the window

---

## 🔧 Issue: Cannot find file selection window

**Possible causes**:
1. Window is behind other windows → Check taskbar
2. tkinter is not installed → Run diagnostic tool to check
3. Window is on another monitor → Check all screens

**Backup plan**:
- The program will automatically detect if tkinter is available
- If not available, will switch to command line input mode
- Simply enter file path manually in the command line

---

## 🔧 Issue: Virtual environment not found

**Symptom**: Bat file shows "Virtual environment does not exist, using system Python..."

**This is normal** if:
- System Python has all required packages installed
- Program can run normally

**If you want to use virtual environment**:
1. Open command line in project directory
2. Run: `py -3.13 -m venv .venv`
3. Activate: `.venv\Scripts\activate`
4. Install dependencies: `pip install pandas openpyxl`

---

## 🔧 Issue: Excel file read failed

**Possible causes and solutions**:

1. **File is open in another program**
   - Close Excel or other programs using this file

2. **File path contains special characters**
   - Rename file to avoid special characters

3. **File format is incorrect**
   - Make sure it's .xlsx, .xls, or .csv format

4. **File encoding issue (CSV)**
   - Program will automatically try multiple encodings
   - If still fails, open in Excel then save as UTF-8 encoded CSV

---

## 🔧 Issue: Column names do not match

**Error message**: "Missing required columns: xxx"

**Solution**:
1. Open Excel file
2. Check whether column names exactly match (including case, spaces):
   - `sect_num`
   - `sect_ttl`
   - `OUTFILE`
   - `Output Type (Table, Listing, Figure)`
   - `tocnumber`
   - `Title`

3. If column names are different, modify Excel file's column names

---

## 📞 Still cannot resolve?

1. **Check full error message**
   - Enhanced bat file displays detailed errors
   - Take screenshot and send to technical support

2. **Check full documentation**
   - [README_GENERATE_BATCH_XML.md](README_GENERATE_BATCH_XML.md)
   - [QUICK_START_GENERATE_XML.md](QUICK_START_GENERATE_XML.md)

3. **Manually run Python script**
   ```bash
   python generate_batch_xml.py
   ```
   View detailed error messages

---

## ✅ Quick Verification Checklist

Verify before running:

- [ ] Python is installed (recommend 3.6+)
- [ ] pandas is installed
- [ ] openpyxl is installed
- [ ] Excel/CSV file is ready
- [ ] Excel file contains all required columns
- [ ] Excel file is not open in other programs

**After all verified, run**: `run_generate_batch_xml.bat`

```