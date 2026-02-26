```markdown
# AegisFlow - Quick Start for Batch XML

**Transform, Validate, Deliver from a Single TOC**

## 🚀 Get Started in 5 Minutes

**New Feature**: Use graphical window to select files without manual path entry! 🎉

### Step 1: Prepare Excel File

Ensure your Excel/CSV file contains the following columns:

| Required Column | Description | Example |
|--------|------|------|
| sect_num | Section number | 14.1 |
| sect_ttl | Section title | Study Population |
| OUTFILE | File name (without .rtf) | t_ds_comb |
| Output Type (Table, Listing, Figure) | Output type | Table |
| tocnumber | TOC number | 14.1.1 |
| Title | Title | Disposition |

### Step 2: Run Tool

Double-click `run_generate_batch_xml.bat`

### Step 3: Enter Information

Follow prompts (steps with ✅ will pop up windows):

```
1. ✅ Select Excel/CSV file
   - Select file in popup window (default opens 02_output directory)

2. Header text (default: AZD0901 CSR DR2)
   - Press Enter to use default, or enter custom text

3. File location (default: root/cdar/d980/d9802c00001/ar/dr2/tlf/dev/output/)
   - Press Enter to use default, or enter custom path

4. ✅ Select output location and file name
   - Select save location in popup window (default opens 03_xml directory)
   - Default file name with timestamp

5. Starting page number (default: 2)
   - Press Enter to use default, or enter another number
```

### Step 4: Complete

Tool will:
- ✅ Automatically read Excel data
- ✅ Check character encoding
- ✅ Group and sort by section
- ✅ Generate XML file

Output file location: `03_xml/_batch_list_YYYYMMDD_HHMMSS.xml`

---

## 📋 Complete Example

### Input Excel Example

| sect_num | sect_ttl | OUTFILE | Output Type | tocnumber | Title |
|----------|----------|---------|-------------|-----------|-------|
| 14.1 | Study Population | t_ds_comb | Table | 14.1.1 | Disposition |
| 14.1 | Study Population | t_aztoncsp16_itt | Table | 14.1.2 | Recruitment per region |
| 14.2.1 | Primary Endpoint - PFS | t_aztoncef04_pfs_bicr_itt | Table | 14.2.1.1 | PFS by BICR |

### Output XML Example

```xml
<?xml version="1.0" encoding="UTF-8"?>
<pdf-builder-metadata>
<!-- input files total to less than 100MB -->
    <ruleset>
        <headers>
            <header text="AZD0901 CSR DR2" startNumber="2" />
        </headers>
        <page orientation="landscape" size="letter" measurementUnit="in" ... />
        <font fontName="CourierNew" style="normal" size="9" />
        <!-- <character-encoding type="ascii" /> -->
        <document-heading text="AZD0901 CSR DR2" fontName="Times New Roman" />
    </ruleset>
    <sectionset>
        <section name="14.1 Study Population">
            <source-file filename="t_ds_comb.rtf" fileLocation="root/cdar/d980/d9802c00001/ar/dr2/tlf/dev/output/" number="Table 14.1.1" title="Disposition" />
            <source-file filename="t_aztoncsp16_itt.rtf" fileLocation="root/cdar/d980/d9802c00001/ar/dr2/tlf/dev/output/" number="Table 14.1.2" title="Recruitment per region" />
        </section>
        <section name="14.2.1 Primary Endpoint - PFS">
            <source-file filename="t_aztoncef04_pfs_bicr_itt.rtf" fileLocation="root/cdar/d980/d9802c00001/ar/dr2/tlf/dev/output/" number="Table 14.2.1.1" title="PFS by BICR" />
        </section>
    </sectionset>
    <output-pdf filename="CG01_DR2.pdf">
        <pdf-import path="root/cdar/d980/d9802c00001/ar/dr2/tlf/doc/" />
    </output-pdf>
    <output-audit filename="CG01_DR2_audit.pdf">
        <audit-import path="root/cdar/d980/d9802c00001/ar/dr2/tlf/doc/" />
    </output-audit>
</pdf-builder-metadata>
```

---

## ⚠️ Common Issues and Solutions

### Issue 0: Batch file window closes on run

**Symptom**: After double-clicking the batch file, the window disappears immediately

**Troubleshooting Steps**:
1. **Run test script first**: Double-click `test_environment.bat`
   - This will check if Python and all required modules are installed
   - The window will remain open, showing test results

2. **Review test results**:
   - If pandas or openpyxl is not installed, run: `pip install pandas openpyxl`
   - If tkinter is not installed, refer to the solution below

3. **If tkinter is not installed**:
   - Windows: tkinter usually comes with Python; reinstall Python and check the "tcl/tk" option
   - Alternatively, the program will automatically switch to command-line input mode (manual path entry)

### Issue 1: File selection window does not appear

**Symptom**: After running the tool, no file selection window appears

**Solution**:
1. The window may be behind other windows; check the taskbar
2. Check if there are error messages in the command line
3. If tkinter is not installed, the program will automatically switch to command-line input mode

### Issue 2: Non-Latin1 character warning detected

**Symptom**:
```
⚠ Warning: Non-Latin1 characters found!
  Row 5, Column 'Title':
    Content: Proportion of participants with maintained...
    Problem character: ≥
```

**Solution**:
1. Return to the Excel file
2. Replace `≥` with `>=`
3. Replace `≤` with `<=`
4. Rerun the tool

### Issue 3: Section order is incorrect

**Check**:
- Whether the sect_num column format in Excel is consistent
- Whether there are extra spaces

**Tip**: The tool automatically sorts by number (14.2.10 will be sorted after 14.2.2)

### Issue 4: Missing required columns

**Check column names** for exact matches (note capitalization and spaces):
- `sect_num` (not sect_number)
- `sect_ttl` (not sect_title)
- `OUTFILE` (all uppercase)
- `Output Type (Table, Listing, Figure)` (include parentheses)
- `tocnumber` (not toc_number)
- `Title` (capitalized)

---

## 📞 Need Help?

View complete documentation: [README_GENERATE_BATCH_XML.md](README_GENERATE_BATCH_XML.md)

---

## ✅ Checklist

Before using, confirm:

- [ ] Excel/CSV file is ready
- [ ] Contains all required columns
- [ ] Column names match exactly
- [ ] sect_num format is consistent
- [ ] Python and dependencies are installed (pandas, openpyxl)

Ready? Run: `run_generate_batch_xml.bat`
```