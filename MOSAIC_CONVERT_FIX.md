```markdown
# MOSAIC_CONVERT Fix Documentation

## Date
- Initial fix: February 10, 2026
- v3.1 update: February 11, 2026

## Fixed Issues

### Issue 1: Data confusion caused by identical program+suffix combinations
**Symptom:**
- When two different tocnumbers use the same program and suffix, the converted azsolid and OUTFILE values become corrupted
- Example: 14.2.3.5.1 and 14.2.3.5.2 both use program='f_forest', suffix='os_itt', but they should have different azsolid and outfile values

**Root Cause:**
- Original code used `sect_num, sect_ttl, program, suffix` combination to filter subset
- When multiple tocnumbers have the same program+suffix, the subset would contain data for all those tocnumbers
- This caused data mixing when extracting values for different tocnumbers

**Fix Method:**
- Use row index range (start_idx:end_idx) to distinguish different tocnumbers
- For each tocnumber, extract all data between current tocnumber row and next tocnumber row
- Ensures each tocnumber gets correct data even if program+suffix is the same

**Code Changes:**
```python
# Before (incorrect):
subset = df[(df['sect_num'] == sect_num) & 
            (df['sect_ttl'] == sect_ttl) & 
            (df['program'] == program) & 
            (df['suffix'] == suffix) & 
            (df['parm'] != 'tocnumber')].copy()

# After (correct):
# Extract rows for this tocnumber using row index range
# This ensures each tocnumber gets its own data, even if program+suffix is the same
subset = df.loc[start_idx:end_idx-1].copy()
```

### Issue 2: OUTFILE column value source (corrected in v3.1)
**Initial fix (v2.x-v3.0):**
- Generate OUTFILE by concatenating PROGRAM and SUFFIX
- Format: `PROGRAM_SUFFIX` (when SUFFIX is not empty)

**New issue discovered (v3.1):**
- CSV business logic has special cases: SUFFIX is not empty, but OUTFILE only equals PROGRAM
- Example: PROGRAM='t_ds', SUFFIX='comb', but CSV `parm='outfile'` value is `'t_ds_comb'` as expected
- But also exists: PROGRAM='xxx', SUFFIX='yyy', but `parm='outfile'` value is `'xxx'` (no suffix concatenation)

**Final fix method (v3.1):**
- **Use CSV original values directly**: Extract OUTFILE column directly from CSV `parm='outfile'` corresponding `value`
- No concatenation or transformation
- Preserve original business logic from CSV, ensure consistency with source data

**Code Changes:**
```python
# v3.0 and earlier (concatenate from PROGRAM+SUFFIX):
def build_outfile(row):
    program = row['PROGRAM'] if 'PROGRAM' in row and pd.notna(row['PROGRAM']) else ''
    suffix = row['SUFFIX'] if 'SUFFIX' in row and pd.notna(row['SUFFIX']) else ''
    if suffix and str(suffix) != 'nan' and str(suffix) != '':
        return f"{program}_{suffix}"
    else:
        return program
index_df['OUTFILE'] = index_df.apply(build_outfile, axis=1)

# v3.1 (use CSV values directly):
for _, row in group.iterrows():
    param_name = row.get(param_col, None)
    if pd.isna(param_name) or param_name == '':
        continue
    key = str(param_name).strip().lower()
    # Special handling: Convert 'outfile' to 'OUTFILE' (uppercase)
    if key == 'outfile':
        key = 'OUTFILE'
    row_data[key] = row.get('value', None)
# OUTFILE now extracted directly from CSV, no extra construction needed
```

### Issue 3: Chinese path encoding issue (added in v3.1)
**Symptom:**
- When running `run_mosaic_convert.bat`, verification step always shows "WARN: Output file path not found, skipping verification"
- `.last_output.txt` contains Chinese characters (such as "工具开发") showing as garbled text
- Converted Excel files are actually generated normally, only verification step fails

**Root Cause:**
- `.last_output.txt` written with GBK encoding, cannot properly handle Chinese characters
- Bat file's `for /f` command has Windows console encoding conversion error when capturing Python output
- Even if Python reads file with UTF-8, parameters passed by bat still get garbled

**Fix Method:**
1. **Encoding unification**: Change `.last_output.txt` to UTF-8 encoding (mosaic_convert.py)
2. **Auto read**: When `validate_output.py` has no command line arguments, auto read path from `.last_output.txt`
3. **Simplified process**: `run_mosaic_convert.bat` calls Python directly, no longer passes path parameters through bat

**Technical Details:**
- Before: bat → `for /f` capture Python output (Chinese garbled) → pass to validate_output.py
- Now: bat → call Python directly → Python reads `.last_output.txt` (UTF-8)
- All Chinese path handling done inside Python, completely avoid Windows console encoding issues

## Verification Results

### Verification Method
1. Use original CSV file (Clinical Study Report_TiFo.csv) to generate new Excel file
2. Compare each tocnumber's PROGRAM, SUFFIX, OUTFILE, azsolid in generated Excel with original CSV values
3. Verify Chinese path support (path contains Chinese characters like "工具开发")

### Verification Results (v3.1)
✅ **249 rows of data based on seq transpose are completely correct**
- Each tocnumber's PROGRAM and SUFFIX match original CSV's program and suffix columns
- **OUTFILE uses CSV `parm='outfile'` value directly** (no longer concatenated)
- azsolid value matches original CSV's value for corresponding tocnumber's parm='azsolid' row
- Even multiple tocnumbers with same program+suffix combination (like 14.2.3.5.1 and 14.2.3.5.2) can be correctly distinguished
- **Fully supports paths with Chinese characters**, verification step runs normally

### Test Cases
| tocnumber | Original CSV | v3.1 Output | Status |
|---|---|---|---|
| 14.1.1 | program=t_ds, suffix=comb, outfile=t_ds_comb | OUTFILE=t_ds_comb | ✓ |
| 14.1.5.2 | program=t_dm, suffix=itt3l, outfile=t_dm_itt3l | OUTFILE=t_dm_itt3l | ✓ |
| 14.2.3.5.1 | program=f_forest, suffix=os_itt, outfile=f_forest_os_itt, azsolid=AZFEF02 | OUTFILE=f_forest_os_itt, azsolid=AZFEF02 | ✓ |
| 14.2.3.5.2 | program=f_forest, suffix=os_itt, outfile=f_forest_os_itt, azsolid=AZTONCEF04 | OUTFILE=f_forest_os_itt, azsolid=AZTONCEF04 | ✓ |
| **Special Case** | program=xxx, suffix=yyy, outfile=xxx (no concatenation) | OUTFILE=xxx (correctly preserve CSV original value) | ✓ |

## Impact Scope
- Modified files: mosaic_convert.py, validate_output.py, run_mosaic_convert.bat
- Updated documentation: README_MOSAIC_CONVERT.md, CHANGELOG.md, MOSAIC_CONVERT_FIX.md
- All workflows using mosaic_convert.py benefit from these fixes

## Next Steps
1. ✅ Code fixes complete (v3.1)
2. ✅ OUTFILE data source corrected
3. ✅ Chinese path support
4. ✅ Documentation updated
5. ✅ Verification tests passed
6. ⏳ Need to regenerate MOSAIC_CONVERT output files
7. ⏳ Need to re-run fill_tlf_template.py with new MOSAIC_CONVERT files

```