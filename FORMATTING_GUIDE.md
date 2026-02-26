# Excel Formatting Guide (v2.5)

## 📋 Overview

The Excel files output by MOSAIC_CONVERT now include comprehensive formatting and intelligent labeling functionality. This guide explains the meaning of each formatting style and its application scenarios.

## 🎨 Formatting Types

### 1. Font Formatting

#### Global Font: DengXian
- **Applications**: All cells
- **Purpose**: Uniform appearance and improved readability
- **Features**:
  - Headers: DengXian + Bold
  - Data: DengXian
  - Rich text sections (containing red text): DengXian + Red

### 2. Background Colors

#### Yellow Background (FFFF00)

**Applied to**:
- H1:J1 (3 header columns)
- Tocnumber/PROGRAM/SUFFIX columns for shared program+suffix combinations

**Meaning**:
- These columns require special attention
- Certain tables share the same program+suffix combinations (31 shared combinations total)
- 66 tocnumbers are affected

**Example**:
| Tocnumber | Program | Suffix | Description |
|-----------|---------|--------|-------------|
| 14.1.5.2 (Yellow) | t_dm (Yellow) | itt3l (Yellow) | Shared with 14.1.5.3 |
| 14.1.5.3 (Yellow) | t_dm (Yellow) | itt3l (Yellow) | Shared with 14.1.5.2 |
| 14.1.9.1 | t_dm | dischar_itt | Unique |

#### Green Background (92D050)

**Applied to**:
- Cells containing non-Latin1 characters

**Meaning**:
- This cell contains character encoding issues
- Problematic characters are highlighted in red
- Requires manual review and correction

**Example**:
- "Primary Endpoint - Progression" (green background, "-" in red)
- Possible character issues:
  - Dashes (--, --)
  - Special symbols (C, R, TM)
  - Non-ASCII characters

### 3. Font Colors

#### Red Font

**Applied to**:
- Non-Latin1 characters (individual character only)

**Meaning**:
- Marks specific characters with encoding issues
- Need to be replaced with equivalent ASCII characters

**Replacement Suggestions**:
| Problem Character | Recommended Replacement |
|-------------------|-------------------------|
| - (en dash) | - (hyphen-minus) |
| - (em dash) | - (hyphen-minus) |
| ' (right single quote) | ' (apostrophe) |
| " (left/right double quote) | " (quotation mark) |

## 📊 Sorting Method

### Numeric Sort

**Sorting Rules**:
1. Split tocnumber by "."
2. Convert each part to integer
3. Compare recursively, left to right

**Example Sequence**:
```
14.1.1
14.1.2
14.1.2.1
14.1.2.2
14.1.3
14.1.10    <- 10 comes after 3 (numeric sort)
14.1.10.1
14.2
14.2.1.1.1
14.2.1.1.2
14.2.1.2
...
```

**Comparison (Character Sort - Incorrect)**:
```
14.1.1
14.1.10    <- Character sort would place 10 before 2
14.1.2     <- This is incorrect!
14.1.3
```

## 🔍 Usage Examples

### Identifying Shared Tables

**Scenario**: 10 tables need to use the same program+suffix

```
1. Search for all yellow background rows
2. View all tocnumbers under the same program+suffix
3. These tocnumbers share structure

Example (program=t_dm, suffix=dischar_itt):
- 14.1.9.1 (Yellow)
- 14.1.9.2 (Yellow)
- 14.1.9.3 (Yellow)
```

### Correcting Non-Latin1 Characters

**Scenario**: Found a cell with green background

```
1. Locate the green cell
2. View the red character
3. Perform replacement according to the table above
4. Re-import or manually correct
```

## 📝 Data Statistics

### Current Data

| Type | Quantity | Notes |
|------|----------|-------|
| Total Tables | 249 | Based on seq sequence transpose |
| Original Rows | 3,211 | CSV format |
| Shared Combinations | 31 | program+suffix |
| Affected Rows | 66 | Shared tocnumbers |
| Red-marked Cells | 59 | Non-Latin1 characters |

### Common Shared Combinations

| Program | Suffix | Number of Tocnumbers | Example |
|---------|--------|----------------------|---------|
| t_dm | dischar_itt | 3 | 14.1.9.1-3 |
| f_km | pfs_itt | 4 | 14.2.1.3.1-4 |
| t_dm | itt3l | 2 | 14.1.5.2-3 |
| t_cm | subct_itt3l | 2 | 14.1.18.2-3 |

## 💡 Best Practices

### 1. Quick Review
- Use yellow background to quickly locate rows requiring special handling
- Use green background to quickly locate encoding issues

### 2. Batch Processing
- Classify according to yellow markers
- Process in batches by program+suffix

### 3. Data Validation
- Pause at green cells to check if correction is needed
- Verify that all yellow-marked rows are present and complete

### 4. Sort Verification
- Check that tocnumber sequences are correct
- Ensure 14.1.10 comes after 14.1.2

## 🔧 Technical Details

### Rich Text Implementation

Using openpyxl's RichText functionality:

```python
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

# Create specific characters with red font
red_font = InlineFont(rFont='DengXian', color='FF0000')
default_font = InlineFont(rFont='DengXian')

# Organize into RichText
rich_text = CellRichText(
    TextBlock(default_font, "Primary Endpoint "),
    TextBlock(red_font, "-"),
    TextBlock(default_font, " Progression")
)
```

### Sorting Implementation

```python
def tocnumber_sort_key(tocnum):
    """Convert tocnumber to sortable tuple"""
    if pd.isna(tocnum):
        return (float('inf'),)
    parts = str(tocnum).split('.')
    return tuple(
        int(p) if p.isdigit() else float('inf') 
        for p in parts
    )

# Apply sorting
index_final['_sort_key'] = index_final['tocnumber'].apply(tocnumber_sort_key)
index_final = index_final.sort_values('_sort_key')
```

## FAQ

### Q: Why are some tables marked as yellow?
**A**: This indicates they share the same program+suffix combination with other tables. This is a data characteristic, not an error.

### Q: Can non-Latin1 characters be converted automatically?
**A**: No. These characters require manual review and correction. Red markers help identify them.

### Q: What is the sorting rule?
**A**: Numeric sorting is used, not character sorting. Therefore, 14.1.2 comes before 14.1.10.

### Q: Can background colors be modified?
**A**: Yes. However, we recommend keeping these markers for quick identification of data characteristics.

### Q: Why are the headers also yellow?
**A**: Columns H, I, and J headers are yellow, indicating that these columns are important and frequently marked.

## Support

For formatting-related questions, please refer to:
- README_MOSAIC_CONVERT.md - Complete documentation
- SELECT_ROWS.md - Feature summary
- CHANGELOG.md - Version history

---

**Version**: 2.5  
**Date**: 2026-02-10  
**Last Updated**: Added numeric sorting and rich text formatting