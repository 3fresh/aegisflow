# AegisFlow - XML Structure Rules

**Transform, Validate, Deliver from a Single TOC**

## 📋 XML Structure Description

The generated XML file strictly follows Adobe PDF Builder's batch list format, consistent with the reference file `03_xml/reference/_batch_list.xml`.

---

## ✏️ User-Customizable Content

### 1. Header Text
**Location**: 
- `<header text="..." />`
- `<document-heading text="..." />`

**Source**: User input in interactive interface

**Default Value**: `AZD0901 CSR DR2`

**Description**: Both locations use the same text

---

### 2. Starting Page Number
**Location**: `<header startNumber="..." />`

**Source**: User input in interactive interface

**Default Value**: `2`

---

### 3. Section Information
**Location**: `<section name="..." >`

**Source**: Concatenation of two columns from Excel file
- `sect_num` column (e.g.: 14.1, 14.2.1)
- `sect_ttl` column (e.g.: Study Population)

**Format**: `sect_num + space + sect_ttl`

**Example**: `14.1 Study Population`

---

### 4. Source-file Attributes

#### a) filename
**Location**: `<source-file filename="..." />`

**Source**: Excel `OUTFILE` column + `.rtf` extension

**Example**: 
- Excel: `t_ds_comb`
- XML: `t_ds_comb.rtf`

#### b) fileLocation
**Location**: `<source-file fileLocation="..." />`

**Source**: User input in interactive interface

**Default Value**: `root/cdar/d980/d9802c00001/ar/dr2/tlf/dev/output/`

**Description**: All source-file elements use the same fileLocation

#### c) number
**Location**: `<source-file number="..." />`

**Source**: Concatenation of two Excel columns
- `Output Type (Table, Listing, Figure)` column (e.g.: Table, Figure)  
- `tocnumber` column (e.g.: 14.1.1)

**Format**: `Output Type + space + tocnumber`

**Example**: `Table 14.1.1`, `Figure 14.2.1.3`

#### d) title
**Location**: `<source-file title="..." />`

**Source**: Excel `Title` column

**Example**: `Disposition`

---

## 🔒 Fixed and Unchanging Content

The following content in the XML is **completely fixed**; the program generates it automatically and **should not and must not** be modified:

### 1. XML Declaration
```xml
<?xml version="1.0" encoding="UTF-8"?>
```
✅ **Preserved**: UTF-8 encoding declaration

---

### 2. Root Element
```xml
<pdf-builder-metadata>
<!-- input files total to less than 100MB -->
```
✅ **Preserved**: 
- No `xmlns` attribute
- No `job-name` attribute
- Includes comment

---

### 3. Ruleset Structure
```xml
<ruleset>
    <headers>
        <header text="【User-Customizable】" startNumber="【User-Customizable】" />
    </headers>
    <page
        orientation="landscape"
        size="letter"
        measurementUnit="in"
        marginTop="           0"
        marginLeft="           0"
        marginRight="           0"
        marginBottom="           0" />
    <font
        fontName="CourierNew"
        style="normal"
        size="9" />
    <!-- <character-encoding type="ascii" /> -->
    <document-heading text="【User-Customizable】" fontName="Times New Roman" />
</ruleset>
```

**Fixed Content**:
- ✅ All attributes of the `<page>` element (landscape, letter, margins, etc.)
- ✅ All attributes of the `<font>` element (CourierNew, size 9)
- ✅ Comment `<!-- <character-encoding type="ascii" /> -->`
- ✅ `<document-heading>` attribute `fontName="Times New Roman"`

---

### 4. Output PDF Configuration
```xml
<output-pdf filename="CG01_DR2.pdf">
    <pdf-import path="root/cdar/d980/d9802c00001/ar/dr2/tlf/doc/" />
</output-pdf>
```

**Fixed Content**:
- ✅ filename: `CG01_DR2.pdf`
- ✅ path: `root/cdar/d980/d9802c00001/ar/dr2/tlf/doc/`

---

### 5. Output Audit Configuration
```xml
<output-audit filename="CG01_DR2_audit.pdf">
    <audit-import path="root/cdar/d980/d9802c00001/ar/dr2/tlf/doc/" />
</output-audit>
```

**Fixed Content**:
- ✅ filename: `CG01_DR2_audit.pdf`
- ✅ path: `root/cdar/d980/d9802c00001/ar/dr2/tlf/doc/`

---

## 📊 Data Flow Diagram

```
Excel File Data
    ├─ sect_num ────┐
    ├─ sect_ttl ────┤─→ section name
    │               └─ (concatenation)
    │
    ├─ OUTFILE ─────→ filename (+ .rtf)
    │
    ├─ Output Type ─┐
    ├─ tocnumber ───┤─→ number
    │               └─ (concatenation)
    │
    └─ Title ───────→ title

User Input
    ├─ Header text ──→ header text & document-heading text
    ├─ Starting page ────→ startNumber
    └─ File location ────→ fileLocation

Program Fixed
    ├─ XML declaration (UTF-8)
    ├─ Ruleset structure
    │   ├─ page configuration
    │   ├─ font configuration
    │   └─ comments
    ├─ output-pdf configuration
    └─ output-audit configuration
```

---

## ⚠️ Important Reminders

1. **Do not delete fixed content**: 
   - page, font configurations in ruleset
   - output-pdf and output-audit sections
   - All comments

2. **Do not modify attribute names**: 
   - `encoding="UTF-8"` cannot be changed to other encodings
   - `<pdf-builder-metadata>` cannot have additional attributes

3. **Maintain format consistency**:
   - All tags use self-closing format `<tag ... />`
   - Indentation is 4 spaces
   - Keep completely consistent with reference XML

---

## 🔍 Validation Method

After generating XML, you can compare with the reference file:
- Reference file: `03_xml/reference/_batch_list.xml`
- Check if fixed sections are completely identical
- Check if user-customizable sections are correctly filled

---

## 📝 Summary

| Content Type | Quantity | Source |
|---------|------|------|
| User Input | 2 items | Header text, file location |
| Excel Data | 6 fields per row | sect_num, sect_ttl, OUTFILE, Output Type, tocnumber, Title |
| Fixed Content | ~20 lines | XML declaration, ruleset, output configuration |

**Principle**: Minimize user input, maximize automation, ensure output format fully complies with Adobe PDF Builder requirements.