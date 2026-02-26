# TLF Template Filler Script - User Guide

## Overview

This project contains five main Python scripts for automating clinical research data processing and report template filling:

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
或
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

/* ====== Program Execution Commands ====== */

%runpgm(pgm=t_ds, error_override=y);
%runpgm(pgm=t_dm, error_override=y);
```

**Related Documentation:**
- README_EXTRACT_PROGRAMS.md

---

### 3. fill_tlf_template.py - TLF Template Filling

**Features:**
- Consume MOSAIC_CONVERT output
- Merge data based on people_management.xlsx structure
- Three-tier cascading match for personnel information (Output Name → Program Name → Unmatched marking)
- Preserve all sheets and columns from personnel file
- Generate new Excel file and save to user-specified location

**使用方法：**
```bash
.venv\Scripts\python.exe fill_tlf_template.py
```

或使用批处理文件：
```bash
run_fill_tlf_template.bat
```

**工作流程：**
1. 系统提示选择MOSAIC_CONVERT输出文件（Index sheet）
2. 系统提示选择people_management.xlsx文件
3. 脚本自动执行：
   - 数据列映射（MOSAIC列 → people_management列）
   - **第一优先级**：使用Output Name进行匹配
   - **第二优先级**：使用Program Name补充未匹配的行
   - **第三优先级**：标记完全未匹配的行（黄色高亮）和Tier 2匹配的行（绿色高亮）
   - 保留people_management中的所有sheet
   - 在目标sheet中填充合并后的数据
4. 系统提示选择输出文件的保存位置和文件名
5. 完成后显示匹配统计信息

## 列映射说明

MOSAIC输出列 → people_management列：

| 源列（MOSAIC） | 目标列（people_management） | 说明 |
|---|---|---|
| Output Type (Table, Listing, Figure) | Output Type (Table, Listing, Figure) | 输出类型 |
| tocnumber | Output # | 输出编号 |
| Title | Title | 标题 |
| sect_num | Section # | 部分号 |
| sect_ttl | Section Title | 部分标题 |
| azsolid | Standard Template Reference | 标准模板参考 |
| PROGRAM | Program Name | 程序名称 |
| OUTFILE | Output Name | 输出文件名 |

## 三级联动匹配说明

### 匹配优先级

**第一优先级：Output Name匹配**
- 使用MOSAIC的Output Name列与people_management中的Output Name列进行精确匹配
- 若匹配成功，直接填充Programmer、QC Program、QC Programmer三列

**第二优先级：Program Name匹配（补充）**
- 仅对第一优先级未匹配的行进行
- 使用MOSAIC的Program Name列与people_management中的Program Name列进行精确匹配
- 若匹配成功，填充Programmer、QC Program、QC Programmer三列

**第三优先级：标记（高亮）**
- 第一、二优先级都未匹配的行使用**黄色**高亮标记，需要手动补充
- 仅通过第二优先级匹配的行使用**绿色**高亮标记，便于识别
- 高亮仅应用于Programmer、QC Program、QC Programmer三列

### 输出结果

- 保留people_management文件的所有sheet（行头信息、各级别管理等）
- 在TLF sheet中填充新的MOSAIC数据
- 所有原有列都保留，未受影响的列保持不变
- 生成新文件，原文件不被修改

## 人员数据文件结构

people_management.xlsx应包含以下列：

| 列名 | 说明 | 示例 |
|---|---|---|
| Program Name | SAS程序名称 | t_dm |
| Programmer | 负责程序员 | John Doe |
| QC Program | QC程序名称 | t_dm |
| QC Programmer | 负责QC人员 | Jane Smith |

合并规则：使用people_management.xlsx中的"Program Name"作为匹配键，映射"Programmer"、"QC Program"和"QC Programmer"。

## 输出结果

### Template Test File

位置：`02_output\2026-02-09\Oncology Internal Validation Template and Guidance_TEST.xlsx`

**内容：**
- 行1：位置标题
- 行2：列标题（24个列）
- 行3+：填充的245行数据，包括：
  - 输出类型、编号、标题
  - 部分号和标题
  - 程序名称和输出文件名
  - 程序员和QC程序员（来自people_management）
  - 其他元数据列（预留供后续填充）

---

### 4. fill_tlf_status.py - TLF状态填充

**功能：**
- 将TFL Status文件中的Comparison Status合并到people_management的QC Status列
- 自动状态预处理（Match→Pass, Mismatch→Fail）
- 基于Dataset（tfl_status）和Output Name（people_management）精确匹配
- 保留所有原有sheet和列
- 生成统计报告

**使用方法：**
```bash
.venv\Scripts\python.exe fill_tlf_status.py
```

或使用批处理文件：
```bash
run_fill_tlf_status.bat
```

**工作流程：**
1. 选择已修改的people_management.xlsx文件
2. 选择tfl_status.xlsx文件
3. 自动执行状态预处理和匹配
4. 更新QC Status列
5. 选择输出位置和文件名
6. 显示统计信息（总数/Pass/Fail/空值/匹配率）

**状态转换规则：**

| tfl_status中的值 | people_management中的值 |
|---|---|
| Match | Pass |
| Mismatch | Fail |
| (未匹配) | (空值) |

**详细文档：**
参见 [README_FILL_TLF_STATUS.md](README_FILL_TLF_STATUS.md) 获取完整说明。

---

### 5. generate_batch_xml.py - Batch List XML生成

**功能：**
- 从Excel/CSV文件自动生成Adobe PDF Builder格式的batch list XML
- 按section智能分组和排序（支持正确的数字排序：14.2.10排在14.2.2之后）
- 验证latin1字符兼容性并给出警告
- 支持自定义header文本、文件位置、输出路径
- 交互式用户界面

**使用方法：**
```bash
.venv\Scripts\python.exe generate_batch_xml.py
```

或使用批处理文件：
```bash
run_generate_batch_xml.bat
```

**工作流程：**
1. 在弹出窗口中选择包含TLF清单的Excel/CSV文件
2. 输入自定义header文本（默认：AZD0901 CSR DR2）
3. 输入文件位置前缀（默认：root/cdar/d980/d9802c00001/ar/dr2/tlf/dev/output/）
4. 在弹出窗口中选择输出XML文件保存位置和文件名
5. 输入起始页码（默认：2）
6. 自动生成XML文件

**输入文件要求：**

必须包含以下列：
- `sect_num` - Section编号（如：14.1, 14.2.1）
- `sect_ttl` - Section标题
- `OUTFILE` - 输出文件名（不含.rtf扩展名）
- `Output Type (Table, Listing, Figure)` - 输出类型
- `tocnumber` - TOC编号
- `Title` - 标题

**输出格式示例：**
```xml
<?xml version="1.0" encoding="UTF-8"?>
<pdf-builder-metadata>
<!-- input files total to less than 100MB -->
    <ruleset>
        <headers>
            <header text="AZD0901 CSR DR2" startNumber="2" />
        </headers>
        <page orientation="landscape" size="letter" ... />
        <font fontName="CourierNew" style="normal" size="9" />
        <!-- <character-encoding type="ascii" /> -->
        <document-heading text="AZD0901 CSR DR2" fontName="Times New Roman" />
    </ruleset>
    <sectionset>
        <section name="14.1 Study Population">
            <source-file filename="t_ds_comb.rtf" fileLocation="root/cdar/d980/d9802c00001/ar/dr2/tlf/dev/output/" number="Table 14.1.1" title="Disposition" />
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

**特殊功能：**
- ✅ 图形化文件选择窗口（无需手动输入路径）
- ✅ 智能数字排序（14.1, 14.2, 14.10, 14.20）
- ✅ Latin1字符验证和警告
- ✅ 自定义输出路径和文件名
- ✅ 时间戳自动命名
- ✅ 自动定位到相关文件夹（02_output和03_xml）
- ✅ **完整XML结构** - 包含ruleset、page、font等所有Adobe PDF Builder必需的元素
- ✅ **UTF-8编码** - 保留 `encoding="UTF-8"` 声明

**详细文档：**
- 快速入门：[QUICK_START_GENERATE_XML.md](QUICK_START_GENERATE_XML.md)
- 完整指南：[README_GENERATE_BATCH_XML.md](README_GENERATE_BATCH_XML.md)
- XML结构说明：[XML_STRUCTURE_RULES.md](XML_STRUCTURE_RULES.md)

## 常见问题

### Q1: 脚本显示"Permission denied"错误
**A:** 这通常表示Excel文件被打开或被其他程序锁定。
- 解决方案：关闭任何打开的Excel文件，重新运行脚本

### Q2: 程序员数据未能合并
**A:** 检查以下几点：
- people_management.xlsx中的Program Name是否与MOSAIC数据中的PROGRAM列对应
- 确保people_management.xlsx有"Programmer"、"QC Program"和"QC Programmer"列

### Q3: 模板中有空行或缺失数据
**A:** 确保：
- MOSAIC_CONVERT输出文件是test8.xlsx（包含245行）
- 所有必需的源列都存在于MOSAIC输出中

## 技术细节

### 数据处理
1. **去重：** 使用tocnumber作为唯一标识符
2. **映射：** 使用pandas字典映射进行列重命名
3. **合并：** 使用pandas.map()进行程序员信息查询
4. **保存：** 使用openpyxl逐行写入Excel

### 性能指标
- 处理245行数据：< 5秒
- 模板更新：< 3秒
- 总执行时间：< 10秒

## 文件清单

### 必需文件
- ✅ mosaic_convert.py - MOSAIC转换脚本（已验证）
- ✅ fill_tlf_template.py - 模板填充脚本（更新完成）
- ✅ Clinical Study Report_TiFo_MOSAIC_CONVERT_test8.xlsx - MOSAIC输出（245行）
- ✅ Oncology Internal Validation Template and Guidance.xlsx - 模板
- ✅ people_management_sample.xlsx - 人员数据示例

### 可选文件
- run_fill_tlf_template.bat - Windows批处理启动脚本
- test_fill_tlf_template.py - 测试脚本（用于验证功能）

## 后续改进建议

1. **自动化批处理**
   - 扩展支持批量处理多个输入文件
   - 添加定时执行功能

2. **数据验证**
   - 添加数据完整性检查
   - 验证所有必需列的存在性

3. **错误处理**
   - 增强异常捕获和错误报告
   - 添加日志记录功能

4. **格式化增强**
   - 保留MOSAIC中的Excel格式化（字体、颜色等）
   - 条件格式化规则的复制

## 支持和维护

如有问题或需要修改，请检查：
1. 脚本中的注释说明
2. 输入文件的格式和内容
3. 虚拟环境中的包版本

---
**更新日期：** 2026年2月11日  
**mosaic_convert.py 版本：** 3.0  
**fill_tlf_template.py 版本：** 1.2（支持三级匹配、结构保留、双色高亮）  
**fill_tlf_status.py 版本：** 1.0（状态合并、统计报告）  
**generate_batch_xml.py 版本：** 1.0（XML生成、智能排序、Latin1验证）  
**系统状态：** ✅ 生产就绪
