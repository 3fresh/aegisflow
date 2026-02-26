# Generate Batch List XML Tool

## 功能说明 / Overview

根据Excel/CSV表格自动生成符合Adobe PDF Builder格式的batch list XML文件。

Generate batch list XML files for Adobe PDF Builder from Excel/CSV input.

---

## 主要功能 / Features

1. ✅ 从Excel或CSV读取数据
2. ✅ **图形化文件选择** - 无需手动输入路径，使用弹出窗口选择文件
3. ✅ 自动按section分组和排序（支持正确的数字排序，如14.2.10排在14.2.2之后）
4. ✅ 生成标准的XML结构（**完整保留**所有Adobe PDF Builder必需的元素）
5. ✅ **保留UTF-8编码**和所有固定内容（ruleset、page、font等）
6. ✅ 验证latin1字符兼容性并给出警告
7. ✅ 支持自定义header、文件位置、输出路径
8. ✅ 交互式用户界面，自动定位到相关文件夹

**详细说明**:
- XML结构规则: [XML_STRUCTURE_RULES.md](XML_STRUCTURE_RULES.md)
- 快速入门: [QUICK_START_GENERATE_XML.md](QUICK_START_GENERATE_XML.md)

---

## 使用方法 / Usage

### 方法1: 双击批处理文件（推荐）

1. 双击 `run_generate_batch_xml.bat`
2. 按提示输入信息
3. 等待生成完成

### 方法2: 命令行运行

```bash
python generate_batch_xml.py
```

---

## 输入文件要求 / Input Requirements

### 必需的列 / Required Columns

Excel/CSV文件必须包含以下列：

| 列名 | 说明 | 示例 |
|------|------|------|
| `sect_num` | Section编号 | 14.1, 14.2.1, 14.2.10 |
| `sect_ttl` | Section标题 | Study Population |
| `OUTFILE` | 输出文件名（不含扩展名） | t_ds_comb |
| `Output Type (Table, Listing, Figure)` | 输出类型 | Table, Figure, Listing |
| `tocnumber` | TOC编号 | 14.1.1, 14.2.1.1 |
| `Title` | 标题 | Disposition |

### 文件格式支持

- ✅ Excel文件: `.xlsx`, `.xls`
- ✅ CSV文件: `.csv` (支持多种编码: UTF-8, GBK, GB2312, latin1)

---

## 交互式输入说明 / Interactive Prompts

运行工具后，会依次提示：

### 1. 选择输入文件
```
Step 1/6: Select Input File
Please select Excel or CSV file in the popup window...
```
- File selection dialog will appear
- Default opens `02_output` directory
- Supports filtering for .xlsx, .xls, .csv files

### 2. Header文本 (Header Text)
```
Step 2/6: Set Header Text
This will be used for: <header text="..."> and <document-heading text="...">
Enter header text [Default: AZD0901 CSR DR2 Batch 1 Listings]:
```
- Press Enter to use default value
- Or enter custom text, e.g.: `My Study CSR Batch 1 Tables`
- **Used for**: XML header display and document heading

### 3. 输出文件名 (Output Filename) - ⭐ NEW
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

### 4. 文件位置前缀 (File Location)
```
Step 4/6: Set File Location
Enter file location [Default: root/cdar/d980/d9802c00001/ar/dr2/tlf/dev/output/]:
```
- Press Enter to use default value
- Or enter custom path

### 5. 输出XML路径 (Output XML Path)
```
Step 5/6: Select Output Location
Please select save location and filename in the popup window...
```
- File save dialog will appear
- Default opens `03_xml` directory
- Default filename with timestamp: `_batch_list_20260211_143025.xml`
- Can modify filename and save location

### 6. 起始页码 (Starting Page Number)
```
Step 6/6: Set Starting Page Number
Enter starting page number [Default: 2]:
```
- Press Enter to use default value 2
- Or enter other number

---

## 输出格式 / Output Format

生成的XML文件格式示例：

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

**注意**: 除了以下用户自定义的内容外，其他所有内容（如ruleset结构、page、font等）都是固定的：
- `header text` 和 `document-heading text`: 用户自定义 (Step 2)
- `header startNumber`: 用户自定义 (Step 6, 默认2)
- `output-pdf filename`: 用户自定义 (Step 3) - **NEW**
- `output-audit filename`: 用户自定义 (Step 3) + "_audit" - **NEW**
- `pdf-import` 和 `audit-import` path: 自动从file location提取
- `section name`: 从Excel的sect_num和sect_ttl列拼接
- `source-file filename`: 从Excel的OUTFILE列 + ".rtf"
- `source-file fileLocation`: 用户自定义 (Step 4)
- `source-file number`: 从Excel的"Output Type"和"tocnumber"列拼接
- `source-file title`: 从Excel的Title列

---

## 特殊功能 / Special Features

### 1. 智能数字排序

工具会正确处理section编号的数字排序：
- ✅ 14.1, 14.2, 14.10, 14.20 (正确)
- ❌ 不会出现: 14.1, 14.10, 14.2, 14.20 (错误的字符串排序)

### 2. Latin1字符检查

在生成XML前，工具会检查所有文本是否只包含latin1字符：

- ✅ 如果全部兼容，直接生成
- ⚠ 如果发现非latin1字符：
  - 列出所有问题位置
  - 显示具体的问题字符
  - 询问是否继续

**常见非latin1字符**:
- 中文字符: 如 "分析"
- 特殊符号: ≥, ≤, —, ±, °, μ 等
- 非ASCII引号: " " ' '

**解决方法**:
- 将中文替换为英文
- 将特殊符号替换为ASCII等效字符:
  - `≥` → `>=`
  - `≤` → `<=`
  - `—` → `-`
  - `±` → `+/-`

### 3. 数据验证

工具会自动验证：
- ✅ 必需列是否存在
- ✅ 文件是否可读
- ✅ 数据完整性
- ✅ 字符编码兼容性

---

## 常见问题 / FAQ

### Q0: bat文件运行时窗口闪退怎么办？

**A**: 这通常表示某些依赖包未安装。请按以下步骤诊断：

1. **运行测试脚本**: 双击 `test_environment.bat`
2. **查看哪些模块缺失**
3. **安装缺失的包**: 
   ```bash
   pip install pandas openpyxl
   ```
4. **重新运行** `run_generate_batch_xml.bat`

**注意**: 增强版的bat文件现在会显示错误信息并保持窗口打开。

### Q1: 提示"缺少必需的列"怎么办？

**A**: 检查Excel文件的列名是否完全匹配（包括大小写和空格）：
- `sect_num`
- `sect_ttl`
- `OUTFILE`
- `Output Type (Table, Listing, Figure)`
- `tocnumber`
- `Title`

### Q2: CSV文件读取失败怎么办？

**A**: 工具会自动尝试多种编码。如果仍失败，请：
1. 用Excel打开CSV
2. 另存为UTF-8编码的CSV
3. 重新运行工具

### Q3: 发现非latin1字符警告，怎么处理？

**A**: 有三个选择：
1. **取消生成**，返回Excel修改内容
2. **继续生成**，但XML可能在某些系统中显示异常
3. **修改后重新生成**（推荐）

### Q4: Section顺序不对怎么办？

**A**: 检查Excel中的`sect_num`列：
- 确保格式一致（都是数字，如14.1）
- 确保没有额外空格
- 工具会自动进行数字排序

### Q5: 想修改XML模板怎么办？

**A**: 编辑 `generate_batch_xml.py` 文件中的 `generate_xml()` 方法。

### Q6: Output filename中包含空格怎么办？ ⭐ NEW

**A**: 工具会自动验证并拒绝包含空格的文件名：
- ❌ 输入包含空格 → 显示错误并要求重新输入
- ✅ 使用下划线 `_` 替代空格
- 示例:
  ```
  错误: My Study DR2
  正确: My_Study_DR2
  ```
- 如果直接按回车使用默认值，工具会自动将header text中的空格转换为下划线

---

## 文件结构 / File Structure

```
03_mastertoc/
├── generate_batch_xml.py          # 主程序
├── run_generate_batch_xml.bat     # Windows批处理文件
├── test_environment.bat            # 环境测试脚本（诊断用）
├── test_imports.py                 # 导入测试脚本（诊断用）
├── README_GENERATE_BATCH_XML.md   # 本文档
│
├── 02_output/                     # 输入文件夹（存放Excel/CSV）
│   └── 2026-02-11/
│       └── Clinical Study Report_TiFo.csv
│
└── 03_xml/                        # 输出文件夹（存放生成的XML）
    └── _batch_list_YYYYMMDD_HHMMSS.xml
```

### 诊断工具

**test_environment.bat** - 环境诊断工具
- 检查Python版本
- 检查所有必需的模块是否已安装
- 如果遇到问题，先运行这个工具

---

## 技术细节 / Technical Details

### 依赖库 / Dependencies

```
pandas
openpyxl (for Excel files)
tkinter (文件选择对话框，Python通常自带)
```

### 安装依赖

```bash
pip install pandas openpyxl
```

**注意**: `tkinter` 通常随Python一起安装。如果遇到 `ModuleNotFoundError: No module named 'tkinter'`，请参考Python官方文档安装。

### Python版本要求

- Python 3.6+

---

## 更新日志 / Changelog

### Version 1.1.0 (2026-02-12) ⭐ NEW
- ✅ **新增**: 独立的Output PDF Filename输入步骤
- ✅ **新增**: 文件名空格验证（不允许空格，必须使用下划线）
- ✅ **改进**: 从5步流程扩展为6步流程，更清晰
- ✅ **改进**: Header text和Output filename分离，提供更大灵活性
- ✅ **改进**: 自动从file location提取基础路径用于pdf-import/audit-import
- ✅ **改进**: 全英文界面和错误提示

### Version 1.0.0 (2026-02-11)
- ✅ 初始版本发布
- ✅ 支持Excel和CSV输入
- ✅ 自动数字排序
- ✅ Latin1字符验证
- ✅ 交互式界面

---

## 作者 / Author

Generated by GitHub Copilot Assistant

---

## 许可 / License

内部使用工具 / Internal use only
