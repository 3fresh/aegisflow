# Quick Start - Generate Batch List XML Tool

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

## 📋 完整示例

### 输入Excel示例

| sect_num | sect_ttl | OUTFILE | Output Type | tocnumber | Title |
|----------|----------|---------|-------------|-----------|-------|
| 14.1 | Study Population | t_ds_comb | Table | 14.1.1 | Disposition |
| 14.1 | Study Population | t_aztoncsp16_itt | Table | 14.1.2 | Recruitment per region |
| 14.2.1 | Primary Endpoint - PFS | t_aztoncef04_pfs_bicr_itt | Table | 14.2.1.1 | PFS by BICR |

### 输出XML示例

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

## ⚠️ 常见问题处理

### 问题0: bat文件运行时窗口闪退

**现象**: 双击bat文件后，窗口一闪就消失了

**诊断步骤**:
1. **先运行测试脚本**: 双击 `test_environment.bat`
   - 这会检查Python和所有必需的模块是否已安装
   - 窗口会停留，显示测试结果

2. **查看测试结果**:
   - 如果 pandas 或 openpyxl 未安装，运行: `pip install pandas openpyxl`
   - 如果 tkinter 未安装，参考下面的解决方法

3. **如果tkinter未安装**:
   - Windows: tkinter通常随Python一起安装，重新安装Python并勾选"tcl/tk"选项
   - 或者，程序会自动切换到命令行输入模式（手动输入路径）

### 问题1: 文件选择窗口没有显示

**现象**: 运行工具后没有看到文件选择窗口

**解决方法**:
1. 窗口可能在其他窗口后面，请检查任务栏
2. 查看命令行是否有错误提示
3. 如果提示tkinter未安装，程序会自动切换到命令行输入模式

### 问题2: 发现非latin1字符警告

**现象**:
```
⚠ 警告: 发现非latin1字符!
  行 5, 列 'Title':
    内容: Proportion of participants with maintained...
    问题字符: ≥
```

**解决方法**:
1. 返回Excel文件
2. 将 `≥` 替换为 `>=`
3. 将 `≤` 替换为 `<=`
4. 重新运行工具

### 问题2: Section顺序不对

**检查**:
- Excel中sect_num列格式是否一致
- 是否有额外空格

**提示**: 工具会自动按数字排序（14.2.10会排在14.2.2之后）

### 问题3: 缺少必需的列

**检查列名**是否完全匹配（注意大小写和空格）:
- `sect_num` （不是sect_number）
- `sect_ttl` （不是sect_title）
- `OUTFILE` （全大写）
- `Output Type (Table, Listing, Figure)` （包括括号）
- `tocnumber` （不是toc_number）
- `Title` （首字母大写）

---

## 📞 需要帮助？

查看完整文档: [README_GENERATE_BATCH_XML.md](README_GENERATE_BATCH_XML.md)

---

## ✅ 检查清单

使用前确认：

- [ ] Excel/CSV文件准备好
- [ ] 包含所有必需的列
- [ ] 列名完全匹配
- [ ] sect_num格式统一
- [ ] 已安装Python和依赖库（pandas, openpyxl）

准备就绪？运行: `run_generate_batch_xml.bat`
