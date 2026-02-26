# XML生成规则说明

## 📋 XML结构说明

生成的XML文件严格遵循Adobe PDF Builder的batch list格式，与参考文件 `03_xml/reference/_batch_list.xml` 保持一致。

---

## ✏️ 用户可自定义的内容

### 1. Header文本
**位置**: 
- `<header text="..." />`
- `<document-heading text="..." />`

**来源**: 用户在交互界面中输入

**默认值**: `AZD0901 CSR DR2`

**说明**: 这两个位置会使用相同的文本

---

### 2. 起始页码
**位置**: `<header startNumber="..." />`

**来源**: 用户在交互界面中输入

**默认值**: `2`

---

### 3. Section信息
**位置**: `<section name="..." >`

**来源**: Excel文件中的两列拼接
- `sect_num` 列（如: 14.1, 14.2.1）
- `sect_ttl` 列（如: Study Population）

**格式**: `sect_num + 空格 + sect_ttl`

**示例**: `14.1 Study Population`

---

### 4. Source-file 属性

#### a) filename
**位置**: `<source-file filename="..." />`

**来源**: Excel的 `OUTFILE` 列 + `.rtf` 扩展名

**示例**: 
- Excel中: `t_ds_comb`
- XML中: `t_ds_comb.rtf`

#### b) fileLocation
**位置**: `<source-file fileLocation="..." />`

**来源**: 用户在交互界面中输入

**默认值**: `root/cdar/d980/d9802c00001/ar/dr2/tlf/dev/output/`

**说明**: 所有source-file使用相同的fileLocation

#### c) number
**位置**: `<source-file number="..." />`

**来源**: Excel的两列拼接
- `Output Type (Table, Listing, Figure)` 列（如: Table, Figure）  
- `tocnumber` 列（如: 14.1.1）

**格式**: `Output Type + 空格 + tocnumber`

**示例**: `Table 14.1.1`, `Figure 14.2.1.3`

#### d) title
**位置**: `<source-file title="..." />`

**来源**: Excel的 `Title` 列

**示例**: `Disposition`

---

## 🔒 固定不变的内容

以下内容在XML中是**完全固定**的，程序会自动生成，**不需要也不应该**修改：

### 1. XML声明
```xml
<?xml version="1.0" encoding="UTF-8"?>
```
✅ **保留**: UTF-8编码声明

---

### 2. 根元素
```xml
<pdf-builder-metadata>
<!-- input files total to less than 100MB -->
```
✅ **保留**: 
- 没有 `xmlns` 属性
- 没有 `job-name` 属性
- 包含注释

---

### 3. Ruleset结构
```xml
<ruleset>
    <headers>
        <header text="【用户自定义】" startNumber="【用户自定义】" />
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
    <document-heading text="【用户自定义】" fontName="Times New Roman" />
</ruleset>
```

**固定内容**:
- ✅ `<page>` 元素的所有属性（landscape, letter, 边距等）
- ✅ `<font>` 元素的所有属性（CourierNew, size 9）
- ✅ 注释 `<!-- <character-encoding type="ascii" /> -->`
- ✅ `<document-heading>` 的 `fontName="Times New Roman"`

---

### 4. Output PDF配置
```xml
<output-pdf filename="CG01_DR2.pdf">
    <pdf-import path="root/cdar/d980/d9802c00001/ar/dr2/tlf/doc/" />
</output-pdf>
```

**固定内容**:
- ✅ filename: `CG01_DR2.pdf`
- ✅ path: `root/cdar/d980/d9802c00001/ar/dr2/tlf/doc/`

---

### 5. Output Audit配置
```xml
<output-audit filename="CG01_DR2_audit.pdf">
    <audit-import path="root/cdar/d980/d9802c00001/ar/dr2/tlf/doc/" />
</output-audit>
```

**固定内容**:
- ✅ filename: `CG01_DR2_audit.pdf`
- ✅ path: `root/cdar/d980/d9802c00001/ar/dr2/tlf/doc/`

---

## 📊 数据流图

```
Excel文件数据
    ├─ sect_num ────┐
    ├─ sect_ttl ────┤─→ section name
    │               └─ (拼接)
    │
    ├─ OUTFILE ─────→ filename (+ .rtf)
    │
    ├─ Output Type ─┐
    ├─ tocnumber ───┤─→ number
    │               └─ (拼接)
    │
    └─ Title ───────→ title

用户输入
    ├─ Header文本 ──→ header text & document-heading text
    ├─ 起始页码 ────→ startNumber
    └─ 文件位置 ────→ fileLocation

程序固定
    ├─ XML声明 (UTF-8)
    ├─ Ruleset结构
    │   ├─ page配置
    │   ├─ font配置
    │   └─ 注释
    ├─ output-pdf配置
    └─ output-audit配置
```

---

## ⚠️ 重要提示

1. **不要删除固定内容**: 
   - ruleset中的page、font等配置
   - output-pdf和output-audit部分
   - 所有注释

2. **不要修改属性名**: 
   - `encoding="UTF-8"` 不能改为其他编码
   - `<pdf-builder-metadata>` 不能添加额外属性

3. **保持格式一致**:
   - 所有标签使用自关闭格式 `<tag ... />`
   - 缩进为4个空格
   - 保持与参考XML完全一致

---

## 🔍 验证方法

生成XML后，可以与参考文件对比：
- 参考文件: `03_xml/reference/_batch_list.xml`
- 检查固定部分是否完全相同
- 检查自定义部分是否正确填充

---

## 📝 总结

| 内容类型 | 数量 | 来源 |
|---------|------|------|
| 用户输入 | 2项 | Header文本、文件位置 |
| Excel数据 | 每行6个字段 | sect_num, sect_ttl, OUTFILE, Output Type, tocnumber, Title |
| 固定内容 | ~20行 | XML声明、ruleset、output配置 |

**原则**: 最小化用户输入，最大化自动化，确保输出格式完全符合Adobe PDF Builder要求。
