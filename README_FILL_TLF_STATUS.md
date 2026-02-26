# AegisFlow - Fill TLF Status

**Transform, Validate, Deliver from a Single TOC**

## 📋 概述

`fill_tlf_status.py` 是一个自动化工具，用于将TFL Status文件中的Comparison Status数据合并到People Management文件的QC Status列中。

## 🎯 功能特性

### 核心功能

1. **文件选择**
   - 选择已修改的people_management.xlsx文件
   - 选择tfl_status.xlsx文件

2. **状态预处理**
   - 自动将"Match"转换为"Pass"
   - 自动将"Mismatch"转换为"Fail"

3. **精确匹配合并**
   - 基于`Dataset`（tfl_status）和`Output Name`（people_management）进行精确匹配
   - 仅在完全匹配时才合并数据
   - 未匹配的行QC Status列置空

4. **结构保留**
   - 保留people_management中的所有sheet
   - 保留所有原有列
   - 仅更新QC Status列，其他列不做任何改动

5. **统计报告**
   - TLF总数目
   - Status为"Pass"的数目
   - Status为"Fail"的数目
   - Status为空的数目
   - 匹配率百分比

## 📁 文件结构

### 输入文件

1. **people_management.xlsx**
   - 必需sheet: `TLF`
   - 必需列: `Output Name`
   - 目标列: `QC Status (Not Started, Ongoing, QC Pending, Fail, Pass)`

2. **tfl_status.xlsx**
   - 必需sheet: `Overview`
   - 必需列: `Dataset`, `Comparison Status`

### 输出文件

- 默认文件名: `people_management_with_status.xlsx`
- 用户可自定义文件名和保存路径

## 🚀 使用方法

### 方法1：使用批处理文件（推荐）

```bash
run_fill_tlf_status.bat
```

### 方法2：直接运行Python脚本

```bash
py -3.13 fill_tlf_status.py
```

## 📊 工作流程

```
输入1: people_management.xlsx (TLF sheet)
输入2: tfl_status.xlsx (Overview sheet)
    ↓
[Step 1] 读取people_management文件
[Step 2] 读取tfl_status文件
[Step 3] 预处理Comparison Status
         - Match → Pass
         - Mismatch → Fail
    ↓
[Step 4] 基于Dataset和Output Name精确匹配
         - 匹配成功：填充QC Status
         - 匹配失败：QC Status置空
    ↓
[Step 5] 更新Excel文件（仅QC Status列）
[Step 6] 用户选择输出路径和文件名
    ↓
输出: people_management_with_status.xlsx
统计: 总数/Pass数/Fail数/空值数/匹配率
```

## 📝 列映射说明

| 源文件 | 源列 | 目标文件 | 目标列 | 操作 |
|--------|------|----------|--------|------|
| tfl_status.xlsx | Dataset | people_management.xlsx | Output Name | 匹配键 |
| tfl_status.xlsx | Comparison Status | people_management.xlsx | QC Status | 合并值（预处理后） |

### 值转换规则

| 原始值（tfl_status） | 转换后值（people_management） |
|---------------------|-------------------------------|
| Match | Pass |
| Mismatch | Fail |
| (其他值) | (保持原样) |

## ⚠️ 注意事项

### 文件要求

1. **people_management.xlsx**
   - 必须包含`TLF` sheet
   - `TLF` sheet的第1行为标题，第2行为列名
   - 必须包含`Output Name`列
   - 如果不存在`QC Status (Not Started, Ongoing, QC Pending, Fail, Pass)`列，脚本会自动创建

2. **tfl_status.xlsx**
   - 必须包含`Overview` sheet
   - 必须包含`Dataset`和`Comparison Status`列

### 运行前检查

- [ ] 确保Excel中没有打开输入文件
- [ ] 确认文件路径正确
- [ ] 确认Python 3.13可用（或已正确创建 `.venv`）
- [ ] 确保有足够的磁盘空间

### 常见错误

#### 错误1: Permission denied
```
❌ 无法读取people_management文件（文件可能被Excel打开）
```
**解决方案**: 关闭Excel中打开的文件，重新运行脚本

#### 错误2: Sheet not found
```
❌ 错误：people_management文件中未找到'TLF' sheet
```
**解决方案**: 检查people_management.xlsx是否包含TLF sheet

#### 错误3: Column not found
```
❌ 错误：tfl_status文件的Overview sheet中未找到'Dataset'列
```
**解决方案**: 检查tfl_status.xlsx的Overview sheet是否包含必需列

## 📈 输出示例

运行成功后，会显示如下统计信息：

```
================================================================================
✓✓✓ 填充完成！
================================================================================
输出文件: C:\path\to\people_management_with_status.xlsx

统计信息：
  - TLF总数目: 249
  - Status为'Pass'的数目: 230
  - Status为'Fail'的数目: 15
  - Status为空的数目: 4
  - 匹配率: 245/249 (98.4%)

提示：可以直接打开Excel文件查看结果
      所有其他列和sheet都已保留，未做任何改动
```

## 🔧 技术细节

### 文件锁定处理

- 自动重试机制（3次，每次间隔1秒）
- 友好的错误提示

### 数据完整性

- 使用pandas进行数据读取和处理
- 使用openpyxl保留Excel格式和结构
- 仅更新目标列，其他数据完全保留

### 性能指标

- 读取people_management: < 2秒
- 读取tfl_status: < 1秒
- 数据处理和合并: < 2秒
- 文件保存: < 2秒

## 🆘 支持和维护

如遇问题，请检查：

1. Python版本 ≥ 3.7
   - 推荐: Python 3.13
2. 依赖包已安装（pandas, openpyxl）
3. 文件格式正确
4. 文件未被其他程序打开
5. 虚拟环境已激活

## 📚 相关文档

- [README.md](README.md) - 项目总览
- [QUICK_START.md](QUICK_START.md) - 快速开始指南
- [fill_tlf_template.py](fill_tlf_template.py) - TLF模板填充脚本

---

**创建日期**: 2026年2月11日  
**版本**: 1.0  
**状态**: ✅ 生产就绪
