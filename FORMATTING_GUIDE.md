# Excel 格式化指南 (v2.5)

## 📋 概览

MOSAIC_CONVERT 输出的Excel文件现已包含完整的格式化和智能标记功能。本指南详细说明每种格式化的含义和应用场景。

## 🎨 格式化类型

### 1. 字体格式

#### 全局字体：等线
- **应用范围**：所有单元格
- **目的**：统一外观，提高可读性
- **功能**: 
  - 表头：等线 + 加粗
  - 数据：等线
  - 富文本部分（内含红字时）：等线 + 红色

### 2. 背景颜色

#### 黄色背景 (FFFF00)

**应用于**：
- H1:J1 (表头的3列)
- 共享program+suffix的tocnumber/PROGRAM/SUFFIX列

**含义**：
- 这些列需要特别关注
- 某些表格共享相同的program+suffix组合（共31个组合）
- 共66个tocnumber受到影响

**示例**：
| Tocnumber | Program | Suffix | 说明 |
|-----------|---------|--------|------|
| 14.1.5.2 (黄) | t_dm (黄) | itt3l (黄) | 与14.1.5.3共享 |
| 14.1.5.3 (黄) | t_dm (黄) | itt3l (黄) | 与14.1.5.2共享 |
| 14.1.9.1 | t_dm | dischar_itt | 唯一 |

#### 绿色背景 (92D050)

**应用于**：
- 包含非latin1字符的单元格

**含义**：
- 此单元格包含字符编码问题
- 问题字符已标红
- 需要手动检查和修正

**示例**：
- "Primary Endpoint – Progression" (绿色背景，"–"为红色)
- 可能的字符问题：
  - 破折号（–、—）
  - 特殊符号（©、®、™）
  - 非ASCII字符

### 3. 字体颜色

#### 红色字体

**应用于**：
- 非latin1字符（仅该字符）

**含义**：
- 标记编码问题的具体字符
- 需要替换为等效的ASCII字符

**替换建议**：
| 问题字符 | 推荐替换 |
|---------|--------|
| – (en dash) | - (hyphen-minus) |
| — (em dash) | - (hyphen-minus) |
| ' (right single quote) | ' (apostrophe) |
| " (left/right double quote) | " (quotation mark) |

## 📊 排序方式

### 数值排序 (Numeric Sort)

**排序规则**：
1. 按"."分割tocnumber
2. 将每部分转换为整数
3. 递归比较，从左到右

**示例序列**：
```
14.1.1
14.1.2
14.1.2.1
14.1.2.2
14.1.3
14.1.10    ← 10在3之后（数值排序）
14.1.10.1
14.2
14.2.1.1.1
14.2.1.1.2
14.2.1.2
...
```

**对比（字符排序 - 错误）**：
```
14.1.1
14.1.10    ← 字符排序会把10放在1之后
14.1.2     ← 这是错误的！
14.1.3
```

## 🔍 使用示例

### 识别共享的表格

**场景**：10张表需要使用相同的program+suffix

```
1. 搜索所有黄色背景的行
2. 查看同一program+suffix下的所有tocnumber
3. 这些tocnumber共享结构

示例 (program=t_dm, suffix=dischar_itt):
- 14.1.9.1 (黄)
- 14.1.9.2 (黄)
- 14.1.9.3 (黄)
```

### 修复非latin1字符

**场景**：发现单元格为绿色背景

```
1. 定位绿色单元格
2. 查看红色字符
3. 根据上表进行替换
4. 重新导入或手动修正
```

## 📝 数据统计

### 当前数据

| 类型 | 数量 | 备注 |
|------|------|------|
| 总表格 | 249 | 基于seq序列转置 |
| 原始行 | 3,211 | CSV格式 |
| 共享组合 | 31 | program+suffix |
| 受影响行 | 66 | 共享tocnumber |
| 标红单元格 | 59 | 非latin1字符 |

### 常见的共享组合

| Program | Suffix | Tocnumber数量 | 示例 |
|---------|--------|-------------|------|
| t_dm | dischar_itt | 3 | 14.1.9.1-3 |
| f_km | pfs_itt | 4 | 14.2.1.3.1-4 |
| t_dm | itt3l | 2 | 14.1.5.2-3 |
| t_cm | subct_itt3l | 2 | 14.1.18.2-3 |

## 💡 最佳实践

### 1. 快速浏览
- 使用黄色背景快速定位需要特殊处理的行
- 使用绿色背景快速定位编码问题

### 2. 批量处理
- 按照黄色标记进行分类
- 按照program+suffix进行批量处理

### 3. 数据验证
- 在绿色单元格处停顿，检查是否需要修正
- 验证黄色标记行是否有缺失

### 4. 排序验证
- 检查tocnumber序列是否正确
- 确保14.1.10在14.1.2之后

## 🔧 技术细节

### 富文本实现

使用openpyxl的RichText功能：

```python
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont

# 创建带有红色字体的特定字符
red_font = InlineFont(rFont='等线', color='FF0000')
default_font = InlineFont(rFont='等线')

# 组织成RichText
rich_text = CellRichText(
    TextBlock(default_font, "Primary Endpoint "),
    TextBlock(red_font, "–"),
    TextBlock(default_font, " Progression")
)
```

### 排序实现

```python
def tocnumber_sort_key(tocnum):
    """将tocnumber转换为可排序的元组"""
    if pd.isna(tocnum):
        return (float('inf'),)
    parts = str(tocnum).split('.')
    return tuple(
        int(p) if p.isdigit() else float('inf') 
        for p in parts
    )

# 应用排序
index_final['_sort_key'] = index_final['tocnumber'].apply(tocnumber_sort_key)
index_final = index_final.sort_values('_sort_key')
```

## ❓ 常见问题

### Q: 为什么有些表格被标记为黄色？
**A**: 这表示它们与其他表格共享相同的program+suffix组合。这是数据的特点，不是错误。

### Q: 非latin1字符可以自动转换吗？
**A**: 不能。这些字符需要手动查看和修正。使用红色标记帮助识别。

### Q: 排序规则是什么？
**A**: 使用数值排序，不是字符排序。所以14.1.2在14.1.10之前。

### Q: 可以修改背景颜色吗？
**A**: 可以。但建议保留这些标记以便快速识别数据特性。

### Q: 表头为什么也是黄色？
**A**: H、I、J列表头都是黄色，提示这些列很重要且经常被标记。

## 📞 支持

如有格式化相关问题，请参考：
- README_MOSAIC_CONVERT.md - 完整文档
- UPDATE_SUMMARY.md - 功能总结
- CHANGELOG.md - 版本历史

---

**版本**: 2.5  
**日期**: 2026-02-10  
**最后更新**: 添加数值排序和富文本格式化
