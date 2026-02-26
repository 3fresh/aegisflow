# EXTRACT_PROGRAMS - SAS程序列表提取工具

## 功能说明

从MOSAIC_CONVERT生成的Excel文件中提取PROGRAM列表，统计分析后生成SAS运行脚本格式的文本文件。

## 主要功能

1. **Excel文件分析**: 读取MOSAIC_CONVERT生成的Excel文件（Index工作表）
2. **程序统计**: 统计每个唯一PROGRAM值及其对应的tocnumber数量
3. **顺序保留**: 按Excel中首次出现的顺序排列（保持原始顺序）
4. **脚本生成**: 生成标准的%runpgm格式的SAS脚本，注释统一在顶部
5. **图形界面**: 提供文件选择对话框，无需手动修改路径

## 输出格式

生成的文本文件格式示例：

```sas
/* Generated SAS Program Execution Script */
/* Programs ordered by first appearance in Excel file */

/* Program Statistics: */
/*   t_ds: 3 table(s) */
/*   t_aztoncsp16: 2 table(s) */
/*   t_dm: 2 table(s) */
/*   f_forest: 1 table(s) */

/* ====== Program Execution Commands ====== */

%runpgm(pgm=t_ds, error_override=y);
%runpgm(pgm=t_aztoncsp16, error_override=y);
%runpgm(pgm=t_dm, error_override=y);
%runpgm(pgm=f_forest, error_override=y);
```

所有程序统计信息集中在文件顶部，执行命令在统计信息之后。

## 使用方法

### 方法1: 使用批处理文件（推荐）

```bash
run_extract_programs.bat
```

批处理文件会：
1. 检查Python虚拟环境
2. 打开文件选择对话框，选择输入Excel文件
3. 自动分析并显示统计信息
4. 打开保存对话框，选择输出文件位置和名称
5. 生成SAS脚本文件

### 方法2: 直接运行Python脚本

```bash
python extract_programs.py
```

## 工作流程

```
1. 选择Excel文件
   ↓
2. 读取Index工作表
   ↓
3. 提取PROGRAM和tocnumber列
   ↓
4. 统计每个PROGRAM的tocnumber数量
   ↓
5. 按Excel中首次出现顺序排列
   ↓
6. 生成注释统计信息（集中在顶部）
   ↓
7. 生成%runpgm执行命令
   ↓
8. 保存到指定位置
```

## 输出统计信息

程序运行时会在控制台显示：

```
================================================================================
PROGRAM统计信息 (按Excel中首次出现的顺序)
================================================================================
序号   PROGRAM                        图表数量  
--------------------------------------------------------------------------------
1      t_ds                           3         
2      t_aztoncsp16                   2         
3      t_dm                           2         
4      f_forest                       1         
...
--------------------------------------------------------------------------------
总计: 85 个唯一程序
================================================================================
```

## 输入要求

- **文件格式**: MOSAIC_CONVERT生成的.xlsx文件
- **必需工作表**: Index
- **必需列**: 
  - `PROGRAM`: SAS程序名称
  - `tocnumber`: 目录编号

## 输出说明

- **文件格式**: .txt文本文件
- **编码**: UTF-8
- **行格式**: `%runpgm(pgm=<程序名>, error_override=y);`
- **排序**: 按Excel中首次出现的顺序（保持原始顺序）
- **注释**: 所有程序统计信息集中在文件顶部

## 应用场景

1. **批量运行SAS程序**: 将生成的脚本复制到SAS中批量执行
2. **顺序执行**: 按Excel中的顺序执行，符合业务逻辑
3. **程序清单管理**: 快速了解项目中所有使用的SAS程序及统计信息
4. **质量控制**: 检查是否有遗漏或重复的程序

## 依赖包

```bash
pip install pandas openpyxl
```

或使用虚拟环境：
```bash
py -3.13 -m venv .venv
.venv\Scripts\activate
pip install pandas openpyxl
```

## 示例输出

假设Excel文件中有以下数据：

| tocnumber | PROGRAM | SUFFIX |
|-----------|---------|--------|
| 14.1.1    | t_ds    | comb   |
| 14.1.2    | t_ds    | pop    |
| 14.1.3    | t_ds    | itt    |
| 14.2.1    | t_dm    | itt    |
| 14.2.2    | t_dm    | saf    |

生成的脚本将是：

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

## 注意事项

1. **程序去重**: 自动去除重复的PROGRAM值
2. **空值处理**: 跳过PROGRAM或tocnumber为空的行
3. **中文路径支持**: 完全支持包含中文字符的文件路径
4. **顺序保留**: 严格按照程序在Excel中首次出现的顺序输出
5. **注释分离**: 所有统计信息注释集中在文件顶部，便于阅读

## 错误处理

程序会处理以下错误情况：

- ❌ **找不到Index工作表**: 提示错误并退出
- ❌ **缺少PROGRAM列**: 提示错误并退出
- ❌ **缺少tocnumber列**: 提示错误并退出
- ❌ **文件读取失败**: 显示详细错误信息

## 版本历史

### v1.0 (2026-02-12)
- ✅ 初始版本发布
- ✅ 支持Excel文件分析
- ✅ 按Excel首次出现顺序排列
- ✅ 注释统一放在文件顶部
- ✅ 生成%runpgm格式脚本
- ✅ 图形文件选择界面
- ✅ 中文路径支持

## 相关工具

- **mosaic_convert.py**: 将CSV转换为Excel（生成EXTRACT_PROGRAMS的输入文件）
- **validate_output.py**: 验证Excel文件结构
- **run_mosaic_convert.bat**: 运行MOSAIC_CONVERT转换

## 技术实现

- **语言**: Python 3.x
- **核心库**: pandas, openpyxl, tkinter
- **输入**: Excel工作簿 (.xlsx)
- **输出**: 文本文件 (.txt)
- **界面**: Tkinter图形对话框
