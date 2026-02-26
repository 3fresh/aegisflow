# MOSAIC_CONVERT - VBA宏转Python脚本

## 功能说明

将VBA宏 `MOSAIC_CONVERT` 转换为Python脚本，用于处理TiFo CSV文件并输出为Excel格式（.xlsx）。

## 主要功能

该脚本实现了原VBA宏的所有功能，并添加了多项增强：

1. **tocnumber唯一性检查**: 验证没有两张表使用相同的toc number，若检测到重复则报错并停止
2. **seq序列生成**: 从CSV按出现的outfile行自动生成seq（从1开始，每遇outfile则+1），用于转置
3. **按seq转置**: 将每个seq（对应一张图表的多行）转换为一行数据
4. **重复program+suffix检测**: 转置后检查是否有多张表使用相同的program+suffix组合，触发则警告并在输出中高亮标黄
5. **输出类型识别**: 自动识别Table、Figure或Listing类型
6. **标题提取**: 从title5字段提取干净的标题文本
7. **CSV预处理**: 仅在footnote值中替换第一个和最后一个单引号为双引号（title保持不变）
8. **Excel格式化**: 
   - 所有字体使用"等线"系列
   - 创建Index和Original两个工作表
   - 标题行加粗
   - H、I、J列标题黄色高亮
   - H2单元格冻结窗格
   - 非latin1字符标红+绿色背景（不转换原文本内容）
   - 重复program+suffix的PROGRAM/SUFFIX列标黄
   - 正确的数值排序（14.1.2在14.1.10之前）

## 关键改进

### v3.1 - OUTFILE修正和中文路径支持
- ✅ **OUTFILE数据源修正**: 直接使用CSV中`parm='outfile'`的`value`值，不再从PROGRAM+SUFFIX拼接
- ✅ **中文路径支持**: 完全解决包含中文字符的路径问题
  - `.last_output.txt`改用UTF-8编码
  - `validate_output.py`自动读取路径文件
  - `run_mosaic_convert.bat`简化验证流程
- ✅ **业务逻辑保真**: 保留CSV中的实际业务逻辑（SUFFIX不为空时OUTFILE也可能只等于PROGRAM）

### v3.0 - 按seq转置和完整的检查机制
- ✅ 新增tocnumber唯一性检查（发现重复则报错并停止）
- ✅ 采用seq序列进行转置（基于outfile行数）
- ✅ 转置后检测重复program+suffix并警告
- ✅ 重复program+suffix行在输出中高亮标黄
- ✅ 非latin1字符只标红+标绿，不转换原文本内容
- ✅ 数据行数为249(基于seq转置得出唯一行数)

### v2.5 - 数值排序和富文本格式化
- ✅ 修复tocnumber排序（从字符串排序改为数值排序）
- ✅ 实现非latin1字符的富文本格式化（只对问题字符标红）
- ✅ 绿色背景标记非latin1字符所在单元格
- ✅ 保留黄色背景用于重复program+suffix的标记

### v2.4 - 完整数据恢复和共享标记
- ✅ 从210行恢复到245行（所有表都包含）
- ✅ 使用tocnumber作为唯一标识符
- ✅ 自动过滤模板脚注占位符
- ✅ 标记31个共享的program+suffix组合（影响66个tocnumber）

### v2.3 - Footnote预处理
- ✅ 仅对footnote字段进行quote替换
- ✅ title字段的quotes保持不变

### v2.2 - 文件选择对话框
- ✅ 动态输入文件选择
- ✅ 动态输出文件保存位置

## 使用方法

### 方法1: 使用批处理文件（推荐）

```bash
run_mosaic_convert.bat
```

批处理文件会：
1. 自动打开文件选择对话框选择输入CSV文件
2. 自动生成输出XLSX文件
3. 自动验证输出文件（249行数据）
4. 完全支持包含中文字符的路径

### 方法2: 直接运行Python脚本

```bash
python mosaic_convert.py
```

脚本会打开图形化文件选择对话框：
1. 选择输入CSV文件
2. 选择输出XLSX文件位置和名称
3. 自动处理并生成结果

### 方法3: 在代码中调用

```python
from mosaic_convert import mosaic_convert

# 指定输入和输出文件
input_file = "path/to/your/input.csv"
output_file = "path/to/your/output.xlsx"

mosaic_convert(input_file, output_file)
```

### 方法4: 修改脚本中的默认路径

编辑 `mosaic_convert.py` 文件末尾的：

```python
input_file = os.path.join(script_dir, "02_output", "2026-02-09", "Clinical Study Report_TiFo.csv")
```

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

## 输入文件格式

CSV文件应包含以下6列：
- `sect_num`: 章节编号
- `sect_ttl`: 章节标题
- `program`: 程序名称
- `suffix`: 后缀
- `parm`: 参数名称（如outfile, outtype, title1, footnote1等）
- `value`: 参数值

## 输出文件结构

生成的Excel文件包含两个工作表：

### Index工作表（主要结果）
包含以下列：
- sect_num, sect_ttl, outtype, azsolid, Core, tocnumber
- Output Type (Table, Listing, Figure), Title
- PROGRAM, SUFFIX, OUTFILE
- title1-7, footnote1-9

**格式化特性**：
- **所有内容**：使用"等线"字体
- **表头行**：加粗显示
- **列H-J**：黄色背景
- **非latin1字符**：红色字体 + 绿色背景
- **共享program+suffix**：tocnumber/PROGRAM/SUFFIX列黄色背景
- **排序**：按tocnumber的数值排序（14.1.2在14.1.10之前）

### Original工作表
保留原始CSV数据的6列结构

## 数据特性说明

### 转置逻辑
- 基于CSV中'outfile'行来定义seq（每遇outfile就seq+1）
- 按seq转置生成最终数据框架
- 共有**249张独特的表格**（基于seq的唯一行数）
- 8行存在重复program+suffix组合（在输出中标黄予以警告）

### 示例
| 组合 | Tocnumber数量 | 示例tocnumber |
|-----|-------------|-------------|
| t_dm + dischar_itt | 3 | 14.1.9.1, 14.1.9.2, 14.1.9.3 |
| f_km + pfs_itt | 4 | 14.2.1.3.1-14.2.1.3.4 |

### 非latin1字符处理
- 自动检测包含非ASCII字符的单元格（如"–"破折号）
- 仅对问题字符使用红色字体（保留原文本内容）
- 整个单元格使用绿色背景便于识别
- 字符被标红但不会被转换或删除

## 排序算法

使用数值排序而不是字符串排序：
```
正确顺序：14.1.1 → 14.1.2 → 14.1.10 → 14.2
字符排序：14.1.1 → 14.1.10 → 14.1.2 → 14.2  (错误)
```

实现逻辑：
1. 将每个tocnumber按"."分割
2. 将各部分转换为整数
3. 按数值进行元组比较排序

## 输出文件命名

- 如果未指定输出文件名，默认在输入文件同目录下生成
- 文件名格式：`<原文件名>_MOSAIC_CONVERT.xlsx`

例如：
- 输入：`Clinical Study Report_TiFo.csv`
- 输出：`Clinical Study Report_TiFo_MOSAIC_CONVERT.xlsx`

## 与VBA宏的对比

| 功能 | VBA宏 | Python脚本 | 说明 |
|------|-------|-----------|------|
| 输入格式 | CSV (需手动导入Excel) | CSV | 直接读取CSV |
| 输出格式 | xlsm (宏文件) | xlsx | 标准Excel文件 |
| 执行速度 | 较慢 | 快速 | Python处理更高效 |
| 依赖 | Excel应用程序 | Python + pandas | 无需安装Excel |
| 自动化 | 需手动点击运行 | 命令行执行 | 易于自动化 |

## 注意事项

1. 确保CSV文件编码为UTF-8（如包含中文字符）
2. 输出目录必须有写入权限
3. 如果输出文件已存在，会被覆盖
4. 脚本会输出处理的行数和生成的索引行数

## 错误处理

如果脚本运行出错，请检查：
- [ ] CSV文件路径是否正确
- [ ] CSV文件格式是否符合要求（6列）
- [ ] 是否已安装所需的Python包
- [ ] 输出目录是否存在且可写

## 示例输出

```
✓ 转换完成！输出文件: Clinical Study Report_TiFo_MOSAIC_CONVERT.xlsx
  - 处理了 3211 行原始数据
  - 生成了 210 行索引数据

成功！请查看输出文件：Clinical Study Report_TiFo_MOSAIC_CONVERT.xlsx
```

## 技术细节

- 使用pandas进行数据转换和分组操作
- 使用openpyxl进行Excel格式化
- 保留了VBA宏的所有逻辑步骤
- 优化了数据查找和匹配算法

## 错误处理

### 停止条件

1. **tocnumber重复错误**
   ```
   ERROR: 有图表使用同样的toc number
   ```
   - 程序立即停止，不生成输出文件
   - 需要修复CSV中的重复tocnumber

2. **seq序列生成失败**
   ```
   ERROR: Missing required 'outfile' rows to build seq
   ```
   - CSV缺少outfile行
   - 需要检查CSV结构

### 警告

**program+suffix重复警告**
```
ERROR: 有图表使用同样的program+suffix,请更新MOSAIC
```
- 程序继续运行，生成输出文件
- 重复行在PROGRAM和SUFFIX列用黄色标注
- 需要更新MOSAIC宏中的program或suffix定义

## 相关工具

- **extract_programs.py**: 从生成的Excel文件中提取PROGRAM列表，生成SAS脚本（见README_EXTRACT_PROGRAMS.md）
- **fill_tlf_template.py**: 将MOSAIC输出与人员数据合并（见README.md）
- **validate_output.py**: 验证Excel文件结构
- **run_mosaic_convert.bat**: 运行MOSAIC_CONVERT转换

## 版本信息

- Python版本：3.7+
- pandas版本：1.0+
- openpyxl版本：3.0+

---

**作者**: AI优化自VBA宏  
**日期**: 2026-02-11  
**版本**: v3.0 (seq转置+完整检查)  
**用途**: TiFo数据处理和MOSAIC格式转换
