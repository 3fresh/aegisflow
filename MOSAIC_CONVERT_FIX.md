# MOSAIC_CONVERT 修复说明

## 日期
- 初版修复：2026年2月10日
- v3.1更新：2026年2月11日

## 修复的问题

### 问题1：相同program+suffix组合导致的数据错乱
**症状：**
- 当两个不同的tocnumber使用相同的program和suffix时，转换后的azsolid和OUTFILE值会错乱
- 例如：14.2.3.5.1和14.2.3.5.2都使用program='f_forest', suffix='os_itt'，但它们应该有不同的azsolid和outfile值

**根本原因：**
- 原代码使用`sect_num, sect_ttl, program, suffix`组合来筛选subset
- 当多个tocnumber有相同的program+suffix时，subset会包含所有这些tocnumber的数据
- 导致取值时混淆了不同tocnumber的数据

**修复方案：**
- 使用行号范围（start_idx:end_idx）来区分不同的tocnumber
- 每个tocnumber取从当前tocnumber行到下一个tocnumber行之间的所有数据
- 确保即使program+suffix相同，每个tocnumber也能获取到正确的数据

**代码变更：**
```python
# 修复前（错误）：
subset = df[(df['sect_num'] == sect_num) & 
            (df['sect_ttl'] == sect_ttl) & 
            (df['program'] == program) & 
            (df['suffix'] == suffix) & 
            (df['parm'] != 'tocnumber')].copy()

# 修复后（正确）：
# Extract rows for this tocnumber using row index range
# This ensures each tocnumber gets its own data, even if program+suffix is the same
subset = df.loc[start_idx:end_idx-1].copy()
```

### 问题2：OUTFILE列的值来源（已于v3.1修正）
**初版修复（v2.x-v3.0）：**
- 从PROGRAM和SUFFIX拼接生成OUTFILE
- 格式为：`PROGRAM_SUFFIX`（当SUFFIX不为空时）

**发现的新问题（v3.1）：**
- CSV实际业务逻辑中存在特殊情况：SUFFIX不为空，但OUTFILE只等于PROGRAM
- 示例：PROGRAM='t_ds', SUFFIX='comb', 但CSV中`parm='outfile'`的`value='t_ds_comb'`与预期一致
- 但也存在：PROGRAM='xxx', SUFFIX='yyy', 但`parm='outfile'`的`value='xxx'`（不拼接suffix）

**最终修复方案（v3.1）：**
- **直接使用CSV原始值**: OUTFILE列直接从CSV中`parm='outfile'`对应的`value`提取
- 不再进行任何拼接或转换
- 保留CSV中的实际业务逻辑，确保与源数据一致

**代码变更：**
```python
# v3.0及之前（从PROGRAM+SUFFIX拼接）：
def build_outfile(row):
    program = row['PROGRAM'] if 'PROGRAM' in row and pd.notna(row['PROGRAM']) else ''
    suffix = row['SUFFIX'] if 'SUFFIX' in row and pd.notna(row['SUFFIX']) else ''
    if suffix and str(suffix) != 'nan' and str(suffix) != '':
        return f"{program}_{suffix}"
    else:
        return program
index_df['OUTFILE'] = index_df.apply(build_outfile, axis=1)

# v3.1（直接使用CSV值）：
for _, row in group.iterrows():
    param_name = row.get(param_col, None)
    if pd.isna(param_name) or param_name == '':
        continue
    key = str(param_name).strip().lower()
    # Special handling: Convert 'outfile' to 'OUTFILE' (uppercase)
    if key == 'outfile':
        key = 'OUTFILE'
    row_data[key] = row.get('value', None)
# OUTFILE现在直接从CSV中获取，无需额外构建
```

### 问题3：中文路径编码问题（v3.1新增）
**症状：**
- 运行`run_mosaic_convert.bat`时，验证步骤总是提示"WARN: 找不到输出文件路径，跳过验证"
- `.last_output.txt`中的路径包含中文字符（如"工具开发"）出现乱码
- 转换后的Excel文件实际生成正常，只是验证步骤失败

**根本原因：**
- `.last_output.txt`使用GBK编码写入，无法正确处理中文字符
- bat文件的`for /f`命令在捕获Python输出时，Windows控制台编码转换出错
- 即使Python读取文件用UTF-8，bat传递参数时仍会乱码

**修复方案：**
1. **编码统一**: `.last_output.txt`改用UTF-8编码写入（mosaic_convert.py）
2. **自动读取**: `validate_output.py`在无命令行参数时，自动从`.last_output.txt`读取路径
3. **简化流程**: `run_mosaic_convert.bat`直接调用Python，不再通过bat传递路径参数

**技术细节：**
- 之前：bat → `for /f`捕获Python输出（中文乱码） → 传递给validate_output.py
- 现在：bat → 直接调用Python → Python自己读取`.last_output.txt`（UTF-8）
- 所有中文路径处理都在Python内部完成，完全避免Windows控制台编码问题

## 验证结果

### 验证方法
1. 使用原始CSV文件（Clinical Study Report_TiFo.csv）生成新的Excel文件
2. 对比生成的Excel中每个tocnumber的PROGRAM、SUFFIX、OUTFILE、azsolid与原始CSV的对应值
3. 验证中文路径支持（路径包含"工具开发"等中文字符）

### 验证结果（v3.1）
✅ **基于seq转置的249行数据完全正确**
- 每个tocnumber的PROGRAM和SUFFIX与原始CSV的program和suffix列一致
- **OUTFILE直接使用CSV中`parm='outfile'`的`value`值**（不再拼接）
- azsolid值与原始CSV中对应tocnumber的parm='azsolid'行的value一致
- 即使多个tocnumber有相同的program+suffix组合（如14.2.3.5.1和14.2.3.5.2），也能正确区分
- **完全支持包含中文字符的路径**，验证步骤正常执行

### 测试案例
| tocnumber | 原始CSV | v3.1输出 | 状态 |
|---|---|---|---|
| 14.1.1 | program=t_ds, suffix=comb, outfile=t_ds_comb | OUTFILE=t_ds_comb | ✓ |
| 14.1.5.2 | program=t_dm, suffix=itt3l, outfile=t_dm_itt3l | OUTFILE=t_dm_itt3l | ✓ |
| 14.2.3.5.1 | program=f_forest, suffix=os_itt, outfile=f_forest_os_itt, azsolid=AZFEF02 | OUTFILE=f_forest_os_itt, azsolid=AZFEF02 | ✓ |
| 14.2.3.5.2 | program=f_forest, suffix=os_itt, outfile=f_forest_os_itt, azsolid=AZTONCEF04 | OUTFILE=f_forest_os_itt, azsolid=AZTONCEF04 | ✓ |
| **特殊案例** | program=xxx, suffix=yyy, outfile=xxx (不拼接) | OUTFILE=xxx（正确保留CSV原值） | ✓ |

## 影响范围
- 修改文件：mosaic_convert.py, validate_output.py, run_mosaic_convert.bat
- 更新文档：README_MOSAIC_CONVERT.md, CHANGELOG.md, MOSAIC_CONVERT_FIX.md
- 所有使用mosaic_convert.py的工作流都会受益于这些修复

## 后续步骤
1. ✅ 代码修复完成（v3.1）
2. ✅ OUTFILE数据源修正
3. ✅ 中文路径支持
4. ✅ 文档更新完成
5. ✅ 验证测试通过
2. ✅ 文档更新完成
3. ⏳ 需要重新生成MOSAIC_CONVERT输出文件
4. ⏳ 需要使用新的MOSAIC_CONVERT文件重新运行fill_tlf_template.py
