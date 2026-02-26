# AegisFlow - XML Troubleshooting Guide

**Transform, Validate, Deliver from a Single TOC**

## 🔧 问题: bat文件闪退

### 步骤1: 运行诊断工具

双击运行 `test_environment.bat`

这将检查:
- ✅ Python是否已安装
- ✅ Python版本
- ✅ pandas模块
- ✅ openpyxl模块
- ✅ tkinter模块
- ✅ xml模块

### 步骤2: 根据诊断结果修复

**如果显示 pandas 未安装:**
```bash
pip install pandas
```

**如果显示 openpyxl 未安装:**
```bash
pip install openpyxl
```

**如果显示 tkinter 未安装:**
- Windows用户: 重新安装Python，确保勾选"tcl/tk and IDLE"选项
- 或者: 程序会自动切换到命令行模式（手动输入路径）

**一次性安装所有依赖:**
```bash
pip install pandas openpyxl
```

### 步骤3: 重新运行

修复后，再次双击 `run_generate_batch_xml.bat`

---

## 🔧 问题: 提示"未选择文件"

**原因**: 在文件选择窗口中点击了"取消"

**解决**: 重新运行程序，在窗口中选择文件

---

## 🔧 问题: 找不到文件选择窗口

**可能原因**:
1. 窗口在其他窗口后面 → 检查任务栏
2. tkinter未安装 → 运行诊断工具检查
3. 窗口在另一个显示器上 → 检查所有屏幕

**备用方案**:
- 程序会自动检测tkinter是否可用
- 如不可用，会切换到命令行输入模式
- 在命令行中手动输入文件路径即可

---

## 🔧 问题: 虚拟环境未找到

**现象**: bat文件显示"虚拟环境不存在，使用系统Python..."

**这是正常的**，如果:
- 系统Python已安装所需的包
- 可以正常运行程序

**如果想使用虚拟环境**:
1. 在项目目录打开命令行
2. 运行: `py -3.13 -m venv .venv`
3. 激活: `.venv\Scripts\activate`
4. 安装依赖: `pip install pandas openpyxl`

---

## 🔧 问题: Excel文件读取失败

**可能原因和解决方法**:

1. **文件正在被其他程序打开**
   - 关闭Excel或其他正在使用该文件的程序

2. **文件路径包含特殊字符**
   - 重命名文件，避免使用特殊字符

3. **文件格式不正确**
   - 确保是 .xlsx, .xls, 或 .csv 格式

4. **文件编码问题（CSV）**
   - 程序会自动尝试多种编码
   - 如果仍失败，用Excel打开后另存为UTF-8编码的CSV

---

## 🔧 问题: 列名不匹配

**错误信息**: "缺少必需的列: xxx"

**解决方法**:
1. 打开Excel文件
2. 检查列名是否完全匹配（包括大小写、空格）:
   - `sect_num`
   - `sect_ttl`
   - `OUTFILE`
   - `Output Type (Table, Listing, Figure)`
   - `tocnumber`
   - `Title`

3. 如果列名不同，修改Excel文件的列名

---

## 📞 仍然无法解决？

1. **查看完整错误信息**
   - 增强版bat文件会显示详细错误
   - 截图发送给技术支持

2. **查看完整文档**
   - [README_GENERATE_BATCH_XML.md](README_GENERATE_BATCH_XML.md)
   - [QUICK_START_GENERATE_XML.md](QUICK_START_GENERATE_XML.md)

3. **手动运行Python脚本**
   ```bash
   python generate_batch_xml.py
   ```
   查看详细错误信息

---

## ✅ 快速检查清单

运行前确认:

- [ ] Python已安装（建议3.6+）
- [ ] pandas已安装
- [ ] openpyxl已安装
- [ ] Excel/CSV文件准备好
- [ ] Excel文件包含所有必需的列
- [ ] Excel文件未被其他程序打开

**全部确认后，运行**: `run_generate_batch_xml.bat`
