---
description: "Translate all Chinese text in source files to English. Use for code comments, string literals, print statements, UI labels, and documentation. File paths containing Chinese characters are allowed and must NOT be changed."
name: "Translate to English"
argument-hint: "File or folder to translate (e.g. aegisflow_app.py, or leave blank for all files)"
agent: "agent"
tools: [read_file, grep_search, file_search, replace_string_in_file, multi_replace_string_in_file]
---

Translate all Chinese text in the specified file(s) to English. If no file is specified, scan the entire workspace.

## Rules

1. **Translate**: All Chinese characters in code comments, string literals, `print()` / `logging` messages, UI label text, docstrings, and Markdown documentation.
2. **Do NOT translate**: File paths, directory names, or any string whose sole purpose is to reference a filesystem path — even if that path contains Chinese characters.
3. **Preserve**: Code structure, indentation, logic, variable names, and all non-Chinese content exactly as-is.
4. **Accuracy**: Translations must be natural, technically precise English appropriate for software development context.

## Process

1. Search the target file(s) for any text matching Chinese Unicode range `[\u4e00-\u9fff]`.
2. For each match, determine if it's in a file path context — if so, skip it.
3. Translate the Chinese text to English in place.
4. After all edits, verify no Chinese characters remain (outside of file paths).

## Example

**Before:**
```python
# 填充空值
print("  根据Excel生成Batch List XML工具")
```

**After:**
```python
# Fill empty values
print("  Batch List XML generation tool based on Excel")
```
