# html_to_markdown.py 开发日志

**日期**：2026-05-05
**文件**：`scripts/html_to_markdown.py`

---

## 问题描述

在转换"项目风险预测与防范"章节的表格时，生成的 Markdown 文件存在两个问题：

1. 单元格行内换行处出现了 3 个 `<br>`，实际只需 1 个（甚至不需要）
2. HTML `<td></td>` 标签内的内容，转换后出现了不应有的换行

---

## 根因分析

问题均出在原 `get_cell_text` 函数：

```python
# 原代码
text = td.get_text(separator=chr(10)).strip()
return text.replace(chr(10), '<br>')
```

- `get_text(separator='\n')` 会在每个子元素之间插入换行符
- `.replace('\n', '<br>')` 将每个换行都变成一个 `<br>`，多个子元素就产生多个 `<br>`

---

## 修复方案

将 `get_cell_text` 的文本提取逻辑改为：

```python
text = td.get_text(separator=' ').strip()
text = re.sub(r'\s+', ' ', text)
return text
```

- 用空格作分隔符，避免换行
- `re.sub` 合并多余空白，保证单元格内容为单行、无 `<br>`

---

## 修改位置

`scripts/html_to_markdown.py`，`get_cell_text` 函数末尾两行（原第 19-20 行）
