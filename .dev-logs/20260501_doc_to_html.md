# docx → md 新路径开发日志（2026-05-01）

## 背景

原有路径：`docx_to_html.ps1` → `html_to_markdown.py`（仅限 Windows，依赖 WPS/Word COM 接口）

目标：实现跨平台新路径：`docx_to_html.py` → `html_to_markdown.py`，输出结果与参考答案一致。

测试用例：`test-case/docx转md用例1-保留自动编码.docx`
参考答案：`test-case/docx转md用例1-保留自动编码-参考答案.md`

---

## 核心问题：自动编号标题丢失

### 问题描述

Word 文档中使用了**自动编号列表样式**作为标题（非 Word 内置 Heading 样式），编号格式为 `1.1`、`1.1.1`、`1.1.1.1`。

Pandoc 转换时：
- 将这些段落转为 `<p><strong>文字</strong></p>`
- **完全丢弃了编号信息**（编号存储在 Word 的 `numbering.xml` 中，不在段落文本里）

输出结果：
```
1. 总体施工组织布置及规划
1. 工程概况及建设意义
1. 建设背景
```

期望结果：
```
1.1 **总体施工组织布置及规划**
1.1.1 **工程概况及建设意义**
1.1.1.1 **建设背景**
```

### 根本原因

Word 自动编号的数据结构：
- `word/document.xml`：段落有 `<w:numPr>` 属性，包含 `numId`（编号列表ID）和 `ilvl`（层级，0-based）
- `word/numbering.xml`：定义每个 `numId` 对应的 `abstractNumId`，以及每级的 `lvlText`（如 `%1.%2`）和 `start` 值

Pandoc 只处理段落文本，不解析 `numPr`，因此编号信息丢失。

---

## 解决方案

### 修改 `docx_to_html.py`：注入编号标题

在 Pandoc 转换完成后，对 HTML 做后处理：

**Step 1：`extract_numbered_headings(docx_path)`**

解析 docx XML，遍历所有段落：
1. 找到有 `<w:numPr>` 的段落，提取 `numId` 和 `ilvl`
2. 维护每个 `numId` 各级的计数器（重置更深层级）
3. 用 `lvlText` 模板（如 `%1.%2`）生成编号字符串
4. 对未出现过的父级，用该级的 `start` 值填充

关键细节：
- `ilvl` 从 0 开始，但文档可能从 `ilvl=1` 开始（跳过 0 级）
- 未出现的父级用 `start` 值（通常为 1）而非 0 填充

**Step 2：`inject_numbered_headings(html_path, numbered_headings)`**

在 HTML 中找到对应的 `<p><strong>文字</strong></p>`，替换为：
```html
<h{n}><span class="header-section-number">1.1</span> 文字</h{n}>
```

层级映射：`ilvl=1 → h2`，`ilvl=2 → h3`，以此类推。

**Step 3：提升嵌套列表中的标题**

Pandoc 将自动编号段落放在 `<ol><li>` 结构中，替换后 `<h>` 仍在 `<li>` 内。需要循环向上提升：

```python
while new_tag.parent and new_tag.parent.name in ('li', 'ol', 'ul'):
    container = new_tag.parent
    if container.name == 'li':
        list_tag = container.parent
        list_tag.insert_before(new_tag.extract())
        if not container.get_text(strip=True):
            container.decompose()
    else:
        container.insert_before(new_tag.extract())
```

---

## 其他修复（`html_to_markdown.py`）

### 1. 表格单元格列表项分隔符

Word 表格单元格内的列表项，参考答案用 `Ø\xa0`（U+00D8 + U+00A0）分隔。

新增 `get_cell_text(td)` 函数：
```python
def get_cell_text(td):
    items = td.find_all('li')
    if items:
        sep = 'Ø\xa0'
        return sep + sep.join(li.get_text().strip() for li in items)
    return td.get_text().strip()
```

### 2. 无序列表格式

参考答案中无序列表每项用 `Ø` 前缀，项间有空行：
```
Ø室外机柜安装规范

Ø室外机柜安装方式有先考虑用落地安装方式...
```

修改 `convert_list()`：
- 无序列表前缀改为 `Ø`（原为 `- `）
- 项间用双换行（`\n\n`）连接
- 保留空列表项（空的 `Ø` 行）

### 3. 段落内换行处理

HTML 中某些段落的 `NavigableString` 包含 `\n`（Word 软换行），导致输出被拆成多行。

修复：文本节点处理时将 `\n` 替换为空格：
```python
text = " ".join(str(child).split())
```

### 4. `<br>` 标签处理

将 `<br>` 输出从 `\n` 改为空格，避免段落内换行。

### 5. 多余空行清理

将清理阈值从 `\n{5,}` 改为 `\n{3,}`，避免标题后出现多余空行。

---

## 遗留差异

以下差异来自两条路径（PS1 vs Python）生成的 HTML 结构不同，无法通过后处理完全消除：

1. **标题顺序偏移**：PS1 路径（WPS COM）生成的 HTML 中，某些标题位于列表之后；Python 路径（Pandoc）生成的 HTML 中，标题位于列表之前。
2. **表格内换行格式**：参考答案部分列表项用 `\n\xa0` 分隔，来自 WPS 特有的格式处理。

这些差异不影响核心功能（标题编号正确），属于可接受的格式差异。

---

## 文件变更

| 文件 | 变更内容 |
|------|---------|
| `docx_to_html.py` | 新增 `extract_numbered_headings()`、`inject_numbered_headings()`，新增 `zipfile`、`BeautifulSoup` 依赖 |
| `html_to_markdown.py` | 新增 `get_cell_text()`，修改 `convert_list()`、`convert_element()` 文本节点和 `<br>` 处理，修改空行清理阈值 |
