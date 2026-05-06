# 2026-05-06 开发日志：项目结构重构 & pyproject.toml 引入

## 背景与目标

将散落在项目根目录的 Python 脚本重构为标准 `src/` 布局的 Python 包，引入 `pyproject.toml` 管理元数据、依赖和 CLI 入口点，使得工具既可作为独立命令使用，也可作为第三方库被 import。

## 变更摘要

| 变更类型 | 文件 | 说明 |
|---------|------|------|
| 新增 | `pyproject.toml` | 包元数据、依赖声明、CLI 入口点 |
| 新增 | `src/docx_md_tools/__init__.py` | 包初始化，导出公共 API |
| 新增 | `.gitignore`（追加） | 新增 `dist/`、`*.egg-info/` 忽略规则 |
| 移动 | `docx_to_html.py` → `src/docx_md_tools/docx_to_html.py` | 纳入包目录 |
| 移动 | `html_to_markdown.py` → `src/docx_md_tools/html_to_markdown.py` | 纳入包目录 |
| 移动 | `md_to_docx.py` → `src/docx_md_tools/md_to_docx.py` | 纳入包目录 |
| 移动 | `docx_to_html.ps1` → `scripts/docx_to_html.ps1` | 脚本归入 scripts/ |
| 删除 | `requirements.txt` | 依赖已迁移至 pyproject.toml |
| 更新 | `README.md` | 反映新项目结构和三种使用方式 |

## 详细变更

### 1. pyproject.toml（新增）

- 构建系统：`hatchling`
- 包名：`docx-md-tools`，版本 `0.1.0`
- Python >= 3.9，MIT 协议
- 依赖：beautifulsoup4, chardet, lxml, pypandoc, python-docx
- CLI 入口点（`[project.scripts]`）：
  - `docx2html` → `docx_md_tools.docx_to_html:main`
  - `html2md` → `docx_md_tools.html_to_markdown:main`
  - `md2docx` → `docx_md_tools.md_to_docx:main`

### 2. src/docx_md_tools/ 包结构

**`__init__.py`** — 导出三个公共函数：
- `docx_to_html` — docx → HTML 转换
- `html_to_markdown` — HTML → Markdown 转换
- `markdown_to_docx` — Markdown → docx 转换

**`docx_to_html.py`** — 重构：
- 拆分 `main()` 和 `docx_to_html()`：核心逻辑提取为独立函数，可被 import 调用
- `docx_to_html()` 参数：`input_file`（str）、`output_file`（str=""）
- `docx_to_html()` 返回值：成功返回 `output_file` 路径
- `docx_to_html()` 抛异常：`FileNotFoundError`（文件不存在）、`ValueError`（非 .docx 后缀）
- `main()` 仅负责 CLI 参数解析和异常处理

**`html_to_markdown.py`** — 重构：
- `html_to_markdown()` 返回值从 `bool` 改为 `md_path`（str）或 `None`（失败）
- `convert_table()` / `convert_list()` / `convert_element()` 移除中文注释
- 所有 print 消息英文化

**`md_to_docx.py`** — 重构：
- 移除中文注释
- `main()` 使用 argparse，帮助文本英文化
- 修复导入间距，代码格式化

### 3. 项目文件结构（新）

```
docx-html-md-scripts/
├── pyproject.toml
├── src/docx_md_tools/
│   ├── __init__.py
│   ├── docx_to_html.py
│   ├── html_to_markdown.py
│   └── md_to_docx.py
├── scripts/
│   └── docx_to_html.ps1
├── README.md
├── .gitignore
├── .dev-logs/
└── test-case/
```

### 4. README.md 更新

- 目录结构更新为 src/ 布局
- 安装说明增加三种方式：本地 editable、Git 安装、作为库导入
- 使用方式增加三种途径：CLI 命令、Python import、直接运行模块文件
- 新增版本管理与更新同步章节
- PowerShell 路径从根目录更新为 scripts/

## 影响与权衡

### 优点
- Standard src layout，符合 Python 社区最佳实践
- `pip install -e .` 开发模式：修改代码立即生效
- 可作为第三方库被其他项目 import
- CLI 命令名简洁（`docx2html`、`html2md`、`md2docx`）
- 删除 `requirements.txt` 避免与 `pyproject.toml` 维护两份依赖列表

### 注意事项
- 旧路径 `docx_to_html.py` 等的直接调用方式不再生效，需使用 CLI 命令或新路径
- PowerShell 脚本路径从 `docx_to_html.ps1` 变为 `scripts/docx_to_html.ps1`
- `html_to_markdown()` 返回值从 `bool` 变为 `str | None`，对已有调用方有 breaking change
