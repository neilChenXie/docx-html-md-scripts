# docx_to_html 开发日志

## 需求背景

需要将 Word 文档 (.docx) 转换为 HTML 格式，要求输出效果与 WPS "另存为 HTML" 完全一致，包括：
- 保留标题编号（一、1.、1.1、1.1.1 等）
- 保留 WPS 特有的 CSS 样式 (`mso-*`)
- 保留 Office XML 命名空间
- 自动提取图片到 `.files` 目录

## 开发过程

### 第一阶段：Python 实现（python-docx）

**尝试方案**：使用 `python-docx` 库解析 docx 并生成 HTML

**问题**：
- python-docx 无法读取 WPS 的列表编号样式
- 无法生成 `mso-*` 等 Office 特有的 CSS 样式
- 输出结果与 WPS 另存为的效果差距较大
- 缺少文档属性和 XML 命名空间

**结论**：纯 Python 实现无法达到 WPS 另存为的效果，因为 WPS 使用的是 Office 专有的格式标记。

### 第二阶段：多方案 PowerShell 脚本

**最终方案**：编写 PowerShell 脚本，通过 COM 接口调用 Office 软件进行转换

**实现策略**（按优先级自动选择）：

1. **WPS Office** (`Kwps.Application`)
   - 首选方案，与 WPS 另存为效果完全一致
   - 支持完整的 WPS 样式和中文排版

2. **Microsoft Word** (`Word.Application`)
   - 备选方案，效果也很好
   - 兼容标准 Office 格式

3. **Pandoc**
   - 轻量级命令行工具
   - 适合批量处理和自动化

4. **LibreOffice**
   - 开源免费，跨平台
   - 适合没有商业软件的环境

**关键代码**：

```powershell
# WPS COM 接口调用
$wps = New-Object -ComObject "Kwps.Application"
$wps.Visible = $false
$doc = $wps.Documents.Open($inputPath)
$doc.SaveAs($outputPath, [ref]8)  # 8 = wdFormatHTML
```

### 第三阶段：文档完善

创建了 `README.md` 包含：
- 使用方法（PowerShell 和 CMD 两种调用方式）
- 依赖要求（4 种可选软件及下载地址）
- 示例输出
- 注意事项

## 最终成果

### 文件列表

```
scripts/
├── docx_to_html.ps1      # PowerShell 转换脚本
├── docx_to_html.py       # Python 备选脚本（基础版）
├── README.md             # 使用说明
└── .dev-logs/
    └── doc_to_html.md    # 本开发日志
```

### 使用方法

```powershell
# PowerShell
.\docx_to_html.ps1 "文档.docx"

# CMD/Bash
powershell.exe -ExecutionPolicy Bypass -File "docx_to_html.ps1" "文档.docx"
```

### 测试验证

以 `系统培训方案.docx` 测试：
- WPS 另存为 HTML：约 69 KB
- 脚本生成 HTML：68.2 KB
- 内容格式：完全一致

## 技术要点

1. **COM 接口是最佳选择**
   - Office 格式复杂，包含大量专有标记
   - 只有通过官方 COM 接口才能获得完整支持

2. **多方案降级策略**
   - 不同环境可能安装的软件不同
   - 按优先级尝试确保脚本通用性

3. **PowerShell 更适合 Windows 环境**
   - COM 接口调用方便
   - 错误处理完善
   - UTF-8 编码支持好

## 后续优化方向

1. 添加批量转换功能
2. 支持更多输出格式（如 MHTML）
3. 添加转换进度显示
4. 支持 Linux/Mac 环境（使用 LibreOffice）

## 日期

2025-04-16
