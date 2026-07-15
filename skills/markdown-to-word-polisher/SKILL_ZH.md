---
name: markdown-to-word-polisher
description: 在 Windows 上进行 Markdown 与 Word .docx 双向转换。用户要求 Markdown 转为抛光后的 Word、Word 转 Markdown 并导出图片资源、基于 Word 模板转换、检查样式映射或验证转换后的 Word 文档时使用。Markdown 转 Word 必须先运行 mdtodocx.exe，再立即运行内置 Word COM 后处理脚本后才能交付。
---
# Markdown To Word Polisher

把这个 skill 当作流程指引，而不是独立功能清单。

Markdown 转 Word 的默认交付流程是：

```text
Markdown 输入
-> 选择模板和 style-map
-> bin/mdtodocx.exe to-docx
-> 立即运行 Word COM 后处理
-> 验证 .docx
-> 返回最终文件路径
```

Word 转 Markdown 的流程是：

```text
.docx 输入
-> bin/mdtodocx.exe to-md
-> Markdown + 导出的 assets
-> 验证图片链接、表格和公式语法
```

所有 PowerShell 脚本都必须兼容 Windows PowerShell 5.1。运行脚本时优先使用 `powershell.exe -ExecutionPolicy Bypass -File ...`。

## 资源

- `bin/mdtodocx.exe`：自包含 C# 转换器，支持 `to-docx` 和 `to-md`。此二进制由发布脚本注入 release 包，不存放在源码仓库中。
- `templates/template.docx`：Markdown 转 Word 的默认模板。
- `config/style-map.json`：默认 Markdown 到 Word 样式映射。
- `scripts/Add-DocxCaptions.ps1`：Markdown 转 Word 后必须运行的后处理脚本。它会将表格按窗口自动调整、更新域，并可插入表格/图片题注。
- `scripts/Get-WordTemplateStyles.ps1`：检查用户提供的 Word 模板中的段落、字符和表格样式。
- `examples/demo.md`：小型转换示例。

## Markdown 转 Word 流程

当用户给出 Markdown 并要求输出 Word 时，必须完成整个流程。不要在 `mdtodocx.exe to-docx` 后停止。

1. 解析输入 Markdown 路径，并选择输出路径，通常放在 `out/`。
2. 按下方模板情景选择模板和 style-map。
3. 运行 `mdtodocx.exe to-docx`。
4. 立即对生成的 `.docx` 运行 `scripts/Add-DocxCaptions.ps1`。
5. 确认 `.docx` 存在；条件允许时，用 Word 打开或用 Word COM 检查，确认没有修复提示。
6. 返回最终 `.docx` 路径。

默认命令：

```powershell
.\bin\mdtodocx.exe to-docx `
  --input .\input.md `
  --output .\out\input.docx `
  --template .\templates\template.docx `
  --style-map .\config\style-map.json
```

随后必须运行后处理：

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\Add-DocxCaptions.ps1 `
  -DocumentPath .\out\input.docx
```

除非用户明确说不要后处理，否则 Markdown 转 Word 后必须运行这一步。即使没有题注 JSON，也要运行，因为脚本仍会执行表格按窗口自动调整和域更新等必要清理。

支持短参数：

```powershell
.\bin\mdtodocx.exe to-docx -i .\input.md -o .\out\input.docx -t .\template.docx -s .\style-map.json
```

为了兼容旧用法，省略子命令时默认执行 `to-docx`。

转换器支持标题、段落、表格、有序列表和无序列表、引用块、代码块、链接、行内代码、斜体、加粗、数学文本、数学块和本地内联图片。相对图片路径会从 Markdown 文件所在目录解析。

## Markdown 转 Word 的模板情景

模板和 style-map 选择属于 Markdown 转 Word 流程的一部分。除非用户只要求检查模板，否则不要把它当作独立任务。

### 用户未提供模板

使用内置默认资源：

- `templates/template.docx`
- `config/style-map.json`

完成转换后，立即运行后处理。

### 用户提供模板

将用户模板作为 `--template`。

如果用户同时提供 style-map，就使用用户提供的映射。如果映射不明确或缺失，先用 Word COM 检查模板样式：

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\Get-WordTemplateStyles.ps1 `
  -TemplatePath .\user-template.docx `
  -OutputPath .\out\template-styles.json
```

优先使用 Word 样式显示名；当模板使用中文 Word 界面样式名时，优先使用中文名称。列表样式应映射到模板中的列表段落样式，让 Word 自动处理编号或项目符号。

style-map 核心结构：

```json
{
  "version": 1,
  "blockStyles": {
    "heading.1": "标题 1",
    "paragraph": "正文",
    "table": "表格文本",
    "unordered_list": "列表段落",
    "ordered_list": "有序列表段落",
    "math_block": "公式"
  },
  "tableStyles": {
    "table": "Merit"
  },
  "inlineStyles": {
    "inline_code": "行内代码"
  }
}
```

使用选定模板和 style-map 完成转换后，仍然必须运行 `Add-DocxCaptions.ps1`。

## 后处理中的题注

插入题注是可选能力，但 Markdown 转 Word 后运行后处理脚本本身不是可选项。

如果需要表题注，准备 JSON：

```json
{
  "tables": [
    { "title": "调研成果类型与建议形式" },
    { "title": "真实问答场景缺失问题及影响" }
  ]
}
```

带表题注文件运行后处理：

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\Add-DocxCaptions.ps1 `
  -DocumentPath .\out\input.docx `
  -TableCaptionsPath .\out\input-table-captions.json
```

只有当用户要求图片题注，或任务明显需要图片题注时，才添加 `-AddImageCaptions`。

题注规则：

- 表题注插入在表格上方。
- 图片题注插入在内联图片下方。
- 使用 Word 域，不写死编号。
- 使用模板中的 `题注` 样式。
- 表题注按一级标题进行章节编号。
- 图片题注默认使用全局编号，除非用户另有要求。

## Word 转 Markdown 流程

当用户给出 `.docx` 并要求输出 Markdown 时，使用 `mdtodocx.exe to-md`。

```powershell
.\bin\mdtodocx.exe to-md `
  --input .\report.docx `
  --output .\out\report.md `
  --assets-dir .\out\report.assets
```

支持短参数：

```powershell
.\bin\mdtodocx.exe to-md -i .\report.docx -o .\out\report.md -a .\out\report.assets
```

如果省略 `--assets-dir`，默认使用 `<输出文件名>.assets`。

`to-md` 是零配置转换。不要要求用户提供反向 style-map，应从 Word 文档结构中推断语义：

- Word 标题转换为 Markdown 标题。
- Word 有序列表和无序列表转换为 Markdown 列表。
- 嵌入图片导出到 assets 目录，并在 Markdown 中以相对路径引用。
- 简单表格输出为 Markdown pipe table。
- 涉及合并单元格的复杂表格输出为 HTML table。
- Word 公式输出为 Markdown 数学语法：行内 `$...$`，段落 `$$...$$`。

## 验证

Markdown 转 Word 后：

- 确认 `.docx` 存在。
- 确认后处理已经运行。
- 条件允许时，用 Word COM 打开或检查文档，确认没有修复提示。
- 当表格布局重要时，确认表格已按窗口自动调整。

Word 转 Markdown 后：

- 确认 Markdown 文件存在。
- 确认图片导出到了 assets 目录。
- 确认 Markdown 中的图片链接是相对路径。
- 确认简单表格是 Markdown pipe table。
- 确认合并单元格表格是 HTML table。
- 确认公式使用 `$...$` 或 `$$...$$`。

如果在此仓库中调用 Python 做验证或辅助工作，请根据本地项目说明使用 `uv`。
