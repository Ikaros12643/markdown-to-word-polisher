---
name: markdown-to-word-polisher
description: 在 Windows 上进行 Markdown 与 Word .docx 双向转换。使用 release 包内置的 mdtodocx.exe 将 Markdown 转为带模板样式的 Word，或将 Word 转为 Markdown 并导出图片资源；简单表格输出 Markdown 表格，复杂表格输出 HTML 表格，公式输出 Markdown 数学语法。使用 PowerShell 5.1 Word COM 脚本进行模板样式检查和基于 Word 域的图表题注后处理。
---

# Markdown To Word Polisher

当用户需要 Markdown 与 Word `.docx` 转换、Word 模板样式映射、Word 输出验证、模板样式检查或题注后处理时，使用此 skill。

release 包中的主要转换器是自包含 C# 可执行文件：

```text
Markdown -> bin/mdtodocx.exe to-docx -> .docx
.docx -> bin/mdtodocx.exe to-md -> Markdown + assets
```

所有 PowerShell 脚本都必须保持兼容 Windows PowerShell 5.1。运行脚本时优先使用 `powershell.exe -ExecutionPolicy Bypass -File ...`。

## 资源布局

- `bin/mdtodocx.exe`：release 包内置的自包含转换器，支持 `to-docx` 和 `to-md`。此二进制由发布脚本注入 release 包，不存放在源码仓库中。
- `config/style-map.json`：默认 Markdown 到 Word 样式映射。
- `templates/template.docx`：内置默认 Word 模板。
- `scripts/Add-DocxCaptions.ps1`：对生成的 Word 文档进行后处理，添加表格和图片题注，并将表格设置为按窗口自动调整。
- `scripts/Get-WordTemplateStyles.ps1`：检查用户提供的 Word 模板中的段落、字符和表格样式。
- `examples/demo.md`：小型转换示例。

## Markdown 转 Word

当用户给出 Markdown 并要求输出 Word 时，使用 `mdtodocx.exe to-docx`。

默认资源：

- `templates/template.docx`
- `config/style-map.json`

示例：

```powershell
.\bin\mdtodocx.exe to-docx `
  --input .\examples\demo.md `
  --output .\out\demo.docx `
  --template .\templates\template.docx `
  --style-map .\config\style-map.json
```

支持短参数：

```powershell
.\bin\mdtodocx.exe to-docx -i .\input.md -o .\out\input.docx -t .\template.docx -s .\style-map.json
```

为了兼容旧用法，省略子命令时默认执行 `to-docx`。

转换器支持标题、段落、表格、有序列表和无序列表、引用块、代码块、链接、行内代码、斜体、加粗、数学文本、数学块和本地内联图片。相对图片路径会从 Markdown 文件所在目录解析。

## Word 转 Markdown

当用户给出 `.docx` 并要求输出 Markdown 时，使用 `mdtodocx.exe to-md`。

示例：

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

## 模板与样式映射

对于 `to-docx`，`style-map.json` 将 Markdown 语义映射到 Word 样式显示名。模板使用中文 Word 界面样式名时，优先使用中文样式名。

示例：

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

`blockStyles.table` 控制表格内部文本样式。`tableStyles.table` 控制 Word 表格对象样式。列表样式应映射到模板中的列表段落样式，让 Word 自动处理编号或项目符号。

当用户提供自定义模板且映射不明确时，使用 Word COM 检查模板样式：

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\Get-WordTemplateStyles.ps1 `
  -TemplatePath .\user-template.docx `
  -OutputPath .\out\template-styles.json
```

## 题注与表格后处理

题注应作为转换后的后处理步骤。除非用户明确要求转换器级行为，否则不要把某个文档特定的题注直接写进主转换器。

后处理脚本会对所有表格执行按窗口自动调整，避免表格在 Word 中缩成一团。

对于表格题注：

- 在每个表格上方插入题注。
- 使用模板中的 `题注` 样式。
- 使用 Word 域，不写死编号。
- 格式为 `表 x-x title`。
- 第一个 `x` 是一级标题编号，第二个 `x` 是该一级章节内的表格序号。

准备一个用于表题注标题的 JSON 文件：

```json
{
  "tables": [
    { "title": "调研成果类型与建议形式" },
    { "title": "真实问答场景缺失问题及影响" }
  ]
}
```

对于图片：

- 在每张内联图片下方插入题注。
- 使用 Word 域进行全局编号。
- 格式为 `图 x`。
- 除非用户要求，否则不要给图片题注添加章节编号。

运行：

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\Add-DocxCaptions.ps1 `
  -DocumentPath .\out\input.docx `
  -TableCaptionsPath .\out\input-table-captions.json `
  -AddImageCaptions
```

## 验证

执行 `to-docx` 后，确认输出 `.docx` 存在；条件允许时，用 Word COM 打开或检查文档，确认不会出现修复提示。

执行 `to-md` 后，确认：

- Markdown 文件存在。
- 图片导出到了 assets 目录。
- Markdown 中的图片链接是相对路径。
- 简单表格是 Markdown pipe table。
- 合并单元格表格是 HTML table。
- 公式使用 `$...$` 或 `$$...$$`。

如果在此仓库中调用 Python 做验证或辅助工作，请根据本地项目说明使用 `uv`。
