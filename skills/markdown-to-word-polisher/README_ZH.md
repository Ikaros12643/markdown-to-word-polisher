# Markdown To Word Polisher

这是一个面向交付文档的 Windows-only Codex skill。release 包内置 `bin/mdtodocx.exe`，支持 Markdown 与 Word `.docx` 双向转换。

## 能力

- Markdown 转带模板样式的 Word。
- Word 转 Markdown，并导出图片资源目录。
- 简单表格转 Markdown pipe table。
- 合并单元格等复杂表格转 HTML table。
- Word 公式转 Markdown 数学语法。
- 使用 Word COM 检查模板样式。
- 后处理添加表格/图片题注，并将表格设置为按窗口自动调整。

## 目录

```text
markdown-to-word-polisher/
├─ SKILL.md
├─ SKILL_ZH.md
├─ README_ZH.md
├─ bin/
│  └─ mdtodocx.exe
├─ config/
│  └─ style-map.json
├─ scripts/
│  ├─ Add-DocxCaptions.ps1
│  └─ Get-WordTemplateStyles.ps1
├─ templates/
│  └─ template.docx
└─ examples/
   └─ demo.md
```

`bin/mdtodocx.exe` 只存在于 release 包中，不存放在源码仓库。

## Markdown 转 Word

```powershell
.\bin\mdtodocx.exe to-docx `
  --input .\examples\demo.md `
  --output .\out\demo.docx `
  --template .\templates\template.docx `
  --style-map .\config\style-map.json
```

短参数：

```powershell
.\bin\mdtodocx.exe to-docx -i .\input.md -o .\out\input.docx -t .\template.docx -s .\style-map.json
```

## Word 转 Markdown

```powershell
.\bin\mdtodocx.exe to-md `
  --input .\report.docx `
  --output .\out\report.md `
  --assets-dir .\out\report.assets
```

短参数：

```powershell
.\bin\mdtodocx.exe to-md -i .\report.docx -o .\out\report.md -a .\out\report.assets
```

## 题注和表格后处理

转换后可运行后处理，为表格和图片添加题注，并让所有表格按窗口自动调整：

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\Add-DocxCaptions.ps1 `
  -DocumentPath .\out\demo.docx `
  -TableCaptionsPath .\out\demo-table-captions.json `
  -AddImageCaptions
```

表题注文本来自 JSON：

```json
{
  "tables": [
    { "title": "示例表格标题" }
  ]
}
```

图片题注使用全局 Word 域编号自动生成。
