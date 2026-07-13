---
name: markdown-to-word-polisher
description: Convert between Markdown and Word .docx on Windows. Use the bundled mdtodocx.exe for Markdown -> styled Word and Word -> Markdown with image asset extraction, simple/complex table handling, and Markdown math output; use PowerShell 5.1 Word COM scripts for template inspection and field-based figure/table captions.
---

# Markdown To Word Polisher

Use this skill when the user needs Markdown and Word `.docx` conversion, Word-template style mapping, Word output validation, template style inspection, or caption post-processing.

The release package contains a self-contained C# converter:

```text
Markdown -> bin/mdtodocx.exe to-docx -> .docx
.docx -> bin/mdtodocx.exe to-md -> Markdown + assets
```

All PowerShell scripts must stay compatible with Windows PowerShell 5.1. Prefer `powershell.exe -ExecutionPolicy Bypass -File ...` when running them.

## Resource Layout

- `bin/mdtodocx.exe`: bundled self-contained converter for `to-docx` and `to-md`. This binary is injected into release packages and is not stored in the source repository.
- `config/style-map.json`: default Markdown-to-Word style mapping.
- `templates/template.docx`: bundled default Word template.
- `scripts/Add-DocxCaptions.ps1`: post-process a generated Word document to add table and image captions.
- `scripts/Get-WordTemplateStyles.ps1`: inspect a user-provided Word template's paragraph, character, and table styles.
- `examples/demo.md`: small conversion example.

## Markdown To Word

When the user gives Markdown and asks for Word output, use `mdtodocx.exe to-docx`.

Default resources:

- `templates/template.docx`
- `config/style-map.json`

Example:

```powershell
.\bin\mdtodocx.exe to-docx `
  --input .\examples\demo.md `
  --output .\out\demo.docx `
  --template .\templates\template.docx `
  --style-map .\config\style-map.json
```

Short arguments are supported:

```powershell
.\bin\mdtodocx.exe to-docx -i .\input.md -o .\out\input.docx -t .\template.docx -s .\style-map.json
```

For compatibility, omitting the subcommand defaults to `to-docx`.

The converter supports headings, paragraphs, tables, ordered and unordered lists, blockquotes, code blocks, links, inline code, emphasis, strong text, math text, math blocks, and local inline images. Relative image paths are resolved from the Markdown file directory.

## Word To Markdown

When the user gives a `.docx` and asks for Markdown output, use `mdtodocx.exe to-md`.

Example:

```powershell
.\bin\mdtodocx.exe to-md `
  --input .\report.docx `
  --output .\out\report.md `
  --assets-dir .\out\report.assets
```

Short arguments are supported:

```powershell
.\bin\mdtodocx.exe to-md -i .\report.docx -o .\out\report.md -a .\out\report.assets
```

If `--assets-dir` is omitted, use `<output filename>.assets`.

`to-md` is zero-config. Do not ask the user for a reverse style map. It should infer semantics from the Word document structure:

- Word headings become Markdown headings.
- Word ordered and unordered lists become Markdown lists.
- Embedded images are exported to assets and referenced with relative paths.
- Simple tables become Markdown pipe tables.
- Tables with merged cells become HTML tables.
- Word formulas become Markdown math: inline `$...$`, block `$$...$$`.

## Template And Style Mapping

For `to-docx`, `style-map.json` maps Markdown semantics to Word style display names. Prefer Chinese style names when the template uses Chinese Word UI names.

Example:

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

`blockStyles.table` controls text inside tables. `tableStyles.table` controls the Word table object style. List styles should map to the template's list paragraph styles so Word can apply automatic numbering or bullets.

When the user provides a custom template, inspect styles with Word COM if mappings are ambiguous:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\Get-WordTemplateStyles.ps1 `
  -TemplatePath .\user-template.docx `
  -OutputPath .\out\template-styles.json
```

## Caption Post-Processing

Treat captions as a post-processing step after conversion. Do not add document-specific captions directly to the main converter unless the user explicitly asks for converter-level behavior.

For tables:

- Insert captions above each table.
- Use the template's `题注` style.
- Use Word fields, not hard-coded numbers.
- Format as `表 x-x title`.
- The first `x` is the level-1 heading number, and the second `x` is that table's sequence within the level-1 section.

Prepare table titles as JSON:

```json
{
  "tables": [
    { "title": "调研成果类型与建议形式" },
    { "title": "真实问答场景缺失问题及影响" }
  ]
}
```

For images:

- Insert captions below each inline image.
- Use Word fields with global numbering.
- Format as `图 x`.
- Do not add chapter numbering to image captions unless the user asks.

Run:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\Add-DocxCaptions.ps1 `
  -DocumentPath .\out\input.docx `
  -TableCaptionsPath .\out\input-table-captions.json `
  -AddImageCaptions
```

## Validation

After `to-docx`, verify the output `.docx` exists and, when practical, open or inspect it with Word COM to confirm it opens without repair prompts.

After `to-md`, verify:

- The Markdown file exists.
- Images were exported to the assets directory.
- Image links in Markdown are relative paths.
- Simple tables are Markdown pipe tables.
- Merged-cell tables are HTML tables.
- Formulas use `$...$` or `$$...$$`.

If invoking Python for validation or helper work in this repository, use `uv` according to the local project instructions.
