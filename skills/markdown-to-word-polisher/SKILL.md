---
name: markdown-to-word-polisher
description: Convert between Markdown and Word .docx on Windows. Use when the user asks for Markdown -> polished Word output, Word -> Markdown with image assets, Word-template based conversion, style-map inspection, or validation of converted Word documents. Markdown -> Word must run mdtodocx.exe first and then immediately run the bundled Word COM post-processing script before delivery.
---

# Markdown To Word Polisher

Use this skill as a workflow guide, not as a list of independent utilities. The main delivery path is:

```text
Markdown input
-> choose template and style-map
-> bin/mdtodocx.exe to-docx
-> immediately run Word COM post-processing
-> validate the .docx
-> return the final file path
```

For Word to Markdown:

```text
.docx input
-> bin/mdtodocx.exe to-md
-> Markdown + extracted assets
-> validate links, tables, and math syntax
```

All PowerShell scripts must stay compatible with Windows PowerShell 5.1. Prefer `powershell.exe -ExecutionPolicy Bypass -File ...` when running them.

## Resources

- `bin/mdtodocx.exe`: self-contained C# converter for `to-docx` and `to-md`. It is injected into release packages and is not stored in source control.
- `templates/template.docx`: default Word template for Markdown to Word.
- `config/style-map.json`: default Markdown-to-Word style mapping.
- `scripts/Add-DocxCaptions.ps1`: mandatory post-processing script after Markdown to Word conversion. It auto-fits tables to the window, refreshes fields, and can insert table/image captions.
- `scripts/Get-WordTemplateStyles.ps1`: inspect paragraph, character, and table styles in a user-provided Word template.
- `examples/demo.md`: small conversion example.

## Markdown To Word Workflow

When the user gives Markdown and asks for Word output, complete the whole workflow. Do not stop after `mdtodocx.exe to-docx`.

1. Resolve the input Markdown path and choose an output path, usually under `out/`.
2. Choose template and style-map according to the template scenario below.
3. Run `mdtodocx.exe to-docx`.
4. Immediately run `scripts/Add-DocxCaptions.ps1` on the generated `.docx`.
5. Verify the `.docx` exists and, when practical, opens in Word without repair prompts.
6. Report the final `.docx` path.

Default command:

```powershell
.\bin\mdtodocx.exe to-docx `
  --input .\input.md `
  --output .\out\input.docx `
  --template .\templates\template.docx `
  --style-map .\config\style-map.json
```

Then always run post-processing:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\Add-DocxCaptions.ps1 `
  -DocumentPath .\out\input.docx
```

This post-processing step is mandatory for Markdown to Word delivery unless the user explicitly says not to run it. Even without caption JSON, it still performs required cleanup such as table auto-fit to window and field updates.

Short converter arguments are supported:

```powershell
.\bin\mdtodocx.exe to-docx -i .\input.md -o .\out\input.docx -t .\template.docx -s .\style-map.json
```

For compatibility, omitting the subcommand defaults to `to-docx`.

The converter supports headings, paragraphs, tables, ordered and unordered lists, blockquotes, code blocks, links, inline code, emphasis, strong text, math text, math blocks, and local inline images. Relative image paths are resolved from the Markdown file directory.

## Template Scenarios For Markdown To Word

Template and style-map selection is part of the Markdown to Word workflow. Do not treat it as a separate task unless the user only asks to inspect a template.

### No User Template

Use the bundled defaults:

- `templates/template.docx`
- `config/style-map.json`

Run conversion, then immediately run post-processing.

### User Provides A Template

Use the user template as `--template`.

If the user also provides a style-map, use it. If mappings are unclear or missing, inspect the template styles:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\Get-WordTemplateStyles.ps1 `
  -TemplatePath .\user-template.docx `
  -OutputPath .\out\template-styles.json
```

Prefer Word style display names, especially Chinese names when the template uses Chinese Word UI names. List styles should map to the template's list paragraph styles so Word can apply automatic numbering or bullets.

Core style-map shape:

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

After conversion with the chosen template and style-map, still run `Add-DocxCaptions.ps1`.

## Captions During Post-Processing

Caption insertion is optional, but the post-processing script itself is not optional after Markdown to Word conversion.

If table captions are needed, prepare JSON:

```json
{
  "tables": [
    { "title": "调研成果类型与建议形式" },
    { "title": "真实问答场景缺失问题及影响" }
  ]
}
```

Run post-processing with the table caption file:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\Add-DocxCaptions.ps1 `
  -DocumentPath .\out\input.docx `
  -TableCaptionsPath .\out\input-table-captions.json
```

For image captions, add `-AddImageCaptions` only when the user asks for image captions or the task clearly requires them.

Caption rules:

- Table captions go above tables.
- Image captions go below inline images.
- Use Word fields, not hard-coded numbers.
- Use the template's `题注` style.
- Table captions use chapter numbering by level-1 heading.
- Image captions use global numbering unless the user asks otherwise.

## Word To Markdown Workflow

When the user gives a `.docx` and asks for Markdown output, use `mdtodocx.exe to-md`.

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

`to-md` is zero-config. Do not ask the user for a reverse style-map. Infer semantics from the Word document structure:

- Word headings become Markdown headings.
- Word ordered and unordered lists become Markdown lists.
- Embedded images are exported to assets and referenced with relative paths.
- Simple tables become Markdown pipe tables.
- Tables with merged cells become HTML tables.
- Word formulas become Markdown math: inline `$...$`, block `$$...$$`.

## Validation

After Markdown to Word:

- Confirm the `.docx` exists.
- Confirm post-processing ran.
- When practical, open or inspect with Word COM to confirm there are no repair prompts.
- When table layout matters, confirm tables were auto-fit to the window.

After Word to Markdown:

- Confirm the Markdown file exists.
- Confirm images were exported to the assets directory.
- Confirm image links are relative paths.
- Confirm simple tables are Markdown pipe tables.
- Confirm merged-cell tables are HTML tables.
- Confirm formulas use `$...$` or `$$...$$`.

If invoking Python for validation or helper work in this repository, use `uv` according to the local project instructions.
