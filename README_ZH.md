# Markdown To Word Polisher

中文 | [English](README.md)

Markdown To Word Polisher 是一个 Windows-only 的 Markdown / Word 转换工具与 Codex skill 开发仓库。

项目提供：

- `mdtodocx.exe`：自包含 C# CLI，支持 Markdown -> `.docx` 和 `.docx` -> Markdown。
- Word 模板与 `style-map.json`，用于 Markdown 转 Word 时套用样式。
- Word COM 后处理脚本，用于题注、模板检查和表格自动调整。
- 位于 `skills/markdown-to-word-polisher/` 的可分发 Codex skill 包源码。

源码仓库不提交 exe。发布 skill 时，脚本会临时构建 `mdtodocx.exe`，并注入 release zip。

## 环境要求

- Windows。
- 开发和发布需要 .NET SDK。
- 后处理和 Word 输出验证需要 Microsoft Word。
- 运行 PowerShell 脚本时使用 Windows PowerShell 5.1。

## 构建

构建 C# 转换器：

```powershell
dotnet build .\src\MdToDocx\MdToDocx.csproj
```

发布本地自包含 exe：

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\Publish-MdToDocx.ps1
```

输出位置：

```text
bin\mdtodocx.exe
```

本地 `bin/` 已被 Git 忽略。

## Markdown 转 Word

```powershell
.\bin\mdtodocx.exe to-docx `
  --input .\skills\markdown-to-word-polisher\examples\demo.md `
  --output .\out\demo.docx `
  --template .\skills\markdown-to-word-polisher\templates\template.docx `
  --style-map .\skills\markdown-to-word-polisher\config\style-map.json
```

短参数：

```powershell
.\bin\mdtodocx.exe to-docx -i .\input.md -o .\out\input.docx -t .\template.docx -s .\style-map.json
```

省略子命令时默认执行 `to-docx`。

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

如果省略 `--assets-dir`，转换器使用 `<输出文件名>.assets`。

`to-md` 行为：

- Word 标题转 Markdown 标题。
- Word 列表转 Markdown 有序或无序列表。
- 图片导出到 assets 目录，并使用相对路径引用。
- 简单表格转 Markdown pipe table。
- 合并单元格等复杂表格转 HTML table。
- Word 公式转 Markdown 数学语法：行内 `$...$`，段落 `$$...$$`。

## 题注和表格后处理

后处理脚本会添加表格/图片题注，并把所有表格设置为按窗口自动调整：

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\skills\markdown-to-word-polisher\scripts\Add-DocxCaptions.ps1 `
  -DocumentPath .\out\demo.docx `
  -TableCaptionsPath .\out\demo-table-captions.json `
  -AddImageCaptions
```

## 发布 Skill Release

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\Publish-SkillRelease.ps1 `
  -Version v0.1.0 `
  -Title "v0.1.0" `
  -Notes "Initial release."
```

发布脚本要求 Git 工作区干净，并且本机可调用 `dotnet`、`git` 和已登录的 `gh`。脚本会在临时目录构建 `mdtodocx.exe`，复制 skill 源文件，注入 `bin/mdtodocx.exe`，生成 `out/markdown-to-word-polisher-skill-<version>.zip`，创建并推送 tag，然后上传到 GitHub Release。

## 验证

```powershell
dotnet build .\src\MdToDocx\MdToDocx.csproj
.\bin\mdtodocx.exe to-docx -i .\skills\markdown-to-word-polisher\examples\demo.md -o .\out\demo.docx -t .\skills\markdown-to-word-polisher\templates\template.docx -s .\skills\markdown-to-word-polisher\config\style-map.json
.\bin\mdtodocx.exe to-md -i .\out\demo.docx -o .\out\demo.md -a .\out\demo.assets
```

当 Word 输出发生变化时，用 Microsoft Word 或 Word COM 打开生成的 `.docx`，确认不会出现修复提示。
