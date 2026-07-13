# Markdown To Word Polisher

Windows-only Markdown / Word conversion toolkit for Codex workflows.

The project provides:

- `mdtodocx.exe`: a self-contained C# CLI for Markdown -> `.docx` and `.docx` -> Markdown.
- Word template and `style-map.json` support for Markdown -> Word.
- Word COM post-processing scripts for caption and polishing tasks.
- A distributable Codex skill package under `skills/markdown-to-word-polisher/`.

Executable binaries are not stored in the source repository. Release packages build `mdtodocx.exe` and inject it into the skill zip.

## Requirements

- Windows.
- .NET SDK for development and publishing.
- Microsoft Word for COM-based validation and post-processing.
- Windows PowerShell 5.1 for runtime scripts.

## Build

Build the C# converter:

```powershell
dotnet build .\src\MdToDocx\MdToDocx.csproj
```

Publish the self-contained executable:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\Publish-MdToDocx.ps1
```

The executable is written to:

```text
bin\mdtodocx.exe
```

This local `bin/` directory is ignored by Git.

## Markdown To Word

Use `to-docx` to convert Markdown into a styled Word document:

```powershell
.\bin\mdtodocx.exe to-docx `
  --input .\skills\markdown-to-word-polisher\examples\demo.md `
  --output .\out\demo.docx `
  --template .\skills\markdown-to-word-polisher\templates\template.docx `
  --style-map .\skills\markdown-to-word-polisher\config\style-map.json
```

Short arguments are also supported:

```powershell
.\bin\mdtodocx.exe to-docx -i .\input.md -o .\out\input.docx -t .\template.docx -s .\style-map.json
```

For compatibility, omitting the subcommand still defaults to `to-docx`.

## Word To Markdown

Use `to-md` to convert Word into Markdown and export embedded images:

```powershell
.\bin\mdtodocx.exe to-md `
  --input .\report.docx `
  --output .\out\report.md `
  --assets-dir .\out\report.assets
```

Short arguments:

```powershell
.\bin\mdtodocx.exe to-md -i .\report.docx -o .\out\report.md -a .\out\report.assets
```

If `--assets-dir` is omitted, the converter uses `<output-name>.assets`.

`to-md` behavior:

- Headings become Markdown headings.
- Word lists become Markdown ordered or unordered lists.
- Images are exported to assets and referenced with relative paths.
- Simple tables become Markdown pipe tables.
- Tables with merged cells become HTML tables.
- Word formulas are emitted as Markdown math: inline `$...$`, block `$$...$$`.

## Caption And Table Post-Processing

The distributable skill includes PowerShell COM scripts for template inspection and Word polishing. `Add-DocxCaptions.ps1` adds captions and sets all tables to AutoFit to window:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\skills\markdown-to-word-polisher\scripts\Add-DocxCaptions.ps1 `
  -DocumentPath .\out\demo.docx `
  -TableCaptionsPath .\out\demo-table-captions.json `
  -AddImageCaptions
```

Use this script when Word COM behavior is required for polishing or caption post-processing.

## Release

Create a skill release package and upload it to GitHub Releases:

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\Publish-SkillRelease.ps1 `
  -Version v0.1.0 `
  -Title "v0.1.0" `
  -Notes "Initial release."
```

The release script requires a clean Git working tree, `dotnet`, `git`, and an authenticated `gh`. It builds `mdtodocx.exe` into a temporary staging directory, copies the source skill files, injects `bin/mdtodocx.exe`, creates `out/markdown-to-word-polisher-skill-<version>.zip`, creates/pushes the Git tag, and uploads the zip to the GitHub Release.

## Validation

Recommended checks after converter changes:

```powershell
dotnet build .\src\MdToDocx\MdToDocx.csproj
.\bin\mdtodocx.exe to-docx -i .\skills\markdown-to-word-polisher\examples\demo.md -o .\out\demo.docx -t .\skills\markdown-to-word-polisher\templates\template.docx -s .\skills\markdown-to-word-polisher\config\style-map.json
.\bin\mdtodocx.exe to-md -i .\out\demo.docx -o .\out\demo.md -a .\out\demo.assets
```

When Word output changes, open the generated `.docx` in Microsoft Word or inspect it with Word COM to ensure it opens without repair prompts.
