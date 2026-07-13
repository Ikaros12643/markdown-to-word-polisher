# Agent 指南

## 依赖管理

执行需要第三方依赖的脚本时，先创建虚拟环境：

- **Python**：使用 `uv venv` 创建虚拟环境。

## 需要第三方依赖的 Python 脚本

- 如果当前路径下已有 `pyproject.toml`：
  - 使用 `uv add` 安装依赖。
- 如果当前路径下没有 `pyproject.toml`：
  - 使用 `uv pip install` 安装依赖。
- 优先使用 `uv run` 运行 Python 脚本。
- 如果 Python 脚本不需要任何第三方依赖，可以直接运行 `python` 或 `python3`。

## 路径操作

- 涉及路径操作时，优先使用相对路径，其次使用绝对路径。

## 项目结构与模块组织

本仓库用于开发一个仅支持 Windows 的 Codex 技能和配套命令行工具，用来在 Markdown 与 Word `.docx` 之间转换，并对生成的 Word 文档进行后期抛光。

- `src/MdAst/`：旧版 C# Markdown AST 解析器源码，仅作为代码保留，不再构建或分发到 skill 包。
- `src/MdToDocx/`：C# 双向转换器源码，内部包含 Markdown 解析、OpenXML 文档生成，以及 `.docx` 到 Markdown 的反向转换，可发布为自包含的 `mdtodocx.exe`。
- `scripts/`：仓库根目录下的开发和发布脚本，包括 `Publish-MdToDocx.ps1` 和 `Publish-SkillRelease.ps1`。
- `skills/markdown-to-word-polisher/`：可分发的技能包源码。运行时文件位于此目录，包括 `SKILL.md`、`SKILL_ZH.md`、`README_ZH.md`、`config/`、`scripts/`、`templates/` 和 `examples/`。`bin/mdtodocx.exe` 只在 release 包中由发布脚本注入，不提交到源码仓库。
- `context/`：设计说明和背景文档。
- `resources/`：项目资源和参考资料。
- `out/`：生成的转换输出目录；不要把其中内容纳入源码变更。

## 构建、测试与开发命令

构建 C# Markdown 到 Word 转换器：

```powershell
dotnet build .\src\MdToDocx\MdToDocx.csproj
```

发布自包含的单文件转换器：

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\Publish-MdToDocx.ps1
```

使用 C# 转换器将 Markdown 转为 Word：

```powershell
.\bin\mdtodocx.exe to-docx `
  --input .\skills\markdown-to-word-polisher\examples\demo.md `
  --output .\out\demo-csharp.docx `
  --template .\skills\markdown-to-word-polisher\templates\template.docx `
  --style-map .\skills\markdown-to-word-polisher\config\style-map.json
```

旧写法仍然兼容，未指定子命令时默认执行 `to-docx`。

使用 C# 转换器将 Word 转为 Markdown，并导出图片资源：

```powershell
.\bin\mdtodocx.exe to-md `
  --input .\out\demo-csharp.docx `
  --output .\out\demo-csharp.md `
  --assets-dir .\out\demo-csharp.assets
```

使用可分发技能包结构运行一次转换：

```powershell
.\bin\mdtodocx.exe to-docx `
  --input .\skills\markdown-to-word-polisher\examples\demo.md `
  --output .\out\demo.docx `
  --template .\skills\markdown-to-word-polisher\templates\template.docx `
  --style-map .\skills\markdown-to-word-polisher\config\style-map.json
```

发布 skill release：

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\scripts\Publish-SkillRelease.ps1 `
  -Version v0.1.0 `
  -Title "v0.1.0" `
  -Notes "Initial release."
```

## 代码风格与命名规范

运行时脚本必须兼容 PowerShell 5.1，因为后期抛光流程仍依赖 Windows PowerShell 和 Word COM。示例和脚本中的路径应尽量使用相对路径。C# 代码使用可空引用类型、隐式 `using`、类型和方法采用 PascalCase，局部变量采用 camelCase。JSON 配置键应保持小写或 camelCase，以匹配现有 `style-map.json` 的风格。

### C# 编码规范

- 按职责组织文件：CLI、Markdown 到 Word、Word 到 Markdown、共享模型和工具应分目录放置；一个核心类优先对应一个 `.cs` 文件，避免重新形成巨大的单文件实现。
- 类型、方法、属性使用 PascalCase；局部变量和参数使用 camelCase；私有字段使用 `_camelCase`。
- 普通方法应尽量保持短小；当方法里出现多个明显步骤、较长的 OpenXML 构造过程，或同时处理多种节点类型时，优先拆成有明确名字的私有方法。
- 转换规则应放在对应模块中：Markdown 解析与 AST 转换放在 `MarkdownToWord`，Word 读取与 Markdown 输出放在 `WordToMarkdown`，样式映射和共享数据结构放在 `Shared`。

### 注释规范

- 公共边界优先写 XML 文档注释：对 `public` 类型、方法，以及未来可能被其他项目复用的 `internal` 类型，可使用 `/// <summary>` 说明用途、输入输出和关键限制。
- 私有方法不强制写注释；只有当代码意图无法从方法名和结构直接读出时才补充说明。
- 注释主要解释“为什么”和“约束”，少解释代码已经表达清楚的“做什么”。例如应说明 Word/OpenXML 的特殊行为、样式名兼容策略、公式或表格转换的取舍。
- 避免空泛注释，例如 `// 遍历段落`、`// 创建对象`、`// 返回结果`。如果注释不能帮助后续维护者避开坑，就不要写。
- 对 OpenXML、Word COM、样式映射、复杂表格、公式转换等容易踩坑的逻辑，应在关键位置留下简短注释，说明原因和不可随意改动的条件。
- 注释语言优先使用中文，除非代码上下文、库 API 或异常信息本身使用英文更清晰。

## 测试指南

目前尚无正式测试套件。修改 `to-docx` 后，应使用 `examples/demo.md` 以及至少一份真实文档进行验证，真实文档应包含标题、列表、表格、图片、行内格式和数学块。确认生成的 `.docx` 能在 Microsoft Word 中打开，并且不会出现修复提示。

修改 `to-md` 后，应至少验证：

- 图片能导出到 assets 目录，并在 Markdown 中使用相对路径引用。
- 简单表格输出为 Markdown pipe table。
- 涉及合并单元格的复杂表格输出为 HTML table。
- Word 公式能输出为 Markdown 数学语法：行内公式 `$...$`，段落公式 `$$...$$`。
- `to-docx` 回归仍能生成可被 Word 打开的 `.docx`。
- 运行 `Add-DocxCaptions.ps1` 后，表格应按窗口自动调整，不应缩成一团。

## 提交与 Pull Request 指南

Pull Request 应包含简要摘要、受影响区域（`src/MdToDocx`、根目录脚本或技能包）、验证命令，以及当 Word 输出发生变化时的截图或示例 `.docx` 说明。

## 安全与配置提示

不要提交私人文档、生成输出、本地 Word 模板或可执行文件，除非它们本来就是预期示例。`mdtodocx.exe` 等二进制产物应只进入 GitHub Release，不进入源码仓库。
