# DOCX 转 Markdown 零配置实现方案

## 背景

当前项目已经有 `mdtodocx.exe`，主要负责 Markdown 到 Word `.docx` 的转换。实际使用中，用户也经常需要把 Word 文档转回 Markdown，作为后续编写、整理、版本管理或再次生成文档的输入。

现有通用工具如 markitdown 能提取文本和部分结构，但对 Word 中图片的处理不够满足本项目需求：图片不能稳定导出到本地资源目录，并以相对路径写入 Markdown。

因此计划在同一个 exe 中增加反向转换能力：

```powershell
mdtodocx to-md -i input.docx -o output.md -a output.assets
```

`to-md` 不再引入额外样式映射配置。转换应尽量从 Word 文档自身结构中识别标题、列表、表格、图片等语义。

## 目标

- 支持 `.docx` 转 `.md`。
- 导出 Word 中的图片到本地资源目录。
- Markdown 中使用相对路径引用图片。
- 简单表格输出为 Markdown pipe table。
- 复杂表格输出为 HTML table。
- Word 公式输出为 Markdown 数学语法：行内公式 `$...$`，段落公式 `$$...$$`。
- 不要求用户维护 `docx -> md` 的 style-map。
- 和现有 `md -> docx` 能力共用同一个 exe，但代码边界清晰。

## 非目标

- 不追求完整还原 Word 视觉样式。
- 不把所有 Word 自定义样式都映射成 Markdown 语义。
- 不要求首版完整支持文本框、SmartArt、图表、批注、修订痕迹。
- 不要求首版完整覆盖所有复杂 OMML 公式结构；无法可靠转换的公式可保留占位，但已识别的公式必须按 Markdown 数学语法输出。

## 命令设计

新增子命令：

```powershell
mdtodocx to-md --input input.docx --output output.md --assets-dir output.assets
```

支持短参数：

```powershell
mdtodocx to-md -i input.docx -o output.md -a output.assets
```

参数说明：

- `-i, --input`：输入 `.docx` 文件，必填。
- `-o, --output`：输出 Markdown 文件，必填。
- `-a, --assets-dir`：图片资源输出目录，可选。
- `--assets-relative-path`：Markdown 中使用的图片相对路径，可选。
- `-h, --help`：帮助信息。

默认规则：

- 如果未传 `--assets-dir`，默认使用 `<output 文件名>.assets`。
- 如果未传 `--assets-relative-path`，默认使用 assets 目录相对 Markdown 文件所在目录的路径。
- Markdown 中的路径统一使用 `/`，避免 Windows 反斜杠影响跨平台阅读。

示例：

```powershell
mdtodocx to-md -i .\report.docx -o .\report.md
```

输出：

```text
report.md
report.assets/
  image-001.png
  image-002.jpg
```

Markdown 图片引用：

```markdown
![](report.assets/image-001.png)
```

## 整体架构

推荐采用两段式路线：

```text
DOCX -> 语义 HTML -> Markdown
```

第一段使用 Mammoth for .NET：

```text
DOCX -> HTML
```

Mammoth 擅长把 Word 文档转换成语义 HTML，包括标题、段落、列表、表格、脚注、图片等。它也支持自定义图片处理器，适合把图片保存到指定目录。

第二段使用自定义 HTML 到 Markdown 管线：

```text
HTML -> Markdown
```

可使用 ReverseMarkdown.Net 处理普通 HTML，但表格和图片路径最好由本项目自己控制。

建议流程：

```text
1. 读取 DOCX
2. Mammoth 转 HTML，同时导出图片
3. 解析 HTML DOM
4. 对 table 做简单/复杂分类
5. 简单 table 转 Markdown 表格
6. 复杂 table 保留为 HTML 表格
7. 其余 HTML 转 Markdown
8. 统一整理空行、路径、转义
9. 写出 .md
```

## 代码组织

在现有 `src/MdToDocx/` 内新增反向转换模块：

```text
src/MdToDocx/
  Program.cs
  MarkdownToDocx/
  DocxToMarkdown/
    DocxToMarkdownCommand.cs
    DocxToHtmlConverter.cs
    ImageExporter.cs
    HtmlNormalizer.cs
    TableClassifier.cs
    MarkdownTableWriter.cs
    HtmlTableWriter.cs
    FormulaExtractor.cs
    OmmlToLatexConverter.cs
    MarkdownWriter.cs
```

首版也可以先保持文件较少，但逻辑边界建议按上面拆分，避免 `Program.cs` 继续膨胀。

核心入口：

```csharp
internal sealed record DocxToMarkdownOptions(
    string InputPath,
    string OutputPath,
    string AssetsDir,
    string AssetsRelativePath);
```

核心转换器：

```csharp
internal sealed class DocxToMarkdownConverter
{
    public void Convert(DocxToMarkdownOptions options);
}
```

## 零配置语义识别策略

`to-md` 不使用反向 style-map。语义识别优先使用 Word 结构，其次使用常见内置样式。

### 标题

标题识别优先级：

1. Mammoth 对 Word heading 样式的默认识别。
2. Word 内置样式名，如 `Heading 1` 到 `Heading 6`。
3. 常见中文显示名，如 `标题 1` 到 `标题 6`。
4. 大纲级别 `outline level`。

输出：

```markdown
# 一级标题
## 二级标题
```

不建议用字体大小、加粗等视觉特征猜标题，避免误判。

### 列表

列表识别优先使用 Word 编号结构，而不是样式名。

可靠信号：

- 段落有 `numPr`。
- 段落样式继承或包含 `numPr`。
- Mammoth 输出为 `<ul>` / `<ol>`。

输出：

```markdown
- 无序列表
- 无序列表
```

```markdown
1. 有序列表
2. 有序列表
```

多级列表按缩进输出：

```markdown
1. 一级
   - 二级
   - 二级
```

首版可先依赖 Mammoth 的 `<ul>` / `<ol>` 输出；如果发现模板样式编号无法被 Mammoth 识别，再补 OpenXML 级别的列表识别。

### 段落

普通段落输出为 Markdown 段落。连续段落之间保留一个空行。

行内格式处理：

- 加粗 -> `**text**`
- 斜体 -> `*text*`
- 链接 -> `[text](url)`
- 行内代码 -> `` `code` ``

如果行内代码无法可靠识别，首版可以保留为普通文本。

### 图片

图片必须导出到本地文件，并在 Markdown 中使用相对路径。

图片保存策略：

```text
assets-dir/
  image-001.png
  image-002.jpg
```

命名规则：

- 按文档出现顺序编号。
- 尽量保留原始扩展名。
- 如果 MIME 类型未知，默认 `.bin` 或尝试根据文件头判断。
- 可选增强：对图片内容计算 hash，避免重复导出。

Markdown 输出：

```markdown
![alt text](report.assets/image-001.png)
```

alt 文本来源优先级：

1. Word 图片替代文本。
2. 图片标题。
3. 空字符串。

### 表格

表格由本项目自己判断，不完全交给 HTML 转 Markdown 库。

#### 简单表格

满足以下条件时输出 Markdown pipe table：

- 没有 `rowspan`。
- 没有 `colspan`。
- 没有垂直合并。
- 没有嵌套表格。
- 每行列数一致。
- 单元格内容只包含简单行内内容。
- 单元格内没有列表、图片、多段落、代码块等复杂块级结构。

输出：

```markdown
| 字段 | 说明 |
|---|---|
| input | Markdown 输入 |
| output | Word 输出 |
```

对齐处理：

- 如果 HTML 中能识别 `text-align: center`，输出 `:---:`。
- 如果能识别 `text-align: right`，输出 `---:`。
- 否则输出 `---`。

#### 复杂表格

以下任一情况视为复杂表格：

- 存在 `rowspan`。
- 存在 `colspan`。
- 存在垂直合并。
- 存在嵌套表格。
- 单元格包含图片。
- 单元格包含列表。
- 单元格包含多个块级段落。
- 行列结构不规则。

复杂表格输出 HTML：

```html
<table>
  <tr>
    <td rowspan="2">...</td>
    <td>...</td>
  </tr>
</table>
```

HTML 表格输出时应尽量清理 Word 冗余样式，只保留结构相关属性：

- `rowspan`
- `colspan`
- 必要的 `align`

不保留大量 Word CSS。

### 代码块

零配置下很难完全可靠识别代码块。首版策略：

- Mammoth 如果输出 `<pre>`，转换为 fenced code block。
- 常见代码样式名如 `代码`、`Code`、`HTML Preformatted` 可作为弱识别。

输出：

````markdown
```text
code
```
````

### 引用

引用识别：

- Mammoth 输出 `<blockquote>`。
- Word 内置样式 `Quote`。
- 中文样式名 `引用`。

输出：

```markdown
> 引用内容
```

### 脚注

Mammoth 支持脚注输出。首版可转换为 Markdown 脚注：

```markdown
正文[^1]

[^1]: 脚注内容
```

如果实现成本较高，可以先保留 Mammoth 输出的链接形式，再迭代优化。

### 公式

公式导出目标是尽量生成通用 Markdown 数学语法，而不是 Word 专属标记：

- 行内公式输出为 `$latex...$`。
- 独立成段的公式输出为 `$$` 包裹的段落公式。
- 如果能从 OMML 转换出 LaTeX，则优先输出 LaTeX。
- 如果不能可靠转换，输出 HTML 注释或纯文本占位，并保留原公式所在位置。

示例：

```markdown
这是一个行内公式 $E=mc^2$。

$$
\frac{a}{b}
$$
```

#### 行内公式与段落公式判定

`to-md` 应根据公式在 Word 段落中的位置判断输出形式：

- 一个段落只有一个公式，且除空白外没有其他文本：输出段落公式。
- 公式与普通文字混排：输出行内公式。
- 一个段落包含多个公式和文字：每个公式按行内公式输出。
- 表格单元格内公式同样遵循以上规则；简单表格输出 Markdown 表格时，需要对单元格内的 `$` 和换行做必要转义或压缩。

#### OMML 到 LaTeX

Word 公式底层通常是 OMML。实现上建议把公式转换拆成独立模块 `OmmlToLatexConverter`：

1. 先覆盖高频结构：普通文本、上下标、分数、根号、括号、求和、积分、希腊字母、常见函数。
2. 转换成功时输出纯 LaTeX，不附带 Word 专属 XML。
3. 遇到矩阵、cases、多行对齐等复杂结构时，可以先输出段落公式，并在内部尽量生成 `matrix`、`cases`、`aligned` 等 LaTeX 环境。
4. 转换失败时输出保守占位，例如：

```markdown
<!-- TODO: unsupported Word equation -->
```

后续如果引入第三方转换器，也应保持最终 Markdown 仍是 `$...$` / `$$...$$` 形式。

## HTML 到 Markdown 管线

建议不要直接把完整 HTML 丢给 ReverseMarkdown.Net 后结束。更稳的方式是先做 DOM 预处理：

```text
HTML DOM
  -> 找到 table 节点
  -> 分类并替换为占位符
  -> 图片 src 确保为相对路径
  -> 普通 HTML 转 Markdown
  -> 恢复表格 Markdown/HTML
  -> 后处理空行
```

占位符示例：

```text
@@TABLE_0001@@
```

这样可以避免 ReverseMarkdown.Net 把复杂表格错误压平成 Markdown 表格。

## 输出整理

Markdown 后处理规则：

- 文件使用 UTF-8。
- 标题前后保留空行。
- 列表内部不插入多余空行。
- 表格前后保留空行。
- 图片前后保留空行。
- 连续 3 个以上空行压缩为 2 个。
- 路径统一使用 `/`。

## 错误处理

常见错误：

- 输入文件不存在。
- 输入不是 `.docx`。
- 输出目录不可写。
- 图片资源目录不可写。
- docx 文件损坏或被 Word 占用。

要求：

- 错误信息用中文说明。
- 对单张图片导出失败时，不应让整个转换直接失败；可输出占位文本并给 warning。
- 对文档无法读取时，直接失败。

## 验证方案

准备样例文档：

- 普通标题和段落。
- 多级标题。
- 有序列表和无序列表。
- 嵌套列表。
- 简单表格。
- 合并单元格表格。
- 含图片文档。
- 含链接、加粗、斜体。
- 含脚注。
- 含公式。

验证点：

- `.md` 文件生成成功。
- 图片文件导出到 assets 目录。
- Markdown 中图片路径为相对路径。
- 简单表格为 Markdown pipe table。
- 合并单元格表格为 HTML table。
- Word 列表转换为 Markdown 列表。
- 标题层级正确。
- 行内公式输出为 `$...$`。
- 独立段落公式输出为 `$$...$$`。
- 无法转换的复杂公式保留位置，并给出 warning 或占位。

## 迭代顺序

建议按以下顺序实现：

1. CLI 子命令框架：`to-docx` / `to-md`。
2. `to-md` 基础文件读写。
3. Mammoth 转 HTML。
4. 图片导出和相对路径。
5. 普通 HTML 转 Markdown。
6. 表格分类：简单表格 Markdown，复杂表格 HTML。
7. 公式识别：行内公式 `$...$`，段落公式 `$$...$$`。
8. Word COM 或 OpenXML 验证样例。
9. 脚注、代码块和复杂公式增强。

## 关键决策

- `to-md` 不使用配置文件。
- 用户面对的是 Markdown 结果，不需要理解 Word 内部 `styleId`。
- 表格由项目代码接管，不完全依赖第三方 HTML 转 Markdown 库。
- 图片必须落盘并使用相对路径。
- 公式导出优先生成通用 Markdown 数学语法，而不是 Word 专属 XML。
- 首版以稳定可编辑 Markdown 为目标，不追求 Word 视觉完全还原。
