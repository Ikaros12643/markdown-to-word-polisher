# Markdown 转 Word 方案思路

## 一、目标

给定一份 Markdown 文档与一份预设好正文、标题、公式、表格样式的 Word 模板（`.docx`），通过 PowerShell 脚本调用 Word COM API 将 Markdown 转换为符合模板样式的 Word 文档。

## 二、技术路线

PowerShell 解析 Markdown → 产出有序元素序列 → Word COM API 基于模板新建文档 → 逐元素写入并套用模板内预定义的样式名。

## 三、两阶段架构

### 阶段一：Markdown 解析（PowerShell 自实现）

按行扫描，识别块级元素并产出有序元素列表：

| Markdown 语法 | 元素类型 | 映射样式 |
|---|---|---|
| `# / ## / ###` | Heading | 标题 1 / 标题 2 / 标题 3 … |
| 普通段落 | Paragraph | 正文 |
| ` ``` 代码块 ``` ` | CodeBlock | 代码 / 自定义 |
| `>` | Quote | 引用 |
| `- / * / 1.` | List | 列表 |
| ` \| a \| b \| ` | Table | 表格正文 |
| `$$...$$` | Formula | 公式 |
| `---` | HRule | — |

行内元素（`**粗体**`、`*斜体*`、`` `code` ``、`[链接](url)`、`![图](path)`、`$行内公式$`）在写入段落时通过 `Range.Font.Bold / Italic` 或单独插入处理。

### 阶段二：Word 文档生成（COM API）

1. `$doc = $word.Documents.Add("模板路径.docx")` —— 自动继承模板内全部样式定义
2. 遍历元素列表，向文档末尾追加：
   - 标题 / 正文：`$selection.TypeText()` + `$selection.Style = "标题 1"`
   - 表格：`$doc.Tables.Add($range, rows, cols)`，逐格填内容后整体 `.Range.Style = "表格正文"`
   - 行内格式：记录文本偏移量，定位子 Range 后设置 `.Font.Bold = $true` 等
3. `$doc.SaveAs("输出.docx")`

## 四、关键难点：公式处理

### 4.1 难点本质

COM API **能**新增公式对象（`Range.OMaths.Add()`），难点不在"能不能加"，而在于**输入格式转换**：

- Markdown / 学术写作中公式都是 **LaTeX 语法**（如 `\frac{a}{b}`、`\sum_{i=1}^{n}`）
- COM API **没有** `OMath.SetLatex("...")` 这样的方法
- Word 公式内部数据格式是 **OMML**（Office Math Markup Language，一种 XML）

要写入公式只有三条路：

1. **手写 OMML XML**：每个 `\frac`、`\sqrt`、上下标都要对应一段 XML 节点，语法冗长且文档稀缺，复杂公式极难构造。
2. **`Selection.OMaths.Add` + `BuildUp`**：先插入 LaTeX 纯文本，选中后调用 `OMath.BuildUp()` 让 Word 自动转换。理论可行，但：
   - 需要 Word 版本 ≥ 2016 且启用"公式以 LaTeX 输入"设置
   - 对 LaTeX 兼容有限，遇 `\begin{cases}`、矩阵、`\mathbb` 等常失败
   - COM 调用时机敏感，批量时易丢转换
3. **剪贴板塞 MathML**：把 LaTeX 转 MathML 后放进剪贴板，让 Word 粘贴为公式。需要额外的转换器（如 Pandoc / `latex2mathml`）。

结论：COM API 能"承载"公式对象，但**把 LaTeX 变成 Word 认识的公式**这一步它几乎不帮忙。

### 4.2 采用方案：公式块 + LaTeX 文本 + 手动转换

利用 Word 自身的"LaTeX → 专业型"转换能力，绕开 COM API 无法构造 OMML 的死结。

**工作原理：**

1. COM API 创建公式块（`Range.OMaths.Add()`）
2. 把 LaTeX 文本作为**线性文本写入公式区域内部**（不是普通段落）
3. 保存后退出
4. Word 打开时，每个公式块里就是待转换的 LaTeX
5. 用户全选后点"转换 → LaTeX 转为专业型"即可批量完成

**为什么这样可行：**

- LaTeX 文本进入 OMath zone 后，Word 把它当作"公式区域的线性内容"存储，识别为可转换公式
- 手动触发转换，等价于调用 `OMath.BuildUp()`，由 Word 原生引擎完成 LaTeX→OMML，兼容性最好
- 绕开了 COM API 无法直接构造 OMML 的死结

**两个要注意的坑：**

1. **必须进入"公式输入模式为 LaTeX"**：Word 默认公式输入是 UnicodeMath，需在「公式工具 → 转换」里把当前公式（或全局）切到 LaTeX 模式，转换按钮才会按 LaTeX 解析。否则 `\frac` 会被当字面量。
2. **写入顺序**：先 `OMaths.Add` 建 zone，再让 Selection 落进 zone 内 `TypeText` LaTeX 文本，顺序反了文本会跑到 zone 外。

### 4.3 可选的两种执行模式

- **纯手动**：脚本只负责建公式块 + 塞 LaTeX 文本，用户打开 Word 全选逐个转换
- **半自动**：脚本塞完 LaTeX 后直接调用 `OMath.BuildUp()` 尝试自动转换，转换失败的（复杂公式）留给用户手动处理

> 半自动方案要求 Word ≥ 2016。

## 五、其他注意点

1. **表格对齐**：markdown 的 `:---:` / `:--` / `---:` 需映射为单元格 `ParagraphFormat.Alignment`（居中 / 左 / 右）。
2. **行内格式定位**：需在写入文本时记录起止偏移，再回查 Range 设置 Bold / Italic，否则格式会"溢出"到相邻文本。
3. **模板样式名约定**：脚本中引用的样式名（如 `标题 1`、`正文`、`表格正文`、`公式`）必须与模板中已定义的样式名完全一致，否则 COM API 会报错或套用默认样式。
