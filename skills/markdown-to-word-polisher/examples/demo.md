# Markdown 转 Word 示例

这是一段普通正文，包含 **加粗**、*斜体*、`行内代码`、[链接](https://example.com) 和 $E=mc^2$。

## 表格

| 字段 | 说明 | 对齐 |
|---|:---:|---:|
| input | Markdown 输入文件 | 左 |
| output | Word 输出文件 | 右 |

## 列表

- 解析 Markdown 为 AST
- 读取样式映射 JSON
- 调用 Word COM 生成文档

## 代码

```powershell
.\bin\mdtodocx.exe to-docx -i .\examples\demo.md -o .\out\demo.docx -t .\templates\template.docx -s .\config\style-map.json
```

## 公式

$$
\frac{a}{b}=\sqrt{c}
$$
