using System.Text.RegularExpressions;
using Markdig.Extensions.Tables;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

internal static class AstConverter
{
    private static readonly Regex MathBlockRegex = new(@"^\s*\$\$(?<body>[\s\S]*?)\$\$\s*$", RegexOptions.Compiled);
    private static readonly Regex InlineMathRegex = new(@"\$(?<body>[^$\r\n]+)\$", RegexOptions.Compiled);

    public static AstDocument Convert(MarkdownDocument document, string sourcePath)
    {
        var blocks = new List<AstBlock>();
        foreach (var block in document)
        {
            AppendBlock(block, blocks);
        }

        return new AstDocument(1, sourcePath, blocks);
    }

    private static void AppendBlock(Block block, List<AstBlock> blocks)
    {
        switch (block)
        {
            case HeadingBlock heading:
                blocks.Add(new AstBlock("heading") { Inlines = ConvertInlines(heading.Inline), Level = heading.Level });
                break;
            case ParagraphBlock paragraph:
                AppendParagraph(paragraph, blocks);
                break;
            case QuoteBlock quote:
                blocks.Add(new AstBlock("blockquote") { Blocks = ConvertChildBlocks(quote) });
                break;
            case FencedCodeBlock fenced:
                blocks.Add(new AstBlock("code_block") { Text = fenced.Lines.ToString(), Language = fenced.Info });
                break;
            case CodeBlock code:
                blocks.Add(new AstBlock("code_block") { Text = code.Lines.ToString() });
                break;
            case ThematicBreakBlock:
                blocks.Add(new AstBlock("thematic_break"));
                break;
            case ListBlock list:
                blocks.Add(ConvertList(list));
                break;
            case Markdig.Extensions.Tables.Table table:
                blocks.Add(ConvertTable(table));
                break;
            default:
                if (!string.IsNullOrWhiteSpace(block.ToString()))
                {
                    blocks.Add(new AstBlock("paragraph")
                    {
                        Inlines = [new AstInline("text") { Text = block.ToString() ?? string.Empty }]
                    });
                }
                break;
        }
    }

    private static void AppendParagraph(ParagraphBlock paragraph, List<AstBlock> blocks)
    {
        var text = paragraph.Inline?.FirstChild == null ? string.Empty : ExtractPlainText(paragraph.Inline);
        var match = MathBlockRegex.Match(text);
        if (match.Success)
        {
            blocks.Add(new AstBlock("math_block") { Text = match.Groups["body"].Value.Trim() });
            return;
        }

        blocks.Add(new AstBlock("paragraph") { Inlines = ConvertInlines(paragraph.Inline) });
    }

    private static AstBlock ConvertList(ListBlock list)
    {
        var items = new List<AstListItem>();
        foreach (var item in list.OfType<ListItemBlock>())
        {
            items.Add(new AstListItem(ConvertChildBlocks(item)));
        }

        return new AstBlock("list") { Ordered = list.IsOrdered, Start = list.OrderedStart, Items = items };
    }

    private static AstBlock ConvertTable(Markdig.Extensions.Tables.Table table)
    {
        var rows = new List<AstTableRow>();
        foreach (var rowObject in table)
        {
            if (rowObject is not Markdig.Extensions.Tables.TableRow row)
            {
                continue;
            }

            var cells = new List<AstTableCell>();
            foreach (var cellObject in row)
            {
                if (cellObject is Markdig.Extensions.Tables.TableCell cell)
                {
                    cells.Add(new AstTableCell(ConvertChildBlocks(cell)));
                }
            }

            rows.Add(new AstTableRow(row.IsHeader, cells));
        }

        var columnCount = rows.Count == 0 ? 0 : rows.Max(row => row.Cells.Count);
        var alignments = new List<string>();
        foreach (var definition in table.ColumnDefinitions)
        {
            if (alignments.Count >= columnCount)
            {
                break;
            }

            alignments.Add(definition.Alignment switch
            {
                TableColumnAlign.Center => "center",
                TableColumnAlign.Right => "right",
                TableColumnAlign.Left => "left",
                _ => "default"
            });
        }

        while (alignments.Count < columnCount)
        {
            alignments.Add("default");
        }

        return new AstBlock("table") { Rows = rows, Alignments = alignments };
    }

    private static List<AstBlock> ConvertChildBlocks(ContainerBlock container)
    {
        var blocks = new List<AstBlock>();
        foreach (var child in container)
        {
            AppendBlock(child, blocks);
        }

        return blocks;
    }

    private static List<AstInline> ConvertInlines(ContainerInline? container)
    {
        var result = new List<AstInline>();
        if (container == null)
        {
            return result;
        }

        foreach (var inline in container)
        {
            AppendInline(inline, result);
        }

        return MergeAdjacentText(result);
    }

    private static void AppendInline(Inline inline, List<AstInline> result)
    {
        switch (inline)
        {
            case LiteralInline literal:
                AppendTextWithMath(literal.Content.ToString(), result);
                break;
            case LineBreakInline:
                result.Add(new AstInline("line_break") { Text = "\n" });
                break;
            case CodeInline code:
                result.Add(new AstInline("inline_code") { Text = code.Content });
                break;
            case EmphasisInline emphasis:
                var emphasisType = emphasis.DelimiterCount >= 2 ? "strong" : "emphasis";
                result.Add(new AstInline(emphasisType) { Inlines = ConvertInlines(emphasis) });
                break;
            case LinkInline link when link.IsImage:
                result.Add(new AstInline("image") { Text = ExtractPlainText(link), Url = link.Url, Title = link.Title });
                break;
            case LinkInline link:
                result.Add(new AstInline("link") { Text = ExtractPlainText(link), Url = link.Url, Title = link.Title, Inlines = ConvertInlines(link) });
                break;
            case ContainerInline nested:
                result.AddRange(ConvertInlines(nested));
                break;
            default:
                var text = inline.ToString();
                if (!string.IsNullOrEmpty(text))
                {
                    AppendTextWithMath(text, result);
                }
                break;
        }
    }

    private static void AppendTextWithMath(string text, List<AstInline> result)
    {
        if (string.IsNullOrEmpty(text))
        {
            return;
        }

        var cursor = 0;
        foreach (Match match in InlineMathRegex.Matches(text))
        {
            if (match.Index > cursor)
            {
                result.Add(new AstInline("text") { Text = text[cursor..match.Index] });
            }

            result.Add(new AstInline("math_inline") { Text = match.Groups["body"].Value });
            cursor = match.Index + match.Length;
        }

        if (cursor < text.Length)
        {
            result.Add(new AstInline("text") { Text = text[cursor..] });
        }
    }

    private static List<AstInline> MergeAdjacentText(List<AstInline> inlines)
    {
        var merged = new List<AstInline>();
        foreach (var inline in inlines)
        {
            if (merged.Count > 0 && inline.Type == "text" && merged[^1].Type == "text")
            {
                merged[^1].Text += inline.Text;
                continue;
            }

            merged.Add(inline);
        }

        return merged;
    }

    private static string ExtractPlainText(ContainerInline? container)
    {
        if (container == null)
        {
            return string.Empty;
        }

        var parts = new List<string>();
        foreach (var inline in container)
        {
            switch (inline)
            {
                case LiteralInline literal:
                    parts.Add(literal.Content.ToString());
                    break;
                case CodeInline code:
                    parts.Add(code.Content);
                    break;
                case LineBreakInline:
                    parts.Add("\n");
                    break;
                case ContainerInline nested:
                    parts.Add(ExtractPlainText(nested));
                    break;
                default:
                    parts.Add(inline.ToString() ?? string.Empty);
                    break;
            }
        }

        return string.Concat(parts);
    }
}

