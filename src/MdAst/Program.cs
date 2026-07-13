using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;
using Markdig;
using Markdig.Extensions.Tables;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;

var options = CliOptions.Parse(args);
if (options.ShowHelp)
{
    CliOptions.WriteHelp();
    return 0;
}

if (string.IsNullOrWhiteSpace(options.InputPath))
{
    Console.Error.WriteLine("Missing required argument: --input <path>");
    CliOptions.WriteHelp();
    return 2;
}

try
{
    var inputPath = Path.GetFullPath(options.InputPath);
    if (!File.Exists(inputPath))
    {
        Console.Error.WriteLine("Input file does not exist: " + inputPath);
        return 2;
    }

    var markdown = File.ReadAllText(inputPath);
    var pipeline = new MarkdownPipelineBuilder()
        .UsePipeTables()
        .Build();

    var document = Markdown.Parse(markdown, pipeline);
    var ast = AstConverter.Convert(document, inputPath);

    var jsonOptions = new JsonSerializerOptions
    {
        WriteIndented = true,
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
    };
    var json = JsonSerializer.Serialize(ast, jsonOptions);

    if (string.IsNullOrWhiteSpace(options.OutputPath))
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        Console.WriteLine(json);
    }
    else
    {
        var outputPath = Path.GetFullPath(options.OutputPath);
        var outputDirectory = Path.GetDirectoryName(outputPath);
        if (!string.IsNullOrWhiteSpace(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        File.WriteAllText(outputPath, json, System.Text.Encoding.UTF8);
    }

    return 0;
}
catch (Exception ex)
{
    Console.Error.WriteLine(ex.Message);
    return 1;
}

internal sealed class CliOptions
{
    public string? InputPath { get; private init; }
    public string? OutputPath { get; private init; }
    public bool ShowHelp { get; private init; }

    public static CliOptions Parse(string[] args)
    {
        var options = new CliOptions();
        var inputPath = string.Empty;
        var outputPath = string.Empty;
        var showHelp = false;

        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            if (arg == "-h" || arg == "--help" || arg == "/?")
            {
                showHelp = true;
                continue;
            }

            if (arg == "--input" || arg == "-i")
            {
                inputPath = ReadValue(args, ref i, arg);
                continue;
            }

            if (arg == "--output" || arg == "-o")
            {
                outputPath = ReadValue(args, ref i, arg);
                continue;
            }

            if (string.IsNullOrWhiteSpace(inputPath))
            {
                inputPath = arg;
                continue;
            }

            throw new ArgumentException("Unknown argument: " + arg);
        }

        options = new CliOptions
        {
            InputPath = inputPath,
            OutputPath = outputPath,
            ShowHelp = showHelp
        };
        return options;
    }

    public static void WriteHelp()
    {
        Console.WriteLine("Usage:");
        Console.WriteLine("  mdast --input <markdown-path> [--output <json-path>]");
        Console.WriteLine("  mdast -i <markdown-path> -o <json-path>");
    }

    private static string ReadValue(string[] args, ref int index, string name)
    {
        if (index + 1 >= args.Length)
        {
            throw new ArgumentException("Missing value for " + name);
        }

        index++;
        return args[index];
    }
}

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
                blocks.Add(new AstBlock(
                    Type: "heading",
                    Inlines: ConvertInlines(heading.Inline),
                    Level: heading.Level));
                break;

            case ParagraphBlock paragraph:
                AppendParagraph(paragraph, blocks);
                break;

            case QuoteBlock quote:
                blocks.Add(new AstBlock(
                    Type: "blockquote",
                    Blocks: ConvertChildBlocks(quote)));
                break;

            case FencedCodeBlock fenced:
                blocks.Add(new AstBlock(
                    Type: "code_block",
                    Text: fenced.Lines.ToString(),
                    Language: fenced.Info));
                break;

            case CodeBlock code:
                blocks.Add(new AstBlock(
                    Type: "code_block",
                    Text: code.Lines.ToString()));
                break;

            case ThematicBreakBlock:
                blocks.Add(new AstBlock(Type: "thematic_break"));
                break;

            case ListBlock list:
                blocks.Add(ConvertList(list));
                break;

            case Table table:
                blocks.Add(ConvertTable(table));
                break;

            default:
                if (!string.IsNullOrWhiteSpace(block.ToString()))
                {
                    blocks.Add(new AstBlock(Type: "paragraph", Inlines: new List<AstInline>
                    {
                        new("text", block.ToString() ?? string.Empty)
                    }));
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
            blocks.Add(new AstBlock(Type: "math_block", Text: match.Groups["body"].Value.Trim()));
            return;
        }

        blocks.Add(new AstBlock(Type: "paragraph", Inlines: ConvertInlines(paragraph.Inline)));
    }

    private static AstBlock ConvertList(ListBlock list)
    {
        var items = new List<AstListItem>();
        foreach (var item in list.OfType<ListItemBlock>())
        {
            items.Add(new AstListItem(ConvertChildBlocks(item)));
        }

        return new AstBlock(
            Type: "list",
            Ordered: list.IsOrdered,
            Start: list.OrderedStart,
            Items: items);
    }

    private static AstBlock ConvertTable(Table table)
    {
        var rows = new List<AstTableRow>();
        foreach (var rowObject in table)
        {
            if (rowObject is not TableRow row)
            {
                continue;
            }

            var cells = new List<AstTableCell>();
            foreach (var cellObject in row)
            {
                if (cellObject is not TableCell cell)
                {
                    continue;
                }

                cells.Add(new AstTableCell(ConvertChildBlocks(cell)));
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

        return new AstBlock(Type: "table", Rows: rows, Alignments: alignments);
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
                result.Add(new AstInline("line_break", "\n"));
                break;

            case CodeInline code:
                result.Add(new AstInline("inline_code", code.Content));
                break;

            case EmphasisInline emphasis:
                var emphasisType = emphasis.DelimiterCount >= 2 ? "strong" : "emphasis";
                result.Add(new AstInline(emphasisType, Inlines: ConvertInlines(emphasis)));
                break;

            case LinkInline link when link.IsImage:
                result.Add(new AstInline(
                    Type: "image",
                    Text: ExtractPlainText(link),
                    Url: link.Url,
                    Title: link.Title));
                break;

            case LinkInline link:
                result.Add(new AstInline(
                    Type: "link",
                    Text: ExtractPlainText(link),
                    Url: link.Url,
                    Title: link.Title,
                    Inlines: ConvertInlines(link)));
                break;

            case ContainerInline nested:
                foreach (var child in ConvertInlines(nested))
                {
                    result.Add(child);
                }
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
                result.Add(new AstInline("text", text.Substring(cursor, match.Index - cursor)));
            }

            result.Add(new AstInline("math_inline", match.Groups["body"].Value));
            cursor = match.Index + match.Length;
        }

        if (cursor < text.Length)
        {
            result.Add(new AstInline("text", text.Substring(cursor)));
        }
    }

    private static List<AstInline> MergeAdjacentText(List<AstInline> inlines)
    {
        var merged = new List<AstInline>();
        foreach (var inline in inlines)
        {
            if (merged.Count > 0 && inline.Type == "text" && merged[^1].Type == "text")
            {
                var previous = merged[^1];
                merged[^1] = previous with { Text = previous.Text + inline.Text };
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

internal sealed record AstDocument(int Version, string SourcePath, List<AstBlock> Blocks);

internal sealed record AstBlock(
    string Type,
    List<AstInline>? Inlines = null,
    int? Level = null,
    string? Text = null,
    string? Language = null,
    bool? Ordered = null,
    string? Start = null,
    List<AstListItem>? Items = null,
    List<AstBlock>? Blocks = null,
    List<AstTableRow>? Rows = null,
    List<string>? Alignments = null);

internal sealed record AstInline(
    string Type,
    string? Text = null,
    string? Url = null,
    string? Title = null,
    List<AstInline>? Inlines = null);

internal sealed record AstListItem(List<AstBlock> Blocks);

internal sealed record AstTableRow(bool IsHeader, List<AstTableCell> Cells);

internal sealed record AstTableCell(List<AstBlock> Blocks);
