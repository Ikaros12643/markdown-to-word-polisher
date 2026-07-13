using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig;
using Markdig.Extensions.Tables;
using Markdig.Syntax;
using Markdig.Syntax.Inlines;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

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

if (string.IsNullOrWhiteSpace(options.OutputPath))
{
    Console.Error.WriteLine("Missing required argument: --output <path>");
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

    var outputPath = Path.GetFullPath(options.OutputPath);
    if (options.Command == CliCommand.ToMarkdown)
    {
        var assetsDir = string.IsNullOrWhiteSpace(options.AssetsDir)
            ? Path.Combine(Path.GetDirectoryName(outputPath) ?? Directory.GetCurrentDirectory(), Path.GetFileNameWithoutExtension(outputPath) + ".assets")
            : Path.GetFullPath(options.AssetsDir);
        var assetsRelativePath = string.IsNullOrWhiteSpace(options.AssetsRelativePath)
            ? Path.GetRelativePath(Path.GetDirectoryName(outputPath) ?? Directory.GetCurrentDirectory(), assetsDir)
            : options.AssetsRelativePath;

        DocxToMarkdownConverter.Convert(new DocxToMarkdownOptions(inputPath, outputPath, assetsDir, NormalizeMarkdownPath(assetsRelativePath)));
        Console.WriteLine("Markdown written to: " + outputPath);
        return 0;
    }

    var styleMapPath = string.IsNullOrWhiteSpace(options.StyleMapPath)
        ? Path.Combine(AppContext.BaseDirectory, "config", "style-map.json")
        : Path.GetFullPath(options.StyleMapPath);

    if (!File.Exists(styleMapPath))
    {
        Console.Error.WriteLine("Style map file does not exist: " + styleMapPath);
        return 2;
    }

    var templatePath = string.IsNullOrWhiteSpace(options.TemplatePath) ? null : Path.GetFullPath(options.TemplatePath);
    if (!string.IsNullOrWhiteSpace(templatePath) && !File.Exists(templatePath))
    {
        Console.Error.WriteLine("Template file does not exist: " + templatePath);
        return 2;
    }

    var outputDirectory = Path.GetDirectoryName(outputPath);
    if (!string.IsNullOrWhiteSpace(outputDirectory))
    {
        Directory.CreateDirectory(outputDirectory);
    }

    var markdown = File.ReadAllText(inputPath);
    var pipeline = new MarkdownPipelineBuilder()
        .UsePipeTables()
        .Build();
    var markdownDocument = Markdown.Parse(markdown, pipeline);
    var ast = AstConverter.Convert(markdownDocument, inputPath);
    var styleMap = StyleMap.Load(styleMapPath);

    if (options.WriteAst)
    {
        var astPath = Path.ChangeExtension(outputPath, ".ast.json");
        var jsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
        };
        File.WriteAllText(astPath, JsonSerializer.Serialize(ast, jsonOptions), System.Text.Encoding.UTF8);
        Console.WriteLine("AST written to: " + astPath);
    }

    DocxWriter.Write(ast, styleMap, inputPath, outputPath, templatePath, options.KeepHeadingNumbers);
    Console.WriteLine("DOCX written to: " + outputPath);
    return 0;
}
catch (Exception ex)
{
    Console.Error.WriteLine(ex);
    return 1;
}

static string NormalizeMarkdownPath(string path)
{
    return path.Replace('\\', '/').TrimEnd('/');
}

internal sealed class CliOptions
{
    public CliCommand Command { get; private init; } = CliCommand.ToDocx;
    public string? InputPath { get; private init; }
    public string? OutputPath { get; private init; }
    public string? TemplatePath { get; private init; }
    public string? StyleMapPath { get; private init; }
    public string? AssetsDir { get; private init; }
    public string? AssetsRelativePath { get; private init; }
    public bool KeepHeadingNumbers { get; private init; }
    public bool WriteAst { get; private init; }
    public bool ShowHelp { get; private init; }

    public static CliOptions Parse(string[] args)
    {
        var inputPath = string.Empty;
        var outputPath = string.Empty;
        var templatePath = string.Empty;
        var styleMapPath = string.Empty;
        var assetsDir = string.Empty;
        var assetsRelativePath = string.Empty;
        var keepHeadingNumbers = false;
        var writeAst = false;
        var showHelp = false;
        var command = CliCommand.ToDocx;
        var startIndex = 0;

        if (args.Length > 0)
        {
            if (string.Equals(args[0], "to-md", StringComparison.OrdinalIgnoreCase))
            {
                command = CliCommand.ToMarkdown;
                startIndex = 1;
            }
            else if (string.Equals(args[0], "to-docx", StringComparison.OrdinalIgnoreCase))
            {
                command = CliCommand.ToDocx;
                startIndex = 1;
            }
        }

        for (var i = startIndex; i < args.Length; i++)
        {
            var arg = args[i];
            switch (arg)
            {
                case "-h":
                case "--help":
                case "/?":
                    showHelp = true;
                    break;
                case "--input":
                case "-i":
                    inputPath = ReadValue(args, ref i, arg);
                    break;
                case "--output":
                case "-o":
                    outputPath = ReadValue(args, ref i, arg);
                    break;
                case "--template":
                case "-t":
                    templatePath = ReadValue(args, ref i, arg);
                    break;
                case "--style-map":
                case "-s":
                    styleMapPath = ReadValue(args, ref i, arg);
                    break;
                case "--assets-dir":
                case "-a":
                    assetsDir = ReadValue(args, ref i, arg);
                    break;
                case "--assets-relative-path":
                    assetsRelativePath = ReadValue(args, ref i, arg);
                    break;
                case "--keep-heading-numbers":
                    keepHeadingNumbers = true;
                    break;
                case "--write-ast":
                    writeAst = true;
                    break;
                default:
                    if (string.IsNullOrWhiteSpace(inputPath))
                    {
                        inputPath = arg;
                    }
                    else if (string.IsNullOrWhiteSpace(outputPath))
                    {
                        outputPath = arg;
                    }
                    else
                    {
                        throw new ArgumentException("Unknown argument: " + arg);
                    }
                    break;
            }
        }

        return new CliOptions
        {
            Command = command,
            InputPath = inputPath,
            OutputPath = outputPath,
            TemplatePath = templatePath,
            StyleMapPath = styleMapPath,
            AssetsDir = assetsDir,
            AssetsRelativePath = assetsRelativePath,
            KeepHeadingNumbers = keepHeadingNumbers,
            WriteAst = writeAst,
            ShowHelp = showHelp
        };
    }

    public static void WriteHelp()
    {
        Console.WriteLine("Usage:");
        Console.WriteLine("  mdtodocx to-docx --input <markdown-path> --output <docx-path> [--template <docx-path>] [--style-map <json-path>]");
        Console.WriteLine("  mdtodocx to-md --input <docx-path> --output <markdown-path> [--assets-dir <dir>]");
        Console.WriteLine("  mdtodocx -i <markdown-path> -o <docx-path> -t <template.docx> -s <style-map.json>");
        Console.WriteLine();
        Console.WriteLine("Options:");
        Console.WriteLine("  -i, --input             Input file.");
        Console.WriteLine("  -o, --output            Output file.");
        Console.WriteLine("  -t, --template          Word template for to-docx.");
        Console.WriteLine("  -s, --style-map         Style map JSON for to-docx.");
        Console.WriteLine("  -a, --assets-dir        Asset output directory for to-md.");
        Console.WriteLine("  --assets-relative-path  Asset path prefix written to Markdown for to-md.");
        Console.WriteLine("  --keep-heading-numbers  Preserve leading heading numbers from Markdown.");
        Console.WriteLine("  --write-ast             Also write <output>.ast.json for debugging.");
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

internal enum CliCommand
{
    ToDocx,
    ToMarkdown
}

internal sealed record DocxToMarkdownOptions(string InputPath, string OutputPath, string AssetsDir, string AssetsRelativePath);

internal static class DocxToMarkdownConverter
{
    public static void Convert(DocxToMarkdownOptions options)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(options.OutputPath) ?? Directory.GetCurrentDirectory());
        Directory.CreateDirectory(options.AssetsDir);

        using var document = WordprocessingDocument.Open(options.InputPath, false);
        var mainPart = document.MainDocumentPart ?? throw new InvalidOperationException("DOCX does not contain a main document part.");
        var body = mainPart.Document.Body ?? throw new InvalidOperationException("DOCX does not contain a document body.");
        var context = new DocxReadContext(mainPart, options);
        var writer = new StringBuilder();

        foreach (var child in body.Elements())
        {
            if (child is Wp.Paragraph paragraph)
            {
                AppendParagraph(writer, paragraph, context);
            }
            else if (child is Wp.Table table)
            {
                AppendBlock(writer, TableToMarkdown(table, context), context);
            }
        }

        File.WriteAllText(options.OutputPath, NormalizeMarkdown(writer.ToString()), Encoding.UTF8);
    }

    private static void AppendParagraph(StringBuilder writer, Wp.Paragraph paragraph, DocxReadContext context)
    {
        var content = RenderParagraphContent(paragraph, context);
        if (string.IsNullOrWhiteSpace(content.Markdown))
        {
            return;
        }

        var headingLevel = context.GetHeadingLevel(paragraph);
        if (headingLevel != null)
        {
            AppendBlock(writer, new string('#', headingLevel.Value) + " " + content.Markdown.Trim(), context);
            return;
        }

        var listInfo = context.GetListInfo(paragraph);
        if (listInfo != null)
        {
            var indent = new string(' ', listInfo.Level * 2);
            var marker = listInfo.Ordered ? "1. " : "- ";
            AppendListItem(writer, indent + marker + content.Markdown.Trim(), context);
            return;
        }

        var paragraphStyleName = context.GetParagraphStyleName(paragraph);
        if (ContainsIgnoreCase(paragraphStyleName, "code") || ContainsIgnoreCase(paragraphStyleName, "代码"))
        {
            AppendBlock(writer, "```" + Environment.NewLine + content.Markdown.Trim() + Environment.NewLine + "```", context);
            return;
        }

        if (content.FormulaCount == 1 && string.IsNullOrWhiteSpace(content.NonFormulaText) && !string.IsNullOrWhiteSpace(content.SingleFormulaLatex))
        {
            AppendBlock(writer, "$$" + Environment.NewLine + content.SingleFormulaLatex + Environment.NewLine + "$$", context);
            return;
        }

        AppendBlock(writer, content.Markdown.Trim(), context);
    }

    private static ParagraphRenderResult RenderParagraphContent(Wp.Paragraph paragraph, DocxReadContext context)
    {
        var result = new ParagraphRenderResult();
        foreach (var child in paragraph.ChildElements)
        {
            result.Append(RenderInlineElement(child, context));
        }

        return result;
    }

    private static InlineRenderResult RenderInlineElement(OpenXmlElement element, DocxReadContext context)
    {
        switch (element)
        {
            case Wp.Text text:
                return InlineRenderResult.Text(EscapeMarkdownText(text.Text));
            case Wp.TabChar:
                return InlineRenderResult.Text("\t");
            case Wp.Break:
                return InlineRenderResult.Text("  " + Environment.NewLine);
            case Wp.Drawing drawing:
                return RenderDrawing(drawing, context);
            case Wp.Hyperlink hyperlink:
                return RenderHyperlink(hyperlink, context);
        }

        if (IsMathElement(element))
        {
            var latex = OmmlToLatexConverter.Convert(element);
            if (string.IsNullOrWhiteSpace(latex))
            {
                latex = "% unsupported Word equation";
            }

            return InlineRenderResult.Formula(latex);
        }

        if (element is Wp.Run run)
        {
            var childResult = RenderChildren(run, context);
            return ApplyRunFormatting(run, childResult, context);
        }

        return RenderChildren(element, context);
    }

    private static InlineRenderResult RenderChildren(OpenXmlElement element, DocxReadContext context)
    {
        var result = new InlineRenderResult();
        foreach (var child in element.ChildElements)
        {
            result.Append(RenderInlineElement(child, context));
        }

        return result;
    }

    private static InlineRenderResult ApplyRunFormatting(Wp.Run run, InlineRenderResult result, DocxReadContext context)
    {
        if (string.IsNullOrEmpty(result.Markdown) || result.HasNonTextContent)
        {
            return result;
        }

        var properties = run.RunProperties;
        var styleName = context.GetRunStyleName(properties?.RunStyle?.Val?.Value);
        var isCode = ContainsIgnoreCase(styleName, "code") || ContainsIgnoreCase(styleName, "代码");
        if (isCode)
        {
            result.Markdown = "`" + result.Markdown.Replace("`", "\\`") + "`";
            return result;
        }

        var bold = properties?.Bold != null && properties.Bold.Val?.Value != false;
        var italic = properties?.Italic != null && properties.Italic.Val?.Value != false;
        if (bold)
        {
            result.Markdown = "**" + result.Markdown + "**";
        }

        if (italic)
        {
            result.Markdown = "*" + result.Markdown + "*";
        }

        return result;
    }

    private static InlineRenderResult RenderHyperlink(Wp.Hyperlink hyperlink, DocxReadContext context)
    {
        var content = RenderChildren(hyperlink, context);
        var target = string.Empty;
        if (!string.IsNullOrWhiteSpace(hyperlink.Id?.Value))
        {
            target = context.GetHyperlinkTarget(hyperlink.Id!.Value);
        }
        else if (!string.IsNullOrWhiteSpace(hyperlink.Anchor?.Value))
        {
            target = "#" + hyperlink.Anchor!.Value;
        }

        if (string.IsNullOrWhiteSpace(target) || string.IsNullOrWhiteSpace(content.Markdown))
        {
            return content;
        }

        content.Markdown = "[" + content.Markdown + "](" + EscapeMarkdownUrl(target) + ")";
        content.HasNonTextContent = true;
        return content;
    }

    private static InlineRenderResult RenderDrawing(Wp.Drawing drawing, DocxReadContext context)
    {
        var blip = drawing.Descendants<A.Blip>().FirstOrDefault();
        var relationshipId = blip?.Embed?.Value ?? blip?.Link?.Value;
        if (string.IsNullOrWhiteSpace(relationshipId))
        {
            return InlineRenderResult.Text("[图片]");
        }

        try
        {
            var part = context.MainPart.GetPartById(relationshipId);
            if (part is not ImagePart imagePart)
            {
                return InlineRenderResult.Text("[图片]");
            }

            var alt = drawing.Descendants<DW.DocProperties>().FirstOrDefault()?.Description?.Value
                ?? drawing.Descendants<DW.DocProperties>().FirstOrDefault()?.Name?.Value
                ?? "image";
            var relativePath = context.ExportImage(imagePart);
            return InlineRenderResult.NonText("![" + EscapeMarkdownAlt(alt) + "](" + EscapeMarkdownUrl(relativePath) + ")");
        }
        catch
        {
            return InlineRenderResult.Text("[图片]");
        }
    }

    private static string TableToMarkdown(Wp.Table table, DocxReadContext context)
    {
        return IsSimpleTable(table) ? SimpleTableToMarkdown(table, context) : ComplexTableToHtml(table, context);
    }

    private static bool IsSimpleTable(Wp.Table table)
    {
        foreach (var cell in table.Descendants<Wp.TableCell>())
        {
            var properties = cell.TableCellProperties;
            if (properties?.GridSpan != null || properties?.HorizontalMerge != null || properties?.VerticalMerge != null)
            {
                return false;
            }
        }

        var widths = table.Elements<Wp.TableRow>().Select(row => row.Elements<Wp.TableCell>().Count()).Distinct().ToList();
        return widths.Count <= 1;
    }

    private static string SimpleTableToMarkdown(Wp.Table table, DocxReadContext context)
    {
        var rows = table.Elements<Wp.TableRow>()
            .Select(row => row.Elements<Wp.TableCell>().Select(cell => TableCellToMarkdownText(cell, context)).ToList())
            .Where(row => row.Count > 0)
            .ToList();
        if (rows.Count == 0)
        {
            return string.Empty;
        }

        var columnCount = rows.Max(row => row.Count);
        foreach (var row in rows)
        {
            while (row.Count < columnCount)
            {
                row.Add(string.Empty);
            }
        }

        var builder = new StringBuilder();
        builder.AppendLine("| " + string.Join(" | ", rows[0].Select(EscapeMarkdownTableCell)) + " |");
        builder.AppendLine("| " + string.Join(" | ", Enumerable.Repeat("---", columnCount)) + " |");
        foreach (var row in rows.Skip(1))
        {
            builder.AppendLine("| " + string.Join(" | ", row.Select(EscapeMarkdownTableCell)) + " |");
        }

        return builder.ToString().TrimEnd();
    }

    private static string ComplexTableToHtml(Wp.Table table, DocxReadContext context)
    {
        var builder = new StringBuilder();
        builder.AppendLine("<table>");
        foreach (var row in table.Elements<Wp.TableRow>())
        {
            builder.AppendLine("  <tr>");
            foreach (var cell in row.Elements<Wp.TableCell>())
            {
                var properties = cell.TableCellProperties;
                var attributes = new List<string>();
                var colspan = properties?.GridSpan?.Val?.Value;
                if (colspan != null && colspan > 1)
                {
                    attributes.Add("colspan=\"" + colspan + "\"");
                }

                var verticalMerge = properties?.VerticalMerge;
                if (verticalMerge != null)
                {
                    attributes.Add("data-vmerge=\"" + (verticalMerge.Val?.Value.ToString() ?? "continue") + "\"");
                }

                var openTag = attributes.Count == 0 ? "<td>" : "<td " + string.Join(" ", attributes) + ">";
                builder.Append("    ").Append(openTag);
                builder.Append(TableCellToHtmlText(cell, context));
                builder.AppendLine("</td>");
            }

            builder.AppendLine("  </tr>");
        }

        builder.AppendLine("</table>");
        return builder.ToString().TrimEnd();
    }

    private static string TableCellToMarkdownText(Wp.TableCell cell, DocxReadContext context)
    {
        var parts = new List<string>();
        foreach (var paragraph in cell.Elements<Wp.Paragraph>())
        {
            var content = RenderParagraphContent(paragraph, context);
            if (!string.IsNullOrWhiteSpace(content.Markdown))
            {
                parts.Add(content.Markdown.Trim());
            }
        }

        return string.Join("<br>", parts);
    }

    private static string TableCellToHtmlText(Wp.TableCell cell, DocxReadContext context)
    {
        var parts = new List<string>();
        foreach (var paragraph in cell.Elements<Wp.Paragraph>())
        {
            var content = RenderParagraphContent(paragraph, context);
            if (!string.IsNullOrWhiteSpace(content.Markdown))
            {
                parts.Add(EscapeHtml(content.Markdown.Trim()));
            }
        }

        return string.Join("<br>", parts);
    }

    private static void AppendBlock(StringBuilder writer, string block, DocxReadContext? context = null)
    {
        if (string.IsNullOrWhiteSpace(block))
        {
            return;
        }

        if (writer.Length > 0)
        {
            writer.AppendLine();
            writer.AppendLine();
        }

        writer.Append(block.Trim());
        if (context != null)
        {
            context.LastBlockKind = MarkdownBlockKind.Other;
        }
    }

    private static void AppendListItem(StringBuilder writer, string block, DocxReadContext context)
    {
        if (string.IsNullOrWhiteSpace(block))
        {
            return;
        }

        if (writer.Length > 0)
        {
            writer.AppendLine();
            if (context.LastBlockKind != MarkdownBlockKind.List)
            {
                writer.AppendLine();
            }
        }

        writer.Append(block.TrimEnd());
        context.LastBlockKind = MarkdownBlockKind.List;
    }

    private static bool IsMathElement(OpenXmlElement element)
    {
        return element.NamespaceUri == "http://schemas.openxmlformats.org/officeDocument/2006/math"
            && (element.LocalName == "oMath" || element.LocalName == "oMathPara");
    }

    private static string NormalizeMarkdown(string markdown)
    {
        return Regex.Replace(markdown.TrimEnd(), @"\n{3,}", Environment.NewLine + Environment.NewLine) + Environment.NewLine;
    }

    private static string EscapeMarkdownText(string text)
    {
        return text
            .Replace("\\", "\\\\")
            .Replace("[", "\\[")
            .Replace("]", "\\]")
            .Replace("*", "\\*")
            .Replace("_", "\\_")
            .Replace("`", "\\`");
    }

    private static string EscapeMarkdownAlt(string text)
    {
        return text.Replace("[", "\\[").Replace("]", "\\]");
    }

    private static string EscapeMarkdownUrl(string text)
    {
        return text.Replace("\\", "/").Replace(" ", "%20");
    }

    private static string EscapeMarkdownTableCell(string text)
    {
        return text.Replace("|", "\\|").Replace(Environment.NewLine, "<br>").Trim();
    }

    private static string EscapeHtml(string text)
    {
        return text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;");
    }

    private static bool ContainsIgnoreCase(string? text, string expected)
    {
        return text?.IndexOf(expected, StringComparison.OrdinalIgnoreCase) >= 0;
    }
}

internal sealed class DocxReadContext
{
    private readonly DocxToMarkdownOptions _options;
    private readonly Dictionary<string, StyleInfo> _styles = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, bool> _numberingIsOrdered = new(StringComparer.OrdinalIgnoreCase);
    private int _imageIndex;

    public DocxReadContext(MainDocumentPart mainPart, DocxToMarkdownOptions options)
    {
        MainPart = mainPart;
        _options = options;
        LoadStyles();
        LoadNumbering();
    }

    public MainDocumentPart MainPart { get; }
    public MarkdownBlockKind LastBlockKind { get; set; } = MarkdownBlockKind.None;

    public int? GetHeadingLevel(Wp.Paragraph paragraph)
    {
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (string.IsNullOrWhiteSpace(styleId) || !_styles.TryGetValue(styleId, out var style))
        {
            return null;
        }

        if (style.OutlineLevel is >= 0 and <= 5)
        {
            return style.OutlineLevel.Value + 1;
        }

        var name = style.Name ?? styleId;
        var match = Regex.Match(name, @"(?:Heading|标题)\s*([1-6])", RegexOptions.IgnoreCase);
        return match.Success ? int.Parse(match.Groups[1].Value) : null;
    }

    public ListInfo? GetListInfo(Wp.Paragraph paragraph)
    {
        var properties = paragraph.ParagraphProperties;
        var level = properties?.NumberingProperties?.NumberingLevelReference?.Val?.Value ?? 0;
        var numId = properties?.NumberingProperties?.NumberingId?.Val?.Value;
        if (numId != null)
        {
            var numIdText = numId.Value.ToString();
            return new ListInfo(_numberingIsOrdered.GetValueOrDefault(numIdText, true), level);
        }

        var styleId = properties?.ParagraphStyleId?.Val?.Value;
        if (!string.IsNullOrWhiteSpace(styleId) && _styles.TryGetValue(styleId, out var style))
        {
            var name = style.Name ?? string.Empty;
            if (name.Contains("有序", StringComparison.OrdinalIgnoreCase) || name.Contains("Number", StringComparison.OrdinalIgnoreCase))
            {
                return new ListInfo(true, 0);
            }

            if (name.Contains("列表", StringComparison.OrdinalIgnoreCase) || name.Contains("List", StringComparison.OrdinalIgnoreCase))
            {
                return new ListInfo(false, 0);
            }
        }

        return null;
    }

    public string? GetRunStyleName(string? styleId)
    {
        return !string.IsNullOrWhiteSpace(styleId) && _styles.TryGetValue(styleId, out var style) ? style.Name : null;
    }

    public string? GetParagraphStyleName(Wp.Paragraph paragraph)
    {
        var styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        return !string.IsNullOrWhiteSpace(styleId) && _styles.TryGetValue(styleId, out var style) ? style.Name : null;
    }

    public string GetHyperlinkTarget(string relationshipId)
    {
        return MainPart.HyperlinkRelationships.FirstOrDefault(link => link.Id == relationshipId)?.Uri.ToString() ?? string.Empty;
    }

    public string ExportImage(ImagePart imagePart)
    {
        _imageIndex++;
        var extension = ImageExtension(imagePart.ContentType);
        var fileName = "image-" + _imageIndex.ToString("000") + extension;
        var outputPath = Path.Combine(_options.AssetsDir, fileName);
        using (var source = imagePart.GetStream(FileMode.Open, FileAccess.Read))
        using (var target = File.Create(outputPath))
        {
            source.CopyTo(target);
        }

        return string.IsNullOrWhiteSpace(_options.AssetsRelativePath)
            ? fileName
            : _options.AssetsRelativePath.TrimEnd('/', '\\').Replace('\\', '/') + "/" + fileName;
    }

    private void LoadStyles()
    {
        var styles = MainPart.StyleDefinitionsPart?.Styles;
        if (styles == null)
        {
            return;
        }

        foreach (var style in styles.Elements<Style>())
        {
            var styleId = style.StyleId?.Value;
            if (string.IsNullOrWhiteSpace(styleId))
            {
                continue;
            }

            var outlineLevel = style.StyleParagraphProperties?.OutlineLevel?.Val?.Value;
            _styles[styleId] = new StyleInfo(style.StyleName?.Val?.Value, outlineLevel);
        }
    }

    private void LoadNumbering()
    {
        var numbering = MainPart.NumberingDefinitionsPart?.Numbering;
        if (numbering == null)
        {
            return;
        }

        var abstractFormats = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        foreach (var abstractNum in numbering.Elements<AbstractNum>())
        {
            var abstractNumId = abstractNum.AbstractNumberId?.Value.ToString();
            if (string.IsNullOrWhiteSpace(abstractNumId))
            {
                continue;
            }

            var format = abstractNum.Elements<Level>().FirstOrDefault()?.NumberingFormat?.Val?.Value;
            abstractFormats[abstractNumId] = format != NumberFormatValues.Bullet;
        }

        foreach (var instance in numbering.Elements<NumberingInstance>())
        {
            var numId = instance.NumberID?.Value.ToString();
            var abstractNumId = instance.AbstractNumId?.Val?.Value.ToString();
            if (!string.IsNullOrWhiteSpace(numId) && !string.IsNullOrWhiteSpace(abstractNumId))
            {
                _numberingIsOrdered[numId] = abstractFormats.GetValueOrDefault(abstractNumId, true);
            }
        }
    }

    private static string ImageExtension(string contentType)
    {
        return contentType.ToLowerInvariant() switch
        {
            "image/png" => ".png",
            "image/gif" => ".gif",
            "image/bmp" => ".bmp",
            "image/tiff" => ".tiff",
            "image/x-emf" => ".emf",
            "image/x-wmf" => ".wmf",
            _ => ".jpg"
        };
    }
}

internal sealed record StyleInfo(string? Name, int? OutlineLevel);
internal sealed record ListInfo(bool Ordered, int Level);

internal enum MarkdownBlockKind
{
    None,
    Other,
    List
}

internal sealed class ParagraphRenderResult
{
    public string Markdown { get; private set; } = string.Empty;
    public string NonFormulaText { get; private set; } = string.Empty;
    public int FormulaCount { get; private set; }
    public string? SingleFormulaLatex { get; private set; }

    public void Append(InlineRenderResult inline)
    {
        Markdown += inline.Markdown;
        NonFormulaText += inline.NonFormulaText;
        FormulaCount += inline.FormulaCount;
        if (inline.FormulaCount == 1)
        {
            SingleFormulaLatex = inline.SingleFormulaLatex;
        }
        else if (inline.FormulaCount > 1)
        {
            SingleFormulaLatex = null;
        }
    }
}

internal sealed class InlineRenderResult
{
    public string Markdown { get; set; } = string.Empty;
    public string NonFormulaText { get; private set; } = string.Empty;
    public int FormulaCount { get; private set; }
    public string? SingleFormulaLatex { get; private set; }
    public bool HasNonTextContent { get; set; }

    public static InlineRenderResult Text(string text)
    {
        return new InlineRenderResult { Markdown = text, NonFormulaText = text };
    }

    public static InlineRenderResult NonText(string markdown)
    {
        return new InlineRenderResult { Markdown = markdown, NonFormulaText = markdown, HasNonTextContent = true };
    }

    public static InlineRenderResult Formula(string latex)
    {
        return new InlineRenderResult
        {
            Markdown = "$" + latex + "$",
            FormulaCount = 1,
            SingleFormulaLatex = latex,
            HasNonTextContent = true
        };
    }

    public void Append(InlineRenderResult other)
    {
        Markdown += other.Markdown;
        NonFormulaText += other.NonFormulaText;
        FormulaCount += other.FormulaCount;
        HasNonTextContent |= other.HasNonTextContent;
        if (FormulaCount == 1)
        {
            SingleFormulaLatex ??= other.SingleFormulaLatex;
        }
        else if (FormulaCount > 1)
        {
            SingleFormulaLatex = null;
        }
    }
}

internal static class OmmlToLatexConverter
{
    public static string Convert(OpenXmlElement element)
    {
        return ConvertChildren(element).Trim();
    }

    private static string ConvertElement(OpenXmlElement element)
    {
        if (element is M.Text text)
        {
            return text.Text;
        }

        return element.LocalName switch
        {
            "oMath" or "oMathPara" or "e" or "num" or "den" or "deg" or "sup" or "sub" => ConvertChildren(element),
            "f" => ConvertFraction(element),
            "rad" => ConvertRadical(element),
            "sSup" => ConvertScript(element, "^"),
            "sSub" => ConvertScript(element, "_"),
            "sSubSup" => ConvertSubSup(element),
            "nary" => ConvertNary(element),
            "d" => "\\left(" + ConvertChildren(element) + "\\right)",
            _ => ConvertChildren(element)
        };
    }

    private static string ConvertChildren(OpenXmlElement element)
    {
        return string.Concat(element.ChildElements.Select(ConvertElement));
    }

    private static string ConvertFraction(OpenXmlElement element)
    {
        var numerator = ConvertFirstChild(element, "num");
        var denominator = ConvertFirstChild(element, "den");
        return "\\frac{" + numerator + "}{" + denominator + "}";
    }

    private static string ConvertRadical(OpenXmlElement element)
    {
        var degree = ConvertFirstChild(element, "deg");
        var expression = ConvertFirstChild(element, "e");
        return string.IsNullOrWhiteSpace(degree) ? "\\sqrt{" + expression + "}" : "\\sqrt[" + degree + "]{" + expression + "}";
    }

    private static string ConvertScript(OpenXmlElement element, string marker)
    {
        var expression = ConvertFirstChild(element, "e");
        var script = ConvertFirstChild(element, marker == "^" ? "sup" : "sub");
        return expression + marker + "{" + script + "}";
    }

    private static string ConvertSubSup(OpenXmlElement element)
    {
        var expression = ConvertFirstChild(element, "e");
        var sub = ConvertFirstChild(element, "sub");
        var sup = ConvertFirstChild(element, "sup");
        return expression + "_{" + sub + "}^{" + sup + "}";
    }

    private static string ConvertNary(OpenXmlElement element)
    {
        var chr = element.Descendants().FirstOrDefault(child => child.LocalName == "chr")?.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == "val").Value;
        var op = chr switch
        {
            "∑" => "\\sum",
            "∫" => "\\int",
            "∏" => "\\prod",
            _ => chr ?? string.Empty
        };
        var sub = ConvertFirstChild(element, "sub");
        var sup = ConvertFirstChild(element, "sup");
        var expression = ConvertFirstChild(element, "e");
        if (!string.IsNullOrWhiteSpace(sub))
        {
            op += "_{" + sub + "}";
        }

        if (!string.IsNullOrWhiteSpace(sup))
        {
            op += "^{" + sup + "}";
        }

        return op + " " + expression;
    }

    private static string ConvertFirstChild(OpenXmlElement element, string localName)
    {
        var child = element.ChildElements.FirstOrDefault(child => child.LocalName == localName);
        return child == null ? string.Empty : ConvertChildren(child);
    }
}

internal static class DocxWriter
{
    private const long EmusPerInch = 914400;

    public static void Write(AstDocument ast, StyleMap styleMap, string inputPath, string outputPath, string? templatePath, bool keepHeadingNumbers)
    {
        if (!string.IsNullOrWhiteSpace(templatePath))
        {
            File.Copy(templatePath, outputPath, true);
            using var document = WordprocessingDocument.Open(outputPath, true);
            EnsurePackage(document);
            ClearBody(document);
            WriteBlocks(document, ast.Blocks, styleMap, inputPath, keepHeadingNumbers);
            document.MainDocumentPart!.Document.Save();
            return;
        }

        using var newDocument = WordprocessingDocument.Create(outputPath, WordprocessingDocumentType.Document);
        var mainPart = newDocument.AddMainDocumentPart();
        mainPart.Document = new Wp.Document(new Body());
        AddDefaultStyles(mainPart, styleMap);
        WriteBlocks(newDocument, ast.Blocks, styleMap, inputPath, keepHeadingNumbers);
        mainPart.Document.Save();
    }

    private static void EnsurePackage(WordprocessingDocument document)
    {
        if (document.MainDocumentPart == null)
        {
            var mainPart = document.AddMainDocumentPart();
            mainPart.Document = new Wp.Document(new Body());
        }
        else if (document.MainDocumentPart.Document == null)
        {
            document.MainDocumentPart.Document = new Wp.Document(new Body());
        }
        else if (document.MainDocumentPart.Document.Body == null)
        {
            document.MainDocumentPart.Document.AppendChild(new Body());
        }
    }

    private static void ClearBody(WordprocessingDocument document)
    {
        var body = document.MainDocumentPart!.Document.Body!;
        var sectionProperties = body.Elements<SectionProperties>().LastOrDefault()?.CloneNode(true);
        body.RemoveAllChildren();
        if (sectionProperties != null)
        {
            body.Append(sectionProperties);
        }
    }

    private static void WriteBlocks(WordprocessingDocument document, IReadOnlyList<AstBlock> blocks, StyleMap styleMap, string inputPath, bool keepHeadingNumbers)
    {
        var resolver = new StyleResolver(document.MainDocumentPart!, styleMap);
        var context = new WriteContext(document, resolver, Path.GetDirectoryName(inputPath) ?? Directory.GetCurrentDirectory());
        var body = document.MainDocumentPart!.Document.Body!;
        var sectionProperties = body.Elements<SectionProperties>().LastOrDefault();
        if (sectionProperties != null)
        {
            sectionProperties.Remove();
        }

        foreach (var block in blocks)
        {
            AppendBlock(body, block, context, keepHeadingNumbers);
        }

        if (sectionProperties != null)
        {
            body.Append(sectionProperties);
        }
    }

    private static void AppendBlock(OpenXmlCompositeElement parent, AstBlock block, WriteContext context, bool keepHeadingNumbers)
    {
        switch (block.Type)
        {
            case "heading":
                if (!keepHeadingNumbers && block.Inlines != null)
                {
                    RemoveLeadingHeadingNumber(block.Inlines);
                }

                var level = Math.Clamp(block.Level ?? 1, 1, 6);
                parent.Append(CreateParagraph(block.Inlines, context, context.Styles.Block("heading." + level, "Heading" + level)));
                break;
            case "paragraph":
                parent.Append(CreateParagraph(block.Inlines, context, context.Styles.Block("paragraph", "Normal")));
                break;
            case "blockquote":
                foreach (var child in block.Blocks ?? [])
                {
                    parent.Append(CreateParagraph(child.Inlines, context, context.Styles.Block("blockquote", "Normal")));
                }
                break;
            case "code_block":
                parent.Append(CreateTextParagraph(block.Text ?? string.Empty, context.Styles.Block("code_block", "Normal"), preserveWhitespace: true));
                break;
            case "thematic_break":
                parent.Append(CreateTextParagraph("----------------------------------------", context.Styles.Block("paragraph", "Normal")));
                break;
            case "list":
                AppendList(parent, block, context);
                break;
            case "table":
                parent.Append(CreateTable(block, context));
                break;
            case "math_block":
                parent.Append(CreateMathParagraph(block.Text ?? string.Empty));
                break;
        }
    }

    private static void AppendList(OpenXmlCompositeElement parent, AstBlock block, WriteContext context)
    {
        var styleId = block.Ordered == true
            ? context.Styles.Block("ordered_list", "ListParagraph")
            : context.Styles.Block("unordered_list", "ListParagraph");
        foreach (var item in block.Items ?? [])
        {
            var textParts = item.Blocks.Select(PlainTextFromBlock).Where(text => !string.IsNullOrWhiteSpace(text));
            parent.Append(CreateTextParagraph(string.Join(" ", textParts), styleId));
        }
    }

    private static Wp.Table CreateTable(AstBlock block, WriteContext context)
    {
        var table = new Wp.Table();
        var tableProperties = new TableProperties(
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = 4 },
                new LeftBorder { Val = BorderValues.Single, Size = 4 },
                new BottomBorder { Val = BorderValues.Single, Size = 4 },
                new RightBorder { Val = BorderValues.Single, Size = 4 },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }));
        var tableStyle = context.Styles.Table("table", string.Empty);
        if (!string.IsNullOrWhiteSpace(tableStyle))
        {
            tableProperties.PrependChild(new TableStyle { Val = tableStyle });
        }

        table.Append(tableProperties);
        var columnCount = block.Rows?.Count > 0 ? block.Rows.Max(row => row.Cells.Count) : 0;
        if (columnCount > 0)
        {
            var grid = new TableGrid();
            for (var i = 0; i < columnCount; i++)
            {
                grid.Append(new GridColumn());
            }

            table.Append(grid);
        }

        var cellStyle = context.Styles.Block("table", "Normal");
        foreach (var row in block.Rows ?? [])
        {
            var tableRow = new Wp.TableRow();
            for (var columnIndex = 0; columnIndex < row.Cells.Count; columnIndex++)
            {
                var cell = row.Cells[columnIndex];
                var tableCell = new Wp.TableCell();
                var paragraph = CreateTextParagraph(PlainTextFromCell(cell), cellStyle);
                ApplyTableCellAlignment(paragraph, block.Alignments, columnIndex);
                if (row.IsHeader)
                {
                    foreach (var run in paragraph.Descendants<Run>())
                    {
                        run.RunProperties ??= new RunProperties();
                        run.RunProperties.Bold = new Bold();
                    }
                }

                tableCell.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Auto }));
                tableCell.Append(paragraph);
                tableRow.Append(tableCell);
            }

            table.Append(tableRow);
        }

        return table;
    }

    private static void ApplyTableCellAlignment(Paragraph paragraph, IReadOnlyList<string>? alignments, int columnIndex)
    {
        if (alignments == null || columnIndex >= alignments.Count)
        {
            return;
        }

        var alignment = alignments[columnIndex] switch
        {
            "center" => JustificationValues.Center,
            "right" => JustificationValues.Right,
            _ => JustificationValues.Left
        };
        paragraph.ParagraphProperties ??= new ParagraphProperties();
        paragraph.ParagraphProperties.Justification = new Justification { Val = alignment };
    }

    private static Paragraph CreateParagraph(IReadOnlyList<AstInline>? inlines, WriteContext context, string styleId)
    {
        var paragraph = new Paragraph();
        ApplyParagraphStyle(paragraph, styleId);
        foreach (var inline in inlines ?? [])
        {
            AppendInline(paragraph, inline, context, bold: false, italic: false);
        }

        return paragraph;
    }

    private static Paragraph CreateTextParagraph(string text, string styleId, bool preserveWhitespace = false)
    {
        var paragraph = new Paragraph();
        ApplyParagraphStyle(paragraph, styleId);
        paragraph.Append(CreateRun(text, preserveWhitespace: preserveWhitespace));
        return paragraph;
    }

    private static Wp.Paragraph CreateMathParagraph(string text)
    {
        return new Wp.Paragraph(
            new Wp.ParagraphProperties(
                new Wp.Justification { Val = Wp.JustificationValues.Center }),
            new M.OfficeMath(
                new M.Run(
                    new M.Text(text))));
    }

    private static void AppendInline(OpenXmlCompositeElement parent, AstInline inline, WriteContext context, bool bold, bool italic)
    {
        switch (inline.Type)
        {
            case "text":
                parent.Append(CreateRun(inline.Text ?? string.Empty, bold, italic));
                break;
            case "line_break":
                parent.Append(new Run(new Break()));
                break;
            case "strong":
                foreach (var child in inline.Inlines ?? [])
                {
                    AppendInline(parent, child, context, bold: true, italic);
                }
                break;
            case "emphasis":
                foreach (var child in inline.Inlines ?? [])
                {
                    AppendInline(parent, child, context, bold, italic: true);
                }
                break;
            case "inline_code":
                parent.Append(CreateRun(inline.Text ?? string.Empty, bold, italic, context.Styles.Inline("inline_code", string.Empty), fallbackFont: "Consolas"));
                break;
            case "math_inline":
                parent.Append(CreateRun(inline.Text ?? string.Empty, bold, italic: true, context.Styles.Inline("math_inline", string.Empty)));
                break;
            case "link":
                AppendLink(parent, inline, context, bold, italic);
                break;
            case "image":
                AppendImage(parent, inline, context);
                break;
            default:
                if (inline.Inlines != null)
                {
                    foreach (var child in inline.Inlines)
                    {
                        AppendInline(parent, child, context, bold, italic);
                    }
                }
                else if (!string.IsNullOrEmpty(inline.Text))
                {
                    parent.Append(CreateRun(inline.Text, bold, italic));
                }
                break;
        }
    }

    private static void AppendLink(OpenXmlCompositeElement parent, AstInline inline, WriteContext context, bool bold, bool italic)
    {
        if (string.IsNullOrWhiteSpace(inline.Url))
        {
            foreach (var child in inline.Inlines ?? [])
            {
                AppendInline(parent, child, context, bold, italic);
            }
            return;
        }

        var relationship = context.Document.MainDocumentPart!.AddHyperlinkRelationship(new Uri(inline.Url, UriKind.RelativeOrAbsolute), true);
        var hyperlink = new Hyperlink { Id = relationship.Id, History = OnOffValue.FromBoolean(true) };
        if (inline.Inlines == null || inline.Inlines.Count == 0)
        {
            hyperlink.Append(CreateRun(inline.Text ?? inline.Url, bold, italic));
        }
        else
        {
            foreach (var child in inline.Inlines)
            {
                AppendInline(hyperlink, child, context, bold, italic);
            }
        }

        parent.Append(hyperlink);
    }

    private static void AppendImage(OpenXmlCompositeElement parent, AstInline inline, WriteContext context)
    {
        var imagePath = ResolveImagePath(inline.Url, context.ImageBasePath);
        if (imagePath == null)
        {
            parent.Append(CreateRun(string.IsNullOrWhiteSpace(inline.Text) ? "[图片]" : "[图片: " + inline.Text + "]"));
            return;
        }

        var mainPart = context.Document.MainDocumentPart!;
        var imagePartType = GetImagePartType(imagePath);
        var imagePart = mainPart.AddImagePart(imagePartType);
        using (var stream = File.OpenRead(imagePath))
        {
            imagePart.FeedData(stream);
        }

        var relationshipId = mainPart.GetIdOfPart(imagePart);
        var (cx, cy) = GetImageSizeEmus(imagePath, maxWidthEmus: 6 * EmusPerInch);
        parent.Append(new Run(CreateDrawing(relationshipId, inline.Text ?? Path.GetFileName(imagePath), cx, cy)));
    }

    private static Run CreateRun(string text, bool bold = false, bool italic = false, string? styleId = null, string? fallbackFont = null, bool preserveWhitespace = false)
    {
        var run = new Run();
        var properties = new RunProperties();
        if (bold)
        {
            properties.Bold = new Bold();
        }

        if (italic)
        {
            properties.Italic = new Italic();
        }

        if (!string.IsNullOrWhiteSpace(styleId))
        {
            properties.RunStyle = new RunStyle { Val = styleId };
        }
        else if (!string.IsNullOrWhiteSpace(fallbackFont))
        {
            properties.RunFonts = new RunFonts { Ascii = fallbackFont, HighAnsi = fallbackFont };
        }

        if (properties.HasChildren || properties.Bold != null || properties.Italic != null || properties.RunStyle != null || properties.RunFonts != null)
        {
            run.Append(properties);
        }

        var textElement = new Text(text);
        if (preserveWhitespace || text.StartsWith(' ') || text.EndsWith(' ') || text.Contains('\n'))
        {
            textElement.Space = SpaceProcessingModeValues.Preserve;
        }

        run.Append(textElement);
        return run;
    }

    private static void ApplyParagraphStyle(Paragraph paragraph, string styleId)
    {
        if (string.IsNullOrWhiteSpace(styleId))
        {
            return;
        }

        paragraph.ParagraphProperties ??= new ParagraphProperties();
        paragraph.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId { Val = styleId };
    }

    private static Drawing CreateDrawing(string relationshipId, string name, long cx, long cy)
    {
        return new Drawing(
            new DW.Inline(
                new DW.Extent { Cx = cx, Cy = cy },
                new DW.EffectExtent { LeftEdge = 0, TopEdge = 0, RightEdge = 0, BottomEdge = 0 },
                new DW.DocProperties { Id = 1U, Name = name, Description = name },
                new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = 0U, Name = name },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(
                                new A.Blip { Embed = relationshipId },
                                new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0, Y = 0 },
                                    new A.Extents { Cx = cx, Cy = cy }),
                                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })))
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }))
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            });
    }

    private static string? ResolveImagePath(string? imageUrl, string basePath)
    {
        if (string.IsNullOrWhiteSpace(imageUrl))
        {
            return null;
        }

        if (Uri.TryCreate(imageUrl, UriKind.Absolute, out var uri))
        {
            if (uri.IsFile && File.Exists(uri.LocalPath))
            {
                return uri.LocalPath;
            }

            return null;
        }

        var candidate = Path.IsPathRooted(imageUrl) ? imageUrl : Path.Combine(basePath, imageUrl);
        return File.Exists(candidate) ? Path.GetFullPath(candidate) : null;
    }

    private static PartTypeInfo GetImagePartType(string imagePath)
    {
        return Path.GetExtension(imagePath).ToLowerInvariant() switch
        {
            ".png" => ImagePartType.Png,
            ".gif" => ImagePartType.Gif,
            ".bmp" => ImagePartType.Bmp,
            ".tif" or ".tiff" => ImagePartType.Tiff,
            _ => ImagePartType.Jpeg
        };
    }

    private static (long Cx, long Cy) GetImageSizeEmus(string imagePath, long maxWidthEmus)
    {
        using var image = System.Drawing.Image.FromFile(imagePath);
        var widthEmus = (long)(image.Width / image.HorizontalResolution * EmusPerInch);
        var heightEmus = (long)(image.Height / image.VerticalResolution * EmusPerInch);
        if (widthEmus > maxWidthEmus)
        {
            var ratio = (double)maxWidthEmus / widthEmus;
            widthEmus = maxWidthEmus;
            heightEmus = (long)(heightEmus * ratio);
        }

        return (Math.Max(widthEmus, 1), Math.Max(heightEmus, 1));
    }

    private static string PlainTextFromBlock(AstBlock block)
    {
        if (block.Inlines != null)
        {
            return PlainTextFromInlines(block.Inlines);
        }

        if (!string.IsNullOrEmpty(block.Text))
        {
            return block.Text;
        }

        if (block.Blocks != null)
        {
            return string.Join(" ", block.Blocks.Select(PlainTextFromBlock));
        }

        return string.Empty;
    }

    private static string PlainTextFromCell(AstTableCell cell)
    {
        return string.Join(Environment.NewLine, cell.Blocks.Select(PlainTextFromBlock).Where(text => !string.IsNullOrWhiteSpace(text)));
    }

    private static string PlainTextFromInlines(IEnumerable<AstInline> inlines)
    {
        return string.Concat(inlines.Select(inline =>
        {
            if (inline.Type == "line_break")
            {
                return Environment.NewLine;
            }

            if (inline.Inlines != null)
            {
                return PlainTextFromInlines(inline.Inlines);
            }

            return inline.Text ?? string.Empty;
        }));
    }

    private static bool RemoveLeadingHeadingNumber(List<AstInline> inlines)
    {
        foreach (var inline in inlines)
        {
            if (inline.Type == "text" && inline.Text != null)
            {
                var cleaned = Regex.Replace(inline.Text, @"^\s*(?:(?:\d+(?:\.\d+)+\.?)|(?:\d+[\.、．)]))\s*", string.Empty);
                if (cleaned != inline.Text)
                {
                    inline.Text = cleaned;
                    return true;
                }

                if (!string.IsNullOrWhiteSpace(inline.Text))
                {
                    return false;
                }
            }
            else if (inline.Inlines != null && RemoveLeadingHeadingNumber(inline.Inlines))
            {
                return true;
            }
        }

        return false;
    }

    private static int TryParseInt(string? value, int fallback)
    {
        return int.TryParse(value, out var result) ? result : fallback;
    }

    private static void AddDefaultStyles(MainDocumentPart mainPart, StyleMap styleMap)
    {
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        var styles = new Styles();
        AddParagraphStyle(styles, "Normal", "正文", defaultStyle: true);
        AddParagraphStyle(styles, "BodyText", "正文2");
        AddParagraphStyle(styles, "Quote", "引用");
        AddParagraphStyle(styles, "CodeBlock", "代码");
        AddParagraphStyle(styles, "TableText", "表格文本");
        AddParagraphStyle(styles, "ListParagraph", "列表段落");
        AddParagraphStyle(styles, "OrderedListParagraph", "有序列表段落");
        AddParagraphStyle(styles, "MathBlock", "公式");
        AddCharacterStyle(styles, "InlineCode", "行内代码", "Consolas");
        for (var i = 1; i <= 6; i++)
        {
            AddParagraphStyle(styles, "Heading" + i, "标题 " + i, basedOn: "Normal", outlineLevel: i - 1);
        }

        foreach (var configuredStyle in styleMap.AllStyleNames())
        {
            if (!styles.Elements<Style>().Any(style => string.Equals(style.StyleId, configuredStyle, StringComparison.OrdinalIgnoreCase)
                || string.Equals(style.StyleName?.Val, configuredStyle, StringComparison.OrdinalIgnoreCase)))
            {
                AddParagraphStyle(styles, configuredStyle, configuredStyle);
            }
        }

        stylesPart.Styles = styles;
    }

    private static void AddParagraphStyle(Styles styles, string styleId, string name, bool defaultStyle = false, string? basedOn = null, int? outlineLevel = null)
    {
        var style = new Style { Type = StyleValues.Paragraph, StyleId = styleId, Default = defaultStyle };
        style.Append(new StyleName { Val = name });
        if (!string.IsNullOrWhiteSpace(basedOn))
        {
            style.Append(new BasedOn { Val = basedOn });
        }

        if (outlineLevel != null)
        {
            style.Append(new StyleParagraphProperties(new OutlineLevel { Val = outlineLevel }));
        }

        styles.Append(style);
    }

    private static void AddCharacterStyle(Styles styles, string styleId, string name, string? font = null)
    {
        var style = new Style { Type = StyleValues.Character, StyleId = styleId };
        style.Append(new StyleName { Val = name });
        if (!string.IsNullOrWhiteSpace(font))
        {
            style.Append(new StyleRunProperties(new RunFonts { Ascii = font, HighAnsi = font }));
        }

        styles.Append(style);
    }
}

internal sealed class WriteContext(WordprocessingDocument document, StyleResolver styles, string imageBasePath)
{
    public WordprocessingDocument Document { get; } = document;
    public StyleResolver Styles { get; } = styles;
    public string ImageBasePath { get; } = imageBasePath;
}

internal sealed class StyleResolver
{
    private readonly StyleMap _styleMap;
    private readonly Dictionary<string, string> _paragraphStyles = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _characterStyles = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _tableStyles = new(StringComparer.OrdinalIgnoreCase);

    public StyleResolver(MainDocumentPart mainPart, StyleMap styleMap)
    {
        _styleMap = styleMap;
        var styles = mainPart.StyleDefinitionsPart?.Styles;
        if (styles == null)
        {
            return;
        }

        foreach (var style in styles.Elements<Style>())
        {
            var styleId = style.StyleId?.Value;
            if (string.IsNullOrWhiteSpace(styleId))
            {
                continue;
            }

            var name = style.StyleName?.Val?.Value;
            Dictionary<string, string>? target = null;
            if (style.Type?.Value == StyleValues.Paragraph)
            {
                target = _paragraphStyles;
            }
            else if (style.Type?.Value == StyleValues.Character)
            {
                target = _characterStyles;
            }
            else if (style.Type?.Value == StyleValues.Table)
            {
                target = _tableStyles;
            }
            if (target == null)
            {
                continue;
            }

            target[styleId] = styleId;
            if (!string.IsNullOrWhiteSpace(name))
            {
                target[name] = styleId;
            }
        }
    }

    public string Block(string key, string fallback)
    {
        return Resolve(_styleMap.BlockStyles.GetValueOrDefault(key), fallback, _paragraphStyles);
    }

    public string Inline(string key, string fallback)
    {
        return Resolve(_styleMap.InlineStyles.GetValueOrDefault(key), fallback, _characterStyles);
    }

    public string Table(string key, string fallback)
    {
        return Resolve(_styleMap.TableStyles.GetValueOrDefault(key), fallback, _tableStyles);
    }

    private static string Resolve(string? configured, string fallback, IReadOnlyDictionary<string, string> knownStyles)
    {
        var candidates = new List<string>();
        AddCandidate(candidates, configured);
        foreach (var alias in KnownStyleAliases.ResolveAll(configured))
        {
            AddCandidate(candidates, alias);
        }

        AddCandidate(candidates, fallback);
        foreach (var alias in KnownStyleAliases.ResolveAll(fallback))
        {
            AddCandidate(candidates, alias);
        }

        foreach (var candidate in candidates)
        {
            if (knownStyles.TryGetValue(candidate, out var styleId))
            {
                return styleId;
            }
        }

        return candidates.FirstOrDefault() ?? string.Empty;
    }

    private static void AddCandidate(List<string> candidates, string? candidate)
    {
        if (string.IsNullOrWhiteSpace(candidate))
        {
            return;
        }

        if (!candidates.Contains(candidate, StringComparer.OrdinalIgnoreCase))
        {
            candidates.Add(candidate);
        }
    }
}

internal static class KnownStyleAliases
{
    private static readonly Dictionary<string, string[]> Aliases = new(StringComparer.OrdinalIgnoreCase)
    {
        ["正文"] = ["Normal"],
        ["正文2"] = ["正文2", "Body Text", "BodyText"],
        ["标题 1"] = ["标题 1", "heading 1", "Heading 1", "Heading1"],
        ["标题 2"] = ["标题 2", "heading 2", "Heading 2", "Heading2"],
        ["标题 3"] = ["标题 3", "heading 3", "Heading 3", "Heading3"],
        ["标题 4"] = ["标题 4", "heading 4", "Heading 4", "Heading4"],
        ["标题 5"] = ["标题 5", "heading 5", "Heading 5", "Heading5"],
        ["标题 6"] = ["标题 6", "heading 6", "Heading 6", "Heading6"],
        ["引用"] = ["引用", "Quote"],
        ["代码"] = ["代码", "CodeBlock"],
        ["表格文本"] = ["表格文本", "TableText"],
        ["列表段落"] = ["列表段落", "List Paragraph", "ListParagraph"],
        ["有序列表段落"] = ["有序列表段落", "OrderedListParagraph"],
        ["公式"] = ["公式", "MathBlock"],
        ["行内代码"] = ["行内代码", "InlineCode"]
    };

    public static IEnumerable<string> ResolveAll(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return [];
        }

        return Aliases.GetValueOrDefault(value, []);
    }
}

internal sealed class StyleMap
{
    public int Version { get; init; } = 1;
    public Dictionary<string, string> BlockStyles { get; init; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, string> TableStyles { get; init; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, string> InlineStyles { get; init; } = new(StringComparer.OrdinalIgnoreCase);

    public static StyleMap Load(string path)
    {
        var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
        return JsonSerializer.Deserialize<StyleMap>(File.ReadAllText(path), options) ?? new StyleMap();
    }

    public IEnumerable<string> AllStyleNames()
    {
        return BlockStyles.Values.Concat(TableStyles.Values).Concat(InlineStyles.Values).Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase);
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

internal sealed record AstDocument(int Version, string SourcePath, List<AstBlock> Blocks);

internal sealed class AstBlock(string type)
{
    public string Type { get; } = type;
    public List<AstInline>? Inlines { get; set; }
    public int? Level { get; set; }
    public string? Text { get; set; }
    public string? Language { get; set; }
    public bool? Ordered { get; set; }
    public string? Start { get; set; }
    public List<AstListItem>? Items { get; set; }
    public List<AstBlock>? Blocks { get; set; }
    public List<AstTableRow>? Rows { get; set; }
    public List<string>? Alignments { get; set; }
}

internal sealed class AstInline(string type)
{
    public string Type { get; } = type;
    public string? Text { get; set; }
    public string? Url { get; set; }
    public string? Title { get; set; }
    public List<AstInline>? Inlines { get; set; }
}

internal sealed record AstListItem(List<AstBlock> Blocks);
internal sealed record AstTableRow(bool IsHeader, List<AstTableCell> Cells);
internal sealed record AstTableCell(List<AstBlock> Blocks);
