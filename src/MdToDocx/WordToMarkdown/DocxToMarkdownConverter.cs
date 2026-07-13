using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

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

