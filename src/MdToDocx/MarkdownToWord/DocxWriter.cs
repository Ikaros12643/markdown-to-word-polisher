using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

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
                parent.Append(CreateCodeBlockParagraph(block.Text ?? string.Empty, context.Styles.Block("code_block", "Normal")));
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

    private static Paragraph CreateCodeBlockParagraph(string text, string styleId)
    {
        var paragraph = new Paragraph();
        ApplyParagraphStyle(paragraph, styleId);

        // Word 不会把 w:t 内的 \n 稳定渲染为换行，代码块需要显式写入 w:br。
        var normalizedText = text.Replace("\r\n", "\n").Replace('\r', '\n');
        if (normalizedText.EndsWith('\n'))
        {
            normalizedText = normalizedText[..^1];
        }

        var lines = normalizedText.Split('\n');
        for (var i = 0; i < lines.Length; i++)
        {
            paragraph.Append(CreateRun(lines[i], preserveWhitespace: true));
            if (i < lines.Length - 1)
            {
                paragraph.Append(new Run(new Break()));
            }
        }

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

