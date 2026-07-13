using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Wp = DocumentFormat.OpenXml.Wordprocessing;

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

