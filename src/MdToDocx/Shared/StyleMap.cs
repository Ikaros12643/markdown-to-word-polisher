using System.Text.Json;

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

