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

