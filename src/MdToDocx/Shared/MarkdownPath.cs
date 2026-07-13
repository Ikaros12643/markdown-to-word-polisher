internal static class MarkdownPath
{
    public static string Normalize(string path)
    {
        return path.Replace('\\', '/').TrimEnd('/');
    }
}

