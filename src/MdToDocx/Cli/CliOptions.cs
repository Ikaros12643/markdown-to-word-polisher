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

