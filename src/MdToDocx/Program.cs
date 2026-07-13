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

        DocxToMarkdownConverter.Convert(new DocxToMarkdownOptions(inputPath, outputPath, assetsDir, MarkdownPath.Normalize(assetsRelativePath)));
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

