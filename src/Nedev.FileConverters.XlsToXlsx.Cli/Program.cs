using System;
using System.IO;
using Nedev.FileConverters.XlsToXlsx;
using Nedev.FileConverters;

void ShowHelp()
{
    Console.WriteLine("Usage: xls2xlsx [options] <input> [output]");
    Console.WriteLine();
    Console.WriteLine("Options:");
    Console.WriteLine("  -i, --input <path>      Input file or directory");
    Console.WriteLine("  -o, --output <file>     Output filename for single-file conversion");
    Console.WriteLine("  --dump-colors <file>    Inspect palette/fonts of an XLS file");
    Console.WriteLine("  --version               Show tool version");
    Console.WriteLine("  --help, -h              Show this message");
}

if (args.Length > 0)
{
    if (args[0] == "--help" || args[0] == "-h")
    {
        ShowHelp();
        return;
    }
    if (args[0] == "--version" || args[0] == "-v")
    {
        Console.WriteLine("Nedev.FileConverters.XlsToXlsx 0.1.0");
        return;
    }
}

string? inputArg = null;
string? outputArg = null;
bool dump = false;

for (int i = 0; i < args.Length; i++)
{
    switch (args[i])
    {
        case "--help":
        case "-h":
            ShowHelp();
            return;
        case "--version":
        case "-v":
            Console.WriteLine("Nedev.FileConverters.XlsToXlsx 0.1.0");
            return;
        case "--dump-colors":
            dump = true;
            if (i + 1 < args.Length)
            {
                inputArg = args[++i];
            }
            break;
        case "--input":
        case "-i":
            if (i + 1 < args.Length) inputArg = args[++i];
            break;
        case "--output":
        case "-o":
            if (i + 1 < args.Length) outputArg = args[++i];
            break;
        default:
            if (!args[i].StartsWith("-") && inputArg == null)
            {
                inputArg = args[i];
            }
            else if (!args[i].StartsWith("-") && outputArg == null)
            {
                outputArg = args[i];
            }
            break;
    }
}

if (dump)
{
    if (string.IsNullOrEmpty(inputArg))
    {
        Console.WriteLine("Please supply a file to inspect with --dump-colors");
        return;
    }
    string dumpPath = Path.GetFullPath(inputArg);
    if (!File.Exists(dumpPath))
    {
        Console.WriteLine($"File not found: {dumpPath}");
        return;
    }
    DumpColors(dumpPath);
    return;
}

if (inputArg == null)
{
    inputArg = Path.Combine("..", "tests");
}

string inputPath = Path.GetFullPath(inputArg);

if (Directory.Exists(inputPath))
{
    BatchConvertDirectory(inputPath);
    return;
}

if (!File.Exists(inputPath))
{
    Console.WriteLine($"File not found: {inputPath}");
    Environment.Exit(1);
}

string outputPath = outputArg != null
    ? Path.GetFullPath(outputArg)
    : Path.ChangeExtension(inputPath, ".xlsx");

ConvertSingleFile(inputPath, outputPath);

// helper methods -------------------------------------------------------------

static void DumpColors(string file)
{
    using var fs = File.OpenRead(file);
    var parser = new Nedev.FileConverters.XlsToXlsx.Formats.Xls.XlsParser(fs);
    var wb = parser.Parse();

    Console.WriteLine($"XfList count = {wb.XfList?.Count}");
    if (wb.XfList != null)
    {
        for (int xi = 0; xi < wb.XfList.Count; xi++)
        {
            var xf = wb.XfList[xi];
            string fcol = "(none)";
            if (xf.FontIndex >= 0 && xf.FontIndex < wb.Fonts.Count)
                fcol = wb.Fonts[xf.FontIndex].Color ?? "(none)";
            Console.WriteLine($"  Xf #{xi} fontIdx={xf.FontIndex} color={fcol}");
        }
    }
    for (int si = 0; si < wb.Worksheets.Count; si++)
    {
        var sheet = wb.Worksheets[si];
        Console.WriteLine($"Sheet {si} Xfs count = {sheet.Xfs.Count}");
        if (sheet.MergeCells != null && sheet.MergeCells.Count > 0)
        {
            Console.WriteLine("  original merge cells:");
            foreach (var m in sheet.MergeCells)
                Console.WriteLine($"    {m.StartRow},{m.StartColumn}-{m.EndRow},{m.EndColumn}");
        }
    }
}

static void BatchConvertDirectory(string path)
{
    var xlsFiles = Directory.GetFiles(path, "*.xls");
    if (xlsFiles.Length == 0)
    {
        Console.WriteLine($"No .xls files found in directory: {path}");
        Environment.Exit(1);
    }

    var outputFiles = xlsFiles.Select(p => Path.ChangeExtension(p, ".xlsx")).ToArray();
    try
    {
        XlsToXlsxConverter.BatchConvert(xlsFiles, outputFiles, (pct, msg) => Console.WriteLine($"{pct}% - {msg}"));
        foreach (var outFile in outputFiles)
            Console.WriteLine($"Done: {outFile}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error during batch conversion: {ex.Message}");
        Environment.Exit(1);
    }
}

static void ConvertSingleFile(string inputPath, string outputPath)
{
    try
    {
        using var inStream = File.OpenRead(inputPath);
        using var convStream = Converter.Convert(inStream, "xls", "xlsx");
        using var outStream = File.Create(outputPath);
        convStream.CopyTo(outStream);
        Console.WriteLine($"Done: {outputPath}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error: {ex.Message}");
        Environment.Exit(1);
    }
}
