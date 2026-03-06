using System;
using System.IO;
using Nedev.FileConverters.XlsToXlsx;

// Convert .xls to .xlsx. Usage: dotnet run [input.xls] [output.xlsx]
// If one argument: output is input with .xlsx extension. If no arguments: convert tests\test.xls.

// Determine input and output from command line arguments.
// Usage examples:
//   dotnet run -- [input.xls] [output.xlsx]
//   dotnet run -- [directory]          // converts all .xls files in the directory
//   dotnet run                         // defaults to ../tests/test.xls or treats ../tests as dir if present

string inputArg;
string outputArg = null;

// debug switch: dotnet run -- --dump-colors file.xls
if (args.Length > 0 && args[0] == "--dump-colors")
{
    if (args.Length < 2)
    {
        Console.WriteLine("Usage: --dump-colors file.xls");
        return;
    }
    string dumpPath = Path.GetFullPath(args[1]);
    if (!File.Exists(dumpPath))
    {
        Console.WriteLine($"File not found: {dumpPath}");
        return;
    }
    // use internal parser to inspect palette and fonts
    using (var fs = File.OpenRead(dumpPath))
    {
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
        foreach (var f in wb.Fonts)
            Console.WriteLine($"  name={f.Name} size={f.Size ?? 0} height={f.Height} colorIndex={f.ColorIndex} color={f.Color} bold={f.IsBold} italic={f.IsItalic}");

        Console.WriteLine("Styles (global workbook styles):");
        for (int i = 0; i < wb.Styles.Count; i++)
        {
            var s = wb.Styles[i];
            var c = s.Font?.Color ?? "(none)";
            Console.WriteLine($"  style#{i} fontColor={c} fontName={s.Font?.Name}");
        }

        Console.WriteLine("Cell styles and applied font colors:");
        foreach (var sheet in wb.Worksheets)
        {
            Console.WriteLine($"Sheet: {sheet.Name}");
            foreach (var row in sheet.Rows)
            {
                foreach (var cell in row.Cells ?? new List<Cell>())
                {
                    if (!string.IsNullOrEmpty(cell.StyleId) && int.TryParse(cell.StyleId, out int styleIndex))
                    {
                        string col = "(none)";
                        if (styleIndex >= 0 && styleIndex < wb.Styles.Count)
                        {
                            var style = wb.Styles[styleIndex];
                            col = style?.Font?.Color ?? "(none)";
                        }
                        Console.WriteLine($"  R{cell.RowIndex}C{cell.ColumnIndex} val={cell.Value} styleColor={col} styleId={cell.StyleId}");
                    }
                    if (cell.RichText != null && cell.RichText.Count > 0)
                    {
                        foreach (var run in cell.RichText)
                        {
                            var rcol = run.Font?.Color ?? "(none)";
                            Console.WriteLine($"    run='{run.Text}' color={rcol}");
                        }
                    }
                }
            }
        }
    }
    return;
}

inputArg = args.Length > 0 ? args[0] : Path.Combine("..", "tests");
if (args.Length > 1) outputArg = args[1];

string inputPath = Path.GetFullPath(inputArg);

if (Directory.Exists(inputPath))
{
    // batch convert every .xls file under the directory
    var xlsFiles = Directory.GetFiles(inputPath, "*.xls");
    if (xlsFiles.Length == 0)
    {
        Console.WriteLine($"No .xls files found in directory: {inputPath}");
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
    return;
}

// treat input as single file
if (!File.Exists(inputPath))
{
    Console.WriteLine($"File not found: {inputPath}");
    Environment.Exit(1);
}

string outputPath = outputArg != null
    ? Path.GetFullPath(outputArg)
    : Path.ChangeExtension(inputPath, ".xlsx");

try
{
    XlsToXlsxConverter.Convert(inputPath, outputPath, (pct, msg) => Console.WriteLine($"{pct}% - {msg}"));
    Console.WriteLine($"Done: {outputPath}");
}
catch (Exception ex)
{
    Console.WriteLine($"Error: {ex.Message}");
    Environment.Exit(1);
}
