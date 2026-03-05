using System;
using System.IO;
using Nedev.XlsToXlsx;

// Convert .xls to .xlsx. Usage: dotnet run [input.xls] [output.xlsx]
// If one argument: output is input with .xlsx extension. If no arguments: convert tests\test.xls.

string input = args.Length > 0 ? args[0] : Path.Combine("..", "tests", "test.xls");
string output = args.Length > 1 ? args[1] : Path.ChangeExtension(input, ".xlsx");

input = Path.GetFullPath(input);
output = Path.GetFullPath(output);

if (!File.Exists(input))
{
    Console.WriteLine($"File not found: {input}");
    Environment.Exit(1);
}

try
{
    XlsToXlsxConverter.Convert(input, output, (pct, msg) => Console.WriteLine($"{pct}% - {msg}"));
    Console.WriteLine($"Done: {output}");
}
catch (Exception ex)
{
    Console.WriteLine($"Error: {ex.Message}");
    Environment.Exit(1);
}
