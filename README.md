# Nedev.FileConverters.XlsToXlsx

**Nedev.FileConverters.XlsToXlsx** is a standalone, lightweight, and fast .NET library designed to convert legacy Microsoft Excel files (`.xls` / BIFF8 format) into the modern OpenXML format (`.xlsx`). It performs this conversion entirely in memory and **does not require Microsoft Office or Excel Interop** to be installed.

## 🚀 Features

### ✅ Supported
- **OLE Compound File Parsing**: Fully implements FAT and DIFAT chain parsing to reliably read even large and highly-fragmented `.xls` files.
- **Cell Data**: Extracts formulas, numbers, strings (including Shared String Table), booleans, and errors.
- **Formatting & Styles**: Converts fonts (size, color, bold, italic), cell fills/patterns, borders, and alignment.
- **Worksheet Layout**: Preserves row heights, column widths, fixed/frozen panes, and hidden rows/cols.
- **Merged Cells**: Accurately maps merged cell ranges.  
  *Overlapping or nested ranges are now filtered in parser order: the first merge is kept and any later range that intersects it is discarded.  This avoids Excel warnings while preventing a single huge union cell from appearing in the output.*
- **Hyperlinks**: Preserves cell hyperlinks (external URLs, local files, and email addresses).
- **Comments/Notes**: Extracts basic cell notes/comments.
- **Data Validation**: Retains basic data validation rules (dropdowns, number constraints).
- **VBA Macros**: Preserves existing VBA macros by extracting the raw `vbaProject.bin` from the legacy file and embedding it into a macro-enabled `.xlsm` compatible structure.
- **Page Setup**: Keeps print margins, page orientation, paper size, and fit-to-page scaling.
- **Pivot Tables**: Converts pivot table structure (fields, layout, data source range); output can be refreshed in Excel.
- **AutoFilter**: Preserves filter range (from _FilterDatabase name) and filter column indices; writes `<autoFilter ref="...">` with `<filterColumn>` in XLSX.
- **Worksheet / Workbook Protection**: Preserves sheet/workbook protection flags and 16‑bit password hashes; writes corresponding `sheetProtection` / `workbookProtection` so Excel still prompts for the original password.
- **Document Properties**: Reads OLE SummaryInformation/DocumentSummaryInformation and writes matching `docProps/core.xml` and `docProps/app.xml` (title, subject, author, company, timestamps, etc).
- **External Workbook Links**: Converts EXTERNSHEET/EXTERNBOOK into OOXML `externalLinks` parts and updates 3D formula refs to `[n]Sheet!A1` form.

### ⚠️ Partially Supported (WIP)
- **Formulas**: A custom formula decompiler supports over 170+ standard Excel functions. Shared formulas (`SHAREDFMLA`) and array formulas (`ARRAY`) are supported.
- **Charts**: Can detect and convert basic chart types (bar, line, pie, etc.), but advanced 3D properties and secondary axes are not yet fully mapped.
- **Images & Drawings**: Basic image extraction is supported, but complex Microsoft Office Drawing (Escher) containers are partially parsed.
- **Conditional Formatting**: Detection is supported, but styling rules (like Color Scales and Data Bars) currently use fallback styles instead of the exact embedded binary properties.

### ❌ Not Yet Supported
- *(none so far – all previously listed features have been implemented in this round, except for advanced edge-cases not covered in README.)*

## 📦 Installation

*(To be added when published to NuGet)*
```bash
dotnet add package Nedev.FileConverters.XlsToXlsx
# core infrastructure required by all converters
dotnet add package Nedev.FileConverters.Core
```

## 💻 Usage

This repository contains two deliverables:

* **`Nedev.FileConverters.XlsToXlsx`** – the core library (DLL) implementing the converter.
* **`Nedev.FileConverters.XlsToXlsx.Cli`** – a small console application that wraps the library and provides a command‑line interface. This CLI is what can be packaged as a global tool.

### Running the CLI locally

```powershell
# build and run inside repo
cd src\Nedev.FileConverters.XlsToXlsx
dotnet run -- -i input.xls -o output.xlsx
```

The CLI tool understands the following options:

* `-i|--input <path>` – input file or directory (required)
* `-o|--output <file>` – output file path (only for single-file conversion)
* `--dump-colors` – inspect an XLS palette/fonts instead of converting
* `--version` – display the tool version and exit
* `--help` – show usage information

If the input path is a directory, all `*.xls` files will be batch-converted.

**Example (file)**
```powershell
xls2xlsx -i C:\old.xls -o C:\new.xlsx
```

**Example (folder)**
```powershell
xls2xlsx -i C:\legacy-files
```

**Inspect colors**
```powershell
xls2xlsx -i C:\workbook.xls --dump-colors
```

### as a global tool

After publishing to NuGet the CLI package can be installed globally:
```powershell
dotnet tool install --global Nedev.FileConverters.XlsToXlsx.Cli
# then run via its command name:
xls2xlsx -i file.xls
```

*The conversion API is still available for library consumers; see the example below.*
* **Core integration** – since this package now implements `IFileConverter` from `Nedev.FileConverters.Core`, you can invoke conversion via the shared `Converter.Convert(...)` helper and discover converters automatically.

Converting a `.xls` file to `.xlsx` takes just a few lines of code. you may use the core helper as shown below:

```csharp
using System;
using Nedev.FileConverters;

class Program
{
    static void Main()
    {
        string inputFilePath = @"C:\path\to\your\legacy_file.xls";
        string outputFilePath = @"C:\path\to\your\converted_file.xlsx";

        try
        {
            // the Core library will locate the Xls -> Xlsx converter automatically
            using var inStream = File.OpenRead(inputFilePath);
            using var converted = Converter.Convert(inStream, "xls", "xlsx");
            using var outStream = File.Create(outputFilePath);
            converted.CopyTo(outStream);

            Console.WriteLine("Conversion completed successfully!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to convert: {ex.Message}");
        }
    }
}
```

## 🛠️ Architecture
The conversion process involves standardizing BIFF8 records into an intermediate domain model (Workbooks, Worksheets, Cells, Styles), which is then serialized into standard OpenXML components containing `xl/worksheets/sheet1.xml`, `xl/styles.xml`, etc.

- `OleCompoundFile`: Manages parsing of the Microsoft Compound File Binary stream.
- `BiffRecord` / `XlsParser`: Breaks down binary streams into logically decipherable records.
- `FormulaDecompiler`: Transforms RPN (Reverse Polish Notation) parsed formula bytes (`Ptg`) back into text.
- `XlsxGenerator`: Produces compliant Open XML (`.xlsx`) ZIP archives from the domain model.

## 📄 License
[MIT License](LICENSE)
