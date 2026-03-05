# Nedev.XlsToXlsx

**Nedev.XlsToXlsx** is a standalone, lightweight, and fast .NET library designed to convert legacy Microsoft Excel files (`.xls` / BIFF8 format) into the modern OpenXML format (`.xlsx`). It performs this conversion entirely in memory and **does not require Microsoft Office or Excel Interop** to be installed.

## 🚀 Features

### ✅ Supported
- **OLE Compound File Parsing**: Fully implements FAT and DIFAT chain parsing to reliably read even large and highly-fragmented `.xls` files.
- **Cell Data**: Extracts formulas, numbers, strings (including Shared String Table), booleans, and errors.
- **Formatting & Styles**: Converts fonts (size, color, bold, italic), cell fills/patterns, borders, and alignment.
- **Worksheet Layout**: Preserves row heights, column widths, fixed/frozen panes, and hidden rows/cols.
- **Merged Cells**: Accurately maps merged cell ranges.
- **Hyperlinks**: Preserves cell hyperlinks (external URLs, local files, and email addresses).
- **Comments/Notes**: Extracts basic cell notes/comments.
- **Data Validation**: Retains basic data validation rules (dropdowns, number constraints).
- **VBA Macros**: Preserves existing VBA macros by extracting the raw `vbaProject.bin` from the legacy file and embedding it into a macro-enabled `.xlsm` compatible structure.
- **Page Setup**: Keeps print margins, page orientation, paper size, and fit-to-page scaling.
- **Pivot Tables**: Converts pivot table structure (fields, layout, data source range); output can be refreshed in Excel.

### ⚠️ Partially Supported (WIP)
- **Formulas**: A custom formula decompiler supports over 170+ standard Excel functions. Shared formulas (`SHAREDFMLA`) and array formulas (`ARRAY`) are supported.
- **Charts**: Can detect and convert basic chart types (bar, line, pie, etc.), but advanced 3D properties and secondary axes are not yet fully mapped.
- **Images & Drawings**: Basic image extraction is supported, but complex Microsoft Office Drawing (Escher) containers are partially parsed.
- **Conditional Formatting**: Detection is supported, but styling rules (like Color Scales and Data Bars) currently use fallback styles instead of the exact embedded binary properties.

### ❌ Not Yet Supported
- AutoFilter configurations
- Worksheet and Workbook Password Protection
- Summary Information / Document Properties
- Multi-workbook external references (`EXTERNSHEET`)

## 📦 Installation

*(To be added when published to NuGet)*
```bash
dotnet add package Nedev.XlsToXlsx
```

## 💻 Usage

Converting a `.xls` file to `.xlsx` takes just a few lines of code:

```csharp
using System;
using Nedev.XlsToXlsx;

class Program
{
    static void Main()
    {
        string inputFilePath = @"C:\path\to\your\legacy_file.xls";
        string outputFilePath = @"C:\path\to\your\converted_file.xlsx";

        try
        {
            // Convert with an optional progress callback
            XlsToXlsxConverter.Convert(inputFilePath, outputFilePath, (percentage, message) =>
            {
                Console.WriteLine($"Progress: {percentage}% - {message}");
            });

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
