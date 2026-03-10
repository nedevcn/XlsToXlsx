#nullable enable
using System;
using System.IO;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using Xunit;
using Nedev.FileConverters.XlsToXlsx;
using Nedev.FileConverters.XlsToXlsx.Formats.Xls;

namespace Nedev.FileConverters.XlsToXlsx.Tests
{
    public class ConverterTests
    {
        static ConverterTests()
        {
            // silence informational logging during unit tests by default
            Logger.LogLevel = LogLevel.Warning;
        }

        [Fact]
        public void Convert_MinimalWorkbook_ShouldProduceXlsx()
        {
            // create a minimal XLS workbook with NPOI
            var workbook = new HSSFWorkbook();
            var sheet = workbook.CreateSheet("Sheet1");
            var row = sheet.CreateRow(0);
            var cell = row.CreateCell(0);
            cell.SetCellValue("Hello world");

            using var xlsStream = new MemoryStream();
            workbook.Write(xlsStream);
            xlsStream.Seek(0, SeekOrigin.Begin);

            // convert using the library
            xlsStream.Seek(0, SeekOrigin.Begin);
            var converter = new XlsToXlsxConverter();
            using var result = converter.Convert(xlsStream);
            Assert.NotNull(result);
            Assert.True(result.Length > 0);

            // optionally verify the XLSX contains expected text by reopening via NPOI OOXML
            result.Seek(0, SeekOrigin.Begin);
            var xlsxBook = new NPOI.XSSF.UserModel.XSSFWorkbook(result);
            var outSheet = xlsxBook.GetSheet("Sheet1");
            Assert.NotNull(outSheet);
            var outRow = outSheet.GetRow(0);
            Assert.NotNull(outRow);
            var outCell = outRow.GetCell(0);
            Assert.Equal("Hello world", outCell.StringCellValue);
        }

        [Fact]
        public void Convert_WithFormulas_ShouldKeepFormulaTexts()
        {
            // create an XLS with a couple of formulas
            var workbook = new HSSFWorkbook();
            var sheet = workbook.CreateSheet("Sheet1");
            var row = sheet.CreateRow(0);
            row.CreateCell(0).SetCellValue(5);
            row.CreateCell(1).SetCellValue(10);
            var formulaRow = sheet.CreateRow(1);
            formulaRow.CreateCell(0).CellFormula = "SUM(A1:B1)";
            formulaRow.CreateCell(1).CellFormula = "IF(A1>0,\"Yes\",\"No\")";

            using var xlsStream = new MemoryStream();
            workbook.Write(xlsStream);
            xlsStream.Seek(0, SeekOrigin.Begin);

            // run parser manually and validate formulas before conversion
            xlsStream.Seek(0, SeekOrigin.Begin);
            var parser = new Nedev.FileConverters.XlsToXlsx.Formats.Xls.XlsParser(xlsStream);
            // before parsing, examine raw BIFF PTGs for the problematic cell by scanning records
            xlsStream.Seek(0, SeekOrigin.Begin);
            var binReader = new System.IO.BinaryReader(xlsStream);
            while (xlsStream.Position < xlsStream.Length)
            {
                var rec = Nedev.FileConverters.XlsToXlsx.Formats.Xls.BiffRecord.Read(binReader);
                if (rec.Id == (ushort)Nedev.FileConverters.XlsToXlsx.Formats.Xls.BiffRecordType.CELL_FORMULA)
                {
                    byte[] data = rec.GetAllData();
                    if (data.Length >= 6)
                    {
                        ushort rowIdx = BitConverter.ToUInt16(data, 0);
                        ushort col = BitConverter.ToUInt16(data, 2);
                        Logger.Debug($"FOUND FORMULA RECORD @ row={rowIdx} col={col} len={data.Length}");
                        // our formula is in row 1 col 1 (0-based)
                        if (rowIdx == 1 && col == 1)
                        {
                            if (data.Length >= 22)
                            {
                                int formulaLength = BitConverter.ToUInt16(data, 20);
                                if (data.Length >= 22 + formulaLength)
                                {
                                    byte[] ptgs = new byte[formulaLength];
                                    Array.Copy(data, 22, ptgs, 0, formulaLength);
                                    Logger.Debug("DEBUG PTGS: " + BitConverter.ToString(ptgs));
                                }
                            }
                            break;
                        }
                    }
                }
            }
            xlsStream.Seek(0, SeekOrigin.Begin);
            var workbookModel = parser.Parse();
            var sheetModel = workbookModel.Worksheets[0];
            // debug: dump all parsed formulas
            foreach (var r in sheetModel.Rows.OrderBy(r=>r.RowIndex))
            {
                foreach (var c in r.Cells)
                {
                    Logger.Debug($"Parsed cell R{r.RowIndex}C{c.ColumnIndex} formula='{c.Formula}'");
                }
            }
            var row2 = sheetModel.Rows.Find(r => r.RowIndex == 2);
            Assert.NotNull(row2);
            var cellA2model = row2!.Cells.Find(c => c.ColumnIndex == 1);
            var cellB2model = row2.Cells.Find(c => c.ColumnIndex == 2);
            Assert.Equal("SUM(A1:B1)", cellA2model?.Formula);
            Assert.Equal("IF(A1>0,\"Yes\",\"No\")", cellB2model?.Formula);

            // (style patcher is internal and not relevant for this check)
            // formulas should remain correct at this point
            Assert.Equal("IF(A1>0,\"Yes\",\"No\")", cellB2model?.Formula);

            // generate XLSX manually
            using var output = new MemoryStream();
            var generator = new Nedev.FileConverters.XlsToXlsx.Formats.Xlsx.XlsxGenerator(output);
            generator.Generate(workbookModel);
            // ensure generator didn't mutate the in-memory formulas
            Assert.Equal("IF(A1>0,\"Yes\",\"No\")", cellB2model?.Formula);
            output.Seek(0, SeekOrigin.Begin);
            // inspect worksheet XML to ensure formulas were written
            using var zip = new System.IO.Compression.ZipArchive(output, System.IO.Compression.ZipArchiveMode.Read, leaveOpen:true);
            var entry = zip.GetEntry("xl/worksheets/sheet1.xml");
            Assert.NotNull(entry);
            using var reader = new StreamReader(entry.Open());
            string xml = reader.ReadToEnd();
            // parse the worksheet XML and verify formula text content (unescaped)
            var xdoc = System.Xml.Linq.XDocument.Parse(xml);
            var formulas = xdoc.Descendants().Where(e => e.Name.LocalName == "f").Select(e => e.Value).ToList();
            Assert.True(formulas.Count >= 2, "Expected at least two <f> elements");
            Assert.Equal("SUM(A1:B1)", formulas[0]);
            Assert.Equal("IF(A1>0,\"Yes\",\"No\")", formulas[1]);
        }

        [Fact]
        public void FormulaDecompiler_UnknownIndex_FallsBack()
        {
            var type = typeof(Nedev.FileConverters.XlsToXlsx.Formats.Xls.FormulaDecompiler);
            var method = type.GetMethod("GetFunctionName", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);
            Assert.NotNull(method);
            var result = method.Invoke(null, new object[] { (ushort)999 });
            Assert.Equal("FUNC_999", result);
        }

        [Fact]
        public void FormulaDecompiler_ShouldHandleIfWithStrings()
        {
            // manually construct PTG sequence for IF(A1>0,"Yes","No")
            var bytes = new System.Collections.Generic.List<byte>();

            // push reference A1
            bytes.Add(0x24);
            bytes.AddRange(BitConverter.GetBytes((ushort)0)); // row
            bytes.AddRange(BitConverter.GetBytes((ushort)0)); // col

            // push integer 0
            bytes.Add(0x1E);
            bytes.AddRange(BitConverter.GetBytes((ushort)0));

            // greater-than operator
            bytes.Add(0x0D);

            // push string "Yes"
            var yes = System.Text.Encoding.ASCII.GetBytes("Yes");
            bytes.Add(0x17);
            bytes.Add((byte)yes.Length);
            bytes.Add(0x00); // options: ASCII
            bytes.AddRange(yes);

            // push string "No"
            var no = System.Text.Encoding.ASCII.GetBytes("No");
            bytes.Add(0x17);
            bytes.Add((byte)no.Length);
            bytes.Add(0x00);
            bytes.AddRange(no);

            // function IF with variable arguments: ptg=0x22, argc=3, index=1
            bytes.Add(0x22);
            bytes.Add(3);
            bytes.AddRange(BitConverter.GetBytes((ushort)1));

            string result = FormulaDecompiler.Decompile(bytes.ToArray(), null);
            Assert.Equal("IF(A1>0,\"Yes\",\"No\")", result);
        }

        [Fact]
        public void Convert_BookSample_InvoiceHeaderCells_ShouldKeepBackgroundFill()
        {
            string sourcePath = GetBookSourcePath();

            // open the original XLS as the source of truth for expected styles
            using var sourceStream = File.OpenRead(sourcePath);
            var expectedBook = new HSSFWorkbook(sourceStream);

            var converter = new XlsToXlsxConverter();
            using var result = converter.Convert(File.OpenRead(sourcePath));
            result.Position = 0;
            using var convertedBook = new XSSFWorkbook(result);

            ISheet expectedSheet = FindInvoiceSheet(expectedBook);
            ISheet convertedSheet = FindInvoiceSheet(convertedBook);

            foreach (string address in new[] { "A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1" })
            {
                ICell expectedCell = GetCell(expectedSheet, address);
                ICell convertedCell = GetCell(convertedSheet, address);
                string? expectedFill = GetVisibleFillHex(expectedBook, expectedCell);
                string? convertedFill = GetVisibleFillHex(convertedBook, convertedCell);

                Assert.Equal(expectedFill, convertedFill);
            }
        }

        [Fact]
        public void Convert_BookSample_InvoiceB1FontStyle_ShouldMatchExpectedValues()
        {
            string sourcePath = GetBookSourcePath();

            using var sourceStream = File.OpenRead(sourcePath);
            var expectedBook = new HSSFWorkbook(sourceStream);

            var converter = new XlsToXlsxConverter();
            using var result = converter.Convert(File.OpenRead(sourcePath));
            result.Position = 0;
            using var convertedBook = new XSSFWorkbook(result);

            ICell expectedCell = GetCell(FindInvoiceSheet(expectedBook), "B1");
            ICell convertedCell = GetCell(FindInvoiceSheet(convertedBook), "B1");
            IFont expectedFont = expectedBook.GetFontAt(expectedCell.CellStyle.FontIndex);
            IFont convertedFont = convertedBook.GetFontAt(convertedCell.CellStyle.FontIndex);

            Assert.Equal(expectedFont.FontHeightInPoints, convertedFont.FontHeightInPoints);
            Assert.Equal(expectedFont.IsBold, convertedFont.IsBold);
            Assert.Equal(GetFontColorHex(expectedBook, expectedFont), GetFontColorHex(convertedBook, convertedFont));
        }

        private static string GetBookSourcePath()
        {
            return Path.Combine(FindSamplesDirectory(), "Book.xls");
        }


        private static string FindSamplesDirectory()
        {
            var dir = new DirectoryInfo(AppContext.BaseDirectory);
            while (dir != null)
            {
                var candidate = Path.Combine(dir.FullName, "samples");
                if (Directory.Exists(candidate))
                    return candidate;
                dir = dir.Parent;
            }
            throw new DirectoryNotFoundException("Could not locate 'samples' directory in parent hierarchy.");
        }

        private static ISheet FindInvoiceSheet(IWorkbook workbook)
        {
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                ISheet sheet = workbook.GetSheetAt(i);
                if (!string.IsNullOrEmpty(sheet.SheetName) && sheet.SheetName.Contains("发票", StringComparison.OrdinalIgnoreCase))
                    return sheet;
            }

            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                ISheet sheet = workbook.GetSheetAt(i);
                ICell? candidate = TryGetCell(sheet, "B1");
                if (candidate == null)
                    continue;

                IFont font = workbook.GetFontAt(candidate.CellStyle.FontIndex);
                if (font.FontHeightInPoints == 25 && font.IsBold)
                    return sheet;
            }

            return workbook.GetSheetAt(0);
        }

        private static ICell GetCell(ISheet sheet, string address)
        {
            ICell? cell = TryGetCell(sheet, address);
            Assert.NotNull(cell);
            return cell!;
        }

        private static ICell? TryGetCell(ISheet sheet, string address)
        {
            CellReference reference = new CellReference(address);
            IRow? row = sheet.GetRow(reference.Row);
            return row?.GetCell(reference.Col);
        }

        private static string? GetVisibleFillHex(IWorkbook workbook, ICell cell)
        {
            if (workbook is HSSFWorkbook hssfWorkbook)
            {
                var style = (HSSFCellStyle)cell.CellStyle;
                return GetHssfColorHex(hssfWorkbook, style.FillForegroundColor)
                    ?? GetHssfColorHex(hssfWorkbook, style.FillBackgroundColor);
            }

            if (workbook is XSSFWorkbook xssfWorkbook)
            {
                var style = (XSSFCellStyle)cell.CellStyle;
                return NormalizeRgbHex(style.FillForegroundXSSFColor?.RGB)
                    ?? NormalizeRgbHex(style.FillBackgroundColorColor?.RGB);
            }

            return null;
        }

        private static string? GetFontColorHex(IWorkbook workbook, IFont font)
        {
            if (workbook is HSSFWorkbook hssfWorkbook && font is HSSFFont hssfFont)
                return GetHssfColorHex(hssfWorkbook, hssfFont.Color);

            if (workbook is XSSFWorkbook && font is XSSFFont xssfFont)
                return NormalizeRgbHex(xssfFont.GetXSSFColor()?.RGB);

            return null;
        }

        private static string? GetHssfColorHex(HSSFWorkbook workbook, short colorIndex)
        {
            HSSFColor? color = workbook.GetCustomPalette().GetColor(colorIndex);
            if (color == null)
                return null;

            // NPOI 2.6 changed GetTriplet to return byte[]
            var triplet = color.GetTriplet();
            return string.Concat(triplet.Select(component => component.ToString("X2")));
        }

        private static string? NormalizeRgbHex(byte[]? rgb)
        {
            if (rgb == null || rgb.Length == 0)
                return null;

            if (rgb.Length >= 3)
                return string.Concat(rgb.Take(3).Select(component => component.ToString("X2")));

            return null;
        }
    }
}