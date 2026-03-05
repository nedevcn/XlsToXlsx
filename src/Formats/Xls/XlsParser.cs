using System.IO;
using System.Collections.Generic;
using System;
using System.Linq;
using System.Text.RegularExpressions;
using Nedev.XlsToXlsx;
using Nedev.XlsToXlsx.Exceptions;

namespace Nedev.XlsToXlsx.Formats.Xls
{
    public class XlsParser
    {
        private Stream _rawStream;
        private OleCompoundFile _oleFile = null!;
        private byte[] _workbookData = null!;  // Workbook 流的完整字节数据
        private Stream _stream = null!;        // 当前正在解析的流 (MemoryStream wrapper)
        private BinaryReader _reader = null!;
        private List<string> _sharedStrings;
        private List<Font> _fonts = new List<Font>();
        private List<Xf> _xfList = new List<Xf>();
        private Dictionary<ushort, string> _formats = new Dictionary<ushort, string>();
        private Dictionary<int, string> _palette = new Dictionary<int, string>();
        private Workbook _workbook = null!;

        /// <summary>BIFF8 默认 64 色调色板 (索引 0-63)，无 PALETTE 记录时使用。与 Excel 默认一致。</summary>
        private static readonly IReadOnlyDictionary<int, string> Biff8DefaultPalette = CreateBiff8DefaultPalette();

        private static IReadOnlyDictionary<int, string> CreateBiff8DefaultPalette()
        {
            var d = new Dictionary<int, string>();
            // 0-7: 固定色
            string[] first8 = { "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF" };
            for (int i = 0; i < first8.Length; i++) d[i] = first8[i];
            // 8-63: Excel 默认 56 色 (标准 BIFF8 调色板)
            string[] rest = {
                "800000", "000080", "008000", "008080", "800080", "808000", "C0C0C0", "808080",
                "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
                "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
                "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
                "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
                "003366", "339966", "003300", "333300", "993300", "993366", "999933", "666666",
                "0066FF", "00CCFF", "00FFFF", "00CC00", "00FF00", "99FF00", "99CC00", "999900"
            };
            for (int i = 0; i < rest.Length && (8 + i) <= 63; i++) d[8 + i] = rest[i];
            return d;
        }

        private const long MAX_FILE_SIZE = 100 * 1024 * 1024; // 100MB文件大小限制

        // 工作簿级别的OfficeArt (DggContainer bytes)
        private List<byte[]> _msoDrawingGroupData = new List<byte[]>();

        // BOUNDSHEET 记录中的工作表子流偏移量 (lbPlyPos)
        private List<int> _sheetOffsets = new List<int>();

        /// <summary>
        /// VBA项目大小限制（字节）
        /// </summary>
        public long VbaSizeLimit { get; set; } = 50 * 1024 * 1024;
        private string _currentCFRange = string.Empty; // 当前条件格式的范围

        // 画图对象状态
        private List<byte[]> _msoDrawingData = new List<byte[]>();
        private List<(int Left, int Top, int Width, int Height)> _pendingChartAnchors = new List<(int, int, int, int)>();

        /// <summary>
        /// XLS 文件打开密码
        /// </summary>
        public string Password { get; set; } = "VelvetSweatshop";
        private XlsDecryptor? _decryptor;

        // 数据透视表解析状态
        private PivotTable? _currentPivotTable;
        private PivotField? _currentPivotField;
        /// <summary>当前解析的工作表索引（0-based），用于解析 AutoFilter 时查找 _FilterDatabase 名称。</summary>
        private int _currentSheetIndex;

        // 共享公式：SHAREDFMLA 后的主公式及范围，用于后续 FORMULA 单元格按行列差调整引用
        private string? _sharedFormulaString;
        private int _sharedFormulaBaseRow;
        private int _sharedFormulaBaseCol;
        private int _sharedFormulaLastRow;
        private int _sharedFormulaLastCol;

        public XlsParser(Stream stream)
        {
            // 验证流是否可读
            if (!stream.CanRead)
            {
                throw new XlsToXlsxException("Stream must be readable", 1000, "StreamError");
            }

            // 检查文件大小限制
            if (stream.CanSeek)
            {
                long fileSize = stream.Length;
                if (fileSize > MAX_FILE_SIZE)
                {
                    throw new XlsToXlsxException($"File size exceeds limit of {MAX_FILE_SIZE / (1024 * 1024)}MB", 1002, "FileSizeError");
                }
            }

            _rawStream = stream;
            _sharedStrings = new List<string>();
        }

        public Workbook Parse()
        {
            _workbook = new Workbook();
            var workbook = _workbook;

            try
            {
                Logger.Info("开始解析XLS文件");

                // 1. 使用 OleCompoundFile 解析 OLE 复合文件结构
                _oleFile = new OleCompoundFile(_rawStream);
                Logger.Info("OLE复合文件解析完成");

                // 2. 读取 Workbook 流
                _workbookData = _oleFile.ReadStreamByName("Workbook")
                             ?? _oleFile.ReadStreamByName("Book")  // Excel 5.0/95 兼容
                             ?? throw new XlsParseException("在OLE文件中未找到Workbook或Book流");
                Logger.Info($"Workbook流读取完成: {_workbookData.Length} 字节");

                // 3. 在 Workbook MemoryStream 上解析 BIFF 记录
                _stream = new MemoryStream(_workbookData);
                _reader = new BinaryReader(_stream);

                // 4. 解析全局记录 (BOUNDSHEET, SST, FONT, XF, FORMAT, PALETTE, NAME)
                ParseWorkbookGlobals(workbook);
                Logger.Info($"全局记录解析完成: {workbook.Worksheets.Count} 个工作表, {_sharedStrings.Count} 个共享字符串");

                // 5. 根据 BOUNDSHEET 中记录的偏移量解析各工作表子流
                ParseAllWorksheetSubstreams(workbook);
                Logger.Info("所有工作表子流解析完成");

                // 6. 将解析到的全局数据转移到工作簿对象
                workbook.SharedStrings = _sharedStrings;
                workbook.Fonts = _fonts;
                workbook.XfList = _xfList;
                workbook.NumberFormats = _formats;
                workbook.Palette = _palette;

                // 7. 解析VBA流
                ParseVbaStream(workbook);
                Logger.Info("VBA流解析完成");

                Logger.Info("XLS文件解析成功");
                return workbook;
            }
            catch (XlsToXlsxException)
            {
                throw;
            }
            catch (System.IO.InvalidDataException ex)
            {
                Logger.Error("解析XLS文件时发生数据格式错误", ex);
                throw new XlsParseException($"解析XLS文件时发生数据格式错误: {ex.Message}", ex);
            }
            catch (System.IO.IOException ex)
            {
                Logger.Error("解析XLS文件时发生IO错误", ex);
                throw new XlsParseException($"解析XLS文件时发生IO错误: {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                Logger.Error("解析XLS文件时发生未知错误", ex);
                throw new XlsParseException($"解析XLS文件时发生未知错误: {ex.Message}", ex);
            }
        }

        public async Task<Workbook> ParseAsync()
        {
            // 使用Task.Run在后台线程中执行解析，避免阻塞主线程
            return await Task.Run(() => Parse());
        }
        
        // ===== 全局流解析 =====

        /// <summary>
        /// 解析 Workbook 全局子流（从 BOF 到 EOF），收集 BOUNDSHEET/SST/FONT/XF/FORMAT 等全局记录。
        /// </summary>
        private void ParseWorkbookGlobals(Workbook workbook)
        {
            _stream.Seek(0, SeekOrigin.Begin);
            long streamEnd = _workbookData.Length;

            BiffRecord? previousRecord = null;

            while (_stream.Position < streamEnd)
            {
                try
                {
                    var recordStartPos = _stream.Position;
                    var record = BiffRecord.Read(_reader);

                    // 如果已启用解密，且不是特殊的非加密记录
                    if (_decryptor != null && record.Id != (ushort)BiffRecordType.BOF && record.Id != (ushort)BiffRecordType.FILEPASS)
                    {
                        if (record.Data != null && record.Data.Length > 0)
                        {
                            // 记录体数据的起始位置是记录头之后的 4 字节
                            _decryptor.Decrypt(record.Data, recordStartPos + 4);
                        }
                    }

                    // 自动合并 CONTINUE 记录
                    if (record.Id == (ushort)BiffRecordType.CONTINUE)
                    {
                        if (previousRecord != null && record.Data != null)
                        {
                            previousRecord.Continues.Add(record.Data);
                        }
                        continue;
                    }

                    // 如果我们遇到一个新记录，先处理前一个记录（因为它现在已经收齐了所有的 CONTINUE 分块）
                    if (previousRecord != null)
                    {
                        ProcessWorkbookGlobalRecord(previousRecord, workbook, streamEnd);
                    }

                    // 将当前记录设为 Previous，等待下一个循环决定是否遇到它的 CONTINUE
                    previousRecord = record;

                    if (record.Id == (ushort)BiffRecordType.EOF)
                    {
                        break;
                    }
                }
                catch (EndOfStreamException)
                {
                    break;
                }
                catch (XlsToXlsxException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    Logger.Error($"解析Workbook全局记录时发生错误: {ex.Message}", ex);
                    continue;
                }
            }

            // 处理最后一个记录
            if (previousRecord != null && previousRecord.Id != (ushort)BiffRecordType.EOF)
            {
                ProcessWorkbookGlobalRecord(previousRecord, workbook, streamEnd);
            }
        }

        private void ProcessWorkbookGlobalRecord(BiffRecord record, Workbook workbook, long streamEnd)
        {
            switch (record.Id)
            {
                case (ushort)BiffRecordType.BOF:
                    break;
                case (ushort)BiffRecordType.EOF:
                    return;
                case (ushort)BiffRecordType.SHEET:
                    ParseSheetRecord(record, workbook);
                    break;
                case (ushort)BiffRecordType.SST:
                    ParseSstInfo(record, streamEnd);
                    break;
                case (ushort)BiffRecordType.FONT:
                    ParseFontRecordToGlobal(record);
                    break;
                case (ushort)BiffRecordType.XF:
                    ParseXfRecordToGlobal(record);
                    break;
                case (ushort)BiffRecordType.FORMAT:
                    ParseFormatRecordGlobal(record);
                    break;
                case (ushort)BiffRecordType.PALETTE:
                    ParsePaletteRecordGlobal(record);
                    break;
                case (ushort)BiffRecordType.BORDER:
                    ParseBorderRecord(record, workbook);
                    break;
                case (ushort)BiffRecordType.FILL:
                    ParseFillRecord(record, workbook);
                    break;
                case (ushort)BiffRecordType.NAME:
                    ParseNameRecord(record, workbook);
                    break;
                case (ushort)BiffRecordType.MSODRAWINGGROUP:
                    ParseMsoDrawingGroupGlobal(record, workbook);
                    break;
                case (ushort)BiffRecordType.FILEPASS:
                    if (record.Data != null && record.Data.Length >= 52)
                    {
                        Logger.Info("检测到加密文件，正在初始化解密器");
                        _decryptor = new XlsDecryptor(record.Data, Password);
                    }
                    break;
                case (ushort)BiffRecordType.EXTERNBOOK:
                    ParseExternBookRecord(record, workbook);
                    break;
                case (ushort)BiffRecordType.EXTERNSHEET:
                    ParseExternSheetRecord(record, workbook);
                    break;
                case (ushort)BiffRecordType.EXTERNALNAME:
                    ParseExternalNameRecord(record, workbook);
                    break;
            }
        }

        // ===== 工作表子流解析 =====

        /// <summary>
        /// 使用 BOUNDSHEET 记录中的 lbPlyPos 偏移量定位和解析每个工作表子流。
        /// </summary>
        private void ParseAllWorksheetSubstreams(Workbook workbook)
        {
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                if (i >= _sheetOffsets.Count)
                {
                    Logger.Warn($"工作表 {i} 没有对应的偏移量信息");
                    break;
                }

                int offset = _sheetOffsets[i];
                if (offset < 0 || offset >= _workbookData.Length)
                {
                    Logger.Warn($"工作表 {workbook.Worksheets[i].Name} 的偏移量 {offset} 无效");
                    continue;
                }

                try
                {
                    _currentSheetIndex = i;
                    _stream.Seek(offset, SeekOrigin.Begin);
                    ParseWorksheetSubstream(workbook.Worksheets[i], workbook);
                    Logger.Info($"工作表 {workbook.Worksheets[i].Name} 解析完成: {workbook.Worksheets[i].Rows.Count} 行");
                }
                catch (Exception ex)
                {
                    Logger.Error($"解析工作表 {workbook.Worksheets[i].Name} 时发生错误", ex);
                }
            }
        }

        /// <summary>
        /// 解析单个工作表子流（从 BOF 到 EOF）
        /// </summary>
        private void ParseWorksheetSubstream(Worksheet worksheet, Workbook workbook)
        {
            _msoDrawingData.Clear();
            _pendingChartAnchors.Clear();
            
            long streamEnd = _workbookData.Length;
            Row? currentRow = null;

            BiffRecord? previousRecord = null;

            while (_stream.Position < streamEnd)
            {
                try
                {
                    var recordStartPos = _stream.Position;
                    var record = BiffRecord.Read(_reader);

                    // 解密数据体
                    if (_decryptor != null && record.Id != (ushort)BiffRecordType.BOF)
                    {
                        if (record.Data != null && record.Data.Length > 0)
                        {
                            _decryptor.Decrypt(record.Data, recordStartPos + 4);
                        }
                    }

                    // 自动合并 CONTINUE 记录
                    if (record.Id == (ushort)BiffRecordType.CONTINUE)
                    {
                        if (previousRecord != null && record.Data != null)
                        {
                            previousRecord.Continues.Add(record.Data);
                        }
                        continue;
                    }

                    // Process previous record once all its CONTINUES are fetched
                    if (previousRecord != null)
                    {
                        ProcessWorksheetRecord(previousRecord, worksheet, workbook, ref currentRow!, streamEnd);
                    }

                    previousRecord = record;

                    if (record.Id == (ushort)BiffRecordType.EOF)
                    {
                        break;
                    }
                }
                catch (EndOfStreamException)
                {
                    break;
                }
                catch (XlsToXlsxException)
                {
                    throw;
                }
                catch (Exception ex)
                {
                    Logger.Error($"解析工作表记录时发生错误: {ex.Message}", ex);
                    continue;
                }
            }

            // 处理最后一个记录
            if (previousRecord != null && previousRecord.Id != (ushort)BiffRecordType.EOF)
            {
                ProcessWorksheetRecord(previousRecord, worksheet, workbook, ref currentRow!, streamEnd);
            }
            
            // 在遇到工作表 EOF 后，检查紧随其后的子流是否为图表子流 (BOF type = 0x0020)
            while (_stream.Position < streamEnd)
            {
                long posBeforePeek = _stream.Position;
                try
                {
                    var nextRecord = BiffRecord.Read(_reader);
                    if (nextRecord.Id == (ushort)BiffRecordType.BOF && nextRecord.Data != null && nextRecord.Data.Length >= 4)
                    {
                        ushort bofType = BitConverter.ToUInt16(nextRecord.Data, 2);
                        if (bofType == 0x0020) // 图表子流
                        {
                            Logger.Info($"在工作表 {worksheet.Name} 后发现图表子流，开始解析");
                            ParseChartSubstream(worksheet, workbook);
                            continue; // 图表子流解析完毕后，继续检查是否还有其他图表子流附在该工作表后面
                        }
                    }
                }
                catch
                {
                    // 无法读取或到达流末尾
                }
                
                // 如果不是图表子流，恢复指针位置并退出
                _stream.Position = posBeforePeek;
                break;
            }
        }

        private void ProcessWorksheetRecord(BiffRecord record, Worksheet worksheet, Workbook workbook, ref Row currentRow, long streamEnd)
        {
            switch (record.Id)
            {
                case (ushort)BiffRecordType.BOF:
                    break;
                case (ushort)BiffRecordType.EOF:
                    return;
                        case (ushort)BiffRecordType.DIMENSION:
                            ParseDimensionRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.ROW:
                            var parsedRow = ParseRowRecord(record);
                            var existingRow = GetOrCreateRow(worksheet, ref currentRow, parsedRow.RowIndex);
                            existingRow.Height = parsedRow.Height;
                            existingRow.CustomHeight = parsedRow.CustomHeight;
                            existingRow.DefaultXfIndex = parsedRow.DefaultXfIndex;
                            if (existingRow.RowIndex > worksheet.MaxRow) worksheet.MaxRow = (int)existingRow.RowIndex;
                            break;
                        case (ushort)BiffRecordType.CELL_BLANK:
                        case (ushort)BiffRecordType.CELL_BOOLERR:
                        case (ushort)BiffRecordType.CELL_LABEL:
                        case (ushort)BiffRecordType.CELL_LABELSST:
                        case (ushort)BiffRecordType.CELL_NUMBER:
                        case (ushort)BiffRecordType.CELL_RK:
                        case (ushort)BiffRecordType.CELL_RSTRING: // 旧版富文本
                            var cell = ParseCellRecord(record);
                            if (cell.ColumnIndex >= 1 && cell.ColumnIndex <= 16384)
                            {
                                var targetRow = GetOrCreateRow(worksheet, ref currentRow, cell.RowIndex);
                                targetRow.Cells.Add(cell);
                                if (cell.ColumnIndex > worksheet.MaxColumn) worksheet.MaxColumn = cell.ColumnIndex;
                                if (cell.RowIndex > worksheet.MaxRow) worksheet.MaxRow = cell.RowIndex;
                            }
                            break;
                        case (ushort)BiffRecordType.CELL_FORMULA:
                            var formulaCell = ParseCellRecord(record);
                            if (formulaCell.ColumnIndex >= 1 && formulaCell.ColumnIndex <= 16384)
                            {
                                ApplySharedFormulaToCell(formulaCell);
                                var targetRow2 = GetOrCreateRow(worksheet, ref currentRow, formulaCell.RowIndex);
                                targetRow2.Cells.Add(formulaCell);
                                if (formulaCell.ColumnIndex > worksheet.MaxColumn) worksheet.MaxColumn = formulaCell.ColumnIndex;
                                if (formulaCell.RowIndex > worksheet.MaxRow) worksheet.MaxRow = formulaCell.RowIndex;
                            }
                            break;
                        case (ushort)BiffRecordType.ARRAY:
                            ParseArrayRecord(record, ref currentRow);
                            break;
                        case (ushort)BiffRecordType.SHAREDFMLA:
                            ParseSharedFmlaRecord(record);
                            break;
                        case (ushort)BiffRecordType.STRING:
                            if (currentRow != null && currentRow.Cells.Count > 0)
                            {
                                var lastCell = currentRow.Cells[currentRow.Cells.Count - 1];
                                byte[] strData = record.GetAllData();
                                if (strData.Length > 0)
                                {
                                    int strOffset = 0;
                                    lastCell.Value = ReadBiffString(strData, ref strOffset);
                                    lastCell.DataType = "inlineStr";
                                }
                            }
                            break;
                        case (ushort)BiffRecordType.MULRK:
                            ParseMulRkRecord(record, ref currentRow, worksheet);
                            break;
                        case (ushort)BiffRecordType.MULBLANK:
                            ParseMulBlankRecord(record, ref currentRow, worksheet);
                            break;
                        case (ushort)BiffRecordType.MERGECELLS:
                            ParseMergeCellsRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.COLINFO:
                            ParseColInfoRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.DEFCOLWIDTH:
                            if (record.Data != null && record.Data.Length >= 2)
                                worksheet.DefaultColumnWidth = BitConverter.ToUInt16(record.Data, 0);
                            break;
                        case (ushort)BiffRecordType.DEFAULTROWHEIGHT:
                            if (record.Data != null && record.Data.Length >= 4)
                                worksheet.DefaultRowHeight = BitConverter.ToUInt16(record.Data, 2) / 20.0;
                            break;
                        case (ushort)BiffRecordType.WINDOW2:
                            ParseWindow2Record(record, worksheet);
                            break;
                        case 0x0041: // PANE
                            ParsePaneRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.DV:
                            ParseDVRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.CFHEADER:
                            ParseCFHeaderRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.CF:
                            ParseCFRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.HYPERLINK:
                            ParseHyperlinkRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.NOTE:
                            ParseCommentRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.MSODRAWING:
                            ParseMSODrawingRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.PICTURE:
                            ParsePictureRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.OBJ:
                            ParseObjRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.HEADER:
                            byte[] headerData = record.GetAllData();
                            if (headerData.Length > 0)
                            {
                                int hPos = 0;
                                worksheet.PageSettings.Header = ReadBiffString(headerData, ref hPos);
                            }
                            break;
                        case (ushort)BiffRecordType.FOOTER:
                            byte[] footerData = record.GetAllData();
                            if (footerData.Length > 0)
                            {
                                int fPos = 0;
                                worksheet.PageSettings.Footer = ReadBiffString(footerData, ref fPos);
                            }
                            break;
                        case (ushort)BiffRecordType.LEFTMARGIN:
                            if (record.Data != null && record.Data.Length >= 8)
                                worksheet.PageSettings.LeftMargin = BitConverter.ToDouble(record.Data, 0);
                            break;
                        case (ushort)BiffRecordType.RIGHTMARGIN:
                            if (record.Data != null && record.Data.Length >= 8)
                                worksheet.PageSettings.RightMargin = BitConverter.ToDouble(record.Data, 0);
                            break;
                        case (ushort)BiffRecordType.TOPMARGIN:
                            if (record.Data != null && record.Data.Length >= 8)
                                worksheet.PageSettings.TopMargin = BitConverter.ToDouble(record.Data, 0);
                            break;
                        case (ushort)BiffRecordType.BOTTOMMARGIN:
                            if (record.Data != null && record.Data.Length >= 8)
                                worksheet.PageSettings.BottomMargin = BitConverter.ToDouble(record.Data, 0);
                            break;
                        case (ushort)BiffRecordType.HCENTER:
                            if (record.Data != null && record.Data.Length >= 2)
                                worksheet.PageSettings.HorizontalCenter = BitConverter.ToUInt16(record.Data, 0) != 0;
                            break;
                        case (ushort)BiffRecordType.VCENTER:
                            if (record.Data != null && record.Data.Length >= 2)
                                worksheet.PageSettings.VerticalCenter = BitConverter.ToUInt16(record.Data, 0) != 0;
                            break;
                        case (ushort)BiffRecordType.PAGESETUP:
                            ParsePageSetupRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.FONT:
                            ParseFontRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.XF:
                            ParseXfRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.PALETTE:
                            ParsePaletteRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.FORMAT:
                            ParseFormatRecord(record);
                            break;
                        case (ushort)BiffRecordType.SXVIEW:
                        case (ushort)BiffRecordType.SXVD:
                        case (ushort)BiffRecordType.SXVI:
                        case (ushort)BiffRecordType.SXDX:
                        case (ushort)BiffRecordType.SXFIELD:
                        case (ushort)BiffRecordType.SXPI:
                            ParsePivotTableRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.AUTOFILTERINFO:
                            ParseAutoFilterInfoRecord(record, worksheet, workbook);
                            break;
                        case (ushort)BiffRecordType.AUTOFILTER:
                            ParseAutoFilterRecord(record, worksheet);
                            break;
                case (ushort)BiffRecordType.SST:
                    ParseSstInfo(record, streamEnd);
                    break;
            }
        }


        // ===== 以下为辅助方法 =====
        private Row GetOrCreateRow(Worksheet worksheet, ref Row currentRow, int rowIndex)
    {
        if (currentRow != null && currentRow.RowIndex == rowIndex)
            return currentRow;
            
        // 倒序查找，因为行通常是顺序添加的，从后往前找最快
        for (int i = worksheet.Rows.Count - 1; i >= 0; i--)
        {
            if (worksheet.Rows[i].RowIndex == rowIndex)
            {
                currentRow = worksheet.Rows[i];
                return currentRow;
            }
        }
        
        // 找不到则创建新行
        var newRow = new Row { RowIndex = rowIndex };
        newRow.Cells.Capacity = 20;
        worksheet.Rows.Add(newRow);
        currentRow = newRow;
        return newRow;
    }
    
    private void ParseFormatRecord(BiffRecord record)
        {
            byte[] data = record.GetAllData();
            if (data != null && data.Length >= 18)
            {
                ushort formatIndex = BitConverter.ToUInt16(data, 0);
                byte formatLength = data[2];
                if (data.Length >= 3 + formatLength)
                {
                    string formatString = System.Text.Encoding.ASCII.GetString(data, 3, formatLength);
                    _formats[formatIndex] = formatString;
                }
            }
        }
        
        private void ParseMulRkRecord(BiffRecord record, ref Row currentRow, Worksheet worksheet)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 12)
                return;
            ushort row = BitConverter.ToUInt16(data, 0);
            ushort firstCol = BitConverter.ToUInt16(data, 2);
            ushort lastCol = BitConverter.ToUInt16(data, 4);
            // 数组从偏移 6 开始： [XF(2), RK(4)] * numCells；与 lastCol 一致时 numCells = lastCol - firstCol + 1
            int numCells = (data.Length - 6) / 6;
            if (lastCol >= firstCol)
                numCells = Math.Min(numCells, lastCol - firstCol + 1);
            if (numCells <= 0) return;
            var targetRow = GetOrCreateRow(worksheet, ref currentRow, row + 1);
            for (int j = 0; j < numCells; j++)
            {
                int offset = 6 + j * 6;
                if (offset + 6 > data.Length) break;
                ushort xfIndex = BitConverter.ToUInt16(data, offset);
                int rkValue = BitConverter.ToInt32(data, offset + 2);
                double value = DecodeRKValue(rkValue);
                var cell = new Cell
                {
                    RowIndex = row + 1,
                    ColumnIndex = firstCol + j + 1,
                    Value = value,
                    DataType = "n",
                    StyleId = xfIndex.ToString()
                };
                targetRow.Cells.Add(cell);
                if (cell.ColumnIndex > worksheet.MaxColumn)
                    worksheet.MaxColumn = cell.ColumnIndex;
            }
        }
        
        private void ParseMulBlankRecord(BiffRecord record, ref Row currentRow, Worksheet worksheet)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 8)
                return;
            ushort row = BitConverter.ToUInt16(data, 0);
            ushort firstCol = BitConverter.ToUInt16(data, 2);
            ushort lastCol = BitConverter.ToUInt16(data, 4);
            // 数组从偏移 6 开始： [XF(2)] * numCells
            int numCells = (data.Length - 6) / 2;
            if (lastCol >= firstCol)
                numCells = Math.Min(numCells, lastCol - firstCol + 1);
            if (numCells <= 0) return;
            var targetRow = GetOrCreateRow(worksheet, ref currentRow, row + 1);
            for (int j = 0; j < numCells; j++)
            {
                int offset = 6 + j * 2;
                if (offset + 2 > data.Length) break;
                ushort xfIndex = BitConverter.ToUInt16(data, offset);
                var cell = new Cell
                {
                    RowIndex = row + 1,
                    ColumnIndex = firstCol + j + 1,
                    Value = null,
                    StyleId = xfIndex.ToString()
                };
                targetRow.Cells.Add(cell);
                if (cell.ColumnIndex > worksheet.MaxColumn)
                    worksheet.MaxColumn = cell.ColumnIndex;
            }
        }
        
        private void ParseColInfoRecord(BiffRecord record, Worksheet worksheet)
        {
            if (record.Data == null || record.Data.Length < 10)
                return;
            ushort firstCol = BitConverter.ToUInt16(record.Data, 0);
            ushort lastCol = BitConverter.ToUInt16(record.Data, 2);
            if (firstCol > lastCol)
                return;
            ushort width = BitConverter.ToUInt16(record.Data, 4);
            ushort xfIndex = BitConverter.ToUInt16(record.Data, 6);
            ushort options = BitConverter.ToUInt16(record.Data, 8);
            bool hidden = (options & 0x0001) != 0;
            worksheet.ColumnInfos.Add(new ColumnInfo
            {
                FirstColumn = firstCol,
                LastColumn = lastCol,
                Width = width,
                XfIndex = xfIndex,
                Hidden = hidden
            });
        }
        
        private void ParseWindow2Record(BiffRecord record, Worksheet worksheet)
        {
            // WINDOW2 记录: options(2) + ...
            // options bit 3: 是否冻结窗格 (fFrozen)
            if (record.Data != null && record.Data.Length >= 2)
            {
                ushort options = BitConverter.ToUInt16(record.Data, 0);
                bool isFrozen = (options & 0x0008) != 0;
                
                if (isFrozen && worksheet.FreezePane == null)
                {
                    // 标记为冻结，具体位置由PANE记录设置
                    worksheet.FreezePane = new FreezePane();
                }
            }
        }
        
        private void ParsePaneRecord(BiffRecord record, Worksheet worksheet)
        {
            // PANE 记录: x(2) + y(2) + topRow(2) + leftCol(2) + activePane(1)
            if (record.Data != null && record.Data.Length >= 8)
            {
                ushort x = BitConverter.ToUInt16(record.Data, 0); // 水平分割位置（列数或像素）
                ushort y = BitConverter.ToUInt16(record.Data, 2); // 垂直分割位置（行数或像素）
                ushort topRow = BitConverter.ToUInt16(record.Data, 4);
                ushort leftCol = BitConverter.ToUInt16(record.Data, 6);
                
                if (worksheet.FreezePane != null)
                {
                    worksheet.FreezePane.ColSplit = x;
                    worksheet.FreezePane.RowSplit = y;
                    worksheet.FreezePane.TopRow = topRow + 1; // 转为1-based
                    worksheet.FreezePane.LeftCol = leftCol + 1; // 转为1-based
                }
                else
                {
                    worksheet.FreezePane = new FreezePane
                    {
                        ColSplit = x,
                        RowSplit = y,
                        TopRow = topRow + 1, // 转为1-based
                        LeftCol = leftCol + 1 // 转为1-based
                    };
                }
            }
        }
        
        private void ParseFontRecord(BiffRecord record, Worksheet worksheet)
        {
            byte[] data = record.GetAllData();
            if (data != null && data.Length >= 48)
            {
                var font = new Font();
                font.Height = BitConverter.ToInt16(data, 0);
                font.IsBold = (BitConverter.ToUInt16(data, 2) & 0x0001) != 0;
                font.IsItalic = (BitConverter.ToUInt16(data, 2) & 0x0002) != 0;
                font.IsUnderline = (BitConverter.ToUInt16(data, 2) & 0x0004) != 0;
                font.IsStrikethrough = (BitConverter.ToUInt16(data, 2) & 0x0008) != 0;
                font.ColorIndex = BitConverter.ToUInt16(data, 6);
                string? wsColor = GetColorFromPalette(font.ColorIndex);
                font.Color = string.IsNullOrEmpty(wsColor) ? null : wsColor.Replace("#", "");
                font.Name = System.Text.Encoding.ASCII.GetString(data, 40, data.Length - 40).TrimEnd('\0');

                worksheet.Fonts.Add(font);
            }
        }
        
        private void ParseXfRecord(BiffRecord record, Worksheet worksheet)
        {
            // 解析扩展格式记录
            if (record.Data != null && record.Data.Length >= 28)
            {
                var xf = new Xf();
                xf.FontIndex = BitConverter.ToUInt16(record.Data, 0);
                xf.NumberFormatIndex = BitConverter.ToUInt16(record.Data, 2);
                xf.CellFormatIndex = BitConverter.ToUInt16(record.Data, 4);
                
                // 解析对齐方式
                ushort alignment = BitConverter.ToUInt16(record.Data, 6);
                byte horizontalAlign = (byte)((alignment & 0x000F) >> 0);
                byte verticalAlign = (byte)((alignment & 0x00F0) >> 4);
                
                switch (horizontalAlign)
                {
                    case 0: xf.HorizontalAlignment = "general";
                        break;
                    case 1: xf.HorizontalAlignment = "left";
                        break;
                    case 2: xf.HorizontalAlignment = "center";
                        break;
                    case 3: xf.HorizontalAlignment = "right";
                        break;
                    case 4: xf.HorizontalAlignment = "fill";
                        break;
                    case 5: xf.HorizontalAlignment = "justify";
                        break;
                    case 6: xf.HorizontalAlignment = "centerContinuous";
                        break;
                    case 7: xf.HorizontalAlignment = "distributed";
                        break;
                }
                
                switch (verticalAlign)
                {
                    case 0: xf.VerticalAlignment = "top";
                        break;
                    case 1: xf.VerticalAlignment = "center";
                        break;
                    case 2: xf.VerticalAlignment = "bottom";
                        break;
                    case 3: xf.VerticalAlignment = "justify";
                        break;
                    case 4: xf.VerticalAlignment = "distributed";
                        break;
                }
                
                // 解析缩进
                xf.Indent = (byte)((alignment & 0x0F00) >> 8);
                
                // 解析文本换行
                xf.WrapText = (alignment & 0x1000) != 0;

                // 解析边框 (偏移10-17)
                if (record.Data.Length >= 18)
                {
                    uint border1 = BitConverter.ToUInt32(record.Data, 10);
                    uint border2 = BitConverter.ToUInt32(record.Data, 14);

                    var border = new Border();
                    border.Left = GetBorderLineStyle((byte)(border1 & 0x0F));
                    border.Right = GetBorderLineStyle((byte)((border1 >> 4) & 0x0F));
                    border.Top = GetBorderLineStyle((byte)((border1 >> 8) & 0x0F));
                    border.Bottom = GetBorderLineStyle((byte)((border1 >> 12) & 0x0F));

                    border.LeftColor = GetColorFromPalette((int)((border1 >> 16) & 0x7F));
                    border.RightColor = GetColorFromPalette((int)((border1 >> 23) & 0x7F));

                    border.TopColor = GetColorFromPalette((int)(border2 & 0x7F));
                    border.BottomColor = GetColorFromPalette((int)((border2 >> 7) & 0x7F));
                    border.DiagonalColor = GetColorFromPalette((int)((border2 >> 14) & 0x7F));
                    border.Diagonal = GetBorderLineStyle((byte)((border2 >> 21) & 0x0F));

                    // 添加到全局列表并分配索引
                    _workbook.Borders.Add(border);
                    xf.BorderIndex = _workbook.Borders.Count - 1;
                }

                // 解析填充 (偏移18-21)：低6位=pattern，次7位=icvFore，icvBack 在字节20
                if (record.Data.Length >= 22)
                {
                    ushort fillData = BitConverter.ToUInt16(record.Data, 18);
                    byte pattern = (byte)(fillData & 0x3F);
                    int icvFore = (fillData >> 6) & 0x7F;
                    int icvBack = record.Data.Length > 20 ? (record.Data[20] & 0x7F) : 65;

                    var fill = new Fill();
                    fill.PatternType = GetPatternType(pattern);
                    fill.ForegroundColor = GetColorFromPalette(icvFore);
                    fill.BackgroundColor = GetColorFromPalette(icvBack);

                    _workbook.Fills.Add(fill);
                    xf.FillIndex = _workbook.Fills.Count + 1; // 2-based: 0=none, 1=gray125, 2+=workbook.Fills
                }
                
                // 解析锁定和隐藏状态
                xf.IsLocked = (BitConverter.ToUInt16(record.Data, 26) & 0x0001) != 0;
                xf.IsHidden = (BitConverter.ToUInt16(record.Data, 26) & 0x0002) != 0;
                
                worksheet.Xfs.Add(xf);
            }
        }
        
        private void ParsePaletteRecord(BiffRecord record, Worksheet worksheet)
        {
            // BIFF8 PALETTE: 2 字节起始索引 + 每色 4 字节 (R,G,B,保留)
            if (record.Data == null || record.Data.Length < 6)
                return;
            int startIndex = BitConverter.ToUInt16(record.Data, 0);
            int colorCount = (record.Data.Length - 2) / 4;
            for (int i = 0; i < colorCount; i++)
            {
                int offset = 2 + i * 4;
                if (offset + 3 <= record.Data.Length)
                {
                    byte red = record.Data[offset];
                    byte green = record.Data[offset + 1];
                    byte blue = record.Data[offset + 2];
                    worksheet.Palette[startIndex + i] = $"#{red:X2}{green:X2}{blue:X2}";
                }
            }
        }
        

        private void ParseMSODrawingRecord(BiffRecord record, Worksheet worksheet)
        {
            // 收集绘图数据以供后续 OBJ 记录使用
            if (record.Data != null && record.Data.Length > 0)
            {
                _msoDrawingData.Add(record.GetAllData());
            }
        }

        private void ParseMsoDrawingGroupGlobal(BiffRecord record, Workbook workbook)
        {
            // 收集全局绘图数据
            if (record.Data != null && record.Data.Length > 0)
            {
                _msoDrawingGroupData.Add(record.GetAllData());
            }
        }
        
        private void ParseObjRecord(BiffRecord record, Worksheet worksheet)
        {
            // 解析对象记录，包括图表容器
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 4) return;
            
            // Obj Type is at offset 4 usually, but Obj record format is a series of subrecords (ftCmo=0x15)
            // subrecord ftCmo usually comes first: cmoId(2), cb(2), ot(2), id(2), options(2)
            ushort ft = BitConverter.ToUInt16(data, 0);
            if (ft == 0x0015 && data.Length >= 10) // ftCmo
            {
                ushort objType = BitConverter.ToUInt16(data, 4);
                if (objType == 0x0005) // Chart Data
                {
                    // 尝试从最近的 MsoDrawing 中读取 ClientAnchor
                    try
                    {
                        if (_msoDrawingData.Count > 0)
                        {
                            // 合并最后收集的一批 MsoDrawing 数据
                            int totalLen = _msoDrawingData.Sum(d => d.Length);
                            byte[] drawingData = new byte[totalLen];
                            int offset = 0;
                            foreach (var d in _msoDrawingData)
                            {
                                Array.Copy(d, 0, drawingData, offset, d.Length);
                                offset += d.Length;
                            }
                            
                            var escherRecords = Escher.EscherParser.ParseStream(drawingData);
                            // 在记录中查找 ClientAnchor (0xF010)
                            var anchor = FindClientAnchor(escherRecords);
                            if (anchor != null && anchor.Data != null && anchor.Data.Length >= 16)
                            {
                                // ClientAnchor 布局 (POI/MS-ODRAW): col1(2), dxL(2), row1(2), dyT(2), col2(2), dxR(2), row2(2), dyB(2)
                                ushort col1 = BitConverter.ToUInt16(anchor.Data, 0);
                                ushort dxL = BitConverter.ToUInt16(anchor.Data, 2);
                                ushort row1 = BitConverter.ToUInt16(anchor.Data, 4);
                                ushort dyT = BitConverter.ToUInt16(anchor.Data, 6);
                                ushort col2 = BitConverter.ToUInt16(anchor.Data, 8);
                                ushort dxR = BitConverter.ToUInt16(anchor.Data, 10);
                                ushort row2 = BitConverter.ToUInt16(anchor.Data, 12);
                                ushort dyB = BitConverter.ToUInt16(anchor.Data, 14);
                                int left = dxL;
                                int top = dyT;
                                // 单元格宽 1024 单位，高 256 单位
                                int width = Math.Max(1, (col2 - col1) * 1024 + (dxR - dxL));
                                int height = Math.Max(1, (row2 - row1) * 256 + (dyB - dyT));
                                _pendingChartAnchors.Add((left, top, width, height));
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Error($"解析 OBJ / Chart Anchor 失败: {ex.Message}", ex);
                    }
                    finally
                    {
                        _msoDrawingData.Clear(); // 清理，避免与其他对象混淆
                    }
                }
                else
                {
                    _msoDrawingData.Clear();
                    
                    // 原先的 ParseObjRecord 逻辑：作为普通的 EmbeddedObject 解析
                    var embeddedObject = new EmbeddedObject();
                    embeddedObject.Data = new byte[data.Length];
                    Array.Copy(data, embeddedObject.Data, data.Length);
                    embeddedObject.MimeType = "application/octet-stream";
                    worksheet.EmbeddedObjects.Add(embeddedObject);
                }
            }
        }

        private Escher.EscherRecord? FindClientAnchor(IEnumerable<Escher.EscherRecord> records)
        {
            foreach (var record in records)
            {
                if (record.Type == Escher.EscherParser.ClientAnchor)
                    return record;
                
                if (record.IsContainer && record.Children.Count > 0)
                {
                    var result = FindClientAnchor(record.Children);
                    if (result != null) return result;
                }
            }
            return null;
        }
        private void ParsePictureRecord(BiffRecord record, Worksheet worksheet)
        {
            // 解析图片记录（可能被 CONTINUE 分片，必须用完整数据）
            byte[] fullData = record.GetAllData();
            if (fullData == null || fullData.Length == 0)
                return;
            try
            {
                var picture = new Picture();
                picture.Data = new byte[fullData.Length];
                Array.Copy(fullData, picture.Data, fullData.Length);

                // 识别图片格式
                if (fullData.Length >= 4)
                {
                    if (fullData[0] == 0x89 && fullData[1] == 0x50 && fullData[2] == 0x4E && fullData[3] == 0x47)
                        {
                            // PNG格式
                            picture.MimeType = "image/png";
                            picture.Extension = "png";
                        }
                        else if (fullData[0] == 0xFF && fullData[1] == 0xD8)
                        {
                            picture.MimeType = "image/jpeg";
                            picture.Extension = "jpg";
                        }
                        else if (fullData[0] == 0x47 && fullData[1] == 0x49 && fullData[2] == 0x46)
                        {
                            picture.MimeType = "image/gif";
                            picture.Extension = "gif";
                        }
                        else if (fullData[0] == 0x42 && fullData[1] == 0x4D)
                        {
                            picture.MimeType = "image/bmp";
                            picture.Extension = "bmp";
                        }
                        else if (fullData[0] == 0x52 && fullData[1] == 0x49 && fullData[2] == 0x46 && fullData[3] == 0x46)
                        {
                            if (fullData.Length >= 12 && fullData[8] == 0x57 && fullData[9] == 0x45 && fullData[10] == 0x42 && fullData[11] == 0x50)
                            {
                                picture.MimeType = "image/webp";
                                picture.Extension = "webp";
                            }
                            else
                            {
                                // 默认为BMP格式
                                picture.MimeType = "image/bmp";
                                picture.Extension = "bmp";
                            }
                        }
                        else if (fullData[0] == 0x49 && fullData[1] == 0x49 && fullData[2] == 0x2A && fullData[3] == 0x00)
                        {
                            picture.MimeType = "image/tiff";
                            picture.Extension = "tiff";
                        }
                        else if (fullData[0] == 0x4D && fullData[1] == 0x4D && fullData[2] == 0x00 && fullData[3] == 0x2A)
                        {
                            picture.MimeType = "image/tiff";
                            picture.Extension = "tiff";
                        }
                        else if (fullData[0] == 0x38 && fullData[1] == 0x42 && fullData[2] == 0x50 && fullData[3] == 0x53)
                        {
                            picture.MimeType = "image/vnd.adobe.photoshop";
                            picture.Extension = "psd";
                        }
                        else if (fullData[0] == 0x52 && fullData[1] == 0x49 && fullData[2] == 0x46 && fullData[3] == 0x46 && fullData.Length >= 10)
                        {
                            picture.MimeType = "image/rtf";
                            picture.Extension = "rtf";
                        }
                        else if (fullData[0] == 0x49 && fullData[1] == 0x4D && fullData[2] == 0x47)
                        {
                            picture.MimeType = "image/x-ms-bmp";
                            picture.Extension = "img";
                        }
                        else if (fullData[0] == 0x43 && fullData[1] == 0x57 && fullData[2] == 0x53)
                        {
                            picture.MimeType = "application/x-cws";
                            picture.Extension = "cws";
                        }
                        else
                        {
                            picture.MimeType = "image/bmp";
                            picture.Extension = "bmp";
                        }
                }
                else
                {
                    picture.MimeType = "image/bmp";
                    picture.Extension = "bmp";
                }
                worksheet.Pictures.Add(picture);
            }
            catch (Exception ex)
            {
                throw new ImageProcessingException($"处理图片时发生错误: {ex.Message}", ex);
            }
        }

        
        private void ParseDVRecord(BiffRecord record, Worksheet worksheet)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 16)
                return;
            var dataValidation = new DataValidation();

            ushort options = BitConverter.ToUInt16(data, 0);
            dataValidation.AllowBlank = (options & 0x01) != 0;

            ushort validationType = BitConverter.ToUInt16(data, 2);
                switch (validationType)
                {
                    case 0: dataValidation.Type = "none"; break;
                    case 1: dataValidation.Type = "whole"; break;
                    case 2: dataValidation.Type = "decimal"; break;
                    case 3: dataValidation.Type = "list"; break;
                    case 4: dataValidation.Type = "date"; break;
                    case 5: dataValidation.Type = "time"; break;
                    case 6: dataValidation.Type = "textLength"; break;
                    case 7: dataValidation.Type = "custom"; break;
                }
                
            ushort operatorType = BitConverter.ToUInt16(data, 4);
            switch (operatorType)
            {
                case 0: dataValidation.Operator = "between"; break;
                case 1: dataValidation.Operator = "notBetween"; break;
                case 2: dataValidation.Operator = "equal"; break;
                case 3: dataValidation.Operator = "notEqual"; break;
                case 4: dataValidation.Operator = "greaterThan"; break;
                case 5: dataValidation.Operator = "lessThan"; break;
                case 6: dataValidation.Operator = "greaterThanOrEqual"; break;
                case 7: dataValidation.Operator = "lessThanOrEqual"; break;
            }

            int currentOffset = 6;
            ushort formula1Size = BitConverter.ToUInt16(data, currentOffset); currentOffset += 2;
            ushort formula2Size = BitConverter.ToUInt16(data, currentOffset); currentOffset += 2;

            if (formula1Size > 0 && currentOffset + formula1Size <= data.Length)
            {
                byte[] formula1Bytes = new byte[formula1Size];
                Array.Copy(data, currentOffset, formula1Bytes, 0, formula1Size);
                dataValidation.Formula1 = FormulaDecompiler.Decompile(formula1Bytes);
                currentOffset += formula1Size;
            }
            if (formula2Size > 0 && currentOffset + formula2Size <= data.Length)
            {
                byte[] formula2Bytes = new byte[formula2Size];
                Array.Copy(data, currentOffset, formula2Bytes, 0, formula2Size);
                dataValidation.Formula2 = FormulaDecompiler.Decompile(formula2Bytes);
                currentOffset += formula2Size;
            }
            if (currentOffset + 8 <= data.Length)
            {
                ushort firstRow = BitConverter.ToUInt16(data, currentOffset); currentOffset += 2;
                ushort lastRow = BitConverter.ToUInt16(data, currentOffset); currentOffset += 2;
                ushort firstCol = BitConverter.ToUInt16(data, currentOffset); currentOffset += 2;
                ushort lastCol = BitConverter.ToUInt16(data, currentOffset);
                dataValidation.Range = $"{GetColumnLetter(firstCol)}{firstRow + 1}:{GetColumnLetter(lastCol)}{lastRow + 1}";
            }
            else
            {
                dataValidation.Range = "A1:A10";
            }
            worksheet.DataValidations.Add(dataValidation);
        }
        
        private void ParseCFRecord(BiffRecord record, Worksheet worksheet)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 8)
                return;
            var conditionalFormat = new ConditionalFormat();

            ushort conditionType = BitConverter.ToUInt16(data, 0);
                switch (conditionType)
                {
                    case 0: conditionalFormat.Type = "cellIs"; break;
                    case 1: conditionalFormat.Type = "expression"; break;
                    case 2: conditionalFormat.Type = "colorScale"; break;
                    case 3: conditionalFormat.Type = "dataBar"; break;
                    case 4: conditionalFormat.Type = "iconSet"; break;
                }

            ushort operatorType = BitConverter.ToUInt16(data, 2);
            switch (operatorType)
            {
                case 0: conditionalFormat.Operator = "between"; break;
                case 1: conditionalFormat.Operator = "notBetween"; break;
                case 2: conditionalFormat.Operator = "equal"; break;
                case 3: conditionalFormat.Operator = "notEqual"; break;
                case 4: conditionalFormat.Operator = "greaterThan"; break;
                case 5: conditionalFormat.Operator = "lessThan"; break;
                case 6: conditionalFormat.Operator = "greaterThanOrEqual"; break;
                case 7: conditionalFormat.Operator = "lessThanOrEqual"; break;
                case 8: conditionalFormat.Operator = "containsText"; break;
                case 9: conditionalFormat.Operator = "notContainsText"; break;
                case 10: conditionalFormat.Operator = "beginsWith"; break;
                case 11: conditionalFormat.Operator = "endsWith"; break;
            }
            if (data.Length >= 12)
            {
                int currentOffset = 4;
                ushort formula1Size = BitConverter.ToUInt16(data, currentOffset);
                currentOffset += 2;
                ushort formula2Size = BitConverter.ToUInt16(data, currentOffset);
                currentOffset += 2;
                if (currentOffset + 4 <= data.Length)
                {
                    currentOffset += 4;
                    if (formula1Size > 0 && currentOffset + formula1Size <= data.Length)
                    {
                        byte[] ptg1 = new byte[formula1Size];
                        Array.Copy(data, currentOffset, ptg1, 0, formula1Size);
                        conditionalFormat.Formula = FormulaDecompiler.Decompile(ptg1);
                        currentOffset += formula1Size;
                    }
                    if (formula2Size > 0 && currentOffset + formula2Size <= data.Length)
                    {
                        byte[] ptg2 = new byte[formula2Size];
                        Array.Copy(data, currentOffset, ptg2, 0, formula2Size);
                    }
                }
            }
            conditionalFormat.Range = !string.IsNullOrEmpty(_currentCFRange) ? _currentCFRange : "A1:A10";
            worksheet.ConditionalFormats.Add(conditionalFormat);
        }
        
        private void ParseCFHeaderRecord(BiffRecord record, Worksheet worksheet)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 12)
                return;
            ushort conditionCount = BitConverter.ToUInt16(data, 0);
            ushort firstRow = BitConverter.ToUInt16(data, 2);
            ushort lastRow = BitConverter.ToUInt16(data, 4);
            ushort firstCol = BitConverter.ToUInt16(data, 6);
            ushort lastCol = BitConverter.ToUInt16(data, 8);
            _currentCFRange = $"{GetColumnLetter(firstCol)}{firstRow + 1}:{GetColumnLetter(lastCol)}{lastRow + 1}";
        }

        private void ParseHyperlinkRecord(BiffRecord record, Worksheet worksheet)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 20)
                return;
            var hyperlink = new Hyperlink();
            ushort firstRow = BitConverter.ToUInt16(data, 0);
            ushort lastRow = BitConverter.ToUInt16(data, 2);
            ushort firstCol = BitConverter.ToUInt16(data, 4);
            ushort lastCol = BitConverter.ToUInt16(data, 6);
            hyperlink.Range = $"{GetColumnLetter(firstCol)}{firstRow + 1}:{GetColumnLetter(lastCol)}{lastRow + 1}";
            int urlLength = BitConverter.ToInt16(data, 18);
            if (urlLength > 0 && data.Length >= 20 + urlLength)
                hyperlink.Target = System.Text.Encoding.ASCII.GetString(data, 20, urlLength);
            else if (urlLength > 0 && data.Length > 20)
            {
                int safeLen = Math.Min(urlLength, data.Length - 20);
                hyperlink.Target = System.Text.Encoding.ASCII.GetString(data, 20, safeLen);
            }
            worksheet.Hyperlinks.Add(hyperlink);
        }

        private void ParseCommentRecord(BiffRecord record, Worksheet worksheet)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 12)
                return;
            var comment = new Comment();
            ushort row = BitConverter.ToUInt16(data, 0);
            ushort col = BitConverter.ToUInt16(data, 2);
            comment.RowIndex = row + 1;
            comment.ColumnIndex = col + 1;
            if (data.Length >= 14)
            {
                byte authorLength = data[12];
                if (authorLength > 0 && data.Length >= 13 + authorLength)
                {
                    comment.Author = System.Text.Encoding.ASCII.GetString(data, 13, authorLength);
                    int textOffset = 13 + authorLength;
                    if (data.Length > textOffset)
                        comment.Text = System.Text.Encoding.ASCII.GetString(data, textOffset, data.Length - textOffset);
                }
            }
            worksheet.Comments.Add(comment);
        }

        private string GetColumnLetter(int columnIndex)
        {
            if (columnIndex <= 0) return "A";
            var columnReference = string.Empty;
            int col = columnIndex;
            while (col > 0)
            {
                col--;
                columnReference = (char)('A' + col % 26) + columnReference;
                col /= 26;
            }
            return columnReference;
        }

        private List<RichTextRun> ParseRichText(byte[] data, int offset)
        {
            var richTextRuns = new List<RichTextRun>();
            int currentOffset = offset;
            
            try
            {
                // 读取富文本记录的结构
                // 根据BIFF8格式规范解析富文本数据
                while (currentOffset < data.Length)
                {
                    // 读取文本长度
                    if (currentOffset + 2 <= data.Length)
                    {
                        short textLength = BitConverter.ToInt16(data, currentOffset);
                        currentOffset += 2;
                        
                        // 读取文本类型（0=ASCII, 1=Unicode）
                        byte textType = 0;
                        if (currentOffset < data.Length)
                        {
                            textType = data[currentOffset];
                            currentOffset += 1;
                        }
                        
                        // 读取文本内容
                        string text;
                        if (textType == 0)
                        {
                            // ASCII字符串
                            if (currentOffset + textLength <= data.Length)
                            {
                                text = System.Text.Encoding.ASCII.GetString(data, currentOffset, textLength);
                                currentOffset += textLength;
                            }
                            else
                            {
                                break;
                            }
                        }
                        else
                        {
                            // Unicode字符串
                            if (currentOffset + textLength * 2 <= data.Length)
                            {
                                text = System.Text.Encoding.Unicode.GetString(data, currentOffset, textLength * 2);
                                currentOffset += textLength * 2;
                            }
                            else
                            {
                                break;
                            }
                        }
                        
                        // 读取字体索引
                        short fontIndex = 0;
                        if (currentOffset + 2 <= data.Length)
                        {
                            fontIndex = BitConverter.ToInt16(data, currentOffset);
                            currentOffset += 2;
                        }
                        
                        // 创建富文本运行
                        var run = new RichTextRun
                        {
                            Text = text,
                            // 根据fontIndex获取对应的字体信息
                            Font = GetFontByIndex(fontIndex)
                        };
                        richTextRuns.Add(run);
                    }
                    else
                    {
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("解析富文本时发生错误", ex);
                // 继续处理，返回已解析的部分
            }
            
            return richTextRuns;
        }

        private Font GetFontByIndex(short fontIndex)
        {
            // 根据字体索引获取字体信息
            try
            {
                // 检查字体索引是否有效
                if (fontIndex >= 0 && fontIndex < _fonts.Count)
                {
                    return _fonts[fontIndex];
                }
                else
                {
                    // 返回默认字体
                    return new Font
                    {
                        Name = "Arial",
                        Size = 11,
                        Bold = false,
                        Italic = false,
                        Underline = false,
                        Color = "000000"
                    };
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"获取字体信息时发生错误，字体索引: {fontIndex}", ex);
                // 返回默认字体
                return new Font
                {
                    Name = "Arial",
                    Size = 11,
                    Bold = false,
                    Italic = false,
                    Underline = false,
                    Color = "000000"
                };
            }
        }

        private void ParseSheetRecord(BiffRecord record, Workbook workbook)
        {
            var worksheet = new Worksheet();
            byte[] data = record.GetAllData();
            if (data != null && data.Length >= 8)
            {
                int lbPlyPos = BitConverter.ToInt32(data, 0);
                _sheetOffsets.Add(lbPlyPos);
                Logger.Debug($"BOUNDSHEET: lbPlyPos={lbPlyPos}");
                int nameOffset = 6;
                if (data.Length > nameOffset)
                {
                    byte len = data[nameOffset];
                    int pos = nameOffset + 1;
                    worksheet.Name = ReadBiffStringFromBytes(data, ref pos, len);
                }
            }
            else if (data != null && data.Length >= 4)
            {
                int lbPlyPos = BitConverter.ToInt32(data, 0);
                _sheetOffsets.Add(lbPlyPos);
            }
            if (string.IsNullOrEmpty(worksheet.Name))
                worksheet.Name = "Sheet" + (workbook.Worksheets.Count + 1);
            workbook.Worksheets.Add(worksheet);
        }

        private void ParseSstInfo(BiffRecord record, long workbookStreamEnd)
        {
            if (record.Data == null || record.Data.Length < 8) return;

            int uniqueCount = BitConverter.ToInt32(record.Data, 4);
            // 防止损坏文件中的异常计数值导致超大分配或死循环（Excel 实际限制约 65535 唯一字符串/工作簿）
            const int maxUniqueCount = 2 * 1024 * 1024;
            if (uniqueCount < 0 || uniqueCount > maxUniqueCount)
                uniqueCount = Math.Clamp(uniqueCount, 0, maxUniqueCount);
            _sharedStrings.Capacity = Math.Max(_sharedStrings.Capacity, uniqueCount);

            var stringReader = new BiffStringReader(record, 8); // SST Header size is 8 bytes

            for (int i = 0; i < uniqueCount; i++)
            {
                string str = stringReader.ReadString();
                // Depending on file corruption or incorrect counts, the reader might return empty at EOF 
                // We add it anyway to maintain the index structure, as cells refer to indexes.
                _sharedStrings.Add(str);
            }
        }

        private string ReadBiffString(byte[] data, ref int offset)
        {
            if (offset + 2 > data.Length) return string.Empty;
            ushort charCount = BitConverter.ToUInt16(data, offset);
            offset += 2;
            if (offset >= data.Length) return string.Empty;
            
            byte option = data[offset];
            offset += 1;
            
            bool isUnicode = (option & 0x01) != 0;
            bool hasRichText = (option & 0x08) != 0;
            bool hasExtended = (option & 0x04) != 0;
            
            int runs = 0;
            if (hasRichText)
            {
                if (offset + 2 > data.Length) return string.Empty;
                runs = BitConverter.ToUInt16(data, offset);
                offset += 2;
            }
            
            int extendedSize = 0;
            if (hasExtended)
            {
                if (offset + 4 > data.Length) return string.Empty;
                extendedSize = BitConverter.ToInt32(data, offset);
                offset += 4; // Skip the size header
            }

            string result;
            if (isUnicode)
            {
                int byteCount = charCount * 2;
                if (offset + byteCount > data.Length) byteCount = data.Length - offset;
                result = System.Text.Encoding.Unicode.GetString(data, offset, byteCount);
                offset += byteCount;
            }
            else
            {
                int byteCount = charCount;
                if (offset + byteCount > data.Length) byteCount = data.Length - offset;
                result = System.Text.Encoding.ASCII.GetString(data, offset, byteCount);
                offset += byteCount;
            }

            if (hasRichText)
            {
                offset += runs * 4; // Skip formatting runs
            }
            
            if (hasExtended)
            {
                offset += extendedSize; // Skip the phonetic string data payload
            }
            
            return result;
        }

        private string ReadBiffStringFromBytes(byte[] data, ref int offset, int charCount)
        {
            if (offset >= data.Length) return string.Empty;
            byte option = data[offset];
            offset += 1;
            bool isUnicode = (option & 0x01) != 0;
            string result;
            if (isUnicode)
            {
                int byteCount = charCount * 2;
                if (offset + byteCount > data.Length) byteCount = data.Length - offset;
                result = System.Text.Encoding.Unicode.GetString(data, offset, byteCount);
                offset += byteCount;
            }
            else
            {
                int byteCount = charCount;
                if (offset + byteCount > data.Length) byteCount = data.Length - offset;
                result = System.Text.Encoding.ASCII.GetString(data, offset, byteCount);
                offset += byteCount;
            }
            return result;
        }

        private void ParseFontRecordToGlobal(BiffRecord record)
        {
            byte[] data = record.GetAllData();
            if (data != null && data.Length >= 14)
            {
                var font = new Font();
                font.Height = BitConverter.ToInt16(data, 0);
                ushort grbit = BitConverter.ToUInt16(data, 2);
                font.IsBold = BitConverter.ToUInt16(data, 6) >= 700;
                font.IsItalic = (grbit & 0x0002) != 0;
                font.IsUnderline = (data[10]) != 0;
                font.IsStrikethrough = (grbit & 0x0008) != 0;
                font.ColorIndex = BitConverter.ToUInt16(data, 4);
                string? resolved = GetColorFromPalette(font.ColorIndex);
                font.Color = string.IsNullOrEmpty(resolved) ? null : resolved.Replace("#", "");

                int nameOffset = 14;
                if (data.Length > nameOffset)
                {
                    byte len = data[nameOffset];
                    if (data.Length > nameOffset + 1)
                    {
                        byte opt = data[nameOffset + 1];
                        bool isUni = (opt & 0x01) != 0;
                        if (isUni)
                        {
                           font.Name = System.Text.Encoding.Unicode.GetString(data, nameOffset + 2, Math.Min(len * 2, data.Length - nameOffset - 2));
                        }
                        else
                        {
                           font.Name = System.Text.Encoding.ASCII.GetString(data, nameOffset + 2, Math.Min(len, data.Length - nameOffset - 2));
                        }
                    }
                }
                _fonts.Add(font);
            }
        }

        private void ParseXfRecordToGlobal(BiffRecord record)
        {
            if (record.Data != null && record.Data.Length >= 20)
            {
                var xf = new Xf();
                xf.FontIndex = BitConverter.ToUInt16(record.Data, 0);
                xf.NumberFormatIndex = BitConverter.ToUInt16(record.Data, 2);
                
                // 解析对齐方式 (offset 6-9)
                ushort alignment = BitConverter.ToUInt16(record.Data, 6);
                byte horizontalAlign = (byte)(alignment & 0x07);
                byte verticalAlign = (byte)((alignment & 0x70) >> 4);
                
                xf.HorizontalAlignment = horizontalAlign switch {
                    1 => "left", 2 => "center", 3 => "right", 4 => "fill", 5 => "justify", 6 => "centerContinuous", 7 => "distributed", _ => "general"
                };
                xf.VerticalAlignment = verticalAlign switch {
                    1 => "center", 2 => "bottom", 3 => "justify", 4 => "distributed", _ => "top"
                };
                
                xf.WrapText = (alignment & 0x08) != 0;
                xf.Indent = (byte)((alignment >> 8) & 0x0F);

                // 解析边框 (偏移10-17)
                if (record.Data.Length >= 18)
                {
                    uint border1 = BitConverter.ToUInt32(record.Data, 10);
                    uint border2 = BitConverter.ToUInt32(record.Data, 14);

                    var border = new Border {
                        Left = GetBorderLineStyle((byte)(border1 & 0x0F)),
                        Right = GetBorderLineStyle((byte)((border1 >> 4) & 0x0F)),
                        Top = GetBorderLineStyle((byte)((border1 >> 8) & 0x0F)),
                        Bottom = GetBorderLineStyle((byte)((border1 >> 12) & 0x0F)),
                        LeftColor = GetColorFromPalette((int)((border1 >> 16) & 0x7F)),
                        RightColor = GetColorFromPalette((int)((border1 >> 23) & 0x7F)),
                        TopColor = GetColorFromPalette((int)(border2 & 0x7F)),
                        BottomColor = GetColorFromPalette((int)((border2 >> 7) & 0x7F)),
                        DiagonalColor = GetColorFromPalette((int)((border2 >> 14) & 0x7F)),
                        Diagonal = GetBorderLineStyle((byte)((border2 >> 21) & 0x0F))
                    };

                    // 只有当边框不是全部为 none 时才添加或查找
                    if (border.Left != "none" || border.Right != "none" || border.Top != "none" || border.Bottom != "none" || border.Diagonal != "none")
                    {
                        int existingBorderIdx = _workbook.Borders.FindIndex(b => 
                            b.Left == border.Left && b.Right == border.Right && b.Top == border.Top && b.Bottom == border.Bottom && 
                            b.LeftColor == border.LeftColor && b.RightColor == border.RightColor && b.TopColor == border.TopColor && b.BottomColor == border.BottomColor);
                        
                        if (existingBorderIdx >= 0)
                        {
                            xf.BorderIndex = existingBorderIdx + 1; // 0 是默认
                        }
                        else
                        {
                            _workbook.Borders.Add(border);
                            xf.BorderIndex = _workbook.Borders.Count;
                        }
                    }
                    else
                    {
                        xf.BorderIndex = 0;
                    }
                }

                // 解析填充 (偏移18-21)：低6位=pattern，次7位=icvFore，icvBack 在字节20
                if (record.Data.Length >= 20)
                {
                    ushort fillData = BitConverter.ToUInt16(record.Data, 18);
                    byte pattern = (byte)(fillData & 0x3F);
                    int icvFore = (fillData >> 6) & 0x7F;
                    int icvBack = record.Data.Length > 20 ? (record.Data[20] & 0x7F) : 65;

                    if (pattern > 0)
                    {
                        var fill = new Fill
                        {
                            PatternType = GetPatternType(pattern),
                            ForegroundColor = GetColorFromPalette(icvFore),
                            BackgroundColor = GetColorFromPalette(icvBack)
                        };
                        int existingFillIdx = _workbook.Fills.FindIndex(f =>
                            f.PatternType == fill.PatternType &&
                            f.ForegroundColor == fill.ForegroundColor &&
                            f.BackgroundColor == fill.BackgroundColor);

                        if (existingFillIdx >= 0)
                        {
                            xf.FillIndex = existingFillIdx + 2; // 0 和 1 是默认
                        }
                        else
                        {
                            _workbook.Fills.Add(fill);
                            xf.FillIndex = _workbook.Fills.Count + 1;
                        }
                    }
                    else
                    {
                        xf.FillIndex = 0;
                    }
                }
                
                // 解析锁定和隐藏状态 (偏移26)
                if (record.Data.Length >= 28)
                {
                    ushort options = BitConverter.ToUInt16(record.Data, 26);
                    xf.IsLocked = (options & 0x0001) != 0;
                    xf.IsHidden = (options & 0x0002) != 0;
                }


                _xfList.Add(xf);
            }
        }

        private void ParseFormatRecordGlobal(BiffRecord record)
        {
            byte[] data = record.GetAllData();
            if (data.Length >= 2)
            {
                ushort index = BitConverter.ToUInt16(data, 0);
                int offset = 2;
                if (offset < data.Length)
                {
                    _formats[index] = ReadBiffString(data, ref offset);
                }
            }
        }

        private void ParsePaletteRecordGlobal(BiffRecord record)
        {
            if (record.Data != null && record.Data.Length >= 4)
            {
                int count = BitConverter.ToUInt16(record.Data, 0);
                for (int i = 0; i < count && (2 + i * 4 + 4 <= record.Data.Length); i++)
                {
                    byte r = record.Data[2 + i * 4];
                    byte g = record.Data[2 + i * 4 + 1];
                    byte b = record.Data[2 + i * 4 + 2];
                    _palette[8 + i] = $"#{r:X2}{g:X2}{b:X2}";
                }
            }
        }

        /// <summary>解析 BIFF8 BORDER (0x00B2) 全局边框记录，与 XF 内嵌边框布局一致：8 字节 border1(4)+border2(4)。</summary>
        private void ParseBorderRecord(BiffRecord record, Workbook workbook)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 8) return;
            uint border1 = BitConverter.ToUInt32(data, 0);
            uint border2 = BitConverter.ToUInt32(data, 4);
            var border = new Border
            {
                Left = GetBorderLineStyle((byte)(border1 & 0x0F)),
                Right = GetBorderLineStyle((byte)((border1 >> 4) & 0x0F)),
                Top = GetBorderLineStyle((byte)((border1 >> 8) & 0x0F)),
                Bottom = GetBorderLineStyle((byte)((border1 >> 12) & 0x0F)),
                LeftColor = GetColorFromPalette((int)((border1 >> 16) & 0x7F)),
                RightColor = GetColorFromPalette((int)((border1 >> 23) & 0x7F)),
                TopColor = GetColorFromPalette((int)(border2 & 0x7F)),
                BottomColor = GetColorFromPalette((int)((border2 >> 7) & 0x7F)),
                DiagonalColor = GetColorFromPalette((int)((border2 >> 14) & 0x7F)),
                Diagonal = GetBorderLineStyle((byte)((border2 >> 21) & 0x0F))
            };
            workbook.Borders.Add(border);
        }

        /// <summary>解析 BIFF8 FILL (0x00F5) 全局填充记录，与 XF 内嵌填充一致：低6位=pattern，次7位=icvFore，下一字节=icvBack。</summary>
        private void ParseFillRecord(BiffRecord record, Workbook workbook)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 4) return;
            ushort fillData = BitConverter.ToUInt16(data, 0);
            byte pattern = (byte)(fillData & 0x3F);
            int icvFore = (fillData >> 6) & 0x7F;
            int icvBack = data.Length > 2 ? (data[2] & 0x7F) : 65;
            if (pattern == 0) return;
            var fill = new Fill
            {
                PatternType = GetPatternType(pattern),
                ForegroundColor = GetColorFromPalette(icvFore),
                BackgroundColor = GetColorFromPalette(icvBack)
            };
            workbook.Fills.Add(fill);
        }

        private Row ParseRowRecord(BiffRecord record)
        {
            var row = new Row();
            row.Cells.Capacity = 100;
            
            if (record.Data != null && record.Data.Length >= 2)
            {
                row.RowIndex = BitConverter.ToUInt16(record.Data, 0) + 1; // 转为1-based
                
                // BIFF8 ROW 记录格式: row(2) + firstCol(2) + lastCol(2) + height(2) + ...
                if (record.Data.Length >= 8)
                {
                    ushort rawHeight = BitConverter.ToUInt16(record.Data, 6);
                    // BIFF8 spec: bit 15 of miHeight is fGhost (1 = default height, 0 = custom height)
                    row.Height = (ushort)(rawHeight & 0x7FFF);
                    row.CustomHeight = (rawHeight & 0x8000) == 0;

                    // 偏移 16: ixfe 行默认 XF 索引（整行背景色等）
                    if (record.Data.Length >= 18)
                    {
                        ushort ixfe = BitConverter.ToUInt16(record.Data, 16);
                        row.DefaultXfIndex = ixfe;
                    }

                    // Option flags at offset 12
                    if (record.Data.Length >= 14)
                    {
                        ushort options = BitConverter.ToUInt16(record.Data, 12);
                        // bit 6: fDyZero (hidden row)
                        if ((options & 0x0040) != 0) 
                        {
                            row.CustomHeight = true;
                            row.Height = 0;
                        }
                    }
                }
            }
            else
            {
                row.RowIndex = 1;
            }
            
            return row;
        }

        private Cell ParseCellRecord(BiffRecord record)
        {
            var cell = new Cell();
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 6)
                return cell;

            // 读取行索引
            ushort rowIndex = BitConverter.ToUInt16(data, 0);
            cell.RowIndex = rowIndex + 1; // 转为1-based
            
            // 读取列索引（从1开始）
            ushort colIndex = BitConverter.ToUInt16(data, 2);
            cell.ColumnIndex = colIndex + 1;
            
            // 读取样式索引 (XF index)，0 也要设置，否则第一行用 XF 0 时不会写出 s="0"
            ushort styleIndex = BitConverter.ToUInt16(data, 4);
            cell.StyleId = styleIndex.ToString();
            
            // 根据记录类型解析单元格值
            switch (record.Id)
                {
                    case (ushort)BiffRecordType.CELL_BLANK:
                        // 空单元格
                        cell.Value = null;
                        break;
                    case (ushort)BiffRecordType.CELL_BOOLERR:
                        // BOOLERR 记录：根据标志字节区分布尔值和错误值
                        if (data.Length >= 8)
                        {
                            byte valueOrError = data[6];
                            byte isError = data[7]; // 0 = 布尔值, 1 = 错误值
                            if (isError == 0)
                            {
                                cell.Value = valueOrError != 0;
                                cell.DataType = "b";
                            }
                            else
                            {
                                cell.Value = GetErrorString(valueOrError);
                                cell.DataType = "e";
                            }
                        }
                        break;
                    case (ushort)BiffRecordType.CELL_LABEL:
                        // 文本值 (BIFF8 LABEL record uses XLUnicodeString)
                        if (data.Length > 6)
                        {
                            int offset = 6;
                            cell.Value = ReadBiffString(data, ref offset);
                            cell.DataType = "inlineStr";
                        }
                        break;
                    case (ushort)BiffRecordType.CELL_RSTRING:
                        // 旧版带格式文本或富文本
                        if (data.Length > 8)
                        {
                            // cch (2 bytes) + grbit (1 byte) ...
                            int offset = 6;
                            cell.Value = ReadBiffString(data, ref offset);
                            cell.DataType = "inlineStr";
                        }
                        break;
                    case (ushort)BiffRecordType.CELL_RICH_TEXT:
                        // 富文本值
                        if (data.Length > 6)
                        {
                            // 解析富文本格式（使用完整 data 以支持 CONTINUE）
                            cell.RichText = ParseRichText(data, 6);
                            cell.DataType = "inlineStr";
                            // 同时设置Value为纯文本，确保兼容性
                            cell.Value = string.Join("", cell.RichText.Select(r => r.Text));
                        }
                        break;
                    case (ushort)BiffRecordType.CELL_LABELSST:
                        // 共享字符串表中的索引
                        if (data.Length >= 10)
                        {
                            int sstIndex = BitConverter.ToInt32(data, 6);
                            if (sstIndex >= 0 && sstIndex < _sharedStrings.Count)
                            {
                                cell.Value = _sharedStrings[sstIndex];
                            }
                            cell.DataType = "s";
                        }
                        break;
                    case (ushort)BiffRecordType.CELL_NUMBER:
                        // 数值
                        if (data.Length >= 14)
                        {
                            double value = BitConverter.ToDouble(data, 6);
                            // 检查是否为日期时间值（Excel 日期时间是从 1900-01-01 开始的天数）
                            if (IsDateTimeValue(value))
                            {
                                cell.Value = ExcelDateToDateTime(value);
                                cell.DataType = "d";
                            }
                            else
                            {
                                cell.Value = value;
                                cell.DataType = "n";
                            }
                        }
                        break;
                    case (ushort)BiffRecordType.CELL_RK:
                        // 压缩数值
                        if (data.Length >= 10)
                        {
                            int rkValue = BitConverter.ToInt32(data, 6);
                            double value = DecodeRKValue(rkValue);
                            // 检查是否为日期时间值
                            if (IsDateTimeValue(value))
                            {
                                cell.Value = ExcelDateToDateTime(value);
                                cell.DataType = "d";
                            }
                            else
                            {
                                cell.Value = value;
                                cell.DataType = "n";
                            }
                        }
                        break;
                    case (ushort)BiffRecordType.CELL_FORMULA:
                        // 公式
                        try
                        {
                            if (data.Length >= 22)
                            {
                                // 读取公式结果
                                double result = BitConverter.ToDouble(data, 6);
                                
                                // 读取公式 Ptgs 长度（BIFF8 FORMULA 中通常在偏移 20）
                                int formulaLength = BitConverter.ToUInt16(data, 20);
                                
                                // 读取公式 Ptgs（使用 data 以支持 CONTINUE 分片）
                                if (data.Length >= 22 + formulaLength)
                                {
                                    byte[] ptgs = new byte[formulaLength];
                                    Array.Copy(data, 22, ptgs, 0, formulaLength);
                                    string formula = FormulaDecompiler.Decompile(ptgs);
                                    
                                    cell.Formula = formula;
                                    
                                    // 处理公式结果
                                    if (IsDateTimeValue(result))
                                    {
                                        cell.Value = ExcelDateToDateTime(result);
                                        cell.DataType = "d";
                                    }
                                    else if (IsErrorValue(result))
                                    {
                                        cell.Value = GetErrorString((byte)result);
                                        cell.DataType = "e";
                                    }
                                    else
                                    {
                                        cell.Value = result;
                                        cell.DataType = "f";
                                    }
                                }
                                else
                                {
                                    // 如果无法读取公式字符串，使用结果值
                                    if (IsDateTimeValue(result))
                                    {
                                        cell.Value = ExcelDateToDateTime(result);
                                        cell.DataType = "d";
                                    }
                                    else if (IsErrorValue(result))
                                    {
                                        cell.Value = GetErrorString((byte)result);
                                        cell.DataType = "e";
                                    }
                                    else
                                    {
                                        cell.Value = result;
                                        cell.DataType = "n";
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Error("解析公式时发生错误", ex);
                            // 继续处理，设置默认值
                            cell.Value = "#ERROR!";
                            cell.DataType = "e";
                        }
                        break;
                }
            
            return cell;
        }
        
        private string GetErrorString(byte errorCode)
        {
            // 映射错误代码到错误字符串
            switch (errorCode)
            {
                case 0x00: return "#NULL!";
                case 0x07: return "#DIV/0!";
                case 0x0F: return "#VALUE!";
                case 0x17: return "#REF!";
                case 0x1D: return "#NAME?";
                case 0x24: return "#NUM!";
                case 0x2A: return "#N/A";
                default: return "#ERROR!";
            }
        }
        
        private double DecodeRKValue(int rkValue)
        {
            // 解码RK值（压缩数值）
            // bit 0: 1 = 除以100, 0 = 不变
            // bit 1: 1 = 30位有符号整数, 0 = IEEE 754 double的前30位
            double value;
            if ((rkValue & 0x02) != 0)
            {
                // 30位整数
                value = (double)(rkValue >> 2);
            }
            else
            {
                // IEEE double 的前30位
                long bits = (long)(rkValue & 0xFFFFFFFC) << 32;
                value = BitConverter.Int64BitsToDouble(bits);
            }

            if ((rkValue & 0x01) != 0)
            {
                value /= 100.0;
            }

            return value;
        }
        
        private bool IsDateTimeValue(double value)
        {
            // 检查是否为日期时间值
            // Excel 日期时间范围通常在 25569（1970-01-01）到 44197（2020-12-31）之间
            // 扩展范围以覆盖更多可能的日期
            return value >= 25569 && value <= 730485; // 从1970-01-01到9999-12-31
        }

        private bool IsErrorValue(double value)
        {
            // 判断是否为错误值（Excel错误值通常是特殊的整数值）
            int intValue = (int)value;
            return intValue >= 0 && intValue <= 0x2A && 
                   (intValue == 0x00 || // #NULL!
                    intValue == 0x07 || // #DIV/0!
                    intValue == 0x0F || // #VALUE!
                    intValue == 0x17 || // #REF!
                    intValue == 0x1D || // #NAME?
                    intValue == 0x24 || // #NUM!
                    intValue == 0x2A);  // #N/A
        }
        
        private DateTime ExcelDateToDateTime(double excelDate)
        {
            // 将 Excel 日期时间值转换为 .NET DateTime
            // Excel 日期时间是从 1900-01-01 开始的天数
            // 注意：Excel 使用 1900 年 2 月 29 日作为有效日期，即使 1900 年不是闰年
            DateTime excelBaseDate = new DateTime(1900, 1, 1);
            
            // 调整 1900 年闰年问题
            if (excelDate >= 60)
            {
                excelDate -= 1;
            }
            
            return excelBaseDate.AddDays(excelDate);
        }

        /// <summary>解析 BIFF8 ARRAY (0x0221)：紧跟 FORMULA，前 8 字节为 firstRow(2), lastRow(2), firstCol(2), lastCol(2)，用于标记上一单元格为数组公式并设置范围。</summary>
        private void ParseArrayRecord(BiffRecord record, ref Row? currentRow)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 8) return;
            if (currentRow == null || currentRow.Cells.Count == 0) return;
            ushort firstRow = BitConverter.ToUInt16(data, 0);
            ushort lastRow = BitConverter.ToUInt16(data, 2);
            ushort firstCol = BitConverter.ToUInt16(data, 4);
            ushort lastCol = BitConverter.ToUInt16(data, 6);
            if (firstRow > lastRow || firstCol > lastCol) return;
            var cell = currentRow.Cells[currentRow.Cells.Count - 1];
            cell.IsArrayFormula = true;
            cell.ArrayRef = $"{GetColumnLetter(firstCol)}{firstRow + 1}:{GetColumnLetter(lastCol)}{lastRow + 1}";
        }

        /// <summary>解析 BIFF8 SHAREDFMLA (0x04BC)：前 8 字节为 firstRow, firstCol, lastRow, lastCol，随后为公式 token 数组，用于后续 FORMULA 按行列差展开。</summary>
        private void ParseSharedFmlaRecord(BiffRecord record)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 10) return;
            _sharedFormulaBaseRow = BitConverter.ToUInt16(data, 0);
            _sharedFormulaBaseCol = BitConverter.ToUInt16(data, 2);
            _sharedFormulaLastRow = BitConverter.ToUInt16(data, 4);
            _sharedFormulaLastCol = BitConverter.ToUInt16(data, 6);
            if (_sharedFormulaBaseRow > _sharedFormulaLastRow || _sharedFormulaBaseCol > _sharedFormulaLastCol) return;
            byte[] formulaTokens = new byte[data.Length - 8];
            Array.Copy(data, 8, formulaTokens, 0, formulaTokens.Length);
            _sharedFormulaString = FormulaDecompiler.Decompile(formulaTokens);
        }

        /// <summary>若当前公式单元格落在最近一次 SHAREDFMLA 范围内且非首格，则用主公式按行列差调整引用后写入。</summary>
        private void ApplySharedFormulaToCell(Cell cell)
        {
            if (string.IsNullOrEmpty(_sharedFormulaString)) return;
            int r0 = cell.RowIndex - 1;
            int c0 = cell.ColumnIndex - 1;
            if (r0 < _sharedFormulaBaseRow || r0 > _sharedFormulaLastRow || c0 < _sharedFormulaBaseCol || c0 > _sharedFormulaLastCol) return;
            if (r0 == _sharedFormulaBaseRow && c0 == _sharedFormulaBaseCol) return;
            int dr = r0 - _sharedFormulaBaseRow;
            int dc = c0 - _sharedFormulaBaseCol;
            cell.Formula = AdjustFormulaRefs(_sharedFormulaString, dr, dc);
            if (cell.DataType != "f") cell.DataType = "f";
        }

        /// <summary>将公式中的相对引用按 (dr, dc) 平移。仅处理简单 A1 引用与 A1:B2 范围，不含 $ 的视为相对。</summary>
        private static string AdjustFormulaRefs(string formula, int dr, int dc)
        {
            if (string.IsNullOrEmpty(formula) || (dr == 0 && dc == 0)) return formula;
            return Regex.Replace(formula, @"(\$?)([A-Za-z]+)(\$?)(\d+)", m =>
            {
                bool absCol = m.Groups[1].Value.Length > 0;
                bool absRow = m.Groups[3].Value.Length > 0;
                string colLetters = m.Groups[2].Value;
                int row = int.Parse(m.Groups[4].Value);
                int col0 = ColumnLettersToIndex(colLetters);
                int newCol = absCol ? col0 : (col0 + dc);
                int newRow = absRow ? row : (row + dr);
                if (newCol < 0 || newRow < 1) return m.Value;
                return (absCol ? "$" : "") + GetColumnLetterStatic(newCol) + (absRow ? "$" : "") + newRow;
            });
        }

        private static int ColumnLettersToIndex(string letters)
        {
            int index = 0;
            foreach (char c in letters.ToUpperInvariant())
                index = index * 26 + (c - 'A' + 1);
            return index - 1;
        }

        private static string GetColumnLetterStatic(int columnIndex)
        {
            if (columnIndex < 0) return "A";
            string s = "";
            int col = columnIndex + 1;
            while (col > 0)
            {
                int mod = (col - 1) % 26;
                s = (char)('A' + mod) + s;
                col = (col - mod) / 26;
            }
            return s;
        }

        /// <summary>解析 BIFF8 DIMENSION (0x0200)：firstRow(4), lastRow(4), firstCol(2), lastCol(2), reserved(2)。行列均为 0-based，用于初始化或扩展 MaxRow/MaxColumn。</summary>
        private void ParseDimensionRecord(BiffRecord record, Worksheet worksheet)
        {
            byte[] data = record.Data;
            if (data == null || data.Length < 14)
                return;
            int lastRow = BitConverter.ToInt32(data, 4);
            int lastCol = BitConverter.ToUInt16(data, 10);
            // 空表或无效值：lastRow/lastCol 可能为 -1 或 0xFFFF，只接受有效范围
            if (lastRow >= 0 && lastRow < 1048576)
                worksheet.MaxRow = Math.Max(worksheet.MaxRow, lastRow + 1);
            if (lastCol >= 0 && lastCol < 16384)
                worksheet.MaxColumn = Math.Max(worksheet.MaxColumn, lastCol + 1);
        }
        
        private void ParseMergeCellsRecord(BiffRecord record, Worksheet worksheet)
        {
            // BIFF8 MERGECELLS: count(2) + [startRow(2),startCol(2),endRow(2),endCol(2)]*count，最多1027个范围
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 2)
                return;
            ushort count = BitConverter.ToUInt16(data, 0);
            int maxCount = Math.Min((data.Length - 2) / 8, 2048); // 防止损坏的 count 导致超大循环
            if (count > maxCount) count = (ushort)maxCount;
            for (int i = 0; i < count; i++)
            {
                int offset = 2 + i * 8;
                if (offset + 8 > data.Length)
                    break;
                ushort startRow = BitConverter.ToUInt16(data, offset);
                ushort startCol = BitConverter.ToUInt16(data, offset + 2);
                ushort endRow = BitConverter.ToUInt16(data, offset + 4);
                ushort endCol = BitConverter.ToUInt16(data, offset + 6);
                // 有效范围：startRow<=endRow && startCol<=endCol（BIFF 行列均为 0-based，含首行/首列）
                if (startRow <= endRow && startCol <= endCol)
                {
                    worksheet.MergeCells.Add(new MergeCell
                    {
                        StartRow = startRow + 1,
                        StartColumn = startCol + 1,
                        EndRow = endRow + 1,
                        EndColumn = endCol + 1
                    });
                }
            }
        }
        
        private void ParseVbaStream(Workbook workbook)
        {
            // 使用 OleCompoundFile 查找 VBA 项目
            try
            {
                // 尝试常见的 VBA 存储路径
                byte[]? vbaData = _oleFile.ReadStreamByName("_VBA_PROJECT_CUR")
                               ?? _oleFile.ReadStreamByName("VBA");

                if (vbaData == null)
                {
                    // 尝试在目录中搜索包含 VBA 的条目
                    var entries = _oleFile.DirectoryEntries;
                    foreach (var entry in entries)
                    {
                        if (entry.Name != null && 
                            entry.Name.IndexOf("VBA", StringComparison.OrdinalIgnoreCase) >= 0 &&
                            entry.ObjectType == DirectoryEntryType.Stream &&
                            entry.StreamSize > 0)
                        {
                            vbaData = _oleFile.ReadStream(entry);
                            if (vbaData != null && vbaData.Length > 0)
                            {
                                Logger.Info($"找到VBA流: {entry.Name}, 大小: {vbaData.Length}");
                                break;
                            }
                        }
                    }
                }

                if (vbaData != null && vbaData.Length > 0)
                {
                    if (vbaData.Length > VbaSizeLimit)
                    {
                        Logger.Warn($"VBA项目大小超过{VbaSizeLimit / (1024 * 1024)}MB限制");
                        return;
                    }

                    workbook.VbaProject = vbaData;
                    Logger.Info($"成功读取VBA项目，大小: {vbaData.Length} 字节");
                }
            }
            catch (Exception ex)
            {
                Logger.Error("解析VBA流时发生错误", ex);
            }
        }

        private void ParseNameRecord(BiffRecord record, Workbook workbook)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 14)
                return;
            ushort options = BitConverter.ToUInt16(data, 0);
            byte nameLen = data[3];
            ushort formulaLen = BitConverter.ToUInt16(data, 4);
            bool hidden = (options & 0x0001) != 0;
            // itab 在偏移 8-9：非 0 表示局部名称，值为 1-based 的 BoundSheet8 索引
            ushort itab = data.Length >= 10 ? BitConverter.ToUInt16(data, 8) : (ushort)0;
            int localSheetId = itab > 0 ? (int)(itab - 1) : 0;
            int offset = 14;
            string name = ReadBiffStringFromBytes(data, ref offset, nameLen);
            if (nameLen == 1 && name.Length == 1 && name[0] == '\u000D')
                name = "FilterDatabase";
            byte[] formulaData = new byte[formulaLen];
            if (formulaLen > 0 && offset + formulaLen <= data.Length)
                Array.Copy(data, offset, formulaData, 0, formulaLen);
            string formula = FormulaDecompiler.Decompile(formulaData);
            workbook.DefinedNames.Add(new DefinedName
            {
                Name = name,
                Formula = formula,
                Hidden = hidden,
                LocalSheetId = itab > 0 ? (int?)(localSheetId) : null
            });
        }

        private void ParsePageSetupRecord(BiffRecord record, Worksheet worksheet)
        {
            // PAGESETUP (0x00A1) - BIFF8
            if (record.Data != null && record.Data.Length >= 34)
            {
                var ps = worksheet.PageSettings;
                ps.PaperSize = BitConverter.ToUInt16(record.Data, 0);
                ps.Scale = BitConverter.ToUInt16(record.Data, 2);
                ps.FitToWidth = BitConverter.ToUInt16(record.Data, 6);
                ps.FitToHeight = BitConverter.ToUInt16(record.Data, 8);
                
                ushort options = BitConverter.ToUInt16(record.Data, 10);
                ps.OrientationLandscape = (options & 0x0002) == 0;
                ps.UsePageNumbers = (options & 0x0001) != 0;
            }
        }

        private void ParseChartSubstream(Worksheet worksheet, Workbook workbook)
        {
            long streamEnd = _workbookData.Length;
            BiffRecord? previousRecord = null;
            Chart currentChart = new Chart();
            Series? currentSeries = null;
            
            // 默认图表类型
            currentChart.ChartType = "colChart";

            while (_stream.Position < streamEnd)
            {
                try
                {
                    var recordStartPos = _stream.Position;
                    var record = BiffRecord.Read(_reader);

                    // 解密数据体
                    if (_decryptor != null && record.Id != (ushort)BiffRecordType.BOF)
                    {
                        if (record.Data != null && record.Data.Length > 0)
                        {
                            _decryptor.Decrypt(record.Data, recordStartPos + 4);
                        }
                    }

                    if (record.Id == (ushort)BiffRecordType.CONTINUE)
                    {
                        if (previousRecord != null && record.Data != null)
                        {
                            previousRecord.Continues.Add(record.Data);
                        }
                        continue;
                    }

                    if (previousRecord != null)
                    {
                        ProcessChartRecord(previousRecord, currentChart, ref currentSeries!, worksheet);
                    }

                    previousRecord = record;

                    if (record.Id == (ushort)BiffRecordType.EOF)
                    {
                        break;
                    }
                }
                catch (EndOfStreamException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    Logger.Error($"解析图表记录时发生错误: {ex.Message}", ex);
                    continue;
                }
            }

            if (previousRecord != null && previousRecord.Id != (ushort)BiffRecordType.EOF)
            {
                ProcessChartRecord(previousRecord, currentChart, ref currentSeries!, worksheet);
            }
            
            // 确保图表有默认轴和图例
            if (currentChart.XAxis == null) currentChart.XAxis = new Axis { Visible = true, Title = "X轴" };
            if (currentChart.YAxis == null) currentChart.YAxis = new Axis { Visible = true, Title = "Y轴" };
            if (currentChart.Legend == null) currentChart.Legend = new Legend { Visible = true, Position = "right" };
            
            // 应用等待分配的坐标信息
            if (_pendingChartAnchors.Count > worksheet.Charts.Count)
            {
                var anchor = _pendingChartAnchors[worksheet.Charts.Count];
                currentChart.Left = anchor.Left;
                currentChart.Top = anchor.Top;
            }

            worksheet.Charts.Add(currentChart);
            Logger.Info($"成功向工作表 {worksheet.Name} 添加了 1 个图表 (类型: {currentChart.ChartType}, 系列数: {currentChart.Series.Count})");
        }

        private void ProcessChartRecord(BiffRecord record, Chart chart, ref Series currentSeries, Worksheet worksheet)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length == 0) return;

            switch (record.Id)
            {
                case 0x1017: // Bar
                    chart.ChartType = "barChart";
                    break;
                case 0x1018: // Line
                    chart.ChartType = "lineChart";
                    break;
                case 0x1019: // Pie
                    chart.ChartType = "pieChart";
                    break;
                case 0x101B: // Scatter
                    chart.ChartType = "scatterChart";
                    break;
                case 0x101A: // Area
                    chart.ChartType = "areaChart";
                    break;
                case 0x1020: // Radar
                    chart.ChartType = "radarChart";
                    break;
                case (ushort)BiffRecordType.CHARTSERIES:
                    currentSeries = new Series();
                    // 添加默认范围，实际公式解析较复杂
                    currentSeries.ValuesRange = $"{worksheet.Name}!$B$2:$B$6";
                    currentSeries.CategoriesRange = $"{worksheet.Name}!$A$2:$A$6";
                    currentSeries.LineStyle = new LineStyle { Width = 2 };
                    chart.Series.Add(currentSeries);
                    break;
            }
        }

        private string GetBorderLineStyle(byte styleId)
        {
            switch (styleId)
            {
                case 0: return "none";
                case 1: return "thin";
                case 2: return "medium";
                case 3: return "dashed";
                case 4: return "dotted";
                case 5: return "thick";
                case 6: return "double";
                case 7: return "hair";
                case 8: return "mediumDashed";
                case 9: return "dashDot";
                case 10: return "mediumDashDot";
                case 11: return "dashDotDot";
                case 12: return "mediumDashDotDot";
                case 13: return "slantDashDot";
                default: return "none";
            }
        }

        private string? GetColorFromPalette(int colorIndex)
        {
            if (colorIndex == 64) return null; // System Foreground
            if (colorIndex == 65) return null; // System Background

            if (_workbook.Palette.TryGetValue(colorIndex, out string? color))
                return color.Replace("#", "");

            if (_palette.TryGetValue(colorIndex, out string? palColor))
                return palColor.Replace("#", "");

            if (Biff8DefaultPalette.TryGetValue(colorIndex, out string? defaultColor))
                return defaultColor;

            return null;
        }

        private void ParseExternBookRecord(BiffRecord record, Workbook workbook)
        {
            if (record.Data != null && record.Data.Length >= 4)
            {
                ushort count = BitConverter.ToUInt16(record.Data, 0);
                ushort type = BitConverter.ToUInt16(record.Data, 2);
                
                var extBook = new ExternalBook();
                
                if (type == 0x0401)
                {
                    extBook.IsSelf = true;
                    Logger.Info("找到内部引用 (Self SUPBOOK)");
                }
                else if (type == 0x3A01)
                {
                    extBook.IsAddIn = true;
                    Logger.Info("找到 Add-In 引用");
                }
                else
                {
                    // 外部文件路径
                    int offset = 2;
                    // SUPBOOK 的路径字符串前通常是一个字节的字符数
                    byte pathLen = record.Data[offset];
                    offset++;
                    extBook.FileName = ReadBiffStringFromBytes(record.Data, ref offset, pathLen);
                    Logger.Info($"找到外部工作簿引用: {extBook.FileName}");
                }
                
                workbook.ExternalBooks.Add(extBook);
            }
        }

        private void ParseExternSheetRecord(BiffRecord record, Workbook workbook)
        {
            if (record.Data != null && record.Data.Length >= 2)
            {
                ushort count = BitConverter.ToUInt16(record.Data, 0);
                
                for (int i = 0; i < count; i++)
                {
                    int offset = 2 + i * 6;
                    if (offset + 6 <= record.Data.Length)
                    {
                        var extSheet = new ExternalSheet
                        {
                            ExternalBookIndex = BitConverter.ToUInt16(record.Data, offset),
                            FirstSheetIndex = BitConverter.ToInt16(record.Data, offset + 2),
                            LastSheetIndex = BitConverter.ToInt16(record.Data, offset + 4)
                        };
                        workbook.ExternalSheets.Add(extSheet);
                    }
                }
                Logger.Info($"解析 EXTERNSHEET, 共有 {count} 个引用映射");
            }
        }

        private void ParseExternalNameRecord(BiffRecord record, Workbook workbook)
        {
            if (record.Data != null && record.Data.Length >= 6)
            {
                // BIFF8 EXTERNALNAME
                ushort options = BitConverter.ToUInt16(record.Data, 0);
                byte nameLen = record.Data[3];
                
                int offset = 6;
                string name = ReadBiffStringFromBytes(record.Data, ref offset, nameLen);
                
                if (workbook.ExternalBooks.Count > 0)
                {
                    workbook.ExternalBooks[workbook.ExternalBooks.Count - 1].ExternalNames.Add(name);
                }
                
                Logger.Info($"找到外部工作簿名称引用: {name}");
            }
        }

        private void ParseAutoFilterInfoRecord(BiffRecord record, Worksheet worksheet, Workbook workbook)
        {
            if (record.Data == null || record.Data.Length < 2) return;
            ushort cEntries = BitConverter.ToUInt16(record.Data, 0);
            worksheet.AutoFilterColumnIndices.Clear();
            if (workbook.DefinedNames != null)
            {
                foreach (var dn in workbook.DefinedNames)
                {
                    if (dn?.Name != "FilterDatabase" && dn?.Name != "_xlnm._FilterDatabase") continue;
                    if (dn.LocalSheetId.HasValue && dn.LocalSheetId.Value != _currentSheetIndex) continue;
                    if (string.IsNullOrEmpty(dn.Formula)) continue;
                    worksheet.AutoFilterRange = StripDefinedNameToRange(dn.Formula);
                    break;
                }
            }
            if (string.IsNullOrEmpty(worksheet.AutoFilterRange))
                worksheet.AutoFilterRange = "A1:Z100";
        }

        private static string StripDefinedNameToRange(string formula)
        {
            if (string.IsNullOrEmpty(formula)) return formula;
            int excl = formula.IndexOf('!');
            string rangePart = excl >= 0 ? formula.Substring(excl + 1).Trim() : formula;
            return rangePart.Replace("$", "");
        }

        private void ParseAutoFilterRecord(BiffRecord record, Worksheet worksheet)
        {
            if (record.Data == null || record.Data.Length < 2) return;
            ushort iEntry = BitConverter.ToUInt16(record.Data, 0);
            if (worksheet.AutoFilterColumnIndices == null)
                worksheet.AutoFilterColumnIndices = new List<int>();
            worksheet.AutoFilterColumnIndices.Add((int)iEntry);
        }

        private void ParsePivotTableRecord(BiffRecord record, Worksheet worksheet)
        {
            if (record.Data == null) return;

            switch (record.Id)
            {
                case (ushort)BiffRecordType.SXVIEW:
                    // SXVIEW (0x00B0) - 开始一个新的数据透视表
                    if (record.Data.Length >= 16)
                    {
                        var pivotTable = new PivotTable();
                        // 部分 BIFF 实现中前 16 字节含范围：offset 4-5 rwFirst, 6-7 rwLast, 8-9 colFirst, 10-11 colLast (0-based)
                        if (record.Data.Length >= 12)
                        {
                            ushort rwFirst = BitConverter.ToUInt16(record.Data, 4);
                            ushort rwLast = BitConverter.ToUInt16(record.Data, 6);
                            ushort colFirst = BitConverter.ToUInt16(record.Data, 8);
                            ushort colLast = BitConverter.ToUInt16(record.Data, 10);
                            if (rwLast >= rwFirst && colLast >= colFirst && rwLast < 65535 && colLast < 256)
                            {
                                pivotTable.DataSource = $"{GetColumnLetterStatic(colFirst + 1)}{rwFirst + 1}:{GetColumnLetterStatic(colLast + 1)}{rwLast + 1}";
                            }
                        }
                        int offset = 16;
                        // SXVIEW 中的表名由一个两字节的字符数开始 (BIFF8)
                        if (offset + 2 <= record.Data.Length)
                        {
                            ushort nameLen = BitConverter.ToUInt16(record.Data, offset);
                            offset += 2;
                            pivotTable.Name = ReadBiffStringFromBytes(record.Data, ref offset, nameLen);
                        }
                        _currentPivotTable = pivotTable;
                        worksheet.PivotTables.Add(pivotTable);
                        Logger.Info($"开始解析数据透视表: {pivotTable.Name}");
                    }
                    break;

                case (ushort)BiffRecordType.SXVD:
                    // SXVD (0x00B1) - 字段属性
                    if (_currentPivotTable != null && record.Data.Length >= 2)
                    {
                        ushort sxaxis = BitConverter.ToUInt16(record.Data, 0);
                        var field = new PivotField();
                        
                        // sxaxis: 0=no axis, 1=row, 2=col, 4=page, 8=data
                        if ((sxaxis & 0x01) != 0) field.Type = "row";
                        else if ((sxaxis & 0x02) != 0) field.Type = "column";
                        else if ((sxaxis & 0x04) != 0) field.Type = "page";
                        else if ((sxaxis & 0x08) != 0) field.Type = "data";
                        
                        if (field.Type == "row") _currentPivotTable.RowFields.Add(field);
                        else if (field.Type == "column") _currentPivotTable.ColumnFields.Add(field);
                        else if (field.Type == "page") _currentPivotTable.PageFields.Add(field);
                        else if (field.Type == "data") _currentPivotTable.DataFields.Add(field);
                        _currentPivotField = field;
                    }
                    break;

                case (ushort)BiffRecordType.SXDX:
                    // SXDX (0x00C6) - 数据字段的具体信息
                    if (_currentPivotTable != null && _currentPivotTable.DataFields.Count > 0)
                    {
                        var lastDataField = _currentPivotTable.DataFields[_currentPivotTable.DataFields.Count - 1];
                        if (record.Data.Length >= 4)
                        {
                            ushort iSubtotal = BitConverter.ToUInt16(record.Data, 2);
                            // 0=Sum, 1=Count, 2=Average, etc.
                            string[] funcs = { "sum", "count", "average", "max", "min", "product", "countNums", "stdDev", "stdDevp", "var", "varp" };
                            if (iSubtotal < funcs.Length) lastDataField.Function = funcs[iSubtotal];
                        }
                    }
                    break;

                case (ushort)BiffRecordType.SXVI:
                    // SXVI (0x00B2) - 数据透视表项（在透视表流中，与全局 BORDER 同 ID 按流区分）
                    if (_currentPivotField != null && record.Data != null && record.Data.Length >= 2)
                    {
                        ushort cItem = BitConverter.ToUInt16(record.Data, 0);
                        int off = 2;
                        for (int i = 0; i < cItem && off + 2 <= record.Data.Length; i++)
                        {
                            ushort cch = BitConverter.ToUInt16(record.Data, off);
                            off += 2;
                            if (cch == 0) continue;
                            if (off + cch > record.Data.Length) break;
                            string itemStr = ReadBiffStringFromBytes(record.Data, ref off, cch);
                            _currentPivotField.Items.Add(itemStr);
                        }
                    }
                    break;

                case (ushort)BiffRecordType.SXFIELD:
                    // SXFIELD (0x00CA) - 字段名称及属性
                    if (_currentPivotTable != null && record.Data.Length >= 4)
                    {
                        int offset = 4;
                        // BIFF8 SXFIELD name starts with 2-byte char count
                        if (offset + 2 <= record.Data.Length)
                        {
                            ushort cch = BitConverter.ToUInt16(record.Data, offset);
                            offset += 2;
                            string fieldName = ReadBiffStringFromBytes(record.Data, ref offset, cch);
                            // 这里暂时无法完美回填，因为字段关联逻辑较复杂（通常按顺序）
                            // 简单起见，如果字段列表非空，回填到最后一个未命名的字段
                            var allFields = _currentPivotTable.RowFields
                                .Concat(_currentPivotTable.ColumnFields)
                                .Concat(_currentPivotTable.PageFields)
                                .Concat(_currentPivotTable.DataFields);
                            
                            foreach (var f in allFields)
                            {
                                if (string.IsNullOrEmpty(f.Name))
                                {
                                    f.Name = fieldName;
                                    break;
                                }
                            }
                        }
                    }
                    break;

                case (ushort)BiffRecordType.SXPI:
                    // SXPI (0x00B7) - 页字段项
                    if (_currentPivotField != null && _currentPivotField.Type == "page" && record.Data != null && record.Data.Length >= 2)
                    {
                        ushort cItem = BitConverter.ToUInt16(record.Data, 0);
                        int off = 2;
                        for (int i = 0; i < cItem && off + 2 <= record.Data.Length; i++)
                        {
                            ushort cch = BitConverter.ToUInt16(record.Data, off);
                            off += 2;
                            if (cch == 0) continue;
                            if (off + cch > record.Data.Length) break;
                            string itemStr = ReadBiffStringFromBytes(record.Data, ref off, cch);
                            _currentPivotField.Items.Add(itemStr);
                        }
                    }
                    break;
            }
        }

        private string GetPatternType(byte patternId)
        {
            switch (patternId)
            {
                case 0: return "none";
                case 1: return "solid";
                case 2: return "mediumGray";
                case 3: return "darkGray";
                case 4: return "lightGray";
                case 5: return "darkHorizontal";
                case 6: return "darkVertical";
                case 7: return "darkDown";
                case 8: return "darkUp";
                case 9: return "darkGrid";
                case 10: return "darkTrellis";
                case 11: return "lightHorizontal";
                case 12: return "lightVertical";
                case 13: return "lightDown";
                case 14: return "lightUp";
                case 15: return "lightGrid";
                case 16: return "lightTrellis";
                case 17: return "gray125";
                case 18: return "gray0625";
                default: return "none";
            }
        }
    }
}
