using System.IO;
using System.Collections.Generic;
using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Nedev.FileConverters.XlsToXlsx;
using Nedev.FileConverters.XlsToXlsx.Exceptions;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// XLS文件解析器 - 将BIFF8格式的XLS文件解析为Workbook对象
    /// </summary>
    public partial class XlsParser
    {
        // 核心依赖
        private Stream _rawStream = null!;
        private OleCompoundFile _oleFile = null!;
        private byte[] _workbookData = null!;
        private Stream _stream = null!;
        private BinaryReader _reader = null!;
        private Workbook _workbook = null!;

        // 数据集合
        private List<string> _sharedStrings = new List<string>();
        private List<Font> _fonts = new List<Font>();
        private List<Xf> _xfList = new List<Xf>();
        private Dictionary<ushort, string> _formats = new Dictionary<ushort, string>();
        private Dictionary<int, string> _palette = new Dictionary<int, string>();
        private List<byte[]> _msoDrawingGroupData = new List<byte[]>();
        private List<int> _sheetOffsets = new List<int>();

        // 状态跟踪
        private List<byte[]> _msoDrawingData = new List<byte[]>();
        private List<(int Left, int Top, int Width, int Height)> _pendingChartAnchors = new List<(int, int, int, int)>();
        private int _currentSheetIndex;

        // 专用解析器
        private StyleParser? _styleParser;
        private CellParser? _cellParser;
        private ChartParser? _chartParser;
        private PivotTableParser? _pivotTableParser;
        private ExternalLinkParser? _externalLinkParser;
        private WorksheetConfigParser? _worksheetConfigParser;
        private ConditionalFormatParser? _conditionalFormatParser;
        private AutoFilterParser? _autoFilterParser;
        private DataValidationParser? _dataValidationParser;
        private HyperlinkParser? _hyperlinkParser;
        private CommentParser? _commentParser;
        private MultiCellParser? _multiCellParser;
        private FontXfParser? _fontXfParser;
        private PaletteParser? _paletteParser;
        private DrawingParser? _drawingParser;

        // 专用解析器
        private SstParser? _sstParser;
        private DefinedNameParser? _definedNameParser;
        private GlobalStyleParser? _globalStyleParser;
        private WorkbookStyleParser? _workbookStyleParser;
        private SheetRecordParser? _sheetRecordParser;
        private WorksheetStyleParser? _worksheetStyleParser;

        // 安全设置
        private XlsDecryptor? _decryptor;
        public long VbaSizeLimit { get; set; } = 50 * 1024 * 1024;
        public string Password { get; set; } = "VelvetSweatshop";
        private const long MAX_FILE_SIZE = 100 * 1024 * 1024;

        /// <summary>
        /// BIFF8默认64色调色板
        /// </summary>
        private static readonly IReadOnlyDictionary<int, string> Biff8DefaultPalette = CreateBiff8DefaultPalette();

        private static IReadOnlyDictionary<int, string> CreateBiff8DefaultPalette()
        {
            var d = new Dictionary<int, string>();
            string[] first8 = { "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF" };
            for (int i = 0; i < first8.Length; i++) d[i] = first8[i];
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

        public XlsParser(Stream stream)
        {
            Initialize(stream);
        }

        private void Initialize(Stream stream)
        {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead)
            {
                throw new XlsToXlsxException("Stream must be readable", 1000, "StreamError");
            }

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
            _fonts = new List<Font>();
            _xfList = new List<Xf>();
            _formats = new Dictionary<ushort, string>();
            _palette = new Dictionary<int, string>();
            _msoDrawingGroupData = new List<byte[]>();
            _sheetOffsets = new List<int>();
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

                // 2. 解析 SummaryInformation / DocumentSummaryInformation 文档属性
                var docPropsParser = new DocumentPropertiesParser(_oleFile);
                docPropsParser.Parse(workbook);

                // 3. 读取 Workbook 流
                _workbookData = _oleFile.ReadStreamByName("Workbook")
                             ?? _oleFile.ReadStreamByName("Book")  // Excel 5.0/95 兼容
                             ?? throw new XlsParseException("在OLE文件中未找到Workbook或Book流");
                Logger.Info($"Workbook流读取完成: {_workbookData.Length} 字节");

                // 4. 在 Workbook MemoryStream 上解析 BIFF 记录
                _stream = new MemoryStream(_workbookData);
                _reader = new BinaryReader(_stream);

                // 5. 解析全局记录 (BOUNDSHEET, SST, FONT, XF, FORMAT, PALETTE, NAME)
                ParseWorkbookGlobals(workbook);
                Logger.Info($"全局记录解析完成: {workbook.Worksheets.Count} 个工作表, {_sharedStrings.Count} 个共享字符串");

                // 6. 根据 BOUNDSHEET 中记录的偏移量解析各工作表子流
                ParseAllWorksheetSubstreams(workbook);
                Logger.Info("所有工作表子流解析完成");

                // 7. 将解析到的全局数据转移到工作簿对象
                workbook.SharedStrings = _sharedStrings;
                workbook.Fonts = _fonts;
                workbook.XfList = _xfList;
                workbook.NumberFormats = _formats;
                workbook.Palette = _palette;

                // 7.1 生成全局样式列表并更新单元格的 StyleId
                var stylesBuilder = new WorkbookStylesBuilder();
                stylesBuilder.BuildWorkbookStyles(workbook);

                // 8. 解析VBA流
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


        
        /// <summary>
        /// 解析 Workbook 全局子流（从 BOF 到 EOF），收集 BOUNDSHEET/SST/FONT/XF/FORMAT 等全局记录。
        /// </summary>
        private void ParseWorkbookGlobals(Workbook workbook)
        {
            // 初始化解析器
            _styleParser = new StyleParser(workbook, _palette, _formats, _fonts, _xfList);
            _externalLinkParser = new ExternalLinkParser();
            _fontXfParser = new FontXfParser(
                workbook, _fonts, _xfList, _formats,
                GetColorFromPalette,
                GetBorderLineStyle,
                GetPatternType);
            _paletteParser = new PaletteParser();
            _sstParser = new SstParser(_sharedStrings);
            _definedNameParser = new DefinedNameParser(workbook);
            _globalStyleParser = new GlobalStyleParser(
                _fonts, _xfList, _formats,
                idx => GetColorFromPalette(idx),
                GetBorderLineStyle,
                GetPatternType);
            _workbookStyleParser = new WorkbookStyleParser(
                workbook, _fonts, _xfList, _formats, _palette,
                idx => GetColorFromPalette(idx),
                GetBorderLineStyle,
                GetPatternType);
            _sheetRecordParser = new SheetRecordParser(_sheetOffsets);

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
                    _sheetRecordParser?.ParseSheetRecord(record, workbook);
                    break;
                case (ushort)BiffRecordType.SST:
                    _sstParser?.ParseSstRecord(record, streamEnd);
                    break;
                case (ushort)BiffRecordType.FONT:
                    _workbookStyleParser?.ParseFontRecord(record);
                    break;
                case (ushort)BiffRecordType.XF:
                    _workbookStyleParser?.ParseXfRecord(record);
                    break;
                case (ushort)BiffRecordType.FORMAT:
                    _workbookStyleParser?.ParseFormatRecord(record);
                    break;
                case (ushort)BiffRecordType.PALETTE:
                    _workbookStyleParser?.ParsePaletteRecord(record);
                    break;
                case (ushort)BiffRecordType.BORDER:
                    _workbookStyleParser?.ParseBorderRecord(record);
                    break;
                case (ushort)BiffRecordType.FILL:
                    _workbookStyleParser?.ParseFillRecord(record);
                    break;
                case (ushort)BiffRecordType.NAME:
                    _definedNameParser?.ParseNameRecord(record);
                    break;
                case (ushort)BiffRecordType.MSODRAWINGGROUP:
                    _drawingParser?.ParseMsoDrawingGroupGlobal(record);
                    break;
                case (ushort)BiffRecordType.FILEPASS:
                    if (record.Data != null && record.Data.Length >= 52)
                    {
                        Logger.Info("检测到加密文件，正在初始化解密器");
                        _decryptor = new XlsDecryptor(record.Data, Password);
                    }
                    break;
                case (ushort)BiffRecordType.EXTERNBOOK:
                    _externalLinkParser?.ParseExternBookRecord(record, workbook);
                    break;
                case (ushort)BiffRecordType.EXTERNSHEET:
                    _externalLinkParser?.ParseExternSheetRecord(record, workbook);
                    break;
                case (ushort)BiffRecordType.EXTERNALNAME:
                    _externalLinkParser?.ParseExternalNameRecord(record, workbook);
                    break;
                case (ushort)BiffRecordType.PROTECT:
                    // 全局 PROTECT: 锁定工作簿结构
                    if (record.Data != null && record.Data.Length >= 2)
                    {
                        ushort flags = BitConverter.ToUInt16(record.Data, 0);
                        workbook.IsStructureProtected = (flags & 0x0001) != 0;
                    }
                    break;
                case (ushort)BiffRecordType.WINDOWPROTECT:
                    // 窗口保护
                    if (record.Data != null && record.Data.Length >= 2)
                    {
                        ushort flags = BitConverter.ToUInt16(record.Data, 0);
                        workbook.IsWindowsProtected = (flags & 0x0001) != 0;
                    }
                    break;
                case (ushort)BiffRecordType.PASSWORD:
                    // 全局 PASSWORD：工作簿结构密码
                    if (record.Data != null && record.Data.Length >= 2)
                    {
                        ushort hash = BitConverter.ToUInt16(record.Data, 0);
                        workbook.WorkbookPasswordHash = hash.ToString("X4", System.Globalization.CultureInfo.InvariantCulture);
                    }
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
            // 初始化解析器
            _cellParser = new CellParser(_sharedStrings, _fonts, workbook);
            _chartParser = new ChartParser(workbook, _decryptor, _workbookData, _stream, _reader, _pendingChartAnchors);
            _pivotTableParser = new PivotTableParser();
            _worksheetConfigParser = new WorksheetConfigParser();
            _conditionalFormatParser = new ConditionalFormatParser(workbook);
            _autoFilterParser = new AutoFilterParser(_currentSheetIndex);
            _dataValidationParser = new DataValidationParser(workbook);
            _hyperlinkParser = new HyperlinkParser();
            _commentParser = new CommentParser();
            _multiCellParser = new MultiCellParser();
            _fontXfParser = new FontXfParser(
                workbook, _fonts, _xfList, _formats,
                GetColorFromPalette,
                GetBorderLineStyle,
                GetPatternType);
            _paletteParser = new PaletteParser();
            _drawingParser = new DrawingParser(_msoDrawingData, _msoDrawingGroupData, _pendingChartAnchors);
            _worksheetStyleParser = new WorksheetStyleParser(
                workbook, _formats,
                GetColorFromPalette,
                GetBorderLineStyle,
                GetPatternType);

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
                            _chartParser?.ParseChartSubstream(worksheet);
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
                            _cellParser?.ParseDimensionRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.ROW:
                            var parsedRow = _cellParser?.ParseRowRecord(record);
                            if (parsedRow != null)
                            {
                                var existingRow = RowOperations.GetOrCreateRow(worksheet, ref currentRow, parsedRow.RowIndex);
                                existingRow.Height = parsedRow.Height;
                                existingRow.CustomHeight = parsedRow.CustomHeight;
                                existingRow.DefaultXfIndex = parsedRow.DefaultXfIndex;
                                if (existingRow.RowIndex > worksheet.MaxRow) worksheet.MaxRow = (int)existingRow.RowIndex;
                            }
                            break;
                        case (ushort)BiffRecordType.CELL_BLANK:
                        case (ushort)BiffRecordType.CELL_BOOLERR:
                        case (ushort)BiffRecordType.CELL_LABEL:
                        case (ushort)BiffRecordType.CELL_LABELSST:
                        case (ushort)BiffRecordType.CELL_NUMBER:
                        case (ushort)BiffRecordType.CELL_RK:
                        case (ushort)BiffRecordType.CELL_RSTRING: // 旧版富文本
                            var cell = _cellParser?.ParseCellRecord(record);
                            if (cell != null && cell.ColumnIndex >= 1 && cell.ColumnIndex <= 16384)
                            {
                                _cellParser?.TryApplyPendingArrayFormula(cell);
                                var targetRow = RowOperations.GetOrCreateRow(worksheet, ref currentRow, cell.RowIndex);
                                targetRow.Cells.Add(cell);
                                if (cell.ColumnIndex > worksheet.MaxColumn) worksheet.MaxColumn = cell.ColumnIndex;
                                if (cell.RowIndex > worksheet.MaxRow) worksheet.MaxRow = cell.RowIndex;
                            }
                            break;
                        case (ushort)BiffRecordType.CELL_FORMULA:
                            var formulaCell = _cellParser?.ParseCellRecord(record);
                            if (formulaCell != null && formulaCell.ColumnIndex >= 1 && formulaCell.ColumnIndex <= 16384)
                            {
                                _cellParser?.ApplySharedFormula(formulaCell);
                                _cellParser?.TryApplyPendingArrayFormula(formulaCell);
                                var targetRow2 = RowOperations.GetOrCreateRow(worksheet, ref currentRow, formulaCell.RowIndex);
                                targetRow2.Cells.Add(formulaCell);
                                if (formulaCell.ColumnIndex > worksheet.MaxColumn) worksheet.MaxColumn = formulaCell.ColumnIndex;
                                if (formulaCell.RowIndex > worksheet.MaxRow) worksheet.MaxRow = formulaCell.RowIndex;
                            }
                            break;
                        case (ushort)BiffRecordType.ARRAY:
                            _cellParser?.ParseArrayRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.SHAREDFMLA:
                            _cellParser?.ParseSharedFormulaRecord(record);
                            break;
                        case (ushort)BiffRecordType.STRING:
                            _cellParser?.ParseStringRecord(record, currentRow);
                            break;
                        case (ushort)BiffRecordType.MULRK:
                            _multiCellParser?.ParseMulRkRecord(record, ref currentRow, worksheet);
                            break;
                        case (ushort)BiffRecordType.MULBLANK:
                            _multiCellParser?.ParseMulBlankRecord(record, ref currentRow, worksheet);
                            break;
                        case (ushort)BiffRecordType.MERGECELLS:
                            _cellParser?.ParseMergeCellsRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.COLINFO:
                            _worksheetConfigParser?.ParseColInfoRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.DEFCOLWIDTH:
                            _worksheetConfigParser?.ParseDefColWidthRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.DEFAULTROWHEIGHT:
                            _worksheetConfigParser?.ParseDefaultRowHeightRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.WINDOW2:
                            _worksheetConfigParser?.ParseWindow2Record(record, worksheet);
                            break;
                        case 0x0041: // PANE
                            _worksheetConfigParser?.ParsePaneRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.PROTECT:
                            _worksheetConfigParser?.ParseProtectRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.PASSWORD:
                            _worksheetConfigParser?.ParsePasswordRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.DV:
                            _dataValidationParser?.ParseDVRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.CFHEADER:
                            _conditionalFormatParser?.ParseCFHeaderRecord(record);
                            break;
                        case (ushort)BiffRecordType.CF:
                            _conditionalFormatParser?.ParseCFRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.HYPERLINK:
                            _hyperlinkParser?.ParseHyperlinkRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.NOTE:
                            _commentParser?.ParseCommentRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.MSODRAWING:
                            _drawingParser?.ParseMSODrawingRecord(record);
                            break;
                        case (ushort)BiffRecordType.PICTURE:
                            _drawingParser?.ParsePictureRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.OBJ:
                            _drawingParser?.ParseObjRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.HEADER:
                            _worksheetConfigParser?.ParseHeaderRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.FOOTER:
                            _worksheetConfigParser?.ParseFooterRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.LEFTMARGIN:
                            _worksheetConfigParser?.ParseMarginRecord(record, worksheet, BiffRecordType.LEFTMARGIN);
                            break;
                        case (ushort)BiffRecordType.RIGHTMARGIN:
                            _worksheetConfigParser?.ParseMarginRecord(record, worksheet, BiffRecordType.RIGHTMARGIN);
                            break;
                        case (ushort)BiffRecordType.TOPMARGIN:
                            _worksheetConfigParser?.ParseMarginRecord(record, worksheet, BiffRecordType.TOPMARGIN);
                            break;
                        case (ushort)BiffRecordType.BOTTOMMARGIN:
                            _worksheetConfigParser?.ParseMarginRecord(record, worksheet, BiffRecordType.BOTTOMMARGIN);
                            break;
                        case (ushort)BiffRecordType.HCENTER:
                            _worksheetConfigParser?.ParseCenterRecord(record, worksheet, true);
                            break;
                        case (ushort)BiffRecordType.VCENTER:
                            _worksheetConfigParser?.ParseCenterRecord(record, worksheet, false);
                            break;
                        case (ushort)BiffRecordType.PAGESETUP:
                            _worksheetConfigParser?.ParsePageSetupRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.FONT:
                            _worksheetStyleParser?.ParseFontRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.XF:
                            _worksheetStyleParser?.ParseXfRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.PALETTE:
                            _worksheetStyleParser?.ParsePaletteRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.FORMAT:
                            _worksheetStyleParser?.ParseFormatRecord(record);
                            break;
                        case (ushort)BiffRecordType.SXVIEW:
                        case (ushort)BiffRecordType.SXVD:
                        case (ushort)BiffRecordType.SXVI:
                        case (ushort)BiffRecordType.SXDX:
                        case (ushort)BiffRecordType.SXFIELD:
                        case (ushort)BiffRecordType.SXPI:
                            _pivotTableParser?.ParsePivotTableRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.AUTOFILTERINFO:
                            _autoFilterParser?.ParseAutoFilterInfoRecord(record, worksheet, workbook);
                            break;
                        case (ushort)BiffRecordType.AUTOFILTER:
                            _autoFilterParser?.ParseAutoFilterRecord(record, worksheet);
                            break;
                case (ushort)BiffRecordType.SST:
                    _sstParser?.ParseSstRecord(record, streamEnd);
                    break;
            }
        }


        // ===== 以下为辅助方法 =====
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

        private string GetBorderLineStyle(byte styleId) => ParsingHelpers.GetBorderLineStyle(styleId);

        private string? GetColorFromPalette(int colorIndex, Worksheet? sheet = null)
        {
            if (colorIndex == 64) return null; // System Foreground
            if (colorIndex == 65) return null; // System Background

            // sheet-level palette overrides workbook/global
            if (sheet != null && sheet.Palette.TryGetValue(colorIndex, out string? sheetColor))
                return sheetColor.Replace("#", "");

            if (_workbook.Palette.TryGetValue(colorIndex, out string? color))
                return color.Replace("#", "");

            if (_palette.TryGetValue(colorIndex, out string? palColor))
                return palColor.Replace("#", "");

            if (Biff8DefaultPalette.TryGetValue(colorIndex, out string? defaultColor))
                return defaultColor;

            return null;
        }

        private string GetPatternType(byte patternId) => ParsingHelpers.GetPatternType(patternId);
    }
}
