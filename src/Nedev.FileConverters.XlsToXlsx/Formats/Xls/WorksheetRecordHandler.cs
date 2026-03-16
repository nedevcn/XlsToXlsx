using System;
using System.Collections.Generic;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 工作表记录处理器 - 处理Worksheet级别的BIFF记录
    /// </summary>
    public class WorksheetRecordHandler
    {
        private readonly Worksheet _worksheet;
        private readonly Workbook _workbook;
        private readonly CellParser? _cellParser;
        private readonly MultiCellParser? _multiCellParser;
        private readonly WorksheetConfigParser? _worksheetConfigParser;
        private readonly ConditionalFormatParser? _conditionalFormatParser;
        private readonly AutoFilterParser? _autoFilterParser;
        private readonly DataValidationParser? _dataValidationParser;
        private readonly HyperlinkParser? _hyperlinkParser;
        private readonly CommentParser? _commentParser;
        private readonly PivotTableParser? _pivotTableParser;
        private readonly FontXfParser? _fontXfParser;
        private readonly PaletteParser? _paletteParser;
        private readonly DrawingParser? _drawingParser;

        // 当前行状态
        private Row? _currentRow;

        public WorksheetRecordHandler(
            Worksheet worksheet,
            Workbook workbook,
            CellParser? cellParser,
            MultiCellParser? multiCellParser,
            WorksheetConfigParser? worksheetConfigParser,
            ConditionalFormatParser? conditionalFormatParser,
            AutoFilterParser? autoFilterParser,
            DataValidationParser? dataValidationParser,
            HyperlinkParser? hyperlinkParser,
            CommentParser? commentParser,
            PivotTableParser? pivotTableParser,
            FontXfParser? fontXfParser,
            PaletteParser? paletteParser,
            DrawingParser? drawingParser)
        {
            _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            _cellParser = cellParser;
            _multiCellParser = multiCellParser;
            _worksheetConfigParser = worksheetConfigParser;
            _conditionalFormatParser = conditionalFormatParser;
            _autoFilterParser = autoFilterParser;
            _dataValidationParser = dataValidationParser;
            _hyperlinkParser = hyperlinkParser;
            _commentParser = commentParser;
            _pivotTableParser = pivotTableParser;
            _fontXfParser = fontXfParser;
            _paletteParser = paletteParser;
            _drawingParser = drawingParser;
        }

        /// <summary>
        /// 获取当前行
        /// </summary>
        public Row? CurrentRow => _currentRow;

        /// <summary>
        /// 处理记录
        /// </summary>
        /// <param name="record">BIFF记录</param>
        /// <returns>是否处理了该记录</returns>
        public bool HandleRecord(BiffRecord record)
        {
            switch (record.Id)
            {
                case (ushort)BiffRecordType.BOF:
                case (ushort)BiffRecordType.EOF:
                    return true;

                case (ushort)BiffRecordType.DIMENSION:
                    _cellParser?.ParseDimensionRecord(record, _worksheet);
                    return true;

                case (ushort)BiffRecordType.ROW:
                    ParseRowRecord(record);
                    return true;

                case (ushort)BiffRecordType.CELL_BLANK:
                case (ushort)BiffRecordType.CELL_BOOLERR:
                case (ushort)BiffRecordType.CELL_LABEL:
                case (ushort)BiffRecordType.CELL_LABELSST:
                case (ushort)BiffRecordType.CELL_NUMBER:
                case (ushort)BiffRecordType.CELL_RK:
                case (ushort)BiffRecordType.CELL_RSTRING:
                    ParseCellRecord(record);
                    return true;

                case (ushort)BiffRecordType.CELL_FORMULA:
                    ParseFormulaCellRecord(record);
                    return true;

                case (ushort)BiffRecordType.ARRAY:
                    _cellParser?.ParseArrayRecord(record, _worksheet);
                    return true;

                case (ushort)BiffRecordType.SHAREDFMLA:
                    _cellParser?.ParseSharedFormulaRecord(record);
                    return true;

                case (ushort)BiffRecordType.STRING:
                    ParseStringRecord(record);
                    return true;

                case (ushort)BiffRecordType.MULRK:
                    _multiCellParser?.ParseMulRkRecord(record, ref _currentRow, _worksheet);
                    return true;

                case (ushort)BiffRecordType.MULBLANK:
                    _multiCellParser?.ParseMulBlankRecord(record, ref _currentRow, _worksheet);
                    return true;

                case (ushort)BiffRecordType.MERGECELLS:
                    _cellParser?.ParseMergeCellsRecord(record, _worksheet);
                    return true;

                case (ushort)BiffRecordType.COLINFO:
                    _worksheetConfigParser?.ParseColInfoRecord(record, _worksheet);
                    return true;

                case (ushort)BiffRecordType.DEFCOLWIDTH:
                    _worksheetConfigParser?.ParseDefColWidthRecord(record, _worksheet);
                    return true;

                case (ushort)BiffRecordType.DEFAULTROWHEIGHT:
                    _worksheetConfigParser?.ParseDefaultRowHeightRecord(record, _worksheet);
                    return true;

                case (ushort)BiffRecordType.WINDOW2:
                    _worksheetConfigParser?.ParseWindow2Record(record, _worksheet);
                    return true;

                case 0x0041: // PANE
                    _worksheetConfigParser?.ParsePaneRecord(record, _worksheet);
                    return true;

                case (ushort)BiffRecordType.PROTECT:
                    _worksheetConfigParser?.ParseProtectRecord(record, _worksheet);
                    return true;

                case (ushort)BiffRecordType.PASSWORD:
                    _worksheetConfigParser?.ParsePasswordRecord(record, _worksheet);
                    return true;

                case (ushort)BiffRecordType.CF:
                    _conditionalFormatParser?.ParseCFRecord(record, _worksheet);
                    return true;

                case (ushort)BiffRecordType.AUTOFILTER:
                    _autoFilterParser?.ParseAutoFilterRecord(record, _worksheet);
                    return true;

                case (ushort)BiffRecordType.DV:
                    _dataValidationParser?.ParseDVRecord(record, _worksheet);
                    return true;

                case (ushort)BiffRecordType.NOTE:
                    _commentParser?.ParseCommentRecord(record, _worksheet);
                    return true;

                case (ushort)BiffRecordType.OBJ:
                    _drawingParser?.ParseObjRecord(record, _worksheet);
                    return true;

                case (ushort)BiffRecordType.MSODRAWING:
                    _drawingParser?.ParseMSODrawingRecord(record);
                    return true;

                case (ushort)BiffRecordType.FONT:
                    _fontXfParser?.ParseFontRecord(record, _worksheet);
                    return true;

                case (ushort)BiffRecordType.XF:
                    _fontXfParser?.ParseXfRecord(record, _worksheet);
                    return true;

                case (ushort)BiffRecordType.PALETTE:
                    _paletteParser?.ParsePaletteRecord(record, _worksheet);
                    return true;

                default:
                    return false;
            }
        }

        /// <summary>
        /// 解析行记录
        /// </summary>
        private void ParseRowRecord(BiffRecord record)
        {
            var parsedRow = _cellParser?.ParseRowRecord(record);
            if (parsedRow != null)
            {
                var existingRow = RowOperations.GetOrCreateRow(_worksheet, ref _currentRow, parsedRow.RowIndex);
                existingRow.Height = parsedRow.Height;
                existingRow.CustomHeight = parsedRow.CustomHeight;
                existingRow.DefaultXfIndex = parsedRow.DefaultXfIndex;
                if (existingRow.RowIndex > _worksheet.MaxRow) _worksheet.MaxRow = (int)existingRow.RowIndex;
            }
        }

        /// <summary>
        /// 解析单元格记录
        /// </summary>
        private void ParseCellRecord(BiffRecord record)
        {
            var cell = _cellParser?.ParseCellRecord(record);
            if (cell != null && cell.ColumnIndex >= 1 && cell.ColumnIndex <= 16384)
            {
                _cellParser?.TryApplyPendingArrayFormula(cell);
                var targetRow = RowOperations.GetOrCreateRow(_worksheet, ref _currentRow, cell.RowIndex);
                targetRow.Cells.Add(cell);
                if (cell.ColumnIndex > _worksheet.MaxColumn) _worksheet.MaxColumn = cell.ColumnIndex;
                if (cell.RowIndex > _worksheet.MaxRow) _worksheet.MaxRow = cell.RowIndex;
            }
        }

        /// <summary>
        /// 解析公式单元格记录
        /// </summary>
        private void ParseFormulaCellRecord(BiffRecord record)
        {
            var formulaCell = _cellParser?.ParseCellRecord(record);
            if (formulaCell != null && formulaCell.ColumnIndex >= 1 && formulaCell.ColumnIndex <= 16384)
            {
                _cellParser?.ApplySharedFormula(formulaCell);
                _cellParser?.TryApplyPendingArrayFormula(formulaCell);
                var targetRow = RowOperations.GetOrCreateRow(_worksheet, ref _currentRow, formulaCell.RowIndex);
                targetRow.Cells.Add(formulaCell);
                if (formulaCell.ColumnIndex > _worksheet.MaxColumn) _worksheet.MaxColumn = formulaCell.ColumnIndex;
                if (formulaCell.RowIndex > _worksheet.MaxRow) _worksheet.MaxRow = formulaCell.RowIndex;
            }
        }

        /// <summary>
        /// 解析字符串记录（公式结果）
        /// </summary>
        private void ParseStringRecord(BiffRecord record)
        {
            if (_currentRow != null && _currentRow.Cells.Count > 0)
            {
                var lastCell = _currentRow.Cells[_currentRow.Cells.Count - 1];
                byte[] strData = record.GetAllData();
                if (strData.Length > 0)
                {
                    int strOffset = 0;
                    lastCell.Value = RichTextParser.ReadBiffString(strData, ref strOffset);
                    lastCell.DataType = "inlineStr";
                }
            }
        }
    }
}
