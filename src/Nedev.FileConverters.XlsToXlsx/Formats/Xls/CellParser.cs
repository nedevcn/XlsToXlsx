using System;
using System.Collections.Generic;
using System.Linq;
using Nedev.FileConverters.XlsToXlsx;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 单元格解析器 - 处理所有单元格相关的BIFF记录
    /// </summary>
    public class CellParser
    {
        private readonly List<string> _sharedStrings;
        private readonly List<Font> _fonts;
        private readonly Workbook _workbook;

        // 共享公式状态
        private string? _sharedFormulaString;
        private int _sharedFormulaBaseRow;
        private int _sharedFormulaBaseCol;
        private int _sharedFormulaLastRow;
        private int _sharedFormulaLastCol;

        public CellParser(
            List<string> sharedStrings,
            List<Font> fonts,
            Workbook workbook)
        {
            _sharedStrings = sharedStrings ?? throw new ArgumentNullException(nameof(sharedStrings));
            _fonts = fonts ?? throw new ArgumentNullException(nameof(fonts));
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
        }

        /// <summary>
        /// 解析单元格记录
        /// </summary>
        public Cell ParseCellRecord(BiffRecord record)
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

            // 读取样式索引 (XF index)
            ushort styleIndex = BitConverter.ToUInt16(data, 4);
            cell.StyleId = styleIndex.ToString();

            // 根据记录类型解析单元格值
            switch (record.Id)
            {
                case (ushort)BiffRecordType.CELL_BLANK:
                    cell.Value = null;
                    break;

                case (ushort)BiffRecordType.CELL_BOOLERR:
                    ParseBoolErrCell(data, cell);
                    break;

                case (ushort)BiffRecordType.CELL_LABEL:
                    ParseLabelCell(data, cell);
                    break;

                case (ushort)BiffRecordType.CELL_RSTRING:
                    ParseRStringCell(data, cell);
                    break;

                case (ushort)BiffRecordType.CELL_RICH_TEXT:
                    ParseRichTextCell(record, data, cell);
                    break;

                case (ushort)BiffRecordType.CELL_LABELSST:
                    ParseLabelSstCell(data, cell);
                    break;

                case (ushort)BiffRecordType.CELL_NUMBER:
                    ParseNumberCell(data, cell);
                    break;

                case (ushort)BiffRecordType.CELL_RK:
                    ParseRkCell(data, cell);
                    break;

                case (ushort)BiffRecordType.CELL_FORMULA:
                    ParseFormulaCell(record, data, cell);
                    break;
            }

            return cell;
        }

        /// <summary>
        /// 解析共享公式记录 (SHRFMLA)
        /// </summary>
        public void ParseSharedFormulaRecord(BiffRecord record)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 8) return;

            _sharedFormulaBaseRow = BitConverter.ToUInt16(data, 0);
            _sharedFormulaLastRow = BitConverter.ToUInt16(data, 2);
            _sharedFormulaBaseCol = data[4];
            _sharedFormulaLastCol = data[5];

            int formulaLength = data[6];
            if (data.Length >= 8 + formulaLength)
            {
                byte[] ptgs = new byte[formulaLength];
                Array.Copy(data, 8, ptgs, 0, formulaLength);
                _sharedFormulaString = FormulaDecompiler.Decompile(ptgs, _workbook);
            }
        }

        /// <summary>
        /// 解析数组公式记录 (ARRAY)
        /// </summary>
        public void ParseArrayRecord(BiffRecord record, Worksheet worksheet)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 12) return;

            ushort firstRow = BitConverter.ToUInt16(data, 0);
            ushort lastRow = BitConverter.ToUInt16(data, 2);
            byte firstCol = data[4];
            byte lastCol = data[5];
            ushort options = BitConverter.ToUInt16(data, 6);
            int formulaLength = BitConverter.ToUInt16(data, 8);

            if (data.Length < 12 + formulaLength) return;

            byte[] ptgs = new byte[formulaLength];
            Array.Copy(data, 12, ptgs, 0, formulaLength);
            string formula = FormulaDecompiler.Decompile(ptgs, _workbook);
            
            // 存储数组公式信息，供后续单元格处理使用
            // 注意：数组公式会在解析到具体单元格时应用
            _pendingArrayFormula = formula;
            _pendingArrayFormulaRange = (firstRow, firstCol, lastRow, lastCol);
        }

        // 待处理的数组公式状态
        private string? _pendingArrayFormula;
        private (int firstRow, int firstCol, int lastRow, int lastCol)? _pendingArrayFormulaRange;

        /// <summary>
        /// 检查并应用待处理的数组公式到单元格
        /// </summary>
        public void TryApplyPendingArrayFormula(Cell cell)
        {
            if (_pendingArrayFormula == null || !_pendingArrayFormulaRange.HasValue) return;
            
            var range = _pendingArrayFormulaRange.Value;
            int row0 = cell.RowIndex - 1;
            int col0 = cell.ColumnIndex - 1;
            
            if (row0 >= range.firstRow && row0 <= range.lastRow &&
                col0 >= range.firstCol && col0 <= range.lastCol)
            {
                cell.Formula = _pendingArrayFormula;
                cell.IsArrayFormula = true;
                cell.ArrayRef = $"{ParsingHelpers.ColumnIndexToLetters1Based(range.firstCol + 1)}{range.firstRow + 1}:{ParsingHelpers.ColumnIndexToLetters1Based(range.lastCol + 1)}{range.lastRow + 1}";
            }
        }

        /// <summary>
        /// 应用共享公式到单元格
        /// </summary>
        public void ApplySharedFormula(Cell cell)
        {
            if (string.IsNullOrEmpty(_sharedFormulaString)) return;
            
            int r0 = cell.RowIndex - 1;
            int c0 = cell.ColumnIndex - 1;
            
            if (r0 < _sharedFormulaBaseRow || r0 > _sharedFormulaLastRow || 
                c0 < _sharedFormulaBaseCol || c0 > _sharedFormulaLastCol) return;
            if (r0 == _sharedFormulaBaseRow && c0 == _sharedFormulaBaseCol) return;
            
            int dr = r0 - _sharedFormulaBaseRow;
            int dc = c0 - _sharedFormulaBaseCol;
            cell.Formula = AdjustFormulaRefs(_sharedFormulaString, dr, dc);
            if (cell.DataType != "f") cell.DataType = "f";
        }

        /// <summary>
        /// 将公式中的相对引用按 (dr, dc) 平移
        /// </summary>
        private static string AdjustFormulaRefs(string formula, int dr, int dc)
        {
            if (string.IsNullOrEmpty(formula) || (dr == 0 && dc == 0)) return formula;
            
            return System.Text.RegularExpressions.Regex.Replace(formula, @"(\$?)([A-Za-z]+)(\$?)(\d+)", m =>
            {
                bool absCol = m.Groups[1].Value.Length > 0;
                bool absRow = m.Groups[3].Value.Length > 0;
                string colLetters = m.Groups[2].Value;
                int row = int.Parse(m.Groups[4].Value);
                int col0 = ExcelAddressHelper.LettersToColumnIndex0Based(colLetters);
                int newCol = absCol ? col0 : (col0 + dc);
                int newRow = absRow ? row : (row + dr);
                if (newCol < 0 || newRow < 1) return m.Value;
                return (absCol ? "$" : "") + ParsingHelpers.ColumnIndexToLetters0Based(newCol) + (absRow ? "$" : "") + newRow;
            });
        }

        /// <summary>
        /// 解析行记录 (ROW)
        /// </summary>
        public Row ParseRowRecord(BiffRecord record)
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

                    // 偏移 16: ixfe 行默认 XF 索引
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

        /// <summary>
        /// 解析合并单元格记录 (MERGECELLS)
        /// </summary>
        public void ParseMergeCellsRecord(BiffRecord record, Worksheet worksheet)
        {
            // BIFF8 MERGECELLS: count(2) + [startRow(2),startCol(2),endRow(2),endCol(2)]*count
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 2) return;

            ushort count = BitConverter.ToUInt16(data, 0);
            int maxCount = Math.Min((data.Length - 2) / 8, 2048);
            if (count > maxCount) count = (ushort)maxCount;

            for (int i = 0; i < count; i++)
            {
                int offset = 2 + i * 8;
                if (offset + 8 > data.Length) break;

                ushort startRow = BitConverter.ToUInt16(data, offset);
                ushort endRow = BitConverter.ToUInt16(data, offset + 2);
                ushort startCol = BitConverter.ToUInt16(data, offset + 4);
                ushort endCol = BitConverter.ToUInt16(data, offset + 6);

                // 转换为1-based
                worksheet.MergeCells.Add(new MergeCell
                {
                    StartRow = startRow + 1,
                    StartColumn = startCol + 1,
                    EndRow = endRow + 1,
                    EndColumn = endCol + 1
                });
            }
        }

        /// <summary>
        /// 解析维度记录 (DIMENSION)
        /// </summary>
        public void ParseDimensionRecord(BiffRecord record, Worksheet worksheet)
        {
            byte[] data = record.Data;
            if (data == null || data.Length < 14) return;

            int lastRow = BitConverter.ToInt32(data, 4);
            int lastCol = BitConverter.ToUInt16(data, 10);

            // 只接受有效范围
            if (lastRow >= 0 && lastRow < 1048576)
                worksheet.MaxRow = Math.Max(worksheet.MaxRow, lastRow + 1);
            if (lastCol >= 0 && lastCol < 16384)
                worksheet.MaxColumn = Math.Max(worksheet.MaxColumn, lastCol + 1);
        }

        #region 私有解析方法

        private void ParseBoolErrCell(byte[] data, Cell cell)
        {
            if (data.Length < 8) return;

            byte valueOrError = data[6];
            byte isError = data[7]; // 0 = 布尔值, 1 = 错误值

            if (isError == 0)
            {
                cell.Value = valueOrError != 0;
                cell.DataType = "b";
            }
            else
            {
                cell.Value = ParsingHelpers.GetErrorString(valueOrError);
                cell.DataType = "e";
            }
        }

        private void ParseLabelCell(byte[] data, Cell cell)
        {
            if (data.Length <= 6) return;

            int offset = 6;
            cell.Value = ReadBiffString(data, ref offset);
            cell.DataType = "inlineStr";
        }

        private void ParseRStringCell(byte[] data, Cell cell)
        {
            if (data.Length <= 8) return;

            int offset = 6;
            cell.Value = ReadBiffString(data, ref offset);
            cell.DataType = "inlineStr";
        }

        private void ParseRichTextCell(BiffRecord record, byte[] data, Cell cell)
        {
            if (data.Length <= 6) return;

            var reader = new BiffStringReader(record, 6);
            var (txt, runInfos) = reader.ReadRichTextString();
            cell.Value = txt;
            cell.DataType = "inlineStr";

            if (runInfos != null && runInfos.Count > 0)
            {
                cell.RichText = new List<RichTextRun>();
                for (int ri = 0; ri < runInfos.Count; ri++)
                {
                    int start = runInfos[ri].CharPos;
                    int end = (ri + 1 < runInfos.Count) ? runInfos[ri + 1].CharPos : txt.Length;
                    if (start < 0) start = 0;
                    if (end > txt.Length) end = txt.Length;

                    string runText = txt.Substring(start, end - start);
                    var run = new RichTextRun
                    {
                        Text = runText,
                        Font = GetFontByIndex(runInfos[ri].FontIndex)
                    };
                    cell.RichText.Add(run);
                }
            }
        }

        private void ParseLabelSstCell(byte[] data, Cell cell)
        {
            if (data.Length < 10) return;

            int sstIndex = BitConverter.ToInt32(data, 6);
            if (sstIndex >= 0 && sstIndex < _sharedStrings.Count)
            {
                cell.Value = _sharedStrings[sstIndex];
            }
            cell.DataType = "s";
        }

        private void ParseNumberCell(byte[] data, Cell cell)
        {
            if (data.Length < 14) return;

            double value = BitConverter.ToDouble(data, 6);

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

        private void ParseRkCell(byte[] data, Cell cell)
        {
            if (data.Length < 10) return;

            int rkValue = BitConverter.ToInt32(data, 6);
            double value = ParsingHelpers.DecodeRKValue(rkValue);

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

        private void ParseFormulaCell(BiffRecord record, byte[] data, Cell cell)
        {
            try
            {
                if (data.Length < 22) return;

                // 读取公式结果
                double result = BitConverter.ToDouble(data, 6);

                // 读取公式 Ptgs 长度
                int formulaLength = BitConverter.ToUInt16(data, 20);

                // 读取公式 Ptgs
                if (data.Length >= 22 + formulaLength)
                {
                    byte[] ptgs = new byte[formulaLength];
                    Array.Copy(data, 22, ptgs, 0, formulaLength);

                    Logger.Info($"Formula PTGs raw (len={formulaLength}): {BitConverter.ToString(ptgs)}");
                    string formula = FormulaDecompiler.Decompile(ptgs, _workbook);
                    Logger.Info($"Decompiled formula -> {formula}");

                    cell.Formula = formula;

                    // 处理公式结果
                    if (IsDateTimeValue(result))
                    {
                        cell.Value = ExcelDateToDateTime(result);
                        cell.DataType = "d";
                    }
                    else if (IsErrorValue(result))
                    {
                        cell.Value = ParsingHelpers.GetErrorString((byte)result);
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
                        cell.Value = ParsingHelpers.GetErrorString((byte)result);
                        cell.DataType = "e";
                    }
                    else
                    {
                        cell.Value = result;
                        cell.DataType = "n";
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("解析公式时发生错误", ex);
                cell.Value = "#ERROR!";
                cell.DataType = "e";
            }
        }

        #endregion

        #region 辅助方法

        private static bool IsDateTimeValue(double value)
        {
            // Excel 日期时间范围：从1970-01-01到9999-12-31
            return value >= 25569 && value <= 730485;
        }

        private static bool IsErrorValue(double value)
        {
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

        private static DateTime ExcelDateToDateTime(double excelDate)
        {
            // Excel 日期时间是从 1900-01-01 开始的天数
            // 注意：Excel 使用 1900 年 2 月 29 日作为有效日期，即使 1900 年不是闰年
            DateTime excelBaseDate = new DateTime(1900, 1, 1);

            // 调整 1900 年闰年问题
            if (excelDate >= 60)
            {
                excelDate -= 1;
            }

            return excelBaseDate.AddDays(excelDate - 1);
        }

        private Font? GetFontByIndex(int fontIndex)
        {
            int normalizedIndex = ParsingHelpers.NormalizeBiffFontIndex(fontIndex);
            if (normalizedIndex >= 0 && normalizedIndex < _fonts.Count)
            {
                return _fonts[normalizedIndex];
            }
            return null;
        }

        private static string ReadBiffString(byte[] data, ref int offset)
        {
            if (offset >= data.Length) return string.Empty;

            byte flags = data[offset];
            bool isUnicode = (flags & 0x01) != 0;
            bool hasHighByte = (flags & 0x02) != 0;
            offset++;

            int charCount = data[offset];
            offset++;

            if (hasHighByte)
            {
                charCount += data[offset] << 8;
                offset++;
            }

            if (charCount == 0) return string.Empty;

            int byteCount = isUnicode ? charCount * 2 : charCount;
            if (offset + byteCount > data.Length)
                byteCount = data.Length - offset;

            string result = isUnicode
                ? System.Text.Encoding.Unicode.GetString(data, offset, byteCount)
                : System.Text.Encoding.ASCII.GetString(data, offset, byteCount);

            offset += byteCount;
            return result.TrimEnd('\0');
        }

        #endregion
    }
}
