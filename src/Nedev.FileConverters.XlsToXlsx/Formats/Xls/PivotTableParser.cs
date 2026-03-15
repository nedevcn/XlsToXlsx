using System;
using System.Collections.Generic;
using System.Linq;
using Nedev.FileConverters.XlsToXlsx;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 数据透视表解析器 - 处理数据透视表相关的BIFF记录
    /// </summary>
    public class PivotTableParser
    {
        // 当前解析状态
        private PivotTable? _currentPivotTable;
        private PivotField? _currentPivotField;

        /// <summary>
        /// 解析数据透视表记录
        /// </summary>
        public void ParsePivotTableRecord(BiffRecord record, Worksheet worksheet)
        {
            if (record.Data == null) return;

            switch (record.Id)
            {
                case (ushort)BiffRecordType.SXVIEW:
                    ParseSxViewRecord(record, worksheet);
                    break;

                case (ushort)BiffRecordType.SXVD:
                    ParseSxvdRecord(record);
                    break;

                case (ushort)BiffRecordType.SXDX:
                    ParseSxdxRecord(record);
                    break;

                case (ushort)BiffRecordType.SXVI:
                    ParseSxviRecord(record);
                    break;

                case (ushort)BiffRecordType.SXFIELD:
                    ParseSxfieldRecord(record);
                    break;

                case (ushort)BiffRecordType.SXPI:
                    ParseSxpiRecord(record);
                    break;
            }
        }

        /// <summary>
        /// 解析SXVIEW记录 (0x00B0) - 数据透视表视图定义
        /// </summary>
        private void ParseSxViewRecord(BiffRecord record, Worksheet worksheet)
        {
            if (record.Data.Length < 16) return;

            var pivotTable = new PivotTable();

            // 部分BIFF实现中前16字节含范围：offset 4-5 rwFirst, 6-7 rwLast, 8-9 colFirst, 10-11 colLast (0-based)
            if (record.Data.Length >= 12)
            {
                ushort rwFirst = BitConverter.ToUInt16(record.Data, 4);
                ushort rwLast = BitConverter.ToUInt16(record.Data, 6);
                ushort colFirst = BitConverter.ToUInt16(record.Data, 8);
                ushort colLast = BitConverter.ToUInt16(record.Data, 10);

                if (rwLast >= rwFirst && colLast >= colFirst && rwLast < 65535 && colLast < 256)
                {
                    pivotTable.DataSource = $"{ParsingHelpers.ColumnIndexToLetters1Based(colFirst + 1)}{rwFirst + 1}:{ParsingHelpers.ColumnIndexToLetters1Based(colLast + 1)}{rwLast + 1}";
                }
            }

            int offset = 16;
            // SXVIEW中的表名由一个两字节的字符数开始(BIFF8)
            if (offset + 2 <= record.Data.Length)
            {
                ushort nameLen = BitConverter.ToUInt16(record.Data, offset);
                offset += 2;
                pivotTable.Name = ReadBiffString(record.Data, ref offset, nameLen);
            }

            _currentPivotTable = pivotTable;
            worksheet.PivotTables.Add(pivotTable);
            Logger.Info($"开始解析数据透视表: {pivotTable.Name}");
        }

        /// <summary>
        /// 解析SXVD记录 (0x00B1) - 字段属性
        /// </summary>
        private void ParseSxvdRecord(BiffRecord record)
        {
            if (_currentPivotTable == null || record.Data.Length < 2) return;

            ushort sxaxis = BitConverter.ToUInt16(record.Data, 0);
            var field = new PivotField();

            // sxaxis: 0=no axis, 1=row, 2=col, 4=page, 8=data
            if ((sxaxis & 0x01) != 0) field.Type = "row";
            else if ((sxaxis & 0x02) != 0) field.Type = "column";
            else if ((sxaxis & 0x04) != 0) field.Type = "page";
            else if ((sxaxis & 0x08) != 0) field.Type = "data";

            switch (field.Type)
            {
                case "row":
                    _currentPivotTable.RowFields.Add(field);
                    break;
                case "column":
                    _currentPivotTable.ColumnFields.Add(field);
                    break;
                case "page":
                    _currentPivotTable.PageFields.Add(field);
                    break;
                case "data":
                    _currentPivotTable.DataFields.Add(field);
                    break;
            }

            _currentPivotField = field;
        }

        /// <summary>
        /// 解析SXDX记录 (0x00C6) - 数据字段的具体信息
        /// </summary>
        private void ParseSxdxRecord(BiffRecord record)
        {
            if (_currentPivotTable == null || _currentPivotTable.DataFields.Count == 0) return;

            var lastDataField = _currentPivotTable.DataFields[_currentPivotTable.DataFields.Count - 1];
            if (record.Data.Length < 4) return;

            ushort iSubtotal = BitConverter.ToUInt16(record.Data, 2);
            // 0=Sum, 1=Count, 2=Average, etc.
            string[] funcs = { "sum", "count", "average", "max", "min", "product", "countNums", "stdDev", "stdDevp", "var", "varp" };
            if (iSubtotal < funcs.Length)
            {
                lastDataField.Function = funcs[iSubtotal];
            }
        }

        /// <summary>
        /// 解析SXVI记录 (0x00B2) - 数据透视表项
        /// </summary>
        private void ParseSxviRecord(BiffRecord record)
        {
            if (_currentPivotField == null || record.Data == null || record.Data.Length < 2) return;

            ushort cItem = BitConverter.ToUInt16(record.Data, 0);
            int off = 2;

            for (int i = 0; i < cItem && off + 2 <= record.Data.Length; i++)
            {
                ushort cch = BitConverter.ToUInt16(record.Data, off);
                off += 2;
                if (cch == 0) continue;
                if (off + cch > record.Data.Length) break;

                string itemStr = ReadBiffString(record.Data, ref off, cch);
                _currentPivotField.Items.Add(itemStr);
            }
        }

        /// <summary>
        /// 解析SXFIELD记录 (0x00CA) - 字段名称及属性
        /// </summary>
        private void ParseSxfieldRecord(BiffRecord record)
        {
            if (_currentPivotTable == null || record.Data.Length < 4) return;

            int offset = 4;
            // BIFF8 SXFIELD name starts with 2-byte char count
            if (offset + 2 > record.Data.Length) return;

            ushort cch = BitConverter.ToUInt16(record.Data, offset);
            offset += 2;

            string fieldName = ReadBiffString(record.Data, ref offset, cch);

            // 回填到最后一个未命名的字段
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

        /// <summary>
        /// 解析SXPI记录 (0x00B7) - 页字段项
        /// </summary>
        private void ParseSxpiRecord(BiffRecord record)
        {
            if (_currentPivotField == null || _currentPivotField.Type != "page" || record.Data == null || record.Data.Length < 2) return;

            ushort cItem = BitConverter.ToUInt16(record.Data, 0);
            int off = 2;

            for (int i = 0; i < cItem && off + 2 <= record.Data.Length; i++)
            {
                ushort cch = BitConverter.ToUInt16(record.Data, off);
                off += 2;
                if (cch == 0) continue;
                if (off + cch > record.Data.Length) break;

                string itemStr = ReadBiffString(record.Data, ref off, cch);
                _currentPivotField.Items.Add(itemStr);
            }
        }

        /// <summary>
        /// 重置解析状态（用于开始新的工作表解析）
        /// </summary>
        public void ResetState()
        {
            _currentPivotTable = null;
            _currentPivotField = null;
        }

        #region 辅助方法

        private static string ReadBiffString(byte[] data, ref int offset, uint charCount)
        {
            if (charCount == 0 || offset >= data.Length)
                return string.Empty;

            int bytesToRead = Math.Min((int)charCount, data.Length - offset);
            string result = System.Text.Encoding.ASCII.GetString(data, offset, bytesToRead).TrimEnd('\0');
            offset += bytesToRead;
            return result;
        }

        #endregion
    }
}
