using System;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 多单元格记录解析器 - 处理MULRK和MULBLANK等批量单元格记录
    /// </summary>
    public class MultiCellParser
    {
        /// <summary>
        /// 解析MULRK记录 (0x00BD) - 多单元格RK值
        /// </summary>
        public void ParseMulRkRecord(BiffRecord record, ref Row currentRow, Worksheet worksheet)
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

        /// <summary>
        /// 解析MULBLANK记录 (0x00BE) - 多单元格空白
        /// </summary>
        public void ParseMulBlankRecord(BiffRecord record, ref Row currentRow, Worksheet worksheet)
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

        /// <summary>
        /// 解码RK值为double
        /// </summary>
        private static double DecodeRKValue(int rkValue)
        {
            // RK值编码：整数部分或浮点数
            if ((rkValue & 0x02) != 0)
            {
                // 整数编码
                return rkValue >> 2;
            }
            else
            {
                // 浮点数编码 - 将RK值转换为double
                byte[] bytes = BitConverter.GetBytes((long)rkValue << 34);
                return BitConverter.ToDouble(bytes, 0);
            }
        }

        /// <summary>
        /// 获取或创建行
        /// </summary>
        private static Row GetOrCreateRow(Worksheet worksheet, ref Row currentRow, int rowIndex)
        {
            if (currentRow != null && currentRow.RowIndex == rowIndex)
                return currentRow;

            // 查找现有行
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
    }
}
