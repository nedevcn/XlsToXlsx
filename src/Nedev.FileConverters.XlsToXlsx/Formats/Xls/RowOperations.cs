using System;
using System.Collections.Generic;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 行操作辅助类 - 提供行相关的操作功能
    /// </summary>
    public static class RowOperations
    {
        /// <summary>
        /// 获取或创建工作表中的行
        /// </summary>
        /// <param name="worksheet">工作表</param>
        /// <param name="currentRow">当前行引用（用于缓存优化）</param>
        /// <param name="rowIndex">行索引（1-based）</param>
        /// <returns>找到或新创建的行</returns>
        public static Row GetOrCreateRow(Worksheet worksheet, ref Row? currentRow, int rowIndex)
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

        /// <summary>
        /// 获取或创建行（无currentRow缓存版本）
        /// </summary>
        /// <param name="worksheet">工作表</param>
        /// <param name="rowIndex">行索引（1-based）</param>
        /// <returns>找到或新创建的行</returns>
        public static Row GetOrCreateRow(Worksheet worksheet, int rowIndex)
        {
            Row? dummy = null;
            return GetOrCreateRow(worksheet, ref dummy, rowIndex);
        }

        /// <summary>
        /// 查找行（不创建）
        /// </summary>
        /// <param name="worksheet">工作表</param>
        /// <param name="rowIndex">行索引（1-based）</param>
        /// <returns>找到的行，或null</returns>
        public static Row? FindRow(Worksheet worksheet, int rowIndex)
        {
            for (int i = worksheet.Rows.Count - 1; i >= 0; i--)
            {
                if (worksheet.Rows[i].RowIndex == rowIndex)
                    return worksheet.Rows[i];
            }
            return null;
        }

        /// <summary>
        /// 确保行的单元格列表已初始化
        /// </summary>
        /// <param name="row">行</param>
        /// <param name="capacity">初始容量</param>
        public static void EnsureCellsCapacity(Row row, int capacity = 20)
        {
            if (row.Cells == null)
            {
                row.Cells = new List<Cell>(capacity);
            }
            else if (row.Cells.Capacity < capacity)
            {
                row.Cells.Capacity = capacity;
            }
        }
    }
}
