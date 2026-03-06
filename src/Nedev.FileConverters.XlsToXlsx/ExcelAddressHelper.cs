using System;

namespace Nedev.FileConverters.XlsToXlsx
{
    /// <summary>
    /// 公共的 Excel 地址/列号 工具函数，集中管理列索引与列字母之间的转换逻辑。
    /// </summary>
    public static class ExcelAddressHelper
    {
        /// <summary>
        /// 将 1-based 列索引转换为列字母（1 → A, 2 → B, ..., 27 → AA）。
        /// 超出范围时自动夹在 [1, int.MaxValue]。
        /// </summary>
        public static string ColumnIndexToLetters1Based(int columnIndex)
        {
            if (columnIndex < 1) columnIndex = 1;
            int dividend = columnIndex;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }

        /// <summary>
        /// 将 0-based 列索引转换为列字母（0 → A, 1 → B, ..., 26 → AA）。
        /// </summary>
        public static string ColumnIndexToLetters0Based(int columnIndex0)
        {
            if (columnIndex0 < 0) columnIndex0 = 0;
            return ColumnIndexToLetters1Based(columnIndex0 + 1);
        }

        /// <summary>
        /// 将列字母转换为 0-based 列索引（A → 0, B → 1, ..., AA → 26）。
        /// 非法输入返回 0。
        /// </summary>
        public static int LettersToColumnIndex0Based(string letters)
        {
            if (string.IsNullOrEmpty(letters))
                return 0;

            int index = 0;
            foreach (char c in letters.ToUpperInvariant())
            {
                if (c < 'A' || c > 'Z')
                    continue;
                index = index * 26 + (c - 'A' + 1);
            }

            return Math.Max(0, index - 1);
        }
    }
}

