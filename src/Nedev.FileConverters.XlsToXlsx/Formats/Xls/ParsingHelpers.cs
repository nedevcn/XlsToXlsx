using System;
using System.Collections.Generic;
using System.Text;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// XLS解析辅助工具类
    /// </summary>
    public static class ParsingHelpers
    {
        /// <summary>
        /// 将1-based列索引转换为列字母（1 → A, 2 → B, ..., 27 → AA）
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
        /// 将0-based列索引转换为列字母（0 → A, 1 → B, ..., 26 → AA）
        /// </summary>
        public static string ColumnIndexToLetters0Based(int columnIndex0)
        {
            if (columnIndex0 < 0) columnIndex0 = 0;
            return ColumnIndexToLetters1Based(columnIndex0 + 1);
        }

        /// <summary>
        /// 从字节数组读取BIFF字符串
        /// </summary>
        public static string ReadBiffStringFromBytes(byte[] data, ref int offset, uint charCount)
        {
            if (charCount == 0 || offset >= data.Length)
                return string.Empty;

            int bytesToRead = Math.Min((int)charCount, data.Length - offset);
            string result = Encoding.ASCII.GetString(data, offset, bytesToRead).TrimEnd('\0');
            offset += bytesToRead;
            return result;
        }

        /// <summary>
        /// 解码RK值为double
        /// </summary>
        public static double DecodeRKValue(int rkValue)
        {
            bool isInteger = (rkValue & 0x02) != 0;
            bool isDiv100 = (rkValue & 0x01) != 0;
            double value;

            if (isInteger)
            {
                value = rkValue >> 2;
            }
            else
            {
                byte[] bytes = BitConverter.GetBytes(rkValue & 0xFFFFFFFC);
                value = BitConverter.ToDouble(bytes, 0);
            }

            if (isDiv100) value /= 100.0;
            return value;
        }

        /// <summary>
        /// 获取边框线样式名称
        /// </summary>
        public static string GetBorderLineStyle(byte styleCode)
        {
            return styleCode switch
            {
                0 => "none",
                1 => "thin",
                2 => "medium",
                3 => "dashed",
                4 => "dotted",
                5 => "thick",
                6 => "double",
                7 => "hair",
                8 => "mediumDashed",
                9 => "dashDot",
                10 => "mediumDashDot",
                11 => "dashDotDot",
                12 => "mediumDashDotDot",
                13 => "slantDashDot",
                _ => "none"
            };
        }

        /// <summary>
        /// 获取填充图案类型名称
        /// </summary>
        public static string GetPatternType(byte patternId)
        {
            return patternId switch
            {
                0 => "none",
                1 => "solid",
                2 => "mediumGray",
                3 => "darkGray",
                4 => "lightGray",
                5 => "darkHorizontal",
                6 => "darkVertical",
                7 => "darkDown",
                8 => "darkUp",
                9 => "darkGrid",
                10 => "darkTrellis",
                11 => "lightHorizontal",
                12 => "lightVertical",
                13 => "lightDown",
                14 => "lightUp",
                15 => "lightGrid",
                16 => "lightTrellis",
                17 => "gray125",
                18 => "gray0625",
                _ => "none"
            };
        }

        /// <summary>
        /// 标准化BIFF字体索引（BIFF保留字体槽4）
        /// </summary>
        public static int NormalizeBiffFontIndex(int fontIndex)
        {
            if (fontIndex < 0) return 0;
            return fontIndex > 4 ? fontIndex - 1 : fontIndex;
        }

        /// <summary>
        /// 解析BIFF错误代码为错误字符串
        /// </summary>
        public static string GetErrorString(byte errCode)
        {
            return errCode switch
            {
                0x00 => "#NULL!",
                0x07 => "#DIV/0!",
                0x0F => "#VALUE!",
                0x17 => "#REF!",
                0x1D => "#NAME?",
                0x24 => "#NUM!",
                0x2A => "#N/A",
                _ => "#UNKNOWN!"
            };
        }
    }
}
