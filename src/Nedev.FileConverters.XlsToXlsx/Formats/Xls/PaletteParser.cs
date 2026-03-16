using System;
using System.Collections.Generic;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 调色板解析器 - 处理PALETTE记录
    /// </summary>
    public class PaletteParser
    {
        /// <summary>
        /// 解析PALETTE记录 (0x0092) - 工作表级别
        /// </summary>
        public void ParsePaletteRecord(BiffRecord record, Worksheet worksheet)
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

        /// <summary>
        /// 解析PALETTE记录 (0x0092) - 全局级别
        /// </summary>
        public void ParsePaletteRecordGlobal(BiffRecord record, Dictionary<int, string> palette)
        {
            if (record.Data != null && record.Data.Length >= 4)
            {
                int count = BitConverter.ToUInt16(record.Data, 0);
                for (int i = 0; i < count && (2 + i * 4 + 4 <= record.Data.Length); i++)
                {
                    byte r = record.Data[2 + i * 4];
                    byte g = record.Data[2 + i * 4 + 1];
                    byte b = record.Data[2 + i * 4 + 2];
                    palette[8 + i] = $"#{r:X2}{g:X2}{b:X2}";
                }
            }
        }
    }
}
