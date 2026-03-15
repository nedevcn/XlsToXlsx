using System;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 超链接解析器 - 处理超链接相关的BIFF记录
    /// </summary>
    public class HyperlinkParser
    {
        /// <summary>
        /// 解析HYPERLINK记录 (0x01B8) - 超链接
        /// </summary>
        public void ParseHyperlinkRecord(BiffRecord record, Worksheet worksheet)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 20)
                return;
            var hyperlink = new Hyperlink();
            ushort firstRow = BitConverter.ToUInt16(data, 0);
            ushort lastRow = BitConverter.ToUInt16(data, 2);
            ushort firstCol = BitConverter.ToUInt16(data, 4);
            ushort lastCol = BitConverter.ToUInt16(data, 6);
            hyperlink.Range = $"{ParsingHelpers.ColumnIndexToLetters1Based(firstCol + 1)}{firstRow + 1}:{ParsingHelpers.ColumnIndexToLetters1Based(lastCol + 1)}{lastRow + 1}";
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
    }
}
