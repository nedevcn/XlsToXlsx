using System;
using System.Text;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 批注解析器 - 处理批注相关的BIFF记录
    /// </summary>
    public class CommentParser
    {
        /// <summary>
        /// 解析NOTE记录 (0x001C) - 批注
        /// </summary>
        public void ParseCommentRecord(BiffRecord record, Worksheet worksheet)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 12)
                return;
            var comment = new Comment();
            ushort row = BitConverter.ToUInt16(data, 0);
            ushort col = BitConverter.ToUInt16(data, 2);
            comment.RowIndex = row + 1;
            comment.ColumnIndex = col + 1;
            if (data.Length >= 14)
            {
                byte authorLength = data[12];
                if (authorLength > 0 && data.Length >= 13 + authorLength)
                {
                    comment.Author = Encoding.ASCII.GetString(data, 13, authorLength);
                    int textOffset = 13 + authorLength;
                    if (data.Length > textOffset)
                        comment.Text = Encoding.ASCII.GetString(data, textOffset, data.Length - textOffset);
                }
            }
            worksheet.Comments.Add(comment);
        }
    }
}
