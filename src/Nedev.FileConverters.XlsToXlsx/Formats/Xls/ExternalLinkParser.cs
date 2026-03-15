using System;
using Nedev.FileConverters.XlsToXlsx;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 外部链接解析器 - 处理外部工作簿引用相关的BIFF记录
    /// </summary>
    public class ExternalLinkParser
    {
        /// <summary>
        /// 解析EXTERNBOOK记录 (0x01AE) - 外部工作簿引用
        /// </summary>
        public void ParseExternBookRecord(BiffRecord record, Workbook workbook)
        {
            if (record.Data == null || record.Data.Length < 4) return;

            ushort count = BitConverter.ToUInt16(record.Data, 0);
            ushort type = BitConverter.ToUInt16(record.Data, 2);

            var extBook = new ExternalBook();

            if (type == 0x0401)
            {
                extBook.IsSelf = true;
                Logger.Info("找到内部引用 (Self SUPBOOK)");
            }
            else if (type == 0x3A01)
            {
                extBook.IsAddIn = true;
                Logger.Info("找到 Add-In 引用");
            }
            else
            {
                // 外部文件路径 + 工作表名称列表
                int offset = 4;
                if (offset >= record.Data.Length)
                {
                    workbook.ExternalBooks.Add(extBook);
                    return;
                }

                byte cchPath = record.Data[offset];
                offset++;
                extBook.FileName = ReadBiffString(record.Data, ref offset, cchPath);
                Logger.Info($"找到外部工作簿引用: {extBook.FileName}");

                // 后续count个工作表名称
                for (int i = 0; i < count && offset < record.Data.Length; i++)
                {
                    if (offset + 2 > record.Data.Length) break;
                    byte cchSheet = record.Data[offset];
                    offset++;
                    string sheetName = ReadBiffString(record.Data, ref offset, cchSheet);
                    extBook.SheetNames.Add(sheetName);
                }
            }

            workbook.ExternalBooks.Add(extBook);
        }

        /// <summary>
        /// 解析EXTERNSHEET记录 (0x017) - 外部工作表引用映射
        /// </summary>
        public void ParseExternSheetRecord(BiffRecord record, Workbook workbook)
        {
            if (record.Data == null || record.Data.Length < 2) return;

            ushort count = BitConverter.ToUInt16(record.Data, 0);

            for (int i = 0; i < count; i++)
            {
                int offset = 2 + i * 6;
                if (offset + 6 > record.Data.Length) break;

                var extSheet = new ExternalSheet
                {
                    ExternalBookIndex = BitConverter.ToUInt16(record.Data, offset),
                    FirstSheetIndex = BitConverter.ToInt16(record.Data, offset + 2),
                    LastSheetIndex = BitConverter.ToInt16(record.Data, offset + 4)
                };
                workbook.ExternalSheets.Add(extSheet);
            }

            Logger.Info($"解析 EXTERNSHEET, 共有 {count} 个引用映射");
        }

        /// <summary>
        /// 解析EXTERNALNAME记录 (0x023) - 外部名称引用
        /// </summary>
        public void ParseExternalNameRecord(BiffRecord record, Workbook workbook)
        {
            if (record.Data == null || record.Data.Length < 6) return;

            // BIFF8 EXTERNALNAME
            ushort options = BitConverter.ToUInt16(record.Data, 0);
            byte nameLen = record.Data[3];

            int offset = 6;
            string name = ReadBiffString(record.Data, ref offset, nameLen);

            if (workbook.ExternalBooks.Count > 0)
            {
                workbook.ExternalBooks[workbook.ExternalBooks.Count - 1].ExternalNames.Add(name);
            }

            Logger.Info($"找到外部工作簿名称引用: {name}");
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
