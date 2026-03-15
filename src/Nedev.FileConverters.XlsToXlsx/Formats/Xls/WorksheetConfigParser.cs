using System;
using Nedev.FileConverters.XlsToXlsx;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 工作表配置解析器 - 处理工作表级别的配置记录（窗口、页面设置、列信息等）
    /// </summary>
    public class WorksheetConfigParser
    {
        /// <summary>
        /// 解析COLINFO记录 (0x007D) - 列信息
        /// </summary>
        public void ParseColInfoRecord(BiffRecord record, Worksheet worksheet)
        {
            if (record.Data == null || record.Data.Length < 10)
                return;

            ushort firstCol = BitConverter.ToUInt16(record.Data, 0);
            ushort lastCol = BitConverter.ToUInt16(record.Data, 2);
            if (firstCol > lastCol)
                return;

            ushort width = BitConverter.ToUInt16(record.Data, 4);
            ushort xfIndex = BitConverter.ToUInt16(record.Data, 6);
            ushort options = BitConverter.ToUInt16(record.Data, 8);
            bool hidden = (options & 0x0001) != 0;

            worksheet.ColumnInfos.Add(new ColumnInfo
            {
                FirstColumn = firstCol,
                LastColumn = lastCol,
                Width = width,
                XfIndex = xfIndex,
                Hidden = hidden
            });
        }

        /// <summary>
        /// 解析WINDOW2记录 (0x023E) - 窗口设置
        /// </summary>
        public void ParseWindow2Record(BiffRecord record, Worksheet worksheet)
        {
            // WINDOW2记录: options(2) + ...
            // options bit 3: 是否冻结窗格 (fFrozen)
            if (record.Data == null || record.Data.Length < 2)
                return;

            ushort options = BitConverter.ToUInt16(record.Data, 0);
            bool isFrozen = (options & 0x0008) != 0;

            if (isFrozen && worksheet.FreezePane == null)
            {
                // 标记为冻结，具体位置由PANE记录设置
                worksheet.FreezePane = new FreezePane();
            }
        }

        /// <summary>
        /// 解析PANE记录 (0x0041) - 窗格分割信息
        /// </summary>
        public void ParsePaneRecord(BiffRecord record, Worksheet worksheet)
        {
            // PANE记录: x(2) + y(2) + topRow(2) + leftCol(2) + activePane(1)
            if (record.Data == null || record.Data.Length < 8)
                return;

            ushort x = BitConverter.ToUInt16(record.Data, 0); // 水平分割位置
            ushort y = BitConverter.ToUInt16(record.Data, 2); // 垂直分割位置
            ushort topRow = BitConverter.ToUInt16(record.Data, 4);
            ushort leftCol = BitConverter.ToUInt16(record.Data, 6);

            if (worksheet.FreezePane != null)
            {
                worksheet.FreezePane.ColSplit = x;
                worksheet.FreezePane.RowSplit = y;
                worksheet.FreezePane.TopRow = topRow + 1; // 转为1-based
                worksheet.FreezePane.LeftCol = leftCol + 1; // 转为1-based
            }
            else
            {
                worksheet.FreezePane = new FreezePane
                {
                    ColSplit = x,
                    RowSplit = y,
                    TopRow = topRow + 1,
                    LeftCol = leftCol + 1
                };
            }
        }

        /// <summary>
        /// 解析PAGESETUP记录 (0x00A1) - 页面设置
        /// </summary>
        public void ParsePageSetupRecord(BiffRecord record, Worksheet worksheet)
        {
            if (record.Data == null || record.Data.Length < 34)
                return;

            var ps = worksheet.PageSettings;
            ps.PaperSize = BitConverter.ToUInt16(record.Data, 0);
            ps.Scale = BitConverter.ToUInt16(record.Data, 2);
            ps.FitToWidth = BitConverter.ToUInt16(record.Data, 6);
            ps.FitToHeight = BitConverter.ToUInt16(record.Data, 8);

            ushort options = BitConverter.ToUInt16(record.Data, 10);
            ps.OrientationLandscape = (options & 0x0002) == 0;
            ps.UsePageNumbers = (options & 0x0001) != 0;
        }

        /// <summary>
        /// 解析HEADER记录 (0x0014) - 页眉
        /// </summary>
        public void ParseHeaderRecord(BiffRecord record, Worksheet worksheet)
        {
            byte[] headerData = record.GetAllData();
            if (headerData.Length > 0)
            {
                int hPos = 0;
                worksheet.PageSettings.Header = ReadBiffString(headerData, ref hPos);
            }
        }

        /// <summary>
        /// 解析FOOTER记录 (0x0015) - 页脚
        /// </summary>
        public void ParseFooterRecord(BiffRecord record, Worksheet worksheet)
        {
            byte[] footerData = record.GetAllData();
            if (footerData.Length > 0)
            {
                int fPos = 0;
                worksheet.PageSettings.Footer = ReadBiffString(footerData, ref fPos);
            }
        }

        /// <summary>
        /// 解析边距记录
        /// </summary>
        public void ParseMarginRecord(BiffRecord record, Worksheet worksheet, BiffRecordType marginType)
        {
            if (record.Data == null || record.Data.Length < 8)
                return;

            double margin = BitConverter.ToDouble(record.Data, 0);

            switch (marginType)
            {
                case BiffRecordType.LEFTMARGIN:
                    worksheet.PageSettings.LeftMargin = margin;
                    break;
                case BiffRecordType.RIGHTMARGIN:
                    worksheet.PageSettings.RightMargin = margin;
                    break;
                case BiffRecordType.TOPMARGIN:
                    worksheet.PageSettings.TopMargin = margin;
                    break;
                case BiffRecordType.BOTTOMMARGIN:
                    worksheet.PageSettings.BottomMargin = margin;
                    break;
            }
        }

        /// <summary>
        /// 解析居中记录
        /// </summary>
        public void ParseCenterRecord(BiffRecord record, Worksheet worksheet, bool isHorizontal)
        {
            if (record.Data == null || record.Data.Length < 2)
                return;

            bool center = BitConverter.ToUInt16(record.Data, 0) != 0;

            if (isHorizontal)
                worksheet.PageSettings.HorizontalCenter = center;
            else
                worksheet.PageSettings.VerticalCenter = center;
        }

        /// <summary>
        /// 解析默认列宽记录
        /// </summary>
        public void ParseDefColWidthRecord(BiffRecord record, Worksheet worksheet)
        {
            if (record.Data != null && record.Data.Length >= 2)
                worksheet.DefaultColumnWidth = BitConverter.ToUInt16(record.Data, 0);
        }

        /// <summary>
        /// 解析默认行高记录
        /// </summary>
        public void ParseDefaultRowHeightRecord(BiffRecord record, Worksheet worksheet)
        {
            if (record.Data != null && record.Data.Length >= 4)
                worksheet.DefaultRowHeight = BitConverter.ToUInt16(record.Data, 2) / 20.0;
        }

        /// <summary>
        /// 解析工作表保护记录
        /// </summary>
        public void ParseProtectRecord(BiffRecord record, Worksheet worksheet)
        {
            if (record.Data != null && record.Data.Length >= 2)
            {
                ushort flags = BitConverter.ToUInt16(record.Data, 0);
                worksheet.IsProtected = (flags & 0x0001) != 0;
            }
        }

        /// <summary>
        /// 解析密码记录
        /// </summary>
        public void ParsePasswordRecord(BiffRecord record, Worksheet worksheet)
        {
            if (record.Data != null && record.Data.Length >= 2)
            {
                ushort hash = BitConverter.ToUInt16(record.Data, 0);
                worksheet.SheetPasswordHash = hash.ToString("X4", System.Globalization.CultureInfo.InvariantCulture);
            }
        }

        #region 辅助方法

        private static string ReadBiffString(byte[] data, ref int offset)
        {
            if (offset >= data.Length)
                return string.Empty;

            int charCount = data[offset];
            if (charCount == 0 || offset + 1 >= data.Length)
                return string.Empty;

            byte flags = data[offset + 1];
            bool isUnicode = (flags & 0x01) != 0;

            int startOffset = offset + 2;
            int byteCount = isUnicode ? charCount * 2 : charCount;

            if (startOffset + byteCount > data.Length)
                byteCount = data.Length - startOffset;

            if (byteCount <= 0)
                return string.Empty;

            string result = isUnicode
                ? System.Text.Encoding.Unicode.GetString(data, startOffset, byteCount)
                : System.Text.Encoding.ASCII.GetString(data, startOffset, byteCount);

            offset = startOffset + byteCount;
            return result.TrimEnd('\0');
        }

        #endregion
    }
}
