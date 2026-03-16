using System;
using System.Collections.Generic;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 全局样式解析器 - 处理工作簿级别的样式记录（FONT、XF、FORMAT）
    /// </summary>
    public class GlobalStyleParser
    {
        private readonly List<Font> _fonts;
        private readonly List<Xf> _xfList;
        private readonly Dictionary<ushort, string> _formats;
        private readonly Func<int, string?> _getColorFromPalette;

        public GlobalStyleParser(
            List<Font> fonts,
            List<Xf> xfList,
            Dictionary<ushort, string> formats,
            Func<int, string?> getColorFromPalette,
            Func<byte, string> getBorderLineStyle,
            Func<byte, string> getPatternType)
        {
            _fonts = fonts ?? throw new ArgumentNullException(nameof(fonts));
            _xfList = xfList ?? throw new ArgumentNullException(nameof(xfList));
            _formats = formats ?? throw new ArgumentNullException(nameof(formats));
            _getColorFromPalette = getColorFromPalette ?? throw new ArgumentNullException(nameof(getColorFromPalette));
        }

        /// <summary>
        /// 解析FONT记录（全局级别）
        /// </summary>
        public void ParseFontRecord(BiffRecord record)
        {
            byte[] data = record.GetAllData();
            if (data != null && data.Length >= 14)
            {
                var font = new Font();
                font.Height = BitConverter.ToInt16(data, 0);
                ushort grbit = BitConverter.ToUInt16(data, 2);
                font.IsBold = BitConverter.ToUInt16(data, 6) >= 700;
                font.IsItalic = (grbit & 0x0002) != 0;
                font.IsUnderline = (data[10]) != 0;
                font.IsStrikethrough = (grbit & 0x0008) != 0;
                font.ColorIndex = BitConverter.ToUInt16(data, 4);
                string? resolved = _getColorFromPalette(font.ColorIndex);
                font.Color = string.IsNullOrEmpty(resolved) ? null : resolved.Replace("#", "");

                int nameOffset = 14;
                if (data.Length > nameOffset)
                {
                    byte len = data[nameOffset];
                    if (data.Length > nameOffset + 1)
                    {
                        byte opt = data[nameOffset + 1];
                        bool isUni = (opt & 0x01) != 0;
                        if (isUni)
                        {
                            font.Name = System.Text.Encoding.Unicode.GetString(data, nameOffset + 2, Math.Min(len * 2, data.Length - nameOffset - 2));
                        }
                        else
                        {
                            font.Name = System.Text.Encoding.ASCII.GetString(data, nameOffset + 2, Math.Min(len, data.Length - nameOffset - 2));
                        }
                    }
                }
                _fonts.Add(font);
            }
        }

        /// <summary>
        /// 解析XF记录（全局级别）
        /// </summary>
        public void ParseXfRecord(BiffRecord record)
        {
            if (record.Data != null && record.Data.Length >= 20)
            {
                var xf = new Xf();
                xf.FontIndex = BitConverter.ToUInt16(record.Data, 0);
                xf.NumberFormatIndex = BitConverter.ToUInt16(record.Data, 2);

                // 解析对齐方式 (offset 6-9)
                ushort alignment = BitConverter.ToUInt16(record.Data, 6);
                byte horizontalAlign = (byte)(alignment & 0x07);
                byte verticalAlign = (byte)((alignment & 0x70) >> 4);

                xf.HorizontalAlignment = horizontalAlign switch
                {
                    1 => "left",
                    2 => "center",
                    3 => "right",
                    4 => "fill",
                    5 => "justify",
                    6 => "centerContinuous",
                    7 => "distributed",
                    _ => "general"
                };
                xf.VerticalAlignment = verticalAlign switch
                {
                    1 => "center",
                    2 => "bottom",
                    3 => "justify",
                    4 => "distributed",
                    _ => "top"
                };

                xf.WrapText = (alignment & 0x08) != 0;
                xf.Indent = (byte)((alignment >> 8) & 0x0F);

                // 解析锁定和保护 (偏移9)
                if (record.Data.Length >= 10)
                {
                    byte protection = record.Data[9];
                    xf.IsLocked = (protection & 0x01) != 0;
                    xf.IsHidden = (protection & 0x02) != 0;
                }

                _xfList.Add(xf);
            }
        }

        /// <summary>
        /// 解析FORMAT记录（全局级别）
        /// </summary>
        public void ParseFormatRecord(BiffRecord record)
        {
            byte[] data = record.GetAllData();
            if (data != null && data.Length >= 4)
            {
                ushort index = BitConverter.ToUInt16(data, 0);
                int offset = 2;
                string formatString = RichTextParser.ReadBiffString(data, ref offset);
                _formats[index] = formatString;
            }
        }
    }
}
