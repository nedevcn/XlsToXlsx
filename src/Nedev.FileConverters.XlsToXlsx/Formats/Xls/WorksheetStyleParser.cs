using System;
using System.Collections.Generic;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 工作表样式解析器 - 处理工作表级别的样式记录（FONT、XF、FORMAT、PALETTE）
    /// </summary>
    public class WorksheetStyleParser
    {
        private readonly Workbook _workbook;
        private readonly Dictionary<ushort, string> _formats;
        private readonly Func<int, Worksheet?, string?> _getColorFromPalette;
        private readonly Func<byte, string> _getBorderLineStyle;
        private readonly Func<byte, string> _getPatternType;

        public WorksheetStyleParser(
            Workbook workbook,
            Dictionary<ushort, string> formats,
            Func<int, Worksheet?, string?> getColorFromPalette,
            Func<byte, string> getBorderLineStyle,
            Func<byte, string> getPatternType)
        {
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            _formats = formats ?? throw new ArgumentNullException(nameof(formats));
            _getColorFromPalette = getColorFromPalette ?? throw new ArgumentNullException(nameof(getColorFromPalette));
            _getBorderLineStyle = getBorderLineStyle ?? throw new ArgumentNullException(nameof(getBorderLineStyle));
            _getPatternType = getPatternType ?? throw new ArgumentNullException(nameof(getPatternType));
        }

        /// <summary>
        /// 解析FORMAT记录（工作表级别）
        /// </summary>
        public void ParseFormatRecord(BiffRecord record)
        {
            byte[] data = record.GetAllData();
            if (data != null && data.Length >= 18)
            {
                ushort formatIndex = BitConverter.ToUInt16(data, 0);
                byte formatLength = data[2];
                if (data.Length >= 3 + formatLength)
                {
                    string formatString = System.Text.Encoding.ASCII.GetString(data, 3, formatLength);
                    _formats[formatIndex] = formatString;
                }
            }
        }

        /// <summary>
        /// 解析FONT记录（工作表级别）
        /// </summary>
        public void ParseFontRecord(BiffRecord record, Worksheet worksheet)
        {
            byte[] data = record.GetAllData();
            if (data != null && data.Length >= 48)
            {
                var font = new Font();
                font.Height = BitConverter.ToInt16(data, 0);
                font.IsBold = (BitConverter.ToUInt16(data, 2) & 0x0001) != 0;
                font.IsItalic = (BitConverter.ToUInt16(data, 2) & 0x0002) != 0;
                font.IsUnderline = (BitConverter.ToUInt16(data, 2) & 0x0004) != 0;
                font.IsStrikethrough = (BitConverter.ToUInt16(data, 2) & 0x0008) != 0;
                font.ColorIndex = BitConverter.ToUInt16(data, 6);
                string? wsColor = _getColorFromPalette(font.ColorIndex, worksheet);
                font.Color = string.IsNullOrEmpty(wsColor) ? null : wsColor.Replace("#", "");
                font.Name = System.Text.Encoding.ASCII.GetString(data, 40, data.Length - 40).TrimEnd('\0');

                worksheet.Fonts.Add(font);
            }
        }

        /// <summary>
        /// 解析XF记录（工作表级别）
        /// </summary>
        public void ParseXfRecord(BiffRecord record, Worksheet worksheet)
        {
            if (record.Data != null && record.Data.Length >= 28)
            {
                var xf = new Xf();
                xf.FontIndex = BitConverter.ToUInt16(record.Data, 0);
                xf.NumberFormatIndex = BitConverter.ToUInt16(record.Data, 2);
                xf.CellFormatIndex = BitConverter.ToUInt16(record.Data, 4);

                // 解析对齐方式
                ushort alignment = BitConverter.ToUInt16(record.Data, 6);
                byte horizontalAlign = (byte)((alignment & 0x000F) >> 0);
                byte verticalAlign = (byte)((alignment & 0x00F0) >> 4);

                xf.HorizontalAlignment = horizontalAlign switch
                {
                    0 => "general",
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
                    0 => "top",
                    1 => "center",
                    2 => "bottom",
                    3 => "justify",
                    4 => "distributed",
                    _ => "top"
                };

                // 解析缩进
                xf.Indent = (byte)((alignment & 0x0F00) >> 8);

                // 解析文本换行
                xf.WrapText = (alignment & 0x1000) != 0;

                // 解析边框 (偏移10-17)
                if (record.Data.Length >= 18)
                {
                    uint border1 = BitConverter.ToUInt32(record.Data, 10);
                    uint border2 = BitConverter.ToUInt32(record.Data, 14);

                    var border = new Border
                    {
                        Left = _getBorderLineStyle((byte)(border1 & 0x0F)),
                        Right = _getBorderLineStyle((byte)((border1 >> 4) & 0x0F)),
                        Top = _getBorderLineStyle((byte)((border1 >> 8) & 0x0F)),
                        Bottom = _getBorderLineStyle((byte)((border1 >> 12) & 0x0F)),
                        LeftColor = _getColorFromPalette((int)((border1 >> 16) & 0x7F), worksheet),
                        RightColor = _getColorFromPalette((int)((border1 >> 23) & 0x7F), worksheet),
                        TopColor = _getColorFromPalette((int)(border2 & 0x7F), worksheet),
                        BottomColor = _getColorFromPalette((int)((border2 >> 7) & 0x7F), worksheet),
                        DiagonalColor = _getColorFromPalette((int)((border2 >> 14) & 0x7F), worksheet),
                        Diagonal = _getBorderLineStyle((byte)((border2 >> 21) & 0x0F))
                    };

                    // 添加到全局列表并分配索引
                    _workbook.Borders.Add(border);
                    xf.BorderIndex = _workbook.Borders.Count - 1;
                }

                // 解析填充 (偏移18-21)
                if (record.Data.Length >= 22)
                {
                    ushort fillData = BitConverter.ToUInt16(record.Data, 18);
                    byte pattern = (byte)(fillData & 0x3F);
                    int icvFore = (fillData >> 6) & 0x7F;
                    int icvBack = record.Data.Length > 20 ? (record.Data[20] & 0x7F) : 65;

                    var fill = new Fill
                    {
                        PatternType = _getPatternType(pattern),
                        ForegroundColor = _getColorFromPalette(icvFore, worksheet),
                        BackgroundColor = _getColorFromPalette(icvBack, worksheet)
                    };

                    _workbook.Fills.Add(fill);
                    xf.FillIndex = _workbook.Fills.Count + 1; // 2-based: 0=none, 1=gray125, 2+=workbook.Fills
                }

                // 解析锁定和隐藏状态
                xf.IsLocked = (BitConverter.ToUInt16(record.Data, 26) & 0x0001) != 0;
                xf.IsHidden = (BitConverter.ToUInt16(record.Data, 26) & 0x0002) != 0;

                worksheet.Xfs.Add(xf);
            }
        }

        /// <summary>
        /// 解析PALETTE记录（工作表级别）
        /// </summary>
        public void ParsePaletteRecord(BiffRecord record, Worksheet worksheet)
        {
            if (record.Data != null && record.Data.Length >= 2)
            {
                ushort count = BitConverter.ToUInt16(record.Data, 0);
                int offset = 2;
                for (int i = 0; i < count && offset + 4 <= record.Data.Length; i++)
                {
                    byte r = record.Data[offset];
                    byte g = record.Data[offset + 1];
                    byte b = record.Data[offset + 2];
                    // 调色板颜色从索引8开始
                    worksheet.Palette[8 + i] = $"#{r:X2}{g:X2}{b:X2}";
                    offset += 4;
                }
            }
        }
    }
}
