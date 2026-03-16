using System;
using System.Collections.Generic;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 工作簿样式解析器 - 处理工作簿级别的样式记录（FONT、XF、FORMAT、PALETTE、BORDER、FILL）
    /// </summary>
    public class WorkbookStyleParser
    {
        private readonly Workbook _workbook;
        private readonly List<Font> _fonts;
        private readonly List<Xf> _xfList;
        private readonly Dictionary<ushort, string> _formats;
        private readonly Dictionary<int, string> _palette;
        private readonly Func<int, string?> _getColorFromPalette;
        private readonly Func<byte, string> _getBorderLineStyle;
        private readonly Func<byte, string> _getPatternType;

        public WorkbookStyleParser(
            Workbook workbook,
            List<Font> fonts,
            List<Xf> xfList,
            Dictionary<ushort, string> formats,
            Dictionary<int, string> palette,
            Func<int, string?> getColorFromPalette,
            Func<byte, string> getBorderLineStyle,
            Func<byte, string> getPatternType)
        {
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            _fonts = fonts ?? throw new ArgumentNullException(nameof(fonts));
            _xfList = xfList ?? throw new ArgumentNullException(nameof(xfList));
            _formats = formats ?? throw new ArgumentNullException(nameof(formats));
            _palette = palette ?? throw new ArgumentNullException(nameof(palette));
            _getColorFromPalette = getColorFromPalette ?? throw new ArgumentNullException(nameof(getColorFromPalette));
            _getBorderLineStyle = getBorderLineStyle ?? throw new ArgumentNullException(nameof(getBorderLineStyle));
            _getPatternType = getPatternType ?? throw new ArgumentNullException(nameof(getPatternType));
        }

        /// <summary>
        /// 解析FONT记录到全局字体列表
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
        /// 解析XF记录到全局XF列表
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
                        LeftColor = _getColorFromPalette((int)((border1 >> 16) & 0x7F)),
                        RightColor = _getColorFromPalette((int)((border1 >> 23) & 0x7F)),
                        TopColor = _getColorFromPalette((int)(border2 & 0x7F)),
                        BottomColor = _getColorFromPalette((int)((border2 >> 7) & 0x7F)),
                        DiagonalColor = _getColorFromPalette((int)((border2 >> 14) & 0x7F)),
                        Diagonal = _getBorderLineStyle((byte)((border2 >> 21) & 0x0F))
                    };

                    // 只有当边框不是全部为 none 时才添加或查找
                    if (border.Left != "none" || border.Right != "none" || border.Top != "none" || border.Bottom != "none" || border.Diagonal != "none")
                    {
                        int existingBorderIdx = _workbook.Borders.FindIndex(b =>
                            b.Left == border.Left && b.Right == border.Right && b.Top == border.Top && b.Bottom == border.Bottom &&
                            b.LeftColor == border.LeftColor && b.RightColor == border.RightColor && b.TopColor == border.TopColor && b.BottomColor == border.BottomColor);

                        if (existingBorderIdx >= 0)
                        {
                            xf.BorderIndex = existingBorderIdx + 1; // 0 是默认
                        }
                        else
                        {
                            _workbook.Borders.Add(border);
                            xf.BorderIndex = _workbook.Borders.Count;
                        }
                    }
                    else
                    {
                        xf.BorderIndex = 0;
                    }
                }

                // 解析填充 (偏移18-21)
                if (record.Data.Length >= 20)
                {
                    ushort fillData = BitConverter.ToUInt16(record.Data, 18);
                    byte pattern = (byte)(fillData & 0x3F);
                    int icvFore = (fillData >> 6) & 0x7F;
                    int icvBack = record.Data.Length > 20 ? (record.Data[20] & 0x7F) : 65;

                    if (pattern > 0)
                    {
                        var fill = new Fill
                        {
                            PatternType = _getPatternType(pattern),
                            ForegroundColor = _getColorFromPalette(icvFore),
                            BackgroundColor = _getColorFromPalette(icvBack)
                        };
                        int existingFillIdx = _workbook.Fills.FindIndex(f =>
                            f.PatternType == fill.PatternType &&
                            f.ForegroundColor == fill.ForegroundColor &&
                            f.BackgroundColor == fill.BackgroundColor);

                        if (existingFillIdx >= 0)
                        {
                            xf.FillIndex = existingFillIdx + 2; // 0 和 1 是默认
                        }
                        else
                        {
                            _workbook.Fills.Add(fill);
                            xf.FillIndex = _workbook.Fills.Count + 1;
                        }
                    }
                    else
                    {
                        xf.FillIndex = 0;
                    }
                }

                // 解析锁定和隐藏状态 (偏移26)
                if (record.Data.Length >= 28)
                {
                    ushort options = BitConverter.ToUInt16(record.Data, 26);
                    xf.IsLocked = (options & 0x0001) != 0;
                    xf.IsHidden = (options & 0x0002) != 0;
                }

                _xfList.Add(xf);
            }
        }

        /// <summary>
        /// 解析FORMAT记录到全局格式字典
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

        /// <summary>
        /// 解析PALETTE记录到全局调色板
        /// </summary>
        public void ParsePaletteRecord(BiffRecord record)
        {
            if (record.Data != null && record.Data.Length >= 2)
            {
                ushort count = BitConverter.ToUInt16(record.Data, 0);
                int offset = 2;
                for (int i = 0; i < count && offset + 4 <= record.Data.Length; i++)
                {
                    int colorIndex = 8 + i; // 调色板颜色从索引8开始
                    byte r = record.Data[offset];
                    byte g = record.Data[offset + 1];
                    byte b = record.Data[offset + 2];
                    _palette[colorIndex] = $"#{r:X2}{g:X2}{b:X2}";
                    offset += 4;
                }
            }
        }

        /// <summary>
        /// 解析BORDER记录到全局边框列表
        /// </summary>
        public void ParseBorderRecord(BiffRecord record)
        {
            if (record.Data != null && record.Data.Length >= 8)
            {
                var border = new Border
                {
                    Left = _getBorderLineStyle(record.Data[0]),
                    Right = _getBorderLineStyle(record.Data[1]),
                    Top = _getBorderLineStyle(record.Data[2]),
                    Bottom = _getBorderLineStyle(record.Data[3])
                };

                if (record.Data.Length >= 12)
                {
                    border.LeftColor = _getColorFromPalette(BitConverter.ToUInt16(record.Data, 4));
                    border.RightColor = _getColorFromPalette(BitConverter.ToUInt16(record.Data, 6));
                    border.TopColor = _getColorFromPalette(BitConverter.ToUInt16(record.Data, 8));
                    border.BottomColor = _getColorFromPalette(BitConverter.ToUInt16(record.Data, 10));
                }

                _workbook.Borders.Add(border);
            }
        }

        /// <summary>
        /// 解析FILL记录到全局填充列表
        /// </summary>
        public void ParseFillRecord(BiffRecord record)
        {
            if (record.Data != null && record.Data.Length >= 4)
            {
                var fill = new Fill
                {
                    PatternType = _getPatternType(record.Data[0]),
                    ForegroundColor = _getColorFromPalette(BitConverter.ToUInt16(record.Data, 1)),
                    BackgroundColor = _getColorFromPalette(BitConverter.ToUInt16(record.Data, 3))
                };

                _workbook.Fills.Add(fill);
            }
        }
    }
}
