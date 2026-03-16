using System;
using System.Collections.Generic;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 字体和XF格式解析器 - 处理FONT、XF、FORMAT记录
    /// </summary>
    public class FontXfParser
    {
        private readonly Workbook _workbook;
        private readonly List<Font> _fonts;
        private readonly List<Xf> _xfList;
        private readonly Dictionary<ushort, string> _formats;
        private readonly Func<int, Worksheet?, string?> _getColorFromPalette;
        private readonly Func<byte, string> _getBorderLineStyle;
        private readonly Func<byte, string> _getPatternType;

        public FontXfParser(
            Workbook workbook,
            List<Font> fonts,
            List<Xf> xfList,
            Dictionary<ushort, string> formats,
            Func<int, Worksheet?, string?> getColorFromPalette,
            Func<byte, string> getBorderLineStyle,
            Func<byte, string> getPatternType)
        {
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            _fonts = fonts ?? throw new ArgumentNullException(nameof(fonts));
            _xfList = xfList ?? throw new ArgumentNullException(nameof(xfList));
            _formats = formats ?? throw new ArgumentNullException(nameof(formats));
            _getColorFromPalette = getColorFromPalette ?? throw new ArgumentNullException(nameof(getColorFromPalette));
            _getBorderLineStyle = getBorderLineStyle ?? throw new ArgumentNullException(nameof(getBorderLineStyle));
            _getPatternType = getPatternType ?? throw new ArgumentNullException(nameof(getPatternType));
        }

        #region Worksheet Level Parsing

        /// <summary>
        /// 解析FONT记录 (0x0031) - 工作表级别
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
        /// 解析XF记录 (0x00E0) - 工作表级别
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

                switch (horizontalAlign)
                {
                    case 0: xf.HorizontalAlignment = "general"; break;
                    case 1: xf.HorizontalAlignment = "left"; break;
                    case 2: xf.HorizontalAlignment = "center"; break;
                    case 3: xf.HorizontalAlignment = "right"; break;
                    case 4: xf.HorizontalAlignment = "fill"; break;
                    case 5: xf.HorizontalAlignment = "justify"; break;
                    case 6: xf.HorizontalAlignment = "centerContinuous"; break;
                    case 7: xf.HorizontalAlignment = "distributed"; break;
                }

                switch (verticalAlign)
                {
                    case 0: xf.VerticalAlignment = "top"; break;
                    case 1: xf.VerticalAlignment = "center"; break;
                    case 2: xf.VerticalAlignment = "bottom"; break;
                    case 3: xf.VerticalAlignment = "justify"; break;
                    case 4: xf.VerticalAlignment = "distributed"; break;
                }

                // 解析缩进
                xf.Indent = (byte)((alignment & 0x0F00) >> 8);

                // 解析文本换行
                xf.WrapText = (alignment & 0x1000) != 0;

                // 解析边框 (偏移10-17)
                if (record.Data.Length >= 18)
                {
                    uint border1 = BitConverter.ToUInt32(record.Data, 10);
                    uint border2 = BitConverter.ToUInt32(record.Data, 14);

                    var border = new Border();
                    border.Left = _getBorderLineStyle((byte)(border1 & 0x0F));
                    border.Right = _getBorderLineStyle((byte)((border1 >> 4) & 0x0F));
                    border.Top = _getBorderLineStyle((byte)((border1 >> 8) & 0x0F));
                    border.Bottom = _getBorderLineStyle((byte)((border1 >> 12) & 0x0F));

                    border.LeftColor = _getColorFromPalette((int)((border1 >> 16) & 0x7F), worksheet);
                    border.RightColor = _getColorFromPalette((int)((border1 >> 23) & 0x7F), worksheet);

                    border.TopColor = _getColorFromPalette((int)(border2 & 0x7F), worksheet);
                    border.BottomColor = _getColorFromPalette((int)((border2 >> 7) & 0x7F), worksheet);
                    border.DiagonalColor = _getColorFromPalette((int)((border2 >> 14) & 0x7F), worksheet);
                    border.Diagonal = _getBorderLineStyle((byte)((border2 >> 21) & 0x0F));

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

                    var fill = new Fill();
                    fill.PatternType = _getPatternType(pattern);
                    fill.ForegroundColor = _getColorFromPalette(icvFore, worksheet);
                    fill.BackgroundColor = _getColorFromPalette(icvBack, worksheet);

                    _workbook.Fills.Add(fill);
                    xf.FillIndex = _workbook.Fills.Count + 1;
                }

                // 解析锁定和隐藏状态
                xf.IsLocked = (BitConverter.ToUInt16(record.Data, 26) & 0x0001) != 0;
                xf.IsHidden = (BitConverter.ToUInt16(record.Data, 26) & 0x0002) != 0;

                worksheet.Xfs.Add(xf);
            }
        }

        /// <summary>
        /// 解析FORMAT记录 (0x001E) - 工作表级别
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

        #endregion

        #region Global Level Parsing

        /// <summary>
        /// 解析FONT记录 (0x0031) - 全局级别
        /// </summary>
        public void ParseFontRecordToGlobal(BiffRecord record)
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
                string? resolved = _getColorFromPalette(font.ColorIndex, null);
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
        /// 解析XF记录 (0x00E0) - 全局级别
        /// </summary>
        public void ParseXfRecordToGlobal(BiffRecord record)
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
                        LeftColor = _getColorFromPalette((int)((border1 >> 16) & 0x7F), null),
                        RightColor = _getColorFromPalette((int)((border1 >> 23) & 0x7F), null),
                        TopColor = _getColorFromPalette((int)(border2 & 0x7F), null),
                        BottomColor = _getColorFromPalette((int)((border2 >> 7) & 0x7F), null),
                        DiagonalColor = _getColorFromPalette((int)((border2 >> 14) & 0x7F), null),
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
                            xf.BorderIndex = existingBorderIdx + 1;
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
                            ForegroundColor = _getColorFromPalette(icvFore, null),
                            BackgroundColor = _getColorFromPalette(icvBack, null)
                        };
                        int existingFillIdx = _workbook.Fills.FindIndex(f =>
                            f.PatternType == fill.PatternType &&
                            f.ForegroundColor == fill.ForegroundColor &&
                            f.BackgroundColor == fill.BackgroundColor);

                        if (existingFillIdx >= 0)
                        {
                            xf.FillIndex = existingFillIdx + 2;
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
        /// 解析FORMAT记录 (0x001E) - 全局级别
        /// </summary>
        public void ParseFormatRecordGlobal(BiffRecord record, Func<byte[], int, string> readBiffString)
        {
            byte[] data = record.GetAllData();
            if (data.Length >= 2)
            {
                ushort index = BitConverter.ToUInt16(data, 0);
                int offset = 2;
                if (offset < data.Length)
                {
                    _formats[index] = readBiffString(data, offset);
                }
            }
        }

        #endregion
    }
}
