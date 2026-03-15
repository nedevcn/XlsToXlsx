using System;
using System.Collections.Generic;
using System.Linq;
using Nedev.FileConverters.XlsToXlsx;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 样式解析器 - 处理字体、XF、边框、填充、调色板等样式相关记录
    /// </summary>
    public class StyleParser
    {
        private readonly Workbook _workbook;
        private readonly Dictionary<int, string> _palette;
        private readonly Dictionary<ushort, string> _formats;
        private readonly List<Font> _fonts;
        private readonly List<Xf> _xfList;

        public StyleParser(
            Workbook workbook,
            Dictionary<int, string> palette,
            Dictionary<ushort, string> formats,
            List<Font> fonts,
            List<Xf> xfList)
        {
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            _palette = palette ?? throw new ArgumentNullException(nameof(palette));
            _formats = formats ?? throw new ArgumentNullException(nameof(formats));
            _fonts = fonts ?? throw new ArgumentNullException(nameof(fonts));
            _xfList = xfList ?? throw new ArgumentNullException(nameof(xfList));
        }

        /// <summary>
        /// 解析字体记录 (FONT)
        /// </summary>
        public void ParseFontRecord(BiffRecord record)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 46) return;

            var font = new Font
            {
                Height = BitConverter.ToInt16(data, 0),
                IsBold = (BitConverter.ToUInt16(data, 2) & 0x0001) != 0,
                IsItalic = (BitConverter.ToUInt16(data, 2) & 0x0002) != 0,
                IsUnderline = (BitConverter.ToUInt16(data, 2) & 0x0004) != 0,
                IsStrikethrough = (BitConverter.ToUInt16(data, 2) & 0x0008) != 0,
                ColorIndex = BitConverter.ToUInt16(data, 6)
            };

            font.Color = GetColorFromPalette(font.ColorIndex)?.Replace("#", "");

            // 字体名称从偏移46开始
            if (data.Length > 46)
            {
                int nameOffset = 46;
                font.Name = ReadBiffString(data, ref nameOffset);
            }

            _fonts.Add(font);
        }

        /// <summary>
        /// 解析XF记录 (扩展格式)
        /// </summary>
        public void ParseXfRecord(BiffRecord record)
        {
            if (record.Data == null || record.Data.Length < 20) return;

            var xf = new Xf
            {
                FontIndex = BitConverter.ToUInt16(record.Data, 0),
                NumberFormatIndex = BitConverter.ToUInt16(record.Data, 2),
                CellFormatIndex = BitConverter.ToUInt16(record.Data, 4)
            };

            // 解析对齐方式
            if (record.Data.Length >= 8)
            {
                ushort alignment = BitConverter.ToUInt16(record.Data, 6);
                xf.HorizontalAlignment = ParseHorizontalAlignment((byte)(alignment & 0x000F));
                xf.VerticalAlignment = ParseVerticalAlignment((byte)((alignment & 0x00F0) >> 4));
                xf.Indent = (byte)((alignment & 0x0F00) >> 8);
                xf.WrapText = (alignment & 0x1000) != 0;
            }

            // 解析边框
            if (record.Data.Length >= 18)
            {
                ParseXfBorders(record.Data, xf);
            }

            // 解析填充
            if (record.Data.Length >= 20)
            {
                ParseXfFill(record.Data, xf);
            }

            // 解析锁定和隐藏状态
            if (record.Data.Length >= 28)
            {
                ushort options = BitConverter.ToUInt16(record.Data, 26);
                xf.IsLocked = (options & 0x0001) != 0;
                xf.IsHidden = (options & 0x0002) != 0;
            }

            _xfList.Add(xf);
        }

        /// <summary>
        /// 解析格式记录 (FORMAT)
        /// </summary>
        public void ParseFormatRecord(BiffRecord record)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 2) return;

            ushort index = BitConverter.ToUInt16(data, 0);
            int offset = 2;
            if (offset < data.Length)
            {
                _formats[index] = ReadBiffString(data, ref offset);
            }
        }

        /// <summary>
        /// 解析调色板记录 (PALETTE)
        /// </summary>
        public void ParsePaletteRecord(BiffRecord record)
        {
            if (record.Data == null || record.Data.Length < 4) return;

            int count = BitConverter.ToUInt16(record.Data, 0);
            for (int i = 0; i < count && (2 + i * 4 + 4 <= record.Data.Length); i++)
            {
                byte r = record.Data[2 + i * 4];
                byte g = record.Data[2 + i * 4 + 1];
                byte b = record.Data[2 + i * 4 + 2];
                _palette[8 + i] = $"#{r:X2}{g:X2}{b:X2}";
            }
        }

        /// <summary>
        /// 解析全局边框记录 (BORDER)
        /// </summary>
        public void ParseBorderRecord(BiffRecord record)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 8) return;

            uint border1 = BitConverter.ToUInt32(data, 0);
            uint border2 = BitConverter.ToUInt32(data, 4);

            var border = new Border
            {
                Left = ParsingHelpers.GetBorderLineStyle((byte)(border1 & 0x0F)),
                Right = ParsingHelpers.GetBorderLineStyle((byte)((border1 >> 4) & 0x0F)),
                Top = ParsingHelpers.GetBorderLineStyle((byte)((border1 >> 8) & 0x0F)),
                Bottom = ParsingHelpers.GetBorderLineStyle((byte)((border1 >> 12) & 0x0F)),
                LeftColor = GetColorFromPalette((int)((border1 >> 16) & 0x7F)),
                RightColor = GetColorFromPalette((int)((border1 >> 23) & 0x7F)),
                TopColor = GetColorFromPalette((int)(border2 & 0x7F)),
                BottomColor = GetColorFromPalette((int)((border2 >> 7) & 0x7F)),
                DiagonalColor = GetColorFromPalette((int)((border2 >> 14) & 0x7F)),
                Diagonal = ParsingHelpers.GetBorderLineStyle((byte)((border2 >> 21) & 0x0F))
            };

            _workbook.Borders.Add(border);
        }

        /// <summary>
        /// 解析全局填充记录 (FILL)
        /// </summary>
        public void ParseFillRecord(BiffRecord record)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 4) return;

            ushort fillData = BitConverter.ToUInt16(data, 0);
            byte pattern = (byte)(fillData & 0x3F);
            if (pattern == 0) return;

            int icvFore = (fillData >> 6) & 0x7F;
            int icvBack = data.Length > 2 ? (data[2] & 0x7F) : 65;

            var fill = new Fill
            {
                PatternType = ParsingHelpers.GetPatternType(pattern),
                ForegroundColor = GetColorFromPalette(icvFore),
                BackgroundColor = GetColorFromPalette(icvBack)
            };

            _workbook.Fills.Add(fill);
        }

        /// <summary>
        /// 解析工作表级别的字体记录
        /// </summary>
        public void ParseFontRecordToWorksheet(BiffRecord record, Worksheet worksheet)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 46) return;

            var font = new Font
            {
                Height = BitConverter.ToInt16(data, 0),
                IsBold = (BitConverter.ToUInt16(data, 2) & 0x0001) != 0,
                IsItalic = (BitConverter.ToUInt16(data, 2) & 0x0002) != 0,
                IsUnderline = (BitConverter.ToUInt16(data, 2) & 0x0004) != 0,
                IsStrikethrough = (BitConverter.ToUInt16(data, 2) & 0x0008) != 0,
                ColorIndex = BitConverter.ToUInt16(data, 6)
            };

            string? wsColor = GetColorFromPalette(font.ColorIndex);
            font.Color = wsColor?.Replace("#", "");

            if (data.Length > 46)
            {
                int nameOffset = 46;
                font.Name = ReadBiffString(data, ref nameOffset);
            }

            worksheet.Fonts.Add(font);
        }

        /// <summary>
        /// 解析工作表级别的XF记录
        /// </summary>
        public void ParseXfRecordToWorksheet(BiffRecord record, Worksheet worksheet)
        {
            if (record.Data == null || record.Data.Length < 20) return;

            var xf = new Xf
            {
                FontIndex = BitConverter.ToUInt16(record.Data, 0),
                NumberFormatIndex = BitConverter.ToUInt16(record.Data, 2),
                CellFormatIndex = BitConverter.ToUInt16(record.Data, 4)
            };

            // 解析对齐方式
            if (record.Data.Length >= 8)
            {
                ushort alignment = BitConverter.ToUInt16(record.Data, 6);
                xf.HorizontalAlignment = ParseHorizontalAlignment((byte)(alignment & 0x000F));
                xf.VerticalAlignment = ParseVerticalAlignment((byte)((alignment & 0x00F0) >> 4));
                xf.Indent = (byte)((alignment & 0x0F00) >> 8);
                xf.WrapText = (alignment & 0x1000) != 0;
            }

            // 解析边框
            if (record.Data.Length >= 18)
            {
                ParseXfBorders(record.Data, xf);
            }

            // 解析填充
            if (record.Data.Length >= 20)
            {
                ParseXfFill(record.Data, xf);
            }

            worksheet.Xfs.Add(xf);
        }

        /// <summary>
        /// 解析工作表级别的调色板记录
        /// </summary>
        public void ParsePaletteRecordToWorksheet(BiffRecord record, Worksheet worksheet)
        {
            if (record.Data == null || record.Data.Length < 6) return;

            int startIndex = BitConverter.ToUInt16(record.Data, 0);
            int colorCount = (record.Data.Length - 2) / 4;

            for (int i = 0; i < colorCount; i++)
            {
                int offset = 2 + i * 4;
                if (offset + 3 > record.Data.Length) break;

                byte red = record.Data[offset];
                byte green = record.Data[offset + 1];
                byte blue = record.Data[offset + 2];
                worksheet.Palette[startIndex + i] = $"#{red:X2}{green:X2}{blue:X2}";
            }
        }

        #region 私有辅助方法

        private void ParseXfBorders(byte[] data, Xf xf)
        {
            uint border1 = BitConverter.ToUInt32(data, 10);
            uint border2 = BitConverter.ToUInt32(data, 14);

            var border = new Border
            {
                Left = ParsingHelpers.GetBorderLineStyle((byte)(border1 & 0x0F)),
                Right = ParsingHelpers.GetBorderLineStyle((byte)((border1 >> 4) & 0x0F)),
                Top = ParsingHelpers.GetBorderLineStyle((byte)((border1 >> 8) & 0x0F)),
                Bottom = ParsingHelpers.GetBorderLineStyle((byte)((border1 >> 12) & 0x0F)),
                LeftColor = GetColorFromPalette((int)((border1 >> 16) & 0x7F)),
                RightColor = GetColorFromPalette((int)((border1 >> 23) & 0x7F)),
                TopColor = GetColorFromPalette((int)(border2 & 0x7F)),
                BottomColor = GetColorFromPalette((int)((border2 >> 7) & 0x7F)),
                DiagonalColor = GetColorFromPalette((int)((border2 >> 14) & 0x7F)),
                Diagonal = ParsingHelpers.GetBorderLineStyle((byte)((border2 >> 21) & 0x0F))
            };

            // 只有当边框不是全部为none时才添加
            if (border.Left != "none" || border.Right != "none" || border.Top != "none" ||
                border.Bottom != "none" || border.Diagonal != "none")
            {
                int existingBorderIdx = _workbook.Borders.FindIndex(b =>
                    b.Left == border.Left && b.Right == border.Right &&
                    b.Top == border.Top && b.Bottom == border.Bottom &&
                    b.LeftColor == border.LeftColor && b.RightColor == border.RightColor &&
                    b.TopColor == border.TopColor && b.BottomColor == border.BottomColor);

                xf.BorderIndex = existingBorderIdx >= 0
                    ? existingBorderIdx + 1
                    : _workbook.Borders.Count + 1;

                if (existingBorderIdx < 0)
                {
                    _workbook.Borders.Add(border);
                }
            }
            else
            {
                xf.BorderIndex = 0;
            }
        }

        private void ParseXfFill(byte[] data, Xf xf)
        {
            ushort fillData = BitConverter.ToUInt16(data, 18);
            byte pattern = (byte)(fillData & 0x3F);
            int icvFore = (fillData >> 6) & 0x7F;
            int icvBack = data.Length > 20 ? (data[20] & 0x7F) : 65;

            if (pattern > 0)
            {
                var fill = new Fill
                {
                    PatternType = ParsingHelpers.GetPatternType(pattern),
                    ForegroundColor = GetColorFromPalette(icvFore),
                    BackgroundColor = GetColorFromPalette(icvBack)
                };

                int existingFillIdx = _workbook.Fills.FindIndex(f =>
                    f.PatternType == fill.PatternType &&
                    f.ForegroundColor == fill.ForegroundColor &&
                    f.BackgroundColor == fill.BackgroundColor);

                xf.FillIndex = existingFillIdx >= 0
                    ? existingFillIdx + 2
                    : _workbook.Fills.Count + 2;

                if (existingFillIdx < 0)
                {
                    _workbook.Fills.Add(fill);
                }
            }
            else
            {
                xf.FillIndex = 0;
            }
        }

        private static string ParseHorizontalAlignment(byte alignCode)
        {
            return alignCode switch
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
        }

        private static string ParseVerticalAlignment(byte alignCode)
        {
            return alignCode switch
            {
                0 => "top",
                1 => "center",
                2 => "bottom",
                3 => "justify",
                4 => "distributed",
                _ => "bottom"
            };
        }

        private string? GetColorFromPalette(int colorIndex)
        {
            if (colorIndex == 64) return null; // System Foreground
            if (colorIndex == 65) return null; // System Background

            if (_palette.TryGetValue(colorIndex, out string? color))
                return color;

            return null;
        }

        private static string ReadBiffString(byte[] data, ref int offset)
        {
            if (offset >= data.Length) return string.Empty;

            byte flags = data[offset];
            bool isUnicode = (flags & 0x01) != 0;
            bool hasHighByte = (flags & 0x02) != 0;
            offset++;

            int charCount = data[offset];
            offset++;

            if (hasHighByte)
            {
                charCount += data[offset] << 8;
                offset++;
            }

            if (charCount == 0) return string.Empty;

            int byteCount = isUnicode ? charCount * 2 : charCount;
            if (offset + byteCount > data.Length)
                byteCount = data.Length - offset;

            string result = isUnicode
                ? System.Text.Encoding.Unicode.GetString(data, offset, byteCount)
                : System.Text.Encoding.ASCII.GetString(data, offset, byteCount);

            offset += byteCount;
            return result.TrimEnd('\0');
        }

        #endregion
    }
}
