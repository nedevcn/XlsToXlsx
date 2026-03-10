using System;
using System.Collections.Generic;
using System.Linq;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    internal static class NpoiStylePatcher
    {
        public static void Apply(System.IO.Stream xlsStream, Workbook workbook)
        {
            using var sourceWorkbook = new HSSFWorkbook(xlsStream);

            workbook.Fonts = BuildFonts(sourceWorkbook);
            workbook.Borders.Clear();
            workbook.Fills.Clear();
            workbook.Styles.Clear();
            workbook.XfList.Clear();

            var fillMap = new Dictionary<string, int>(StringComparer.Ordinal);
            var borderMap = new Dictionary<string, int>(StringComparer.Ordinal);

            for (short styleIndex = 0; styleIndex < sourceWorkbook.NumCellStyles; styleIndex++)
            {
                ICellStyle sourceStyle = sourceWorkbook.GetCellStyleAt(styleIndex);
                var xf = new Xf
                {
                    FontIndex = (ushort)sourceStyle.FontIndex,
                    NumberFormatIndex = (ushort)sourceStyle.DataFormat,
                    IsLocked = sourceStyle.IsLocked,
                    IsHidden = sourceStyle.IsHidden,
                    HorizontalAlignment = MapHorizontalAlignment(sourceStyle.Alignment),
                    VerticalAlignment = MapVerticalAlignment(sourceStyle.VerticalAlignment),
                    WrapText = sourceStyle.WrapText,
                    Indent = (byte)Math.Clamp(sourceStyle.Indention, (short)0, (short)byte.MaxValue),
                    ApplyAlignment = sourceStyle.Alignment != HorizontalAlignment.General ||
                                     sourceStyle.VerticalAlignment != VerticalAlignment.Bottom ||
                                     sourceStyle.WrapText ||
                                     sourceStyle.Indention > 0
                };

                Fill? fill = BuildFill(sourceWorkbook, sourceStyle);
                if (fill != null)
                {
                    string fillKey = GetFillKey(fill);
                    if (!fillMap.TryGetValue(fillKey, out int fillIndex))
                    {
                        workbook.Fills.Add(fill);
                        fillIndex = workbook.Fills.Count + 1;
                        fillMap[fillKey] = fillIndex;
                    }

                    xf.FillIndex = fillIndex;
                }

                Border? border = BuildBorder(sourceWorkbook, sourceStyle);
                if (border != null)
                {
                    string borderKey = GetBorderKey(border);
                    if (!borderMap.TryGetValue(borderKey, out int borderIndex))
                    {
                        workbook.Borders.Add(border);
                        borderIndex = workbook.Borders.Count;
                        borderMap[borderKey] = borderIndex;
                    }

                    xf.BorderIndex = borderIndex;
                }

                workbook.XfList.Add(xf);
                workbook.Styles.Add(new Style
                {
                    Id = styleIndex.ToString(),
                    Font = ResolveFont(workbook.Fonts, xf.FontIndex),
                    Fill = fill,
                    Border = border,
                    Alignment = new Alignment
                    {
                        Horizontal = xf.HorizontalAlignment,
                        Vertical = xf.VerticalAlignment,
                        Indent = xf.Indent,
                        WrapText = xf.WrapText
                    }
                });
            }

            for (int sheetIndex = 0; sheetIndex < Math.Min(sourceWorkbook.NumberOfSheets, workbook.Worksheets.Count); sheetIndex++)
            {
                ISheet sourceSheet = sourceWorkbook.GetSheetAt(sheetIndex);
                Worksheet targetSheet = workbook.Worksheets[sheetIndex];
                var rowMap = targetSheet.Rows.ToDictionary(row => row.RowIndex);

                for (int rowIndex = sourceSheet.FirstRowNum; rowIndex <= sourceSheet.LastRowNum; rowIndex++)
                {
                    IRow? sourceRow = sourceSheet.GetRow(rowIndex);
                    if (sourceRow == null)
                        continue;

                    if (!rowMap.TryGetValue(rowIndex + 1, out Row? targetRow))
                        continue;

                    if (sourceRow.RowStyle != null)
                        targetRow.DefaultXfIndex = sourceRow.RowStyle.Index;

                    var cellMap = targetRow.Cells.ToDictionary(cell => cell.ColumnIndex);
                    for (int columnIndex = sourceRow.FirstCellNum; columnIndex >= 0 && columnIndex < sourceRow.LastCellNum; columnIndex++)
                    {
                        ICell? sourceCell = sourceRow.GetCell(columnIndex);
                        if (sourceCell == null)
                            continue;

                        if (cellMap.TryGetValue(columnIndex + 1, out Cell? targetCell))
                            targetCell.StyleId = sourceCell.CellStyle.Index.ToString();
                    }
                }
            }
        }

        private static List<Font> BuildFonts(HSSFWorkbook workbook)
        {
            var fontMap = new Dictionary<int, Font>();

            for (short styleIndex = 0; styleIndex < workbook.NumCellStyles; styleIndex++)
            {
                ICellStyle style = workbook.GetCellStyleAt(styleIndex);
                int rawFontIndex = style.FontIndex;
                int normalizedFontIndex = NormalizeBiffFontIndex(rawFontIndex);

                if (fontMap.ContainsKey(normalizedFontIndex))
                    continue;

                IFont sourceFont = workbook.GetFontAt((short)rawFontIndex);
                fontMap[normalizedFontIndex] = new Font
                {
                    Name = sourceFont.FontName,
                    Size = sourceFont.FontHeightInPoints,
                    Height = (short)(sourceFont.FontHeightInPoints * 20),
                    Bold = sourceFont.IsBold,
                    IsBold = sourceFont.IsBold,
                    Italic = sourceFont.IsItalic,
                    IsItalic = sourceFont.IsItalic,
                    Underline = sourceFont.Underline != FontUnderlineType.None,
                    IsUnderline = sourceFont.Underline != FontUnderlineType.None,
                    IsStrikethrough = sourceFont.IsStrikeout,
                    ColorIndex = (ushort)sourceFont.Color,
                    Color = GetPaletteColorHex(workbook, sourceFont.Color)
                };
            }

            if (fontMap.Count == 0)
                fontMap[0] = new Font { Name = "Calibri", Size = 11, Height = 220, Color = "000000" };

            int maxIndex = fontMap.Keys.Max();
            var fonts = new List<Font>(maxIndex + 1);
            for (int index = 0; index <= maxIndex; index++)
            {
                fonts.Add(fontMap.TryGetValue(index, out Font? font)
                    ? font
                    : new Font { Name = "Calibri", Size = 11, Height = 220, Color = "000000" });
            }

            return fonts;
        }

        private static Font? ResolveFont(List<Font> fonts, ushort rawFontIndex)
        {
            int normalizedFontIndex = NormalizeBiffFontIndex(rawFontIndex);
            return normalizedFontIndex >= 0 && normalizedFontIndex < fonts.Count ? fonts[normalizedFontIndex] : null;
        }

        private static Fill? BuildFill(HSSFWorkbook workbook, ICellStyle style)
        {
            string pattern = MapFillPattern(style.FillPattern);
            string? foreground = GetPaletteColorHex(workbook, style.FillForegroundColor);
            string? background = GetPaletteColorHex(workbook, style.FillBackgroundColor);

            if (pattern == "none" && string.IsNullOrEmpty(foreground) && string.IsNullOrEmpty(background))
                return null;

            return new Fill
            {
                PatternType = pattern,
                ForegroundColor = foreground,
                BackgroundColor = background
            };
        }

        private static Border? BuildBorder(HSSFWorkbook workbook, ICellStyle style)
        {
            var border = new Border
            {
                Left = MapBorderStyle(style.BorderLeft),
                Right = MapBorderStyle(style.BorderRight),
                Top = MapBorderStyle(style.BorderTop),
                Bottom = MapBorderStyle(style.BorderBottom),
                LeftColor = GetPaletteColorHex(workbook, style.LeftBorderColor),
                RightColor = GetPaletteColorHex(workbook, style.RightBorderColor),
                TopColor = GetPaletteColorHex(workbook, style.TopBorderColor),
                BottomColor = GetPaletteColorHex(workbook, style.BottomBorderColor)
            };

            if (border.Left == "none" && border.Right == "none" && border.Top == "none" && border.Bottom == "none")
                return null;

            return border;
        }

        private static int NormalizeBiffFontIndex(int fontIndex)
        {
            if (fontIndex < 0)
                return 0;

            return fontIndex > 4 ? fontIndex - 1 : fontIndex;
        }

        private static string? GetPaletteColorHex(HSSFWorkbook workbook, short colorIndex)
        {
            if (colorIndex == HSSFColor.Automatic.Index || colorIndex == 64 || colorIndex == 65)
                return null;

            HSSFColor? color = workbook.GetCustomPalette().GetColor(colorIndex);
            if (color == null)
                return null;

            byte[] triplet = color.GetTriplet();
            return string.Concat(triplet.Select(component => component.ToString("X2")));
        }

        private static string GetFillKey(Fill fill)
        {
            return $"{fill.PatternType}|{fill.ForegroundColor}|{fill.BackgroundColor}";
        }

        private static string GetBorderKey(Border border)
        {
            return string.Join("|", border.Left, border.LeftColor, border.Right, border.RightColor, border.Top, border.TopColor, border.Bottom, border.BottomColor, border.Diagonal, border.DiagonalColor);
        }

        private static string MapHorizontalAlignment(HorizontalAlignment alignment)
        {
            return alignment switch
            {
                HorizontalAlignment.Left => "left",
                HorizontalAlignment.Center => "center",
                HorizontalAlignment.Right => "right",
                HorizontalAlignment.Fill => "fill",
                HorizontalAlignment.Justify => "justify",
                HorizontalAlignment.CenterSelection => "centerContinuous",
                HorizontalAlignment.Distributed => "distributed",
                _ => "general"
            };
        }

        private static string MapVerticalAlignment(VerticalAlignment alignment)
        {
            return alignment switch
            {
                VerticalAlignment.Top => "top",
                VerticalAlignment.Center => "center",
                VerticalAlignment.Justify => "justify",
                VerticalAlignment.Distributed => "distributed",
                _ => "bottom"
            };
        }

        private static string MapFillPattern(FillPattern pattern)
        {
            return pattern switch
            {
                FillPattern.SolidForeground => "solid",
                FillPattern.FineDots => "gray0625",
                FillPattern.AltBars => "gray125",
                FillPattern.SparseDots => "lightGrid",
                FillPattern.ThickHorizontalBands => "darkHorizontal",
                FillPattern.ThickVerticalBands => "darkVertical",
                FillPattern.ThickBackwardDiagonals => "darkDown",
                FillPattern.ThickForwardDiagonals => "darkUp",
                FillPattern.BigSpots => "darkGrid",
                FillPattern.Bricks => "darkTrellis",
                FillPattern.ThinHorizontalBands => "lightHorizontal",
                FillPattern.ThinVerticalBands => "lightVertical",
                FillPattern.ThinBackwardDiagonals => "lightDown",
                FillPattern.ThinForwardDiagonals => "lightUp",
                FillPattern.Squares => "lightGrid",
                FillPattern.Diamonds => "lightTrellis",
                FillPattern.LessDots => "lightGray",
                FillPattern.LeastDots => "mediumGray",
                _ => "none"
            };
        }

        private static string MapBorderStyle(BorderStyle borderStyle)
        {
            return borderStyle switch
            {
                BorderStyle.Thin => "thin",
                BorderStyle.Medium => "medium",
                BorderStyle.Dashed => "dashed",
                BorderStyle.Dotted => "dotted",
                BorderStyle.Thick => "thick",
                BorderStyle.Double => "double",
                BorderStyle.Hair => "hair",
                BorderStyle.MediumDashed => "mediumDashed",
                BorderStyle.DashDot => "dashDot",
                BorderStyle.MediumDashDot => "mediumDashDot",
                BorderStyle.DashDotDot => "dashDotDot",
                BorderStyle.MediumDashDotDot => "mediumDashDotDot",
                BorderStyle.SlantedDashDot => "slantDashDot",
                _ => "none"
            };
        }
    }
}