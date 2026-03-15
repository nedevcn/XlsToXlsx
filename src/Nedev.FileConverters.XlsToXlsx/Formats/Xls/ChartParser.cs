using System;
using System.Collections.Generic;
using Nedev.FileConverters.XlsToXlsx;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 图表解析器 - 处理图表相关的BIFF记录
    /// </summary>
    public class ChartParser
    {
        private readonly Workbook _workbook;
        private readonly XlsDecryptor? _decryptor;
        private readonly byte[] _workbookData;
        private readonly Stream _stream;
        private readonly BinaryReader _reader;

        // 等待分配的图表锚点信息
        private readonly List<(int Left, int Top, int Width, int Height)> _pendingChartAnchors;

        public ChartParser(
            Workbook workbook,
            XlsDecryptor? decryptor,
            byte[] workbookData,
            Stream stream,
            BinaryReader reader,
            List<(int Left, int Top, int Width, int Height)> pendingChartAnchors)
        {
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            _decryptor = decryptor;
            _workbookData = workbookData ?? throw new ArgumentNullException(nameof(workbookData));
            _stream = stream ?? throw new ArgumentNullException(nameof(stream));
            _reader = reader ?? throw new ArgumentNullException(nameof(reader));
            _pendingChartAnchors = pendingChartAnchors ?? throw new ArgumentNullException(nameof(pendingChartAnchors));
        }

        /// <summary>
        /// 解析图表子流
        /// </summary>
        public void ParseChartSubstream(Worksheet worksheet)
        {
            long streamEnd = _workbookData.Length;
            BiffRecord? previousRecord = null;
            Chart currentChart = new Chart();
            Series? currentSeries = null;

            // 默认图表类型
            currentChart.ChartType = "colChart";

            while (_stream.Position < streamEnd)
            {
                try
                {
                    var recordStartPos = _stream.Position;
                    var record = BiffRecord.Read(_reader);

                    // 解密数据体
                    if (_decryptor != null && record.Id != (ushort)BiffRecordType.BOF)
                    {
                        if (record.Data != null && record.Data.Length > 0)
                        {
                            _decryptor.Decrypt(record.Data, recordStartPos + 4);
                        }
                    }

                    if (record.Id == (ushort)BiffRecordType.CONTINUE)
                    {
                        if (previousRecord != null && record.Data != null)
                        {
                            previousRecord.Continues.Add(record.Data);
                        }
                        continue;
                    }

                    if (previousRecord != null)
                    {
                        ProcessChartRecord(previousRecord, currentChart, ref currentSeries!, worksheet);
                    }

                    previousRecord = record;

                    if (record.Id == (ushort)BiffRecordType.EOF)
                    {
                        break;
                    }
                }
                catch (EndOfStreamException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    Logger.Error($"解析图表记录时发生错误: {ex.Message}", ex);
                    continue;
                }
            }

            if (previousRecord != null && previousRecord.Id != (ushort)BiffRecordType.EOF)
            {
                ProcessChartRecord(previousRecord, currentChart, ref currentSeries!, worksheet);
            }

            // 确保图表有默认轴和图例
            if (currentChart.XAxis == null) currentChart.XAxis = new Axis { Visible = true, Title = "X轴" };
            if (currentChart.YAxis == null) currentChart.YAxis = new Axis { Visible = true, Title = "Y轴" };
            if (currentChart.Legend == null) currentChart.Legend = new Legend { Visible = true, Position = "right" };

            // 应用等待分配的坐标信息
            if (_pendingChartAnchors.Count > worksheet.Charts.Count)
            {
                var anchor = _pendingChartAnchors[worksheet.Charts.Count];
                currentChart.Left = anchor.Left;
                currentChart.Top = anchor.Top;
            }

            worksheet.Charts.Add(currentChart);
            Logger.Info($"成功向工作表 {worksheet.Name} 添加了 1 个图表 (类型: {currentChart.ChartType}, 系列数: {currentChart.Series.Count})");
        }

        /// <summary>
        /// 处理单个图表记录
        /// </summary>
        private void ProcessChartRecord(BiffRecord record, Chart chart, ref Series currentSeries, Worksheet worksheet)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length == 0) return;

            switch (record.Id)
            {
                case (ushort)BiffRecordType.CHART3D:
                    // BIFF8 CHART3D (0x103A) - 标记为3D图表
                    chart.Is3D = true;
                    break;

                case 0x1017: // Bar
                    chart.ChartType = "barChart";
                    break;

                case 0x1018: // Line
                    chart.ChartType = "lineChart";
                    break;

                case 0x1019: // Pie
                    chart.ChartType = "pieChart";
                    break;

                case 0x101B: // Scatter
                    chart.ChartType = "scatterChart";
                    break;

                case 0x101A: // Area
                    chart.ChartType = "areaChart";
                    break;

                case 0x1020: // Radar
                    chart.ChartType = "radarChart";
                    break;

                case (ushort)BiffRecordType.CHARTSERIES:
                    currentSeries = new Series();
                    // 添加默认范围，实际公式解析较复杂
                    currentSeries.ValuesRange = $"{worksheet.Name}!$B$2:$B$6";
                    currentSeries.CategoriesRange = $"{worksheet.Name}!$A$2:$A$6";
                    currentSeries.LineStyle = new LineStyle { Width = 2 };
                    chart.Series.Add(currentSeries);
                    break;

                case 0x1021: // Chart Legend
                    ParseLegendRecord(data, chart);
                    break;

                case 0x1022: // Chart Title
                    ParseTitleRecord(data, chart);
                    break;

                case 0x1023: // Axis
                    ParseAxisRecord(data, chart);
                    break;
            }
        }

        /// <summary>
        /// 解析图例记录
        /// </summary>
        private void ParseLegendRecord(byte[] data, Chart chart)
        {
            if (data.Length < 2) return;

            ushort options = BitConverter.ToUInt16(data, 0);
            chart.Legend = new Legend
            {
                Visible = (options & 0x0001) != 0,
                Position = ((options >> 1) & 0x03) switch
                {
                    0 => "bottom",
                    1 => "corner",
                    2 => "top",
                    3 => "right",
                    _ => "right"
                }
            };
        }

        /// <summary>
        /// 解析标题记录
        /// </summary>
        private void ParseTitleRecord(byte[] data, Chart chart)
        {
            if (data.Length < 2) return;

            ushort options = BitConverter.ToUInt16(data, 0);
            if ((options & 0x0001) != 0 && data.Length > 2)
            {
                chart.Title = ReadBiffString(data, 2);
            }
        }

        /// <summary>
        /// 解析轴记录
        /// </summary>
        private void ParseAxisRecord(byte[] data, Chart chart)
        {
            if (data.Length < 4) return;

            ushort axisType = BitConverter.ToUInt16(data, 0);
            ushort options = BitConverter.ToUInt16(data, 2);

            var axis = new Axis
            {
                Visible = (options & 0x0001) != 0,
                Title = data.Length > 4 ? ReadBiffString(data, 4) : null
            };

            if (axisType == 0) // X轴
            {
                chart.XAxis = axis;
            }
            else if (axisType == 1) // Y轴
            {
                chart.YAxis = axis;
            }
        }

        /// <summary>
        /// 解析OBJ记录中的图表锚点信息
        /// </summary>
        public void ParseObjChartAnchor(BiffRecord record)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 6) return;

            try
            {
                ushort objType = BitConverter.ToUInt16(data, 4);
                if (objType != 0x0005) return; // 不是图表对象

                // 锚点数据通常在OBJ记录之后或嵌入在数据后面
                if (data.Length >= 26)
                {
                    int anchorOffset = 6;

                    // 读取锚点坐标 (left, top, width, height)
                    int left = BitConverter.ToInt32(data, anchorOffset);
                    int top = BitConverter.ToInt32(data, anchorOffset + 4);
                    int width = BitConverter.ToInt32(data, anchorOffset + 8);
                    int height = BitConverter.ToInt32(data, anchorOffset + 12);

                    _pendingChartAnchors.Add((left, top, width, height));
                    Logger.Info($"解析到图表锚点: Left={left}, Top={top}, Width={width}, Height={height}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"解析 OBJ / Chart Anchor 失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 清除待处理的图表锚点
        /// </summary>
        public void ClearPendingAnchors()
        {
            _pendingChartAnchors.Clear();
        }

        #region 辅助方法

        private static string ReadBiffString(byte[] data, int offset)
        {
            if (offset >= data.Length) return string.Empty;

            int charCount = data[offset];
            if (charCount == 0 || offset + 1 >= data.Length) return string.Empty;

            byte flags = data[offset + 1];
            bool isUnicode = (flags & 0x01) != 0;

            int startOffset = offset + 2;
            int byteCount = isUnicode ? charCount * 2 : charCount;

            if (startOffset + byteCount > data.Length)
                byteCount = data.Length - startOffset;

            if (byteCount <= 0) return string.Empty;

            string result = isUnicode
                ? System.Text.Encoding.Unicode.GetString(data, startOffset, byteCount)
                : System.Text.Encoding.ASCII.GetString(data, startOffset, byteCount);

            return result.TrimEnd('\0');
        }

        #endregion
    }
}
