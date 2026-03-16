using System;
using System.Collections.Generic;
using System.Linq;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 绘图解析器 - 处理MSODRAWING、OBJ、PICTURE等绘图相关记录
    /// </summary>
    public class DrawingParser
    {
        private readonly List<byte[]> _msoDrawingData;
        private readonly List<byte[]> _msoDrawingGroupData;
        private readonly List<(int left, int top, int width, int height)> _pendingChartAnchors;

        public DrawingParser(
            List<byte[]> msoDrawingData,
            List<byte[]> msoDrawingGroupData,
            List<(int left, int top, int width, int height)> pendingChartAnchors)
        {
            _msoDrawingData = msoDrawingData ?? throw new ArgumentNullException(nameof(msoDrawingData));
            _msoDrawingGroupData = msoDrawingGroupData ?? throw new ArgumentNullException(nameof(msoDrawingGroupData));
            _pendingChartAnchors = pendingChartAnchors ?? throw new ArgumentNullException(nameof(pendingChartAnchors));
        }

        /// <summary>
        /// 解析MSODRAWING记录（工作表级别）
        /// </summary>
        public void ParseMSODrawingRecord(BiffRecord record)
        {
            if (record.Data != null && record.Data.Length > 0)
            {
                _msoDrawingData.Add(record.GetAllData());
            }
        }

        /// <summary>
        /// 解析MSODRAWINGGROUP记录（全局级别）
        /// </summary>
        public void ParseMsoDrawingGroupGlobal(BiffRecord record)
        {
            if (record.Data != null && record.Data.Length > 0)
            {
                _msoDrawingGroupData.Add(record.GetAllData());
            }
        }

        /// <summary>
        /// 解析OBJ记录
        /// </summary>
        public void ParseObjRecord(BiffRecord record, Worksheet worksheet)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 4) return;

            ushort ft = BitConverter.ToUInt16(data, 0);
            if (ft == 0x0015 && data.Length >= 10) // ftCmo
            {
                ushort objType = BitConverter.ToUInt16(data, 4);
                if (objType == 0x0005) // Chart Data
                {
                    ParseChartAnchor();
                }
                else
                {
                    _msoDrawingData.Clear();
                    ParseEmbeddedObject(data, worksheet);
                }
            }
        }

        private void ParseChartAnchor()
        {
            try
            {
                if (_msoDrawingData.Count > 0)
                {
                    int totalLen = _msoDrawingData.Sum(d => d.Length);
                    byte[] drawingData = new byte[totalLen];
                    int offset = 0;
                    foreach (var d in _msoDrawingData)
                    {
                        Array.Copy(d, 0, drawingData, offset, d.Length);
                        offset += d.Length;
                    }

                    var escherRecords = Escher.EscherParser.ParseStream(drawingData);
                    var anchor = FindClientAnchor(escherRecords);
                    if (anchor != null && anchor.Data != null && anchor.Data.Length >= 16)
                    {
                        ushort col1 = BitConverter.ToUInt16(anchor.Data, 0);
                        ushort dxL = BitConverter.ToUInt16(anchor.Data, 2);
                        ushort row1 = BitConverter.ToUInt16(anchor.Data, 4);
                        ushort dyT = BitConverter.ToUInt16(anchor.Data, 6);
                        ushort col2 = BitConverter.ToUInt16(anchor.Data, 8);
                        ushort dxR = BitConverter.ToUInt16(anchor.Data, 10);
                        ushort row2 = BitConverter.ToUInt16(anchor.Data, 12);
                        ushort dyB = BitConverter.ToUInt16(anchor.Data, 14);
                        int left = dxL;
                        int top = dyT;
                        int width = Math.Max(1, (col2 - col1) * 1024 + (dxR - dxL));
                        int height = Math.Max(1, (row2 - row1) * 256 + (dyB - dyT));
                        _pendingChartAnchors.Add((left, top, width, height));
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"解析 OBJ / Chart Anchor 失败: {ex.Message}", ex);
            }
            finally
            {
                _msoDrawingData.Clear();
            }
        }

        private void ParseEmbeddedObject(byte[] data, Worksheet worksheet)
        {
            var embeddedObject = new EmbeddedObject
            {
                Data = new byte[data.Length],
                MimeType = "application/octet-stream"
            };
            Array.Copy(data, embeddedObject.Data, data.Length);
            worksheet.EmbeddedObjects.Add(embeddedObject);
        }

        private Escher.EscherRecord? FindClientAnchor(IEnumerable<Escher.EscherRecord> records)
        {
            foreach (var record in records)
            {
                if (record.Type == Escher.EscherParser.ClientAnchor)
                    return record;

                if (record.IsContainer && record.Children.Count > 0)
                {
                    var result = FindClientAnchor(record.Children);
                    if (result != null) return result;
                }
            }
            return null;
        }

        /// <summary>
        /// 解析PICTURE记录
        /// </summary>
        public void ParsePictureRecord(BiffRecord record, Worksheet worksheet)
        {
            byte[] fullData = record.GetAllData();
            if (fullData == null || fullData.Length == 0)
                return;

            try
            {
                var picture = new Picture
                {
                    Data = new byte[fullData.Length]
                };
                Array.Copy(fullData, picture.Data, fullData.Length);

                // 识别图片格式
                (picture.MimeType, picture.Extension) = IdentifyImageFormat(fullData);

                worksheet.Pictures.Add(picture);
            }
            catch (Exception ex)
            {
                Logger.Error($"解析图片记录失败: {ex.Message}", ex);
            }
        }

        private (string mimeType, string extension) IdentifyImageFormat(byte[] data)
        {
            if (data.Length < 4)
                return ("application/octet-stream", "bin");

            // PNG
            if (data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47)
                return ("image/png", "png");

            // JPEG
            if (data[0] == 0xFF && data[1] == 0xD8)
                return ("image/jpeg", "jpg");

            // GIF
            if (data[0] == 0x47 && data[1] == 0x49 && data[2] == 0x46)
                return ("image/gif", "gif");

            // BMP
            if (data[0] == 0x42 && data[1] == 0x4D)
                return ("image/bmp", "bmp");

            // RIFF (WEBP or BMP)
            if (data[0] == 0x52 && data[1] == 0x49 && data[2] == 0x46 && data[3] == 0x46)
            {
                if (data.Length >= 12 && data[8] == 0x57 && data[9] == 0x45 && data[10] == 0x42 && data[11] == 0x50)
                    return ("image/webp", "webp");
                return ("image/bmp", "bmp");
            }

            // TIFF (little endian)
            if (data[0] == 0x49 && data[1] == 0x49 && data[2] == 0x2A && data[3] == 0x00)
                return ("image/tiff", "tiff");

            // TIFF (big endian)
            if (data[0] == 0x4D && data[1] == 0x4D && data[2] == 0x00 && data[3] == 0x2A)
                return ("image/tiff", "tiff");

            // ICO
            if (data[0] == 0x00 && data[1] == 0x00 && data[2] == 0x01 && data[3] == 0x00)
                return ("image/x-icon", "ico");

            // WMF
            if (data.Length >= 18 && data[0] == 0xD7 && data[1] == 0xCD && data[2] == 0xC6 && data[3] == 0x9A)
                return ("image/wmf", "wmf");

            // EMF
            if (data.Length >= 44 && data[0] == 0x01 && data[1] == 0x00 && data[2] == 0x00 && data[3] == 0x00)
                return ("image/emf", "emf");

            return ("application/octet-stream", "bin");
        }
    }
}
