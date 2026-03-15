using System;
using System.Collections.Generic;
using System.Text;
using Nedev.FileConverters.XlsToXlsx;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 文档属性解析器 - 解析OLE文档属性流
    /// </summary>
    public class DocumentPropertiesParser
    {
        private readonly OleCompoundFile _oleFile;

        public DocumentPropertiesParser(OleCompoundFile oleFile)
        {
            _oleFile = oleFile ?? throw new ArgumentNullException(nameof(oleFile));
        }

        /// <summary>
        /// 解析文档属性并填充到Workbook
        /// </summary>
        public void Parse(Workbook workbook)
        {
            try
            {
                var summary = _oleFile.ReadStreamByName("\u0005SummaryInformation");
                if (summary != null && summary.Length >= 48)
                {
                    ParseSummaryInformation(summary, workbook);
                }

                var docSummary = _oleFile.ReadStreamByName("\u0005DocumentSummaryInformation");
                if (docSummary != null && docSummary.Length >= 48)
                {
                    ParseDocumentSummaryInformation(docSummary, workbook);
                }
            }
            catch (Exception ex)
            {
                Logger.Warn($"解析文档属性时发生错误: {ex.Message}");
            }
        }

        private static void ParseSummaryInformation(byte[] data, Workbook workbook)
        {
            if (!TryReadPropertySection(data, out var props, out int codePage))
                return;

            foreach (var p in props)
            {
                uint vt = BitConverter.ToUInt32(data, p.ValueOffset);
                switch (p.Id)
                {
                    case 2: // Title
                        workbook.Title ??= ReadPropertyString(data, p.ValueOffset, codePage, vt);
                        break;
                    case 3: // Subject
                        workbook.Subject ??= ReadPropertyString(data, p.ValueOffset, codePage, vt);
                        break;
                    case 4: // Author
                        workbook.Author ??= ReadPropertyString(data, p.ValueOffset, codePage, vt);
                        break;
                    case 5: // Keywords
                        workbook.Keywords ??= ReadPropertyString(data, p.ValueOffset, codePage, vt);
                        break;
                    case 6: // Comments
                        workbook.Comments ??= ReadPropertyString(data, p.ValueOffset, codePage, vt);
                        break;
                    case 8: // LastAuthor
                        workbook.LastAuthor ??= ReadPropertyString(data, p.ValueOffset, codePage, vt);
                        break;
                    case 12: // Create time
                        workbook.CreatedUtc ??= ReadFileTimeUtc(data, p.ValueOffset, vt);
                        break;
                    case 13: // Last save time
                        workbook.ModifiedUtc ??= ReadFileTimeUtc(data, p.ValueOffset, vt);
                        break;
                }
            }
        }

        private static void ParseDocumentSummaryInformation(byte[] data, Workbook workbook)
        {
            if (!TryReadPropertySection(data, out var props, out int codePage))
                return;

            foreach (var p in props)
            {
                uint vt = BitConverter.ToUInt32(data, p.ValueOffset);
                switch (p.Id)
                {
                    case 2: // Category
                        workbook.Category ??= ReadPropertyString(data, p.ValueOffset, codePage, vt);
                        break;
                    case 14: // Manager
                        workbook.Manager ??= ReadPropertyString(data, p.ValueOffset, codePage, vt);
                        break;
                    case 15: // Company
                        workbook.Company ??= ReadPropertyString(data, p.ValueOffset, codePage, vt);
                        break;
                }
            }
        }

        private readonly struct PropertyRef
        {
            public readonly uint Id;
            public readonly int ValueOffset;
            public PropertyRef(uint id, int valueOffset)
            {
                Id = id;
                ValueOffset = valueOffset;
            }
        }

        /// <summary>
        /// 解析OLE PropertySetStream的首个section
        /// </summary>
        private static bool TryReadPropertySection(byte[] data, out List<PropertyRef> props, out int codePage)
        {
            props = new List<PropertyRef>();
            codePage = 1252;
            
            if (data == null || data.Length < 48) return false;

            ushort byteOrder = BitConverter.ToUInt16(data, 0);
            if (byteOrder != 0xFFFE) return false;

            int cSections = BitConverter.ToInt32(data, 24);
            if (cSections <= 0) return false;

            int sectionListOffset = 28;
            if (data.Length < sectionListOffset + 20) return false;

            int sectionOffset = BitConverter.ToInt32(data, sectionListOffset + 16);
            if (sectionOffset <= 0 || sectionOffset + 8 > data.Length) return false;

            int cProps = BitConverter.ToInt32(data, sectionOffset + 4);
            if (cProps <= 0) return false;

            int propListOffset = sectionOffset + 8;
            for (int i = 0; i < cProps; i++)
            {
                if (propListOffset + 8 > data.Length) break;
                uint propId = BitConverter.ToUInt32(data, propListOffset);
                int offset = BitConverter.ToInt32(data, propListOffset + 4);
                int valueOffset = sectionOffset + offset;
                if (valueOffset >= 0 && valueOffset + 8 <= data.Length)
                    props.Add(new PropertyRef(propId, valueOffset));
                propListOffset += 8;
            }

            // 查找CodePage
            foreach (var p in props)
            {
                if (p.Id != 1) continue;
                if (p.ValueOffset + 8 > data.Length) continue;
                uint vt = BitConverter.ToUInt32(data, p.ValueOffset);
                if (vt == 0x0002 && p.ValueOffset + 6 <= data.Length)
                {
                    codePage = BitConverter.ToUInt16(data, p.ValueOffset + 4);
                }
                break;
            }

            return props.Count > 0;
        }

        private static string? ReadPropertyString(byte[] data, int valueOffset, int codePage, uint vt)
        {
            try
            {
                if (valueOffset + 8 > data.Length) return null;
                
                if (vt == 0x1E) // VT_LPSTR
                {
                    int len = BitConverter.ToInt32(data, valueOffset + 4);
                    if (len <= 0) return null;
                    int bytesLen = len;
                    int start = valueOffset + 8;
                    if (start + bytesLen > data.Length)
                        bytesLen = Math.Max(0, data.Length - start);
                    if (bytesLen <= 0) return null;
                    var enc = Encoding.GetEncoding(codePage <= 0 ? 1252 : codePage);
                    string s = enc.GetString(data, start, bytesLen);
                    return s.TrimEnd('\0');
                }
                else if (vt == 0x1F) // VT_LPWSTR
                {
                    int cch = BitConverter.ToInt32(data, valueOffset + 4);
                    if (cch <= 0) return null;
                    int bytesLen = cch * 2;
                    int start = valueOffset + 8;
                    if (start + bytesLen > data.Length)
                        bytesLen = Math.Max(0, data.Length - start);
                    if (bytesLen <= 0) return null;
                    string s = Encoding.Unicode.GetString(data, start, bytesLen);
                    return s.TrimEnd('\0');
                }
            }
            catch
            {
                // ignore malformed
            }
            return null;
        }

        private static DateTime? ReadFileTimeUtc(byte[] data, int valueOffset, uint vt)
        {
            try
            {
                if (vt != 0x40) return null; // VT_FILETIME
                if (valueOffset + 12 > data.Length) return null;
                long fileTime = BitConverter.ToInt64(data, valueOffset + 4);
                if (fileTime <= 0) return null;
                return DateTime.FromFileTimeUtc(fileTime);
            }
            catch
            {
                return null;
            }
        }
    }
}
