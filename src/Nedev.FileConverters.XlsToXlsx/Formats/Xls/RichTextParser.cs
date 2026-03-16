using System;
using System.Collections.Generic;
using System.Text;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 富文本解析器 - 处理富文本和BIFF字符串
    /// </summary>
    public class RichTextParser
    {
        private readonly Func<short, Font?> _getFontByIndex;

        public RichTextParser(Func<short, Font?> getFontByIndex)
        {
            _getFontByIndex = getFontByIndex ?? throw new ArgumentNullException(nameof(getFontByIndex));
        }

        /// <summary>
        /// 解析富文本数据
        /// </summary>
        /// <param name="data">字节数据</param>
        /// <param name="offset">起始偏移</param>
        /// <returns>富文本运行列表</returns>
        public List<RichTextRun> ParseRichText(byte[] data, int offset)
        {
            var richTextRuns = new List<RichTextRun>();
            int currentOffset = offset;

            try
            {
                // 读取富文本记录的结构
                // 根据BIFF8格式规范解析富文本数据
                while (currentOffset < data.Length)
                {
                    // 读取文本长度
                    if (currentOffset + 2 <= data.Length)
                    {
                        short textLength = BitConverter.ToInt16(data, currentOffset);
                        currentOffset += 2;

                        // 读取文本类型（0=ASCII, 1=Unicode）
                        byte textType = 0;
                        if (currentOffset < data.Length)
                        {
                            textType = data[currentOffset];
                            currentOffset += 1;
                        }

                        // 读取文本内容
                        string text;
                        if (textType == 0)
                        {
                            // ASCII字符串
                            if (currentOffset + textLength <= data.Length)
                            {
                                text = Encoding.ASCII.GetString(data, currentOffset, textLength);
                                currentOffset += textLength;
                            }
                            else
                            {
                                break;
                            }
                        }
                        else
                        {
                            // Unicode字符串
                            if (currentOffset + textLength * 2 <= data.Length)
                            {
                                text = Encoding.Unicode.GetString(data, currentOffset, textLength * 2);
                                currentOffset += textLength * 2;
                            }
                            else
                            {
                                break;
                            }
                        }

                        // 读取字体索引
                        short fontIndex = 0;
                        if (currentOffset + 2 <= data.Length)
                        {
                            fontIndex = BitConverter.ToInt16(data, currentOffset);
                            currentOffset += 2;
                        }

                        // 创建富文本运行
                        var run = new RichTextRun
                        {
                            Text = text,
                            // 根据fontIndex获取对应的字体信息
                            Font = _getFontByIndex(fontIndex)
                        };
                        richTextRuns.Add(run);
                    }
                    else
                    {
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("解析富文本时发生错误", ex);
            }

            return richTextRuns;
        }

        /// <summary>
        /// 从字节数组读取BIFF字符串（支持ASCII和Unicode）
        /// </summary>
        /// <param name="data">字节数据</param>
        /// <param name="offset">引用偏移（会被更新）</param>
        /// <returns>读取的字符串</returns>
        public static string ReadBiffString(byte[] data, ref int offset)
        {
            if (offset + 2 > data.Length) return string.Empty;
            ushort charCount = BitConverter.ToUInt16(data, offset);
            offset += 2;
            if (offset >= data.Length) return string.Empty;

            byte option = data[offset];
            offset += 1;

            bool isUnicode = (option & 0x01) != 0;
            bool hasRichText = (option & 0x08) != 0;
            bool hasExtended = (option & 0x04) != 0;

            int runs = 0;
            if (hasRichText)
            {
                if (offset + 2 > data.Length) return string.Empty;
                runs = BitConverter.ToUInt16(data, offset);
                offset += 2;
            }

            int extendedSize = 0;
            if (hasExtended)
            {
                if (offset + 4 > data.Length) return string.Empty;
                extendedSize = BitConverter.ToInt32(data, offset);
                offset += 4; // Skip the size header
            }

            string result;
            if (isUnicode)
            {
                int byteCount = charCount * 2;
                if (offset + byteCount > data.Length) byteCount = data.Length - offset;
                result = Encoding.Unicode.GetString(data, offset, byteCount);
                offset += byteCount;
            }
            else
            {
                int byteCount = charCount;
                if (offset + byteCount > data.Length) byteCount = data.Length - offset;
                result = Encoding.ASCII.GetString(data, offset, byteCount);
                offset += byteCount;
            }

            if (hasRichText)
            {
                offset += runs * 4; // Skip formatting runs
            }

            if (hasExtended)
            {
                offset += extendedSize; // Skip the phonetic string data payload
            }

            return result;
        }

        /// <summary>
        /// 从字节数组读取指定字符数的BIFF字符串
        /// </summary>
        /// <param name="data">字节数据</param>
        /// <param name="offset">引用偏移（会被更新）</param>
        /// <param name="charCount">字符数</param>
        /// <returns>读取的字符串</returns>
        public static string ReadBiffStringFromBytes(byte[] data, ref int offset, int charCount)
        {
            if (offset >= data.Length) return string.Empty;
            byte option = data[offset];
            offset += 1;
            bool isUnicode = (option & 0x01) != 0;
            string result;
            if (isUnicode)
            {
                int byteCount = charCount * 2;
                if (offset + byteCount > data.Length) byteCount = data.Length - offset;
                result = Encoding.Unicode.GetString(data, offset, byteCount);
                offset += byteCount;
            }
            else
            {
                int byteCount = charCount;
                if (offset + byteCount > data.Length) byteCount = data.Length - offset;
                result = Encoding.ASCII.GetString(data, offset, byteCount);
                offset += byteCount;
            }
            return result;
        }
    }
}
