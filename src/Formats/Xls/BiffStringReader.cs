using System;
using System.Collections.Generic;
using System.Text;

namespace Nedev.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 处理跨越多个 CONTINUE (0x003C) 记录的 BIFF 字符串读取
    /// </summary>
    public class BiffStringReader
    {
        private readonly BiffRecord _record;
        private int _chunkIndex; // 0 = Data, 1+ = Continues[0...]
        private int _chunkOffset;
        
        public BiffStringReader(BiffRecord record, int startOffset)
        {
            _record = record;
            _chunkIndex = 0;
            _chunkOffset = startOffset;
        }

        public string ReadString()
        {
            if (IsEOF()) return string.Empty;

            // 1. Read Char Count (2 bytes)
            ushort charCount = ReadUInt16();
            if (charCount == 0) return string.Empty;

            // 2. Read Option Flag (1 byte)
            byte option = ReadByte();
            bool isUnicode = (option & 0x01) != 0;
            bool hasRichText = (option & 0x08) != 0;
            bool hasExtended = (option & 0x04) != 0;

            // 3. Read Formatting Runs Count (optional, 2 bytes)
            int runCount = 0;
            if (hasRichText)
            {
                runCount = ReadUInt16();
            }

            // 4. Read Extended Data Size (optional, 4 bytes)
            int extendedSize = 0;
            if (hasExtended)
            {
                extendedSize = ReadInt32();
            }

            // 5. Read String Characters
            StringBuilder sb = new StringBuilder(charCount);
            int charsRemaining = charCount;

            while (charsRemaining > 0)
            {
                // 如果当前 chunk 已经读完，移动到下一个 chunk (CONTINUE 记录)
                if (_chunkOffset >= GetCurrentChunkLength())
                {
                    if (MoveToNextChunk())
                    {
                        // 在跨越 chunk 边界读取字符串内容时，
                        // CONTINUE 记录的第一个字节是新的 Option 标志 (指示此块是 ASCII 还是 Unicode)
                        byte newOption = ReadByte();
                        isUnicode = (newOption & 0x01) != 0;
                    }
                    else
                    {
                        break; // 数据意外结束
                    }
                }

                int availableBytes = GetCurrentChunkLength() - _chunkOffset;
                int bytesPerChar = isUnicode ? 2 : 1;
                int charsInThisChunk = availableBytes / bytesPerChar;
                
                int charsToRead = Math.Min(charsRemaining, charsInThisChunk);

                if (charsToRead > 0)
                {
                    byte[] chunkArray = GetCurrentChunk();
                    if (isUnicode)
                    {
                        sb.Append(Encoding.Unicode.GetString(chunkArray, _chunkOffset, charsToRead * 2));
                        _chunkOffset += charsToRead * 2;
                    }
                    else
                    {
                        sb.Append(Encoding.ASCII.GetString(chunkArray, _chunkOffset, charsToRead));
                        _chunkOffset += charsToRead;
                    }
                    charsRemaining -= charsToRead;
                }
                else
                {
                    // Available bytes < bytesPerChar (例如 Unicode 字符被硬生生切在两个 chunk 里)
                    // Excel 规范中通常不会将一个两字节字符切开在边界处存放，但为了防御性编程，我们跳过这 1 个字节
                    if (availableBytes > 0)
                    {
                        _chunkOffset += availableBytes;
                    }
                }
            }

            // 6. Skip Formatting Runs Data (4 bytes per run)
            if (hasRichText)
            {
                for (int i = 0; i < runCount * 4; i++)
                {
                    ReadByte();
                }
            }

            // 7. Skip Extended String Data
            if (hasExtended)
            {
                for (int i = 0; i < extendedSize; i++)
                {
                    ReadByte();
                }
            }

            return sb.ToString();
        }

        private bool IsEOF()
        {
            return _chunkIndex > _record.Continues.Count || 
                   (_chunkIndex == _record.Continues.Count && _chunkOffset >= GetCurrentChunkLength());
        }

        private byte[] GetCurrentChunk()
        {
            if (_chunkIndex == 0)
                return _record.Data ?? Array.Empty<byte>();
            
            return _record.Continues[_chunkIndex - 1];
        }

        private int GetCurrentChunkLength()
        {
            return GetCurrentChunk().Length;
        }

        private bool MoveToNextChunk()
        {
            _chunkIndex++;
            _chunkOffset = 0;
            return _chunkIndex <= _record.Continues.Count;
        }

        private byte ReadByte()
        {
            while (_chunkOffset >= GetCurrentChunkLength())
            {
                if (!MoveToNextChunk())
                    return 0; // EOF
            }

            byte val = GetCurrentChunk()[_chunkOffset];
            _chunkOffset++;
            return val;
        }

        private ushort ReadUInt16()
        {
            byte b1 = ReadByte();
            byte b2 = ReadByte();
            return (ushort)(b1 | (b2 << 8));
        }

        private int ReadInt32()
        {
            byte b1 = ReadByte();
            byte b2 = ReadByte();
            byte b3 = ReadByte();
            byte b4 = ReadByte();
            return b1 | (b2 << 8) | (b3 << 16) | (b4 << 24);
        }
    }
}
