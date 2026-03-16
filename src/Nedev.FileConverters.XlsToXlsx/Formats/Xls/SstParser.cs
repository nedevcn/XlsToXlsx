using System;
using System.Collections.Generic;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// SST（共享字符串表）解析器 - 处理BIFF8共享字符串表
    /// </summary>
    public class SstParser
    {
        private readonly List<string> _sharedStrings;

        public SstParser(List<string> sharedStrings)
        {
            _sharedStrings = sharedStrings ?? throw new ArgumentNullException(nameof(sharedStrings));
        }

        /// <summary>
        /// 解析SST记录（共享字符串表）
        /// </summary>
        /// <param name="record">SST记录</param>
        /// <param name="workbookStreamEnd">工作簿流结束位置（用于安全限制）</param>
        public void ParseSstRecord(BiffRecord record, long workbookStreamEnd)
        {
            if (record.Data == null || record.Data.Length < 8) return;

            int uniqueCount = BitConverter.ToInt32(record.Data, 4);
            // 防止损坏文件中的异常计数值导致超大分配或死循环（Excel 实际限制约 65535 唯一字符串/工作簿）
            const int maxUniqueCount = 2 * 1024 * 1024;
            if (uniqueCount < 0 || uniqueCount > maxUniqueCount)
                uniqueCount = Math.Clamp(uniqueCount, 0, maxUniqueCount);
            _sharedStrings.Capacity = Math.Max(_sharedStrings.Capacity, uniqueCount);

            var stringReader = new BiffStringReader(record, 8); // SST Header size is 8 bytes

            for (int i = 0; i < uniqueCount; i++)
            {
                string str = stringReader.ReadString();
                // Depending on file corruption or incorrect counts, the reader might return empty at EOF 
                // We add it anyway to maintain the index structure, as cells refer to indexes.
                _sharedStrings.Add(str);
            }
        }
    }
}
