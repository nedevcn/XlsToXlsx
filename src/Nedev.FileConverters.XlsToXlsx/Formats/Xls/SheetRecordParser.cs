using System;
using System.Collections.Generic;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 工作表记录解析器 - 处理SHEET/BOUNDSHEET记录
    /// </summary>
    public class SheetRecordParser
    {
        private readonly List<int> _sheetOffsets;

        public SheetRecordParser(List<int> sheetOffsets)
        {
            _sheetOffsets = sheetOffsets ?? throw new ArgumentNullException(nameof(sheetOffsets));
        }

        /// <summary>
        /// 解析SHEET记录（BOUNDSHEET）
        /// </summary>
        /// <param name="record">SHEET记录</param>
        /// <param name="workbook">工作簿对象</param>
        public void ParseSheetRecord(BiffRecord record, Workbook workbook)
        {
            var worksheet = new Worksheet();
            byte[] data = record.GetAllData();
            
            if (data != null && data.Length >= 8)
            {
                int lbPlyPos = BitConverter.ToInt32(data, 0);
                _sheetOffsets.Add(lbPlyPos);
                Logger.Debug($"BOUNDSHEET: lbPlyPos={lbPlyPos}");
                
                int nameOffset = 6;
                if (data.Length > nameOffset)
                {
                    byte len = data[nameOffset];
                    int pos = nameOffset + 1;
                    worksheet.Name = RichTextParser.ReadBiffStringFromBytes(data, ref pos, len);
                }
            }
            else if (data != null && data.Length >= 4)
            {
                int lbPlyPos = BitConverter.ToInt32(data, 0);
                _sheetOffsets.Add(lbPlyPos);
            }
            
            if (string.IsNullOrEmpty(worksheet.Name))
                worksheet.Name = "Sheet" + (workbook.Worksheets.Count + 1);
                
            workbook.Worksheets.Add(worksheet);
        }
    }
}
