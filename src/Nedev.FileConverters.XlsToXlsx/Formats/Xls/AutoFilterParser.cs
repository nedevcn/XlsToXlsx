using System;
using System.Collections.Generic;
using Nedev.FileConverters.XlsToXlsx;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 自动筛选解析器 - 处理自动筛选相关的BIFF记录
    /// </summary>
    public class AutoFilterParser
    {
        private readonly int _currentSheetIndex;

        public AutoFilterParser(int currentSheetIndex)
        {
            _currentSheetIndex = currentSheetIndex;
        }

        /// <summary>
        /// 解析AUTOFILTERINFO记录 (0x009D) - 自动筛选信息
        /// </summary>
        public void ParseAutoFilterInfoRecord(BiffRecord record, Worksheet worksheet, Workbook workbook)
        {
            if (record.Data == null || record.Data.Length < 2)
                return;

            ushort cEntries = BitConverter.ToUInt16(record.Data, 0);
            worksheet.AutoFilterColumnIndices.Clear();

            // 从定义名称中查找筛选范围
            if (workbook.DefinedNames != null)
            {
                foreach (var dn in workbook.DefinedNames)
                {
                    if (dn?.Name != "FilterDatabase" && dn?.Name != "_xlnm._FilterDatabase")
                        continue;
                    if (dn.LocalSheetId.HasValue && dn.LocalSheetId.Value != _currentSheetIndex)
                        continue;
                    if (string.IsNullOrEmpty(dn.Formula))
                        continue;

                    worksheet.AutoFilterRange = StripDefinedNameToRange(dn.Formula);
                    break;
                }
            }

            // 如果没有找到范围，使用默认值
            if (string.IsNullOrEmpty(worksheet.AutoFilterRange))
                worksheet.AutoFilterRange = "A1:Z100";
        }

        /// <summary>
        /// 解析AUTOFILTER记录 (0x009E) - 自动筛选条件
        /// </summary>
        public void ParseAutoFilterRecord(BiffRecord record, Worksheet worksheet)
        {
            if (record.Data == null || record.Data.Length < 2)
                return;

            ushort iEntry = BitConverter.ToUInt16(record.Data, 0);

            if (worksheet.AutoFilterColumnIndices == null)
                worksheet.AutoFilterColumnIndices = new List<int>();

            worksheet.AutoFilterColumnIndices.Add((int)iEntry);
        }

        /// <summary>
        /// 从定义名称公式中提取范围
        /// </summary>
        private static string StripDefinedNameToRange(string formula)
        {
            if (string.IsNullOrEmpty(formula))
                return formula;

            int excl = formula.IndexOf('!');
            string rangePart = excl >= 0 ? formula.Substring(excl + 1).Trim() : formula;
            return rangePart.Replace("$", "");
        }
    }
}
