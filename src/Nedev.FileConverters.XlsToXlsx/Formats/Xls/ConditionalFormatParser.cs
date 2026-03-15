using System;
using Nedev.FileConverters.XlsToXlsx;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 条件格式解析器 - 处理条件格式相关的BIFF记录
    /// </summary>
    public class ConditionalFormatParser
    {
        private readonly Workbook _workbook;
        private string _currentCFRange = string.Empty;

        public ConditionalFormatParser(Workbook workbook)
        {
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
        }

        /// <summary>
        /// 解析CFHEADER记录 (0x01B0) - 条件格式头
        /// </summary>
        public void ParseCFHeaderRecord(BiffRecord record)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 12)
                return;

            ushort firstRow = BitConverter.ToUInt16(data, 2);
            ushort lastRow = BitConverter.ToUInt16(data, 4);
            ushort firstCol = BitConverter.ToUInt16(data, 6);
            ushort lastCol = BitConverter.ToUInt16(data, 8);

            _currentCFRange = $"{ParsingHelpers.ColumnIndexToLetters1Based(firstCol + 1)}{firstRow + 1}:{ParsingHelpers.ColumnIndexToLetters1Based(lastCol + 1)}{lastRow + 1}";
        }

        /// <summary>
        /// 解析CF记录 (0x01B1) - 条件格式规则
        /// </summary>
        public void ParseCFRecord(BiffRecord record, Worksheet worksheet)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 8)
                return;

            var conditionalFormat = new ConditionalFormat();

            // 解析条件类型
            ushort conditionType = BitConverter.ToUInt16(data, 0);
            conditionalFormat.Type = conditionType switch
            {
                0 => "cellIs",
                1 => "expression",
                2 => "colorScale",
                3 => "dataBar",
                4 => "iconSet",
                _ => "cellIs"
            };

            // 解析操作符
            ushort operatorType = BitConverter.ToUInt16(data, 2);
            conditionalFormat.Operator = operatorType switch
            {
                0 => "between",
                1 => "notBetween",
                2 => "equal",
                3 => "notEqual",
                4 => "greaterThan",
                5 => "lessThan",
                6 => "greaterThanOrEqual",
                7 => "lessThanOrEqual",
                8 => "containsText",
                9 => "notContainsText",
                10 => "beginsWith",
                11 => "endsWith",
                _ => "equal"
            };

            // 解析公式
            if (data.Length >= 12)
            {
                int currentOffset = 4;
                ushort formula1Size = BitConverter.ToUInt16(data, currentOffset);
                currentOffset += 2;
                ushort formula2Size = BitConverter.ToUInt16(data, currentOffset);
                currentOffset += 2;

                if (currentOffset + 4 <= data.Length)
                {
                    currentOffset += 4; // 跳过保留字节

                    if (formula1Size > 0 && currentOffset + formula1Size <= data.Length)
                    {
                        byte[] ptg1 = new byte[formula1Size];
                        Array.Copy(data, currentOffset, ptg1, 0, formula1Size);
                        conditionalFormat.Formula = FormulaDecompiler.Decompile(ptg1, _workbook);
                        currentOffset += formula1Size;
                    }

                    if (formula2Size > 0 && currentOffset + formula2Size <= data.Length)
                    {
                        byte[] ptg2 = new byte[formula2Size];
                        Array.Copy(data, currentOffset, ptg2, 0, formula2Size);
                        conditionalFormat.Formula2 = FormulaDecompiler.Decompile(ptg2, _workbook);
                    }
                }
            }

            conditionalFormat.Range = !string.IsNullOrEmpty(_currentCFRange) ? _currentCFRange : "A1:A10";
            worksheet.ConditionalFormats.Add(conditionalFormat);
        }

        /// <summary>
        /// 重置当前条件格式范围
        /// </summary>
        public void ResetRange()
        {
            _currentCFRange = string.Empty;
        }
    }
}
