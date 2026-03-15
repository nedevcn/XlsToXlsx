using System;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 数据验证解析器 - 处理数据验证相关的BIFF记录
    /// </summary>
    public class DataValidationParser
    {
        private readonly Workbook _workbook;

        public DataValidationParser(Workbook workbook)
        {
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
        }

        /// <summary>
        /// 解析DV记录 (0x01BE) - 数据验证
        /// </summary>
        public void ParseDVRecord(BiffRecord record, Worksheet worksheet)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 16)
                return;
            var dataValidation = new DataValidation();

            ushort options = BitConverter.ToUInt16(data, 0);
            dataValidation.AllowBlank = (options & 0x01) != 0;

            ushort validationType = BitConverter.ToUInt16(data, 2);
            switch (validationType)
            {
                case 0: dataValidation.Type = "none"; break;
                case 1: dataValidation.Type = "whole"; break;
                case 2: dataValidation.Type = "decimal"; break;
                case 3: dataValidation.Type = "list"; break;
                case 4: dataValidation.Type = "date"; break;
                case 5: dataValidation.Type = "time"; break;
                case 6: dataValidation.Type = "textLength"; break;
                case 7: dataValidation.Type = "custom"; break;
            }

            ushort operatorType = BitConverter.ToUInt16(data, 4);
            switch (operatorType)
            {
                case 0: dataValidation.Operator = "between"; break;
                case 1: dataValidation.Operator = "notBetween"; break;
                case 2: dataValidation.Operator = "equal"; break;
                case 3: dataValidation.Operator = "notEqual"; break;
                case 4: dataValidation.Operator = "greaterThan"; break;
                case 5: dataValidation.Operator = "lessThan"; break;
                case 6: dataValidation.Operator = "greaterThanOrEqual"; break;
                case 7: dataValidation.Operator = "lessThanOrEqual"; break;
            }

            int currentOffset = 6;
            ushort formula1Size = BitConverter.ToUInt16(data, currentOffset); currentOffset += 2;
            ushort formula2Size = BitConverter.ToUInt16(data, currentOffset); currentOffset += 2;

            if (formula1Size > 0 && currentOffset + formula1Size <= data.Length)
            {
                byte[] formula1Bytes = new byte[formula1Size];
                Array.Copy(data, currentOffset, formula1Bytes, 0, formula1Size);
                dataValidation.Formula1 = FormulaDecompiler.Decompile(formula1Bytes, _workbook);
                currentOffset += formula1Size;
            }
            if (formula2Size > 0 && currentOffset + formula2Size <= data.Length)
            {
                byte[] formula2Bytes = new byte[formula2Size];
                Array.Copy(data, currentOffset, formula2Bytes, 0, formula2Size);
                dataValidation.Formula2 = FormulaDecompiler.Decompile(formula2Bytes, _workbook);
                currentOffset += formula2Size;
            }
            if (currentOffset + 8 <= data.Length)
            {
                ushort firstRow = BitConverter.ToUInt16(data, currentOffset); currentOffset += 2;
                ushort lastRow = BitConverter.ToUInt16(data, currentOffset); currentOffset += 2;
                ushort firstCol = BitConverter.ToUInt16(data, currentOffset); currentOffset += 2;
                ushort lastCol = BitConverter.ToUInt16(data, currentOffset);
                dataValidation.Range = $"{ParsingHelpers.ColumnIndexToLetters1Based(firstCol + 1)}{firstRow + 1}:{ParsingHelpers.ColumnIndexToLetters1Based(lastCol + 1)}{lastRow + 1}";
            }
            else
            {
                dataValidation.Range = "A1:A10";
            }
            worksheet.DataValidations.Add(dataValidation);
        }
    }
}
