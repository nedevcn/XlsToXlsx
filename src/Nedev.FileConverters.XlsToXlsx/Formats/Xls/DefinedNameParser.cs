using System;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 名称定义解析器 - 处理Excel中的定义名称（Named Ranges）
    /// </summary>
    public class DefinedNameParser
    {
        private readonly Workbook _workbook;

        public DefinedNameParser(Workbook workbook)
        {
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
        }

        /// <summary>
        /// 解析NAME记录（定义名称）
        /// </summary>
        /// <param name="record">NAME记录</param>
        public void ParseNameRecord(BiffRecord record)
        {
            byte[] data = record.GetAllData();
            if (data == null || data.Length < 14)
                return;

            ushort options = BitConverter.ToUInt16(data, 0);
            byte nameLen = data[3];
            ushort formulaLen = BitConverter.ToUInt16(data, 4);
            bool hidden = (options & 0x0001) != 0;

            // itab 在偏移 8-9：非 0 表示局部名称，值为 1-based 的 BoundSheet8 索引
            ushort itab = data.Length >= 10 ? BitConverter.ToUInt16(data, 8) : (ushort)0;
            int localSheetId = itab > 0 ? (int)(itab - 1) : 0;

            int offset = 14;
            string name = RichTextParser.ReadBiffStringFromBytes(data, ref offset, nameLen);

            // 处理特殊名称（如FilterDatabase）
            if (nameLen == 1 && name.Length == 1 && name[0] == '\u000D')
                name = "FilterDatabase";

            // 提取公式数据
            byte[] formulaData = new byte[formulaLen];
            if (formulaLen > 0 && offset + formulaLen <= data.Length)
                Array.Copy(data, offset, formulaData, 0, formulaLen);

            // 反编译公式
            string formula = FormulaDecompiler.Decompile(formulaData, _workbook);

            // 添加到工作簿
            _workbook.DefinedNames.Add(new DefinedName
            {
                Name = name,
                Formula = formula,
                Hidden = hidden,
                LocalSheetId = itab > 0 ? (int?)(localSheetId) : null
            });
        }
    }
}
