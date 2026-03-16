using System;
using System.Collections.Generic;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// 全局记录处理器 - 处理Workbook级别的BIFF记录
    /// </summary>
    public class GlobalRecordHandler
    {
        private readonly Workbook _workbook;
        private readonly StyleParser? _styleParser;
        private readonly ExternalLinkParser? _externalLinkParser;
        private readonly XlsDecryptor? _decryptor;
        private readonly List<uint> _sheetOffsets;
        private readonly List<string> _sharedStrings;
        private readonly Action<BiffRecord, long> _parseSstInfo;
        private readonly Action<BiffRecord> _parseNameRecord;
        private readonly Action<BiffRecord> _parseMsoDrawingGroup;
        private readonly string _password;

        public GlobalRecordHandler(
            Workbook workbook,
            StyleParser? styleParser,
            ExternalLinkParser? externalLinkParser,
            XlsDecryptor? decryptor,
            List<uint> sheetOffsets,
            List<string> sharedStrings,
            Action<BiffRecord, long> parseSstInfo,
            Action<BiffRecord> parseNameRecord,
            Action<BiffRecord> parseMsoDrawingGroup,
            string password)
        {
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            _styleParser = styleParser;
            _externalLinkParser = externalLinkParser;
            _decryptor = decryptor;
            _sheetOffsets = sheetOffsets ?? throw new ArgumentNullException(nameof(sheetOffsets));
            _sharedStrings = sharedStrings ?? throw new ArgumentNullException(nameof(sharedStrings));
            _parseSstInfo = parseSstInfo ?? throw new ArgumentNullException(nameof(parseSstInfo));
            _parseNameRecord = parseNameRecord ?? throw new ArgumentNullException(nameof(parseNameRecord));
            _parseMsoDrawingGroup = parseMsoDrawingGroup ?? throw new ArgumentNullException(nameof(parseMsoDrawingGroup));
            _password = password;
        }

        /// <summary>
        /// 创建记录路由器并注册所有全局记录处理器
        /// </summary>
        /// <param name="streamEnd">流结束位置</param>
        /// <returns>配置好的记录路由器</returns>
        public RecordRouter CreateRouter(long streamEnd)
        {
            var router = new RecordRouter();

            // 特殊记录
            router.Register((ushort)BiffRecordType.BOF, _ => { });
            router.Register((ushort)BiffRecordType.EOF, _ => { });

            // 工作表记录
            router.Register((ushort)BiffRecordType.SHEET, r => ParseSheetRecord(r));

            // 字符串表
            router.Register((ushort)BiffRecordType.SST, r => _parseSstInfo(r, streamEnd));

            // 样式相关
            router.Register((ushort)BiffRecordType.FONT, r => _styleParser?.ParseFontRecord(r));
            router.Register((ushort)BiffRecordType.XF, r => _styleParser?.ParseXfRecord(r));
            router.Register((ushort)BiffRecordType.FORMAT, r => _styleParser?.ParseFormatRecord(r));
            router.Register((ushort)BiffRecordType.PALETTE, r => _styleParser?.ParsePaletteRecord(r));
            router.Register((ushort)BiffRecordType.BORDER, r => _styleParser?.ParseBorderRecord(r));
            router.Register((ushort)BiffRecordType.FILL, r => _styleParser?.ParseFillRecord(r));

            // 名称定义
            router.Register((ushort)BiffRecordType.NAME, r => _parseNameRecord(r));

            // 绘图组
            router.Register((ushort)BiffRecordType.MSODRAWINGGROUP, r => _parseMsoDrawingGroup(r));

            // 加密
            router.Register((ushort)BiffRecordType.FILEPASS, r => ParseFilePassRecord(r));

            // 外部链接
            router.Register((ushort)BiffRecordType.EXTERNBOOK, r => _externalLinkParser?.ParseExternBookRecord(r, _workbook));
            router.Register((ushort)BiffRecordType.EXTERNSHEET, r => _externalLinkParser?.ParseExternSheetRecord(r, _workbook));
            router.Register((ushort)BiffRecordType.EXTERNALNAME, r => _externalLinkParser?.ParseExternalNameRecord(r, _workbook));

            // 保护
            router.Register((ushort)BiffRecordType.PROTECT, r => ParseProtectRecord(r));

            return router;
        }

        /// <summary>
        /// 解析SHEET记录 (BOUNDSHEET)
        /// </summary>
        private void ParseSheetRecord(BiffRecord record)
        {
            var worksheet = new Worksheet();
            byte[] data = record.GetAllData();
            if (data != null && data.Length >= 8)
            {
                int lbPlyPos = BitConverter.ToInt32(data, 0);
                _sheetOffsets.Add((uint)lbPlyPos);
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
                _sheetOffsets.Add((uint)lbPlyPos);
            }
            if (string.IsNullOrEmpty(worksheet.Name))
                worksheet.Name = "Sheet" + (_workbook.Worksheets.Count + 1);
            _workbook.Worksheets.Add(worksheet);
        }

        /// <summary>
        /// 解析FILEPASS记录（加密）
        /// </summary>
        private void ParseFilePassRecord(BiffRecord record)
        {
            // 此方法仅用于注册，实际解密在ParseWorkbookGlobals中处理
            // 这里不需要额外处理，因为解密器已经在ParseWorkbookGlobals中创建
        }

        /// <summary>
        /// 解析PROTECT记录（工作簿保护）
        /// </summary>
        private void ParseProtectRecord(BiffRecord record)
        {
            if (record.Data != null && record.Data.Length >= 2)
            {
                ushort flags = BitConverter.ToUInt16(record.Data, 0);
                _workbook.IsStructureProtected = (flags & 0x0001) != 0;
            }
        }
    }
}
