using System.IO;

namespace Nedev.XlsToXlsx.Formats.Xls
{
    public class BiffRecord
    {
        public ushort Id { get; set; }
        public ushort Length { get; set; }
        public byte[]? Data { get; set; }

        public static BiffRecord Read(BinaryReader reader)
        {
            var record = new BiffRecord();
            record.Id = reader.ReadUInt16();
            record.Length = reader.ReadUInt16();
            record.Data = reader.ReadBytes(record.Length);
            return record;
        }
    }

    public enum BiffRecordType
    {
        // Workbook records
        BOF = 0x0809,
        EOF = 0x000A,
        SHEET = 0x0085,        // BOUNDSHEET - 工作表信息
        
        // Worksheet records - 单元格类型
        ROW = 0x0208,
        CELL_BLANK = 0x0201,
        CELL_BOOLERR = 0x0205, // BOOLERR - 布尔值和错误值
        CELL_LABEL = 0x0204,   // LABEL - 直接包含字符串的单元格
        CELL_LABELSST = 0x00FD, // LABELSST - 引用SST的字符串单元格
        CELL_NUMBER = 0x0203,
        CELL_RK = 0x027E,
        CELL_FORMULA = 0x0006, // 公式记录
        STRING = 0x0207,       // 公式字符串缓存记录
        MULRK = 0x00BD,        // 多值RK记录（连续数值单元格）
        MULBLANK = 0x00BE,     // 多空白单元格记录
        MERGECELLS = 0x00E5,   // 合并单元格记录
        FORMAT = 0x041E,       // BIFF8 格式记录 (注意: 不是 0x001E)
        
        // 样式记录
        FONT = 0x0031,           // 字体记录
        XF = 0x00E0,             // 扩展格式记录
        PALETTE = 0x0092,        // 调色板记录
        BORDER = 0x00B2,         // 边框记录
        FILL = 0x00F5,           // 填充记录
        
        // 图表记录
        CHART = 0x003D,         // 图表记录
        CHARTTITLE = 0x003E,    // 图表标题
        SERIES = 0x0040,        // 数据系列
        AXIS = 0x0041,          // 坐标轴
        
        // 图片和嵌入对象
        MSODRAWING = 0x00EC,    // 图片和绘图对象 (Escher记录)
        PICTURE = 0x0074,       // 图片记录
        OBJ = 0x005D,           // 嵌入对象
        
        // 数据验证和条件格式
        DV = 0x01B2,            // 数据验证
        CF = 0x01B0,            // 条件格式
        CFHEADER = 0x01B1,      // 条件格式头部
        
        // String table
        SST = 0x00FC,           // 共享字符串表 (注意: 不是 0x00FF)
        CONTINUE = 0x003C,
        
        // 工作表设置记录
        COLINFO = 0x007D,        // 列宽信息
        DEFCOLWIDTH = 0x0055,    // 默认列宽
        DEFAULTROWHEIGHT = 0x0225, // 默认行高
        DIMENSION = 0x0200,      // 工作表范围
        WINDOW2 = 0x023E,        // 工作表窗口设置
        
        // 超链接和注释
        HYPERLINK = 0x01B8,      // 超链接记录
        NOTE = 0x001C,           // 注释/批注记录 (BIFF8 NOTE)
        CELL_RICH_TEXT = 0x00D1, // 富文本记录

        // P3 新增: 打印设置与命名区域
        NAME = 0x0018,           // 命名区域
        HEADER = 0x0014,         // 页眉
        FOOTER = 0x0015,         // 页脚
        LEFTMARGIN = 0x0026,     // 左边距
        RIGHTMARGIN = 0x0027,    // 右边距
        TOPMARGIN = 0x0028,      // 上边距
        BOTTOMMARGIN = 0x0029,   // 下边距
        HCENTER = 0x0083,        // 水平居中
        VCENTER = 0x0084,        // 垂直居中
        PAGESETUP = 0x00A1       // 页面设置 (SETUP)
    }
}