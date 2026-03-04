using System;
using System.IO;

namespace Nedev.XlsToXlsx.Formats.Xls
{
    public class BiffRecord
    {
        public ushort Id { get; set; }
        public ushort Length { get; set; }
        public byte[]? Data { get; set; }
        
        /// <summary>
        /// 存储后续所有的 CONTINUE (0x003C) 记录数据
        /// </summary>
        public System.Collections.Generic.List<byte[]> Continues { get; set; } = new System.Collections.Generic.List<byte[]>();

        /// <summary>
        /// 拼接基础 Data 和所有 Continues 的数据为一个完整数组
        /// </summary>
        public byte[] GetAllData()
        {
            if (Data == null) return Array.Empty<byte>();
            if (Continues.Count == 0) return Data;

            int totalLength = Data.Length;
            foreach (var chunk in Continues)
            {
                totalLength += chunk.Length;
            }

            byte[] fullData = new byte[totalLength];
            Array.Copy(Data, 0, fullData, 0, Data.Length);
            
            int offset = Data.Length;
            foreach (var chunk in Continues)
            {
                Array.Copy(chunk, 0, fullData, offset, chunk.Length);
                offset += chunk.Length;
            }
            
            return fullData;
        }

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
        FILEPASS = 0x002F,     // 文件加密信息
        EXTERNSHEET = 0x0017,  // 外部引用映射记录
        EXTERNBOOK = 0x01AE,   // 外部工作簿引用 (SUPBOOK)
        EXTERNALNAME = 0x0023, // 外部名称记录
        
        // Worksheet records - 单元格类型
        ROW = 0x0208,
        CELL_BLANK = 0x0201,
        CELL_BOOLERR = 0x0205, // BOOLERR - 布尔值和错误值
        CELL_LABEL = 0x0204,   // LABEL - 直接包含字符串的单元格
        CELL_LABELSST = 0x00FD, // LABELSST - 引用SST的字符串单元格
        CELL_NUMBER = 0x0203,
        CELL_RK = 0x027E,
        CELL_FORMULA = 0x0006, // 公式记录
        ARRAY = 0x0221,        // 数组公式
        SHAREDFMLA = 0x04BC,   // 共享公式
        STRING = 0x0207,       // 公式字符串缓存记录
        CELL_RSTRING = 0x00D6, // 旧版富文本或带格式文本
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
        
        // 图表记录 (BIFF8 Chart Substream)
        CHARTFORMAT = 0x1014,   // 图表格式记录 (Start of chart)
        CHARTSERIES = 0x1003,   // 图表数据系列
        CHARTTITLE = 0x1025,    // 图表标题
        CHARTLEGEND = 0x1015,   // 图表图例
        CHARTAXIS = 0x101D,     // 坐标轴
        CHARTLINEFORMAT = 0x1007, // 线条格式
        CHARTAREA = 0x101A,     // 图表区/绘图区
        CHARTMARKERFORMAT = 0x1009, // 标记格式
        CHART3D = 0x103A,       // 3D 图表标志
        CHARTFORMATLINK = 0x1022,   // 链接到文本
        SERIESTEXT = 0x100D,    // 系列文本/标题
        CHARTEND = 0x1033,      // Chart结束标
        BRAI = 0x1051,          // 系列数据引用
        
        // 图片和嵌入对象
        MSODRAWINGGROUP = 0x00EB, // Office Art 绘图全局(包含DggContainer)
        MSODRAWING = 0x00EC,      // 图片和绘图对象 (Escher记录)
        PICTURE = 0x0074,         // 图片记录
        OBJ = 0x005D,             // 嵌入对象
        
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