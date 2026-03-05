using System.Collections.Generic;

namespace Nedev.XlsToXlsx
{
    /// <summary>
    /// 工作簿类，包含Excel文件的所有内容
    /// </summary>
    public class Workbook
    {
        /// <summary>
        /// 工作表列表
        /// </summary>
        public List<Worksheet> Worksheets { get; set; } = new List<Worksheet>();
        /// <summary>
        /// 作者
        /// </summary>
        public string? Author { get; set; }
        /// <summary>
        /// 标题
        /// </summary>
        public string? Title { get; set; }
        /// <summary>
        /// 主题
        /// </summary>
        public string? Subject { get; set; }
        /// <summary>
        /// 关键词
        /// </summary>
        public string? Keywords { get; set; }
        /// <summary>
        /// 注释
        /// </summary>
        public string? Comments { get; set; }
        /// <summary>
        /// VBA项目数据
        /// </summary>
        public byte[]? VbaProject { get; set; }
        /// <summary>
        /// 样式列表
        /// </summary>
        public List<Style> Styles { get; set; } = new List<Style>();
        /// <summary>
        /// 超链接列表
        /// </summary>
        public List<Hyperlink> Hyperlinks { get; set; } = new List<Hyperlink>();
        /// <summary>
        /// 共享字符串列表
        /// </summary>
        public List<string> SharedStrings { get; set; } = new List<string>();
        /// <summary>
        /// 字体列表
        /// </summary>
        public List<Font> Fonts { get; set; } = new List<Font>();
        /// <summary>
        /// 扩展格式列表
        /// </summary>
        public List<Xf> XfList { get; set; } = new List<Xf>();
        /// <summary>
        /// 数字格式映射
        /// </summary>
        public Dictionary<ushort, string> NumberFormats { get; set; } = new Dictionary<ushort, string>();
        /// <summary>
        /// 调色板映射
        /// </summary>
        public Dictionary<int, string> Palette { get; set; } = new Dictionary<int, string>();
        /// <summary>
        /// 命名区域列表
        /// </summary>
        public List<DefinedName> DefinedNames { get; set; } = new List<DefinedName>();
        /// <summary>
        /// 边框列表
        /// </summary>
        public List<Border> Borders { get; set; } = new List<Border>();
        /// <summary>
        /// 填充列表
        /// </summary>
        public List<Fill> Fills { get; set; } = new List<Fill>();
        /// <summary>
        /// 外部工作簿引用列表
        /// </summary>
        public List<ExternalBook> ExternalBooks { get; set; } = new List<ExternalBook>();
        /// <summary>
        /// 外部工作表引用索引表
        /// </summary>
        public List<ExternalSheet> ExternalSheets { get; set; } = new List<ExternalSheet>();
    }

    /// <summary>
    /// 工作表类
    /// </summary>
    public class Worksheet
    {
        /// <summary>
        /// 工作表名称
        /// </summary>
        public string? Name { get; set; }
        /// <summary>
        /// 行列表
        /// </summary>
        public List<Row> Rows { get; set; } = new List<Row>();
        /// <summary>
        /// 最大列索引
        /// </summary>
        public int MaxColumn { get; set; }
        /// <summary>
        /// 最大行索引
        /// </summary>
        public int MaxRow { get; set; }
        /// <summary>
        /// 合并单元格列表
        /// </summary>
        public List<MergeCell> MergeCells { get; set; } = new List<MergeCell>();
        /// <summary>
        /// 图表列表
        /// </summary>
        public List<Chart> Charts { get; set; } = new List<Chart>();
        /// <summary>
        /// 图片列表
        /// </summary>
        public List<Picture> Pictures { get; set; } = new List<Picture>();
        /// <summary>
        /// 嵌入对象列表
        /// </summary>
        public List<EmbeddedObject> EmbeddedObjects { get; set; } = new List<EmbeddedObject>();
        /// <summary>
        /// 默认列宽 (字符数)
        /// </summary>
        public double? DefaultColumnWidth { get; set; }
        /// <summary>
        /// 默认行高 (点数)
        /// </summary>
        public double? DefaultRowHeight { get; set; }
        /// <summary>
        /// 数据验证列表
        /// </summary>
        public List<DataValidation> DataValidations { get; set; } = new List<DataValidation>();
        /// <summary>
        /// 条件格式列表
        /// </summary>
        public List<ConditionalFormat> ConditionalFormats { get; set; } = new List<ConditionalFormat>();
        /// <summary>
        /// 超链接列表
        /// </summary>
        public List<Hyperlink> Hyperlinks { get; set; } = new List<Hyperlink>();
        /// <summary>
        /// 注释列表
        /// </summary>
        public List<Comment> Comments { get; set; } = new List<Comment>();
        /// <summary>
        /// 列宽信息列表
        /// </summary>
        public List<ColumnInfo> ColumnInfos { get; set; } = new List<ColumnInfo>();
        /// <summary>
        /// 冻结窗格设置
        /// </summary>
        public FreezePane? FreezePane { get; set; }
        /// <summary>
        /// 字体列表
        /// </summary>
        public List<Font> Fonts { get; set; } = new List<Font>();
        /// <summary>
        /// 扩展格式列表
        /// </summary>
        public List<Xf> Xfs { get; set; } = new List<Xf>();
        /// <summary>
        /// 调色板
        /// </summary>
        public Dictionary<int, string> Palette { get; set; } = new Dictionary<int, string>();
        /// <summary>
        /// 页面设置
        /// </summary>
        public PageSettings PageSettings { get; set; } = new PageSettings();
        /// <summary>
        /// 数据透视表列表
        /// </summary>
        public List<PivotTable> PivotTables { get; set; } = new List<PivotTable>();
    }

    /// <summary>
    /// 数据透视表类
    /// </summary>
    public class PivotTable
    {
        /// <summary>
        /// 数据透视表名称
        /// </summary>
        public string? Name { get; set; }
        /// <summary>
        /// 数据透视表范围
        /// </summary>
        public string? Range { get; set; }
        /// <summary>
        /// 数据源范围
        /// </summary>
        public string? DataSource { get; set; }
        /// <summary>
        /// 行字段列表
        /// </summary>
        public List<PivotField> RowFields { get; set; } = new List<PivotField>();
        /// <summary>
        /// 列字段列表
        /// </summary>
        public List<PivotField> ColumnFields { get; set; } = new List<PivotField>();
        /// <summary>
        /// 数据字段列表
        /// </summary>
        public List<PivotField> DataFields { get; set; } = new List<PivotField>();
        /// <summary>
        /// 页字段列表
        /// </summary>
        public List<PivotField> PageFields { get; set; } = new List<PivotField>();
    }

    /// <summary>
    /// 数据透视表字段类
    /// </summary>
    public class PivotField
    {
        /// <summary>
        /// 字段名称
        /// </summary>
        public string? Name { get; set; }
        /// <summary>
        /// 字段来源
        /// </summary>
        public string? SourceName { get; set; }
        /// <summary>
        /// 字段类型：row, column, data, page
        /// </summary>
        public string? Type { get; set; }
        /// <summary>
        /// 汇总函数：sum, count, average, max, min, product, countNums, stdDev, stdDevp, var, varp
        /// </summary>
        public string? Function { get; set; }
        /// <summary>
        /// 是否可见
        /// </summary>
        public bool Visible { get; set; } = true;
        /// <summary>
        /// 排序方式：ascending, descending, manual
        /// </summary>
        public string? SortType { get; set; }
        /// <summary>
        /// 筛选值
        /// </summary>
        public List<object>? FilterValues { get; set; }
    }

    /// <summary>
    /// 图片类
    /// </summary>
    public class Picture
    {
        /// <summary>
        /// 图片数据
        /// </summary>
        public byte[]? Data { get; set; }
        /// <summary>
        /// MIME类型
        /// </summary>
        public string? MimeType { get; set; }
        /// <summary>
        /// 文件扩展名
        /// </summary>
        public string? Extension { get; set; }
        /// <summary>
        /// 左边距
        /// </summary>
        public int Left { get; set; }
        /// <summary>
        /// 上边距
        /// </summary>
        public int Top { get; set; }
        /// <summary>
        /// 宽度
        /// </summary>
        public int Width { get; set; }
        /// <summary>
        /// 高度
        /// </summary>
        public int Height { get; set; }
    }

    /// <summary>
    /// 嵌入对象类
    /// </summary>
    public class EmbeddedObject
    {
        /// <summary>
        /// 对象名称
        /// </summary>
        public string? Name { get; set; }
        /// <summary>
        /// 对象数据
        /// </summary>
        public byte[]? Data { get; set; }
        /// <summary>
        /// MIME类型
        /// </summary>
        public string? MimeType { get; set; }
    }

    /// <summary>
    /// 数据验证类
    /// </summary>
    public class DataValidation
    {
        /// <summary>
        /// 验证范围
        /// </summary>
        public string? Range { get; set; }
        /// <summary>
        /// 第一个公式
        /// </summary>
        public string? Formula1 { get; set; }
        /// <summary>
        /// 第二个公式
        /// </summary>
        public string? Formula2 { get; set; }
        /// <summary>
        /// 是否允许空白
        /// </summary>
        public bool AllowBlank { get; set; }
        /// <summary>
        /// 数据验证类型：whole, decimal, list, date, time, textLength, custom
        /// </summary>
        public string? Type { get; set; }
        /// <summary>
        /// 操作符：between, notBetween, equal, notEqual, greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual
        /// </summary>
        public string? Operator { get; set; }
    }

    /// <summary>
    /// 条件格式类
    /// </summary>
    public class ConditionalFormat
    {
        /// <summary>
        /// 条件格式范围
        /// </summary>
        public string? Range { get; set; }
        /// <summary>
        /// 条件公式
        /// </summary>
        public string? Formula { get; set; }
        /// <summary>
        /// 格式
        /// </summary>
        public string? Format { get; set; }
        /// <summary>
        /// 条件格式类型：cellIs, expression, colorScale, dataBar, iconSet
        /// </summary>
        public string? Type { get; set; }
        /// <summary>
        /// 操作符：between, notBetween, equal, notEqual, greaterThan, lessThan, greaterThanOrEqual, lessThanOrEqual, containsText, notContainsText, beginsWith, endsWith
        /// </summary>
        public string? Operator { get; set; }
        /// <summary>
        /// 第一个值
        /// </summary>
        public object? Value1 { get; set; }
        /// <summary>
        /// 第二个值
        /// </summary>
        public object? Value2 { get; set; }
    }

    /// <summary>
    /// 超链接类
    /// </summary>
    public class Hyperlink
    {
        /// <summary>
        /// 超链接范围
        /// </summary>
        public string? Range { get; set; }
        /// <summary>
        /// 目标地址
        /// </summary>
        public string? Target { get; set; }
        /// <summary>
        /// 显示文本
        /// </summary>
        public string? DisplayText { get; set; }
    }

    /// <summary>
    /// 注释类
    /// </summary>
    public class Comment
    {
        /// <summary>
        /// 行索引
        /// </summary>
        public int RowIndex { get; set; }
        /// <summary>
        /// 列索引
        /// </summary>
        public int ColumnIndex { get; set; }
        /// <summary>
        /// 作者
        /// </summary>
        public string? Author { get; set; }
        /// <summary>
        /// 注释文本
        /// </summary>
        public string? Text { get; set; }
        /// <summary>
        /// 富文本内容
        /// </summary>
        public List<RichTextRun>? RichText { get; set; }
    }

    /// <summary>
    /// 合并单元格类
    /// </summary>
    public class MergeCell
    {
        /// <summary>
        /// 起始行
        /// </summary>
        public int StartRow { get; set; }
        /// <summary>
        /// 起始列
        /// </summary>
        public int StartColumn { get; set; }
        /// <summary>
        /// 结束行
        /// </summary>
        public int EndRow { get; set; }
        /// <summary>
        /// 结束列
        /// </summary>
        public int EndColumn { get; set; }
    }

    /// <summary>
    /// 图表类
    /// </summary>
    public class Chart
    {
        /// <summary>
        /// 图表标题
        /// </summary>
        public string? Title { get; set; }
        /// <summary>
        /// 图表类型
        /// </summary>
        public string? ChartType { get; set; }
        /// <summary>
        /// 数据系列列表
        /// </summary>
        public List<Series> Series { get; set; } = new List<Series>();
        /// <summary>
        /// X轴标题
        /// </summary>
        public string? XAxisTitle { get; set; }
        /// <summary>
        /// Y轴标题
        /// </summary>
        public string? YAxisTitle { get; set; }
        /// <summary>
        /// 宽度
        /// </summary>
        public int Width { get; set; } = 500;
        /// <summary>
        /// 高度
        /// </summary>
        public int Height { get; set; } = 300;
        /// <summary>
        /// 左边距
        /// </summary>
        public int Left { get; set; } = 10;
        /// <summary>
        /// 上边距
        /// </summary>
        public int Top { get; set; } = 10;
        /// <summary>
        /// 图例
        /// </summary>
        public Legend? Legend { get; set; }
        /// <summary>
        /// 绘图区
        /// </summary>
        public PlotArea? PlotArea { get; set; }
        /// <summary>
        /// X轴
        /// </summary>
        public Axis? XAxis { get; set; }
        /// <summary>
        /// Y轴
        /// </summary>
        public Axis? YAxis { get; set; }
    }

    /// <summary>
    /// 数据系列类
    /// </summary>
    public class Series
    {
        /// <summary>
        /// 系列名称
        /// </summary>
        public string? Name { get; set; }
        /// <summary>
        /// 值范围
        /// </summary>
        public string? ValuesRange { get; set; }
        /// <summary>
        /// 类别范围
        /// </summary>
        public string? CategoriesRange { get; set; }
        /// <summary>
        /// 颜色
        /// </summary>
        public string? Color { get; set; }
        /// <summary>
        /// 标记
        /// </summary>
        public Marker? Marker { get; set; }
        /// <summary>
        /// 线条样式
        /// </summary>
        public LineStyle? LineStyle { get; set; }
        /// <summary>
        /// 填充样式
        /// </summary>
        public FillStyle? FillStyle { get; set; }
        /// <summary>
        /// 填充颜色
        /// </summary>
        public string? FillColor { get; set; }
    }

    /// <summary>
    /// 图例类
    /// </summary>
    public class Legend
    {
        /// <summary>
        /// 是否可见
        /// </summary>
        public bool Visible { get; set; } = true;
        /// <summary>
        /// 位置
        /// </summary>
        public string? Position { get; set; } = "right";
        /// <summary>
        /// 字体
        /// </summary>
        public Font? Font { get; set; }
    }

    /// <summary>
    /// 绘图区类
    /// </summary>
    public class PlotArea
    {
        /// <summary>
        /// 填充样式
        /// </summary>
        public FillStyle? Fill { get; set; }
        /// <summary>
        /// 边框
        /// </summary>
        public Border? Border { get; set; }
    }

    /// <summary>
    /// 坐标轴类
    /// </summary>
    public class Axis
    {
        /// <summary>
        /// 标题
        /// </summary>
        public string? Title { get; set; }
        /// <summary>
        /// 是否可见
        /// </summary>
        public bool Visible { get; set; } = true;
        /// <summary>
        /// 标题字体
        /// </summary>
        public Font? TitleFont { get; set; }
        /// <summary>
        /// 标签字体
        /// </summary>
        public Font? LabelFont { get; set; }
        /// <summary>
        /// 数字格式
        /// </summary>
        public string? NumberFormat { get; set; }
    }

    /// <summary>
    /// 标记类
    /// </summary>
    public class Marker
    {
        /// <summary>
        /// 是否可见
        /// </summary>
        public bool Visible { get; set; } = false;
        /// <summary>
        /// 类型
        /// </summary>
        public string? Type { get; set; } = "circle";
        /// <summary>
        /// 大小
        /// </summary>
        public int Size { get; set; } = 5;
        /// <summary>
        /// 颜色
        /// </summary>
        public string? Color { get; set; }
    }

    /// <summary>
    /// 线条样式类
    /// </summary>
    public class LineStyle
    {
        /// <summary>
        /// 颜色
        /// </summary>
        public string? Color { get; set; }
        /// <summary>
        /// 宽度
        /// </summary>
        public int Width { get; set; } = 1;
        /// <summary>
        ///  dash类型
        /// </summary>
        public string? DashStyle { get; set; } = "solid";
    }

    /// <summary>
    /// 填充样式类
    /// </summary>
    public class FillStyle
    {
        /// <summary>
        /// 颜色
        /// </summary>
        public string? Color { get; set; }
        /// <summary>
        /// 图案
        /// </summary>
        public string? Pattern { get; set; } = "solid";
        /// <summary>
        /// 透明度
        /// </summary>
        public double Transparency { get; set; } = 0;
    }

    /// <summary>
    /// 列宽信息类（对应 COLINFO BIFF 记录）
    /// </summary>
    public class ColumnInfo
    {
        /// <summary>
        /// 起始列（0-based）
        /// </summary>
        public int FirstColumn { get; set; }
        /// <summary>
        /// 结束列（0-based）
        /// </summary>
        public int LastColumn { get; set; }
        /// <summary>
        /// 列宽（单位：1/256 字符宽度）
        /// </summary>
        public int Width { get; set; }
        /// <summary>
        /// XF 索引
        /// </summary>
        public int XfIndex { get; set; }
        /// <summary>
        /// 是否隐藏
        /// </summary>
        public bool Hidden { get; set; }
    }

    /// <summary>
    /// 冻结窗格类
    /// </summary>
    public class FreezePane
    {
        /// <summary>
        /// 冻结行数（垂直分割位置）
        /// </summary>
        public int RowSplit { get; set; }
        /// <summary>
        /// 冻结列数（水平分割位置）
        /// </summary>
        public int ColSplit { get; set; }
        /// <summary>
        /// 右下窗格顶行
        /// </summary>
        public int TopRow { get; set; }
        /// <summary>
        /// 右下窗格左列
        /// </summary>
        public int LeftCol { get; set; }
    }

    /// <summary>
    /// 行类
    /// </summary>
    public class Row
    {
        /// <summary>
        /// 行索引
        /// </summary>
        public int RowIndex { get; set; }
        /// <summary>
        /// 单元格列表
        /// </summary>
        public List<Cell> Cells { get; set; } = new List<Cell>();
        /// <summary>
        /// 行高（单位：1/20点，即 twips）
        /// </summary>
        public ushort Height { get; set; }
        /// <summary>
        /// 是否自定义行高
        /// </summary>
        public bool CustomHeight { get; set; }
        /// <summary>
        /// 行默认 XF 索引（来自 ROW 记录的 ixfe），用于整行格式如背景色
        /// </summary>
        public int? DefaultXfIndex { get; set; }
    }

    /// <summary>
    /// 单元格类
    /// </summary>
    public class Cell
    {
        /// <summary>
        /// 行索引
        /// </summary>
        public int RowIndex { get; set; }
        /// <summary>
        /// 列索引
        /// </summary>
        public int ColumnIndex { get; set; }
        /// <summary>
        /// 单元格值
        /// </summary>
        public object? Value { get; set; }
        /// <summary>
        /// 公式
        /// </summary>
        public string? Formula { get; set; }
        /// <summary>
        /// 数据类型
        /// </summary>
        public string? DataType { get; set; }
        /// <summary>
        /// 样式ID
        /// </summary>
        public string? StyleId { get; set; }
        /// <summary>
        /// 富文本内容
        /// </summary>
        public List<RichTextRun>? RichText { get; set; }
        /// <summary>
        /// 是否为数组公式
        /// </summary>
        public bool IsArrayFormula { get; set; }
        /// <summary>
        /// 数组公式范围（如 "A1:C3"），仅当 IsArrayFormula 为 true 时有效
        /// </summary>
        public string? ArrayRef { get; set; }
    }

    /// <summary>
    /// 富文本运行类
    /// </summary>
    public class RichTextRun
    {
        /// <summary>
        /// 文本
        /// </summary>
        public string? Text { get; set; }
        /// <summary>
        /// 字体
        /// </summary>
        public Font? Font { get; set; }
    }

    /// <summary>
    /// 样式类
    /// </summary>
    public class Style
    {
        /// <summary>
        /// 样式ID
        /// </summary>
        public string? Id { get; set; }
        /// <summary>
        /// 字体
        /// </summary>
        public Font? Font { get; set; }
        /// <summary>
        /// 填充
        /// </summary>
        public Fill? Fill { get; set; }
        /// <summary>
        /// 边框
        /// </summary>
        public Border? Border { get; set; }
        /// <summary>
        /// 对齐方式
        /// </summary>
        public Alignment? Alignment { get; set; }
        /// <summary>
        /// 数字格式
        /// </summary>
        public string? NumberFormat { get; set; }
        /// <summary>
        /// 保护设置
        /// </summary>
        public Protection? Protection { get; set; }
    }

    /// <summary>
    /// 保护设置类
    /// </summary>
    public class Protection
    {
        public bool? Locked { get; set; }
        public bool? Hidden { get; set; }
    }

    /// <summary>
    /// 字体类
    /// </summary>
    public class Font
    {
        /// <summary>
        /// 字体名称
        /// </summary>
        public string? Name { get; set; }
        /// <summary>
        /// 字体大小
        /// </summary>
        public double? Size { get; set; }
        /// <summary>
        /// 是否粗体
        /// </summary>
        public bool? Bold { get; set; }
        /// <summary>
        /// 是否斜体
        /// </summary>
        public bool? Italic { get; set; }
        /// <summary>
        /// 是否下划线
        /// </summary>
        public bool? Underline { get; set; }
        /// <summary>
        /// 颜色
        /// </summary>
        public string? Color { get; set; }
        /// <summary>
        /// 字体高度（单位：1/20点）
        /// </summary>
        public short Height { get; set; }
        /// <summary>
        /// 是否粗体
        /// </summary>
        public bool IsBold { get; set; }
        /// <summary>
        /// 是否斜体
        /// </summary>
        public bool IsItalic { get; set; }
        /// <summary>
        /// 是否下划线
        /// </summary>
        public bool IsUnderline { get; set; }
        /// <summary>
        /// 是否删除线
        /// </summary>
        public bool IsStrikethrough { get; set; }
        /// <summary>
        /// 颜色索引
        /// </summary>
        public ushort ColorIndex { get; set; }
    }

    /// <summary>
    /// 填充类
    /// </summary>
    public class Fill
    {
        /// <summary>
        /// 图案类型
        /// </summary>
        public string? PatternType { get; set; }
        /// <summary>
        /// 前景色
        /// </summary>
        public string? ForegroundColor { get; set; }
        /// <summary>
        /// 背景色
        /// </summary>
        public string? BackgroundColor { get; set; }
    }

    /// <summary>
    /// 边框类
    /// </summary>
    public class Border
    {
        /// <summary>
        /// 左边框
        /// </summary>
        public string? Left { get; set; }
        /// <summary>
        /// 右边框
        /// </summary>
        public string? Right { get; set; }
        /// <summary>
        /// 上边框
        /// </summary>
        public string? Top { get; set; }
        /// <summary>
        /// 上边框颜色
        /// </summary>
        public string? TopColor { get; set; }
        /// <summary>
        /// 下边框
        /// </summary>
        public string? Bottom { get; set; }
        /// <summary>
        /// 下边框颜色
        /// </summary>
        public string? BottomColor { get; set; }
        /// <summary>
        /// 对角线
        /// </summary>
        public string? Diagonal { get; set; }
        /// <summary>
        /// 对角线颜色
        /// </summary>
        public string? DiagonalColor { get; set; }
        /// <summary>
        /// 左边框颜色
        /// </summary>
        public string? LeftColor { get; set; }
        /// <summary>
        /// 右边框颜色
        /// </summary>
        public string? RightColor { get; set; }
    }

    /// <summary>
    /// 对齐方式类
    /// </summary>
    public class Alignment
    {
        /// <summary>
        /// 水平对齐
        /// </summary>
        public string? Horizontal { get; set; }
        /// <summary>
        /// 垂直对齐
        /// </summary>
        public string? Vertical { get; set; }
        /// <summary>
        /// 缩进
        /// </summary>
        public int? Indent { get; set; }
        /// <summary>
        /// 是否自动换行
        /// </summary>
        public bool? WrapText { get; set; }
        /// <summary>
        /// 文本旋转角度
        /// </summary>
        public int? Rotation { get; set; }
    }
    
    /// <summary>
    /// 扩展格式类
    /// </summary>
    public class Xf
    {
        /// <summary>
        /// 字体索引
        /// </summary>
        public ushort FontIndex { get; set; }
        /// <summary>
        /// 数字格式索引
        /// </summary>
        public ushort NumberFormatIndex { get; set; }
        /// <summary>
        /// 单元格格式索引
        /// </summary>
        public ushort CellFormatIndex { get; set; }
        /// <summary>
        /// 是否锁定
        /// </summary>
        public bool IsLocked { get; set; }
        /// <summary>
        /// 是否隐藏
        /// </summary>
        public bool IsHidden { get; set; }
        /// <summary>
        /// 水平对齐方式
        /// </summary>
        public string? HorizontalAlignment { get; set; }
        /// <summary>
        /// 垂直对齐方式
        /// </summary>
        public string? VerticalAlignment { get; set; }
        /// <summary>
        /// 缩进
        /// </summary>
        public byte Indent { get; set; }
        /// <summary>
        /// 是否自动换行
        /// </summary>
        public bool WrapText { get; set; }
        /// <summary>
        /// 是否应用对齐方式
        /// </summary>
        public bool ApplyAlignment { get; set; }
        /// <summary>
        /// 边框索引
        /// </summary>
        public int BorderIndex { get; set; }
        /// <summary>
        /// 填充索引
        /// </summary>
        public int FillIndex { get; set; }
    }

    /// <summary>
    /// 外部工作簿引用
    /// </summary>
    public class ExternalBook
    {
        public string? FileName { get; set; }
        public List<string> SheetNames { get; set; } = new List<string>();
        public List<string> ExternalNames { get; set; } = new List<string>();
        public bool IsSelf { get; set; } // 指向当前工作簿
        public bool IsAddIn { get; set; }
    }

    /// <summary>
    /// 外部工作表引用（映射索引）
    /// </summary>
    public class ExternalSheet
    {
        public int ExternalBookIndex { get; set; }
        public int FirstSheetIndex { get; set; }
        public int LastSheetIndex { get; set; }
    }

    /// <summary>
    /// 命名区域类
    /// </summary>
    public class DefinedName
    {
        public string? Name { get; set; }
        public string? Formula { get; set; }
        public int? LocalSheetId { get; set; } // null for global
        public bool Hidden { get; set; }
    }

    /// <summary>
    /// 页面设置类
    /// </summary>
    public class PageSettings
    {
        public string? Header { get; set; }
        public string? Footer { get; set; }
        public double LeftMargin { get; set; } = 0.7;
        public double RightMargin { get; set; } = 0.7;
        public double TopMargin { get; set; } = 0.75;
        public double BottomMargin { get; set; } = 0.75;
        public double HeaderMargin { get; set; } = 0.3;
        public double FooterMargin { get; set; } = 0.3;
        public bool HorizontalCenter { get; set; }
        public bool VerticalCenter { get; set; }
        public ushort PaperSize { get; set; } = 9; // A4
        public ushort Scale { get; set; } = 100;
        public ushort FitToWidth { get; set; }
        public ushort FitToHeight { get; set; }
        public bool OrientationLandscape { get; set; }
        public bool UsePageNumbers { get; set; }
    }
}