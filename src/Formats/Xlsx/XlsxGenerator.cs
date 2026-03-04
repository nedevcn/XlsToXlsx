using System.IO;
using System.IO.Compression;
using System.Xml;
using Nedev.XlsToXlsx;
using Nedev.XlsToXlsx.Exceptions;

namespace Nedev.XlsToXlsx.Formats.Xlsx
{
    public class XlsxGenerator
    {
        private Stream _stream;
        private const long MAX_OUTPUT_SIZE = 200 * 1024 * 1024; // 200MB输出大小限制
        
        /// <summary>
        /// VBA项目大小限制（字节）
        /// </summary>
        public long VbaSizeLimit { get; set; } = 50 * 1024 * 1024;

        public XlsxGenerator(Stream stream)
        {
            // 验证流是否可写
            if (!stream.CanWrite)
            {
                throw new XlsToXlsxException("Stream must be writable", 2000, "StreamError");
            }
            _stream = stream;
        }

        /// <summary>
        /// 清理XML字符串，移除无效字符
        /// </summary>
        /// <param name="input">输入字符串</param>
        /// <returns>清理后的字符串</returns>
        private string CleanXmlString(string input)
        {
            if (string.IsNullOrEmpty(input))
                return string.Empty;
            
            // 移除XML无效字符
            return new string(input.Where(c => 
                (c >= 0x0020 && c <= 0xD7FF) ||
                (c >= 0xE000 && c <= 0xFFFD) ||
                c == 0x0009 ||
                c == 0x000A ||
                c == 0x000D
            ).ToArray());
        }

        public void Generate(Workbook workbook)
        {
            // 验证workbook对象
            if (workbook == null)
            {
                throw new XlsToXlsxException("Workbook cannot be null", 2001, "NullError");
            }

            // 验证工作表数量
            if (workbook.Worksheets == null)
            {
                throw new XlsToXlsxException("Workbook.Worksheets cannot be null", 2003, "NullError");
            }

            try
            {
                Logger.Info("开始生成XLSX文件");
                
                // 使用LeaveOpen参数确保ZipArchive不关闭底层流
                using (var archive = new ZipArchive(_stream, ZipArchiveMode.Create, true))
                {
                    // 创建[Content_Types].xml
                    CreateContentTypesXml(archive, workbook);
                    Logger.Info("创建[Content_Types].xml完成");

                    // 创建_rels/.rels
                    CreateRelsXml(archive);
                    Logger.Info("创建_rels/.rels完成");

                    // 创建xl/workbook.xml
                    CreateWorkbookXml(archive, workbook);
                    Logger.Info("创建xl/workbook.xml完成");

                    // 创建xl/_rels/workbook.xml.rels
                    CreateWorkbookRelsXml(archive, workbook);
                    Logger.Info("创建xl/_rels/workbook.xml.rels完成");

                    // 创建xl/styles.xml
                    CreateStylesXml(archive, workbook);
                    Logger.Info("创建xl/styles.xml完成");

                    // 创建xl/sharedStrings.xml
                    CreateSharedStringsXml(archive, workbook);
                    Logger.Info("创建xl/sharedStrings.xml完成");

                    // 并行创建工作表
                    int actualWorksheetCount = workbook.Worksheets.Count;
                    if (actualWorksheetCount == 0)
                    {
                        Logger.Info("开始创建1个默认工作表");
                        // 创建一个默认的工作表
                        var defaultWorksheet = new Worksheet { Name = "Sheet1" };
                        CreateWorksheetXml(archive, defaultWorksheet, 1, workbook);
                        CreateWorksheetRelsXml(archive, defaultWorksheet, 1);
                        if (defaultWorksheet.Comments.Count > 0)
                        {
                            CreateCommentsXml(archive, defaultWorksheet, 1);
                        }
                        Logger.Info("默认工作表创建完成");
                    }
                    else
                    {
                        Logger.Info($"开始创建{actualWorksheetCount}个工作表");
                        // 串行创建工作表，避免ZipArchive的线程安全问题
                        for (int i = 0; i < actualWorksheetCount; i++)
                        {
                            try
                            {
                                CreateWorksheetXml(archive, workbook.Worksheets[i], i + 1, workbook);
                                CreateWorksheetRelsXml(archive, workbook.Worksheets[i], i + 1);
                                if (workbook.Worksheets[i].Comments.Count > 0)
                                {
                                    CreateCommentsXml(archive, workbook.Worksheets[i], i + 1);
                                }
                                Logger.Debug($"工作表{i + 1}创建完成");
                            }
                            catch (XlsToXlsxException)
                            {
                                throw;
                            }
                            catch (Exception ex)
                            {
                                Logger.Error($"创建工作表{i + 1}时发生错误", ex);
                                throw new XlsxGenerateException($"Error creating worksheet {i + 1}: {ex.Message}", ex);
                            }
                        }
                        Logger.Info("所有工作表创建完成");
                    }
                    
                    // 创建图片和嵌入对象
                    CreateDrawings(archive, workbook);
                    Logger.Info("创建图片和嵌入对象完成");
                    
                    // 创建VBA项目文件
                    if (workbook.VbaProject != null)
                    {
                        // 验证VBA项目大小
                        if (workbook.VbaProject.Length > VbaSizeLimit)
                        {
                            throw new XlsToXlsxException($"VBA project size exceeds limit of {VbaSizeLimit / (1024 * 1024)}MB", 2004, "FileSizeError");
                        }
                        CreateVbaProjectBin(archive, workbook);
                        Logger.Info("创建VBA项目文件完成");
                    }
                }
                
                // 检查输出大小
                if (_stream.CanSeek)
                {
                    long outputSize = _stream.Length;
                    if (outputSize > MAX_OUTPUT_SIZE)
                    {
                        throw new XlsToXlsxException($"Output file size exceeds limit of {MAX_OUTPUT_SIZE / (1024 * 1024)}MB", 2002, "FileSizeError");
                    }
                }
                
                Logger.Info("XLSX文件生成成功");
            }
            catch (XlsToXlsxException)
            {
                throw;
            }
            catch (Exception ex)
            {
                Logger.Error("生成XLSX文件时发生错误", ex);
                throw new XlsxGenerateException($"Error generating XLSX file: {ex.Message}", ex);
            }
        }

        public async Task GenerateAsync(Workbook workbook)
        {
            // 验证workbook对象
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook), "Workbook cannot be null");
            }

            // 验证工作表数量
            if (workbook.Worksheets == null)
            {
                throw new InvalidDataException("Workbook.Worksheets cannot be null");
            }

            // 使用Task.Run在后台线程中执行生成，避免阻塞主线程
            await Task.Run(() => Generate(workbook));
        }

        private void CreateContentTypesXml(ZipArchive archive, Workbook workbook)
        {
            var entry = archive.CreateEntry("[Content_Types].xml");
            using (var stream = entry.Open())
            using (var writer = XmlWriter.Create(stream, new XmlWriterSettings { Indent = false, OmitXmlDeclaration = false }))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("Types", "http://schemas.openxmlformats.org/package/2006/content-types");
                
                // 默认内容类型
                writer.WriteStartElement("Default");
                writer.WriteAttributeString("Extension", "rels");
                writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
                writer.WriteEndElement();
                
                writer.WriteStartElement("Default");
                writer.WriteAttributeString("Extension", "xml");
                writer.WriteAttributeString("ContentType", "application/xml");
                writer.WriteEndElement();
                
                // 图片扩展名默认类型
                var imageExtensions = new HashSet<string>();
                foreach (var ws in workbook.Worksheets)
                {
                    foreach (var pic in ws.Pictures)
                    {
                        if (!string.IsNullOrEmpty(pic.Extension))
                            imageExtensions.Add(pic.Extension.ToLower());
                    }
                }
                foreach (var ext in imageExtensions)
                {
                    string mimeType = ext switch
                    {
                        "png" => "image/png",
                        "jpg" or "jpeg" => "image/jpeg",
                        "gif" => "image/gif",
                        "bmp" => "image/bmp",
                        "tiff" or "tif" => "image/tiff",
                        _ => "image/png"
                    };
                    writer.WriteStartElement("Default");
                    writer.WriteAttributeString("Extension", ext);
                    writer.WriteAttributeString("ContentType", mimeType);
                    writer.WriteEndElement();
                }
                
                // VML 扩展名（用于注释的 legacy drawing）
                bool hasComments = workbook.Worksheets.Any(ws => ws.Comments != null && ws.Comments.Count > 0);
                if (hasComments)
                {
                    writer.WriteStartElement("Default");
                    writer.WriteAttributeString("Extension", "vml");
                    writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.vmlDrawing");
                    writer.WriteEndElement();
                }
                
                // 工作簿内容类型（VBA 时使用 macro-enabled 类型）
                writer.WriteStartElement("Override");
                writer.WriteAttributeString("PartName", "/xl/workbook.xml");
                if (workbook.VbaProject != null)
                {
                    writer.WriteAttributeString("ContentType", "application/vnd.ms-excel.sheet.macroEnabled.main+xml");
                }
                else
                {
                    writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
                }
                writer.WriteEndElement();
                
                writer.WriteStartElement("Override");
                writer.WriteAttributeString("PartName", "/xl/styles.xml");
                writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml");
                writer.WriteEndElement();
                
                writer.WriteStartElement("Override");
                writer.WriteAttributeString("PartName", "/xl/sharedStrings.xml");
                writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml");
                writer.WriteEndElement();
                
                // 为每个工作表添加内容类型定义
                int actualSheetCount = workbook.Worksheets.Count > 0 ? workbook.Worksheets.Count : 1;
                for (int i = 0; i < actualSheetCount; i++)
                {
                    writer.WriteStartElement("Override");
                    writer.WriteAttributeString("PartName", $"/xl/worksheets/sheet{i + 1}.xml");
                    writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml");
                    writer.WriteEndElement();
                }
                
                // 图表内容类型
                int chartIndex = 1;
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    var ws = workbook.Worksheets[i];
                    if (ws.Charts.Count > 0)
                    {
                        // Drawing
                        writer.WriteStartElement("Override");
                        writer.WriteAttributeString("PartName", $"/xl/drawings/drawing{i + 1}.xml");
                        writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml");
                        writer.WriteEndElement();
                        
                        // Charts
                        for (int j = 0; j < ws.Charts.Count; j++)
                        {
                            writer.WriteStartElement("Override");
                            writer.WriteAttributeString("PartName", $"/xl/charts/chart{chartIndex}.xml");
                            writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
                            writer.WriteEndElement();
                            chartIndex++;
                        }
                    }
                    
                    // Comments
                    if (ws.Comments != null && ws.Comments.Count > 0)
                    {
                        writer.WriteStartElement("Override");
                        writer.WriteAttributeString("PartName", $"/xl/comments{i + 1}.xml");
                        writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml");
                        writer.WriteEndElement();
                    }
                }
                
                // VBA项目内容类型
                if (workbook.VbaProject != null)
                {
                    writer.WriteStartElement("Override");
                    writer.WriteAttributeString("PartName", "/xl/vbaProject.bin");
                    writer.WriteAttributeString("ContentType", "application/vnd.ms-office.vbaProject");
                    writer.WriteEndElement();
                }
                
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        private void CreateRelsXml(ZipArchive archive)
        {
            var entry = archive.CreateEntry("_rels/.rels");
            using (var stream = entry.Open())
            using (var writer = XmlWriter.Create(stream, new XmlWriterSettings { Indent = false, OmitXmlDeclaration = false }))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
                
                writer.WriteStartElement("Relationship");
                writer.WriteAttributeString("Id", "rId1");
                writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
                writer.WriteAttributeString("Target", "xl/workbook.xml");
                writer.WriteEndElement();
                
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        private void CreateWorkbookXml(ZipArchive archive, Workbook workbook)
        {
            var entry = archive.CreateEntry("xl/workbook.xml");
            using (var stream = entry.Open())
            using (var writer = XmlWriter.Create(stream, new XmlWriterSettings { Indent = false, OmitXmlDeclaration = false }))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("workbook", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                
                writer.WriteStartElement("sheets");
                if (workbook.Worksheets.Count == 0)
                {
                    // 如果没有工作表，添加一个默认的工作表
                    writer.WriteStartElement("sheet");
                    writer.WriteAttributeString("name", "Sheet1");
                    writer.WriteAttributeString("sheetId", "1");
                    writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "rId1");
                    writer.WriteEndElement();
                }
                else
                {
                    for (int i = 0; i < workbook.Worksheets.Count; i++)
                    {
                        writer.WriteStartElement("sheet");
                        writer.WriteAttributeString("name", workbook.Worksheets[i].Name);
                        writer.WriteAttributeString("sheetId", (i + 1).ToString());
                        writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "rId" + (i + 1));
                        writer.WriteEndElement();
                    }
                }
                writer.WriteEndElement(); // sheets

                if (workbook.DefinedNames != null && workbook.DefinedNames.Count > 0)
                {
                    writer.WriteStartElement("definedNames");
                    foreach (var dn in workbook.DefinedNames)
                    {
                        writer.WriteStartElement("definedName");
                        writer.WriteAttributeString("name", dn.Name);
                        if (dn.LocalSheetId.HasValue)
                        {
                            writer.WriteAttributeString("localSheetId", dn.LocalSheetId.Value.ToString());
                        }
                        if (dn.Hidden)
                        {
                            writer.WriteAttributeString("hidden", "1");
                        }
                        writer.WriteString(dn.Formula ?? "");
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement(); // definedNames
                }

                writer.WriteEndElement(); // workbook
                writer.WriteEndDocument();
            }
        }

        private void CreateWorkbookRelsXml(ZipArchive archive, Workbook workbook)
        {
            var entry = archive.CreateEntry("xl/_rels/workbook.xml.rels");
            using (var stream = entry.Open())
            using (var writer = XmlWriter.Create(stream, new XmlWriterSettings { Indent = false, OmitXmlDeclaration = false }))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
                
                int worksheetCount = workbook.Worksheets.Count;
                if (worksheetCount == 0)
                {
                    // 如果没有工作表，添加一个默认工作表的关系
                    writer.WriteStartElement("Relationship");
                    writer.WriteAttributeString("Id", "rId1");
                    writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
                    writer.WriteAttributeString("Target", "worksheets/sheet1.xml");
                    writer.WriteEndElement();
                    worksheetCount = 1;
                }
                else
                {
                    for (int i = 0; i < workbook.Worksheets.Count; i++)
                    {
                        writer.WriteStartElement("Relationship");
                        writer.WriteAttributeString("Id", "rId" + (i + 1));
                        writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet");
                        writer.WriteAttributeString("Target", "worksheets/sheet" + (i + 1) + ".xml");
                        writer.WriteEndElement();
                    }
                }
                
                writer.WriteStartElement("Relationship");
                writer.WriteAttributeString("Id", "rId" + (worksheetCount + 1));
                writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
                writer.WriteAttributeString("Target", "styles.xml");
                writer.WriteEndElement();
                
                writer.WriteStartElement("Relationship");
                writer.WriteAttributeString("Id", "rId" + (worksheetCount + 2));
                writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings");
                writer.WriteAttributeString("Target", "sharedStrings.xml");
                writer.WriteEndElement();
                
                // VBA项目关系
                if (workbook.VbaProject != null)
                {
                    writer.WriteStartElement("Relationship");
                    writer.WriteAttributeString("Id", "rId" + (worksheetCount + 3));
                    writer.WriteAttributeString("Type", "http://schemas.microsoft.com/office/2006/relationships/vbaProject");
                    writer.WriteAttributeString("Target", "vbaProject.bin");
                    writer.WriteEndElement();
                }
                
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        private void CreateWorksheetXml(ZipArchive archive, Worksheet worksheet, int sheetIndex, Workbook workbook)
        {
            var entry = archive.CreateEntry($"xl/worksheets/sheet{sheetIndex}.xml");
            using (var stream = entry.Open())
            using (var writer = XmlWriter.Create(stream, new XmlWriterSettings { Indent = false, OmitXmlDeclaration = false }))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("worksheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                
                // dimension（工作表范围）
                writer.WriteStartElement("dimension");
                string maxRef = GetCellReference(Math.Max(1, worksheet.MaxRow), Math.Max(1, worksheet.MaxColumn));
                writer.WriteAttributeString("ref", $"A1:{maxRef}");
                writer.WriteEndElement(); // dimension
                
                // sheetViews（包含冻结窗格）
                writer.WriteStartElement("sheetViews");
                writer.WriteStartElement("sheetView");
                writer.WriteAttributeString("workbookViewId", "0");
                if (sheetIndex == 1)
                {
                    writer.WriteAttributeString("tabSelected", "1");
                }
                
                if (worksheet.FreezePane != null && (worksheet.FreezePane.RowSplit > 0 || worksheet.FreezePane.ColSplit > 0))
                {
                    var fp = worksheet.FreezePane;
                    writer.WriteStartElement("pane");
                    if (fp.ColSplit > 0)
                    {
                        writer.WriteAttributeString("xSplit", fp.ColSplit.ToString());
                    }
                    if (fp.RowSplit > 0)
                    {
                        writer.WriteAttributeString("ySplit", fp.RowSplit.ToString());
                    }
                    string topLeftCell = GetCellReference(fp.TopRow, fp.LeftCol);
                    writer.WriteAttributeString("topLeftCell", topLeftCell);
                    writer.WriteAttributeString("state", "frozen");
                    // 活动窗格
                    string activePane = (fp.RowSplit > 0 && fp.ColSplit > 0) ? "bottomRight" :
                                       (fp.RowSplit > 0) ? "bottomLeft" : "topRight";
                    writer.WriteAttributeString("activePane", activePane);
                    writer.WriteEndElement(); // pane
                    
                    writer.WriteStartElement("selection");
                    writer.WriteAttributeString("activeCell", topLeftCell);
                    writer.WriteAttributeString("sqref", topLeftCell);
                    writer.WriteEndElement(); // selection
                }
                
                writer.WriteEndElement(); // sheetView
                writer.WriteEndElement(); // sheetViews
                
                // sheetFormatPr（默认行高列宽）
                writer.WriteStartElement("sheetFormatPr");
                if (worksheet.DefaultRowHeight.HasValue)
                {
                    writer.WriteAttributeString("defaultRowHeight", worksheet.DefaultRowHeight.Value.ToString("F2", System.Globalization.CultureInfo.InvariantCulture));
                    writer.WriteAttributeString("customHeight", "1");
                }
                else
                {
                    writer.WriteAttributeString("defaultRowHeight", "15");
                }
                if (worksheet.DefaultColumnWidth.HasValue)
                {
                    writer.WriteAttributeString("baseColWidth", worksheet.DefaultColumnWidth.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
                }
                writer.WriteEndElement(); // sheetFormatPr
                
                // cols（列宽信息）
                if (worksheet.ColumnInfos != null && worksheet.ColumnInfos.Count > 0)
                {
                    writer.WriteStartElement("cols");
                    foreach (var colInfo in worksheet.ColumnInfos)
                    {
                        writer.WriteStartElement("col");
                        writer.WriteAttributeString("min", (colInfo.FirstColumn + 1).ToString()); // 1-based
                        writer.WriteAttributeString("max", (colInfo.LastColumn + 1).ToString());
                        // 将 1/256 字符宽u5355位转为 XLSX 的字符宽度（除以 256）
                        double widthInChars = colInfo.Width / 256.0;
                        writer.WriteAttributeString("width", widthInChars.ToString("F2", System.Globalization.CultureInfo.InvariantCulture));
                        if (colInfo.Hidden)
                        {
                            writer.WriteAttributeString("hidden", "1");
                        }
                        writer.WriteAttributeString("customWidth", "1");
                        writer.WriteEndElement(); // col
                    }
                    writer.WriteEndElement(); // cols
                }
                
                writer.WriteStartElement("sheetData");
                foreach (var row in worksheet.Rows)
                {
                    writer.WriteStartElement("row");
                    writer.WriteAttributeString("r", row.RowIndex.ToString());
                    
                    // 自定义行高
                    if (row.CustomHeight && row.Height > 0)
                    {
                        // 将 twips (1/20点) 转为点数
                        double heightInPoints = row.Height / 20.0;
                        writer.WriteAttributeString("ht", heightInPoints.ToString("F2", System.Globalization.CultureInfo.InvariantCulture));
                        writer.WriteAttributeString("customHeight", "1");
                    }
                    
                    // 用于跟踪已处理的列索引，防止重复
                    var processedColumns = new HashSet<int>();
                    
                    foreach (var cell in row.Cells)
                    {
                        // 确保列索引有效且未处理过
                        if (cell.ColumnIndex > 0 && cell.ColumnIndex <= 16384 && !processedColumns.Contains(cell.ColumnIndex))
                        {
                            writer.WriteStartElement("c");
                            writer.WriteAttributeString("r", GetCellReference(row.RowIndex, cell.ColumnIndex));
                            
                            if (!string.IsNullOrEmpty(cell.DataType))
                            {
                                writer.WriteAttributeString("t", cell.DataType);
                            }
                            
                            // 样式ID
                            if (!string.IsNullOrEmpty(cell.StyleId))
                            {
                                writer.WriteAttributeString("s", cell.StyleId);
                            }
                            else if (cell.Value is DateTime)
                            {
                                writer.WriteAttributeString("s", "1");
                            }
                            
                            // 处理公式
                            if (cell.DataType == "f")
                            {
                                writer.WriteStartElement("f");
                                writer.WriteString(cell.Formula ?? "");
                                writer.WriteEndElement();
                                writer.WriteStartElement("v");
                                // 写入公式结果
                                if (cell.Value != null)
                                {
                                    writer.WriteString(cell.Value.ToString() ?? "");
                                }
                                writer.WriteEndElement();
                            }
                            else if (cell.RichText != null && cell.RichText.Count > 0)
                            {
                                // 富文本单元格
                                writer.WriteAttributeString("t", "inlineStr");
                                writer.WriteStartElement("is");
                                
                                foreach (var run in cell.RichText)
                                {
                                    writer.WriteStartElement("r");
                                    if (run.Font != null)
                                    {
                                        writer.WriteStartElement("rPr");
                                        if (!string.IsNullOrEmpty(run.Font.Name))
                                        {
                                            writer.WriteStartElement("rFont");
                                            writer.WriteAttributeString("val", run.Font.Name);
                                            writer.WriteEndElement();
                                        }
                                        if (run.Font.Size > 0)
                                        {
                                            writer.WriteStartElement("sz");
                                            writer.WriteAttributeString("val", run.Font.Size.ToString());
                                            writer.WriteEndElement();
                                        }
                                        if (run.Font.Bold.HasValue && run.Font.Bold.Value)
                                        {
                                            writer.WriteStartElement("b");
                                            writer.WriteEndElement();
                                        }
                                        if (run.Font.Italic.HasValue && run.Font.Italic.Value)
                                        {
                                            writer.WriteStartElement("i");
                                            writer.WriteEndElement();
                                        }
                                        if (run.Font.Underline.HasValue && run.Font.Underline.Value)
                                        {
                                            writer.WriteStartElement("u");
                                            writer.WriteEndElement();
                                        }
                                        if (!string.IsNullOrEmpty(run.Font.Color))
                                        {
                                            writer.WriteStartElement("color");
                                            writer.WriteAttributeString("rgb", run.Font.Color);
                                            writer.WriteEndElement();
                                        }
                                        writer.WriteEndElement();
                                    }
                                    writer.WriteStartElement("t");
                                    writer.WriteAttributeString("xml:space", "preserve");
                                    writer.WriteString(run.Text ?? "");
                                    writer.WriteEndElement();
                                    writer.WriteEndElement();
                                }
                                
                                writer.WriteEndElement();
                            }
                            else
                            {
                                // 处理日期时间类型
                                if (cell.Value is DateTime dateTime)
                                {
                                    // Excel 日期时间是从 1900-01-01 开始的天数
                                    double excelDate = DateTimeToExcelDate(dateTime);
                                    writer.WriteStartElement("v");
                                    writer.WriteString(excelDate.ToString());
                                    writer.WriteEndElement();
                                }
                                else if (cell.DataType == "s")
                                {
                                    // 共享字符串类型，写入索引
                                    var textValue = cell.Value?.ToString();
                                    writer.WriteStartElement("v");
                                    if (!string.IsNullOrEmpty(textValue))
                                    {
                                        // 查找字符串在共享字符串表中的索引
                                        int index = workbook.SharedStrings.IndexOf(textValue);
                                        writer.WriteString(index >= 0 ? index.ToString() : "0");
                                    }
                                    else
                                    {
                                        writer.WriteString("0");
                                    }
                                    writer.WriteEndElement();
                                }
                                else if (cell.Value != null)
                                {
                                    string valStr = CleanXmlString(cell.Value.ToString() ?? "");
                                    if (!string.IsNullOrEmpty(valStr))
                                    {
                                        writer.WriteStartElement("v");
                                        writer.WriteString(valStr);
                                        writer.WriteEndElement();
                                    }
                                }
                                // If cell has no value and no type, skip <v> entirely
                            }
                            
                            writer.WriteEndElement();
                            
                            // 标记该列已处理
                            processedColumns.Add(cell.ColumnIndex);
                        }
                    }
                    
                    writer.WriteEndElement();
                }
                writer.WriteEndElement();
                
                // 写入合并单元格信息
                if (worksheet.MergeCells != null && worksheet.MergeCells.Count > 0)
                {
                    writer.WriteStartElement("mergeCells");
                    writer.WriteAttributeString("count", worksheet.MergeCells.Count.ToString());
                    
                    foreach (var mergeCell in worksheet.MergeCells)
                    {
                        writer.WriteStartElement("mergeCell");
                        string refValue = GetCellReference(mergeCell.StartRow, mergeCell.StartColumn) + ":" + 
                                         GetCellReference(mergeCell.EndRow, mergeCell.EndColumn);
                        writer.WriteAttributeString("ref", refValue);
                        writer.WriteEndElement();
                    }
                    
                    writer.WriteEndElement();
                }
                
                // 写入数据验证信息
                if (worksheet.DataValidations != null && worksheet.DataValidations.Count > 0)
                {
                    writer.WriteStartElement("dataValidations");
                    writer.WriteAttributeString("count", worksheet.DataValidations.Count.ToString());
                    
                    foreach (var dv in worksheet.DataValidations)
                    {
                        writer.WriteStartElement("dataValidation");
                        if (!string.IsNullOrEmpty(dv.Range))
                        {
                            writer.WriteAttributeString("sqref", dv.Range);
                        }
                        if (!string.IsNullOrEmpty(dv.Type))
                        {
                            writer.WriteAttributeString("type", dv.Type);
                        }
                        if (!string.IsNullOrEmpty(dv.Operator))
                        {
                            writer.WriteAttributeString("operator", dv.Operator);
                        }
                        writer.WriteAttributeString("allowBlank", dv.AllowBlank.ToString().ToLower());
                        // formula1/formula2 必须是子元素，不是属性 (OOXML 规范)
                        if (!string.IsNullOrEmpty(dv.Formula1))
                        {
                            writer.WriteElementString("formula1", dv.Formula1);
                        }
                        if (!string.IsNullOrEmpty(dv.Formula2))
                        {
                            writer.WriteElementString("formula2", dv.Formula2);
                        }
                        writer.WriteEndElement();
                    }
                    
                    writer.WriteEndElement();
                }
                
                // 写入条件格式信息
                if (worksheet.ConditionalFormats != null && worksheet.ConditionalFormats.Count > 0)
                {
                    int cfPriority = 1; // 每个 cfRule 必须有唯一的 priority
                    foreach (var cf in worksheet.ConditionalFormats)
                    {
                        if (!string.IsNullOrEmpty(cf.Range))
                        {
                            writer.WriteStartElement("conditionalFormatting");
                            writer.WriteAttributeString("sqref", cf.Range);
                            
                            writer.WriteStartElement("cfRule");
                            writer.WriteAttributeString("priority", cfPriority.ToString());
                            cfPriority++;
                            if (!string.IsNullOrEmpty(cf.Type))
                            {
                                writer.WriteAttributeString("type", cf.Type);
                            }
                            if (!string.IsNullOrEmpty(cf.Operator))
                            {
                                writer.WriteAttributeString("operator", cf.Operator);
                            }
                            if (!string.IsNullOrEmpty(cf.Formula))
                            {
                                writer.WriteElementString("formula", cf.Formula);
                            }
                            
                            // 为不同类型的条件格式添加相应的元素
                            switch (cf.Type)
                            {
                                case "colorScale":
                                    // 添加颜色刻度元素
                                    writer.WriteStartElement("colorScale");
                                    writer.WriteStartElement("cfvo");
                                    writer.WriteAttributeString("type", "min");
                                    writer.WriteEndElement();
                                    writer.WriteStartElement("cfvo");
                                    writer.WriteAttributeString("type", "max");
                                    writer.WriteEndElement();
                                    writer.WriteStartElement("color");
                                    writer.WriteAttributeString("rgb", "FF00FF00"); // 绿色
                                    writer.WriteEndElement();
                                    writer.WriteStartElement("color");
                                    writer.WriteAttributeString("rgb", "FFFFFF00"); // 黄色
                                    writer.WriteEndElement();
                                    writer.WriteStartElement("color");
                                    writer.WriteAttributeString("rgb", "FFFF0000"); // 红色
                                    writer.WriteEndElement();
                                    writer.WriteEndElement();
                                    break;
                                case "dataBar":
                                    // 添加数据条元素
                                    writer.WriteStartElement("dataBar");
                                    writer.WriteAttributeString("minLength", "0");
                                    writer.WriteAttributeString("maxLength", "100");
                                    writer.WriteStartElement("cfvo");
                                    writer.WriteAttributeString("type", "min");
                                    writer.WriteEndElement();
                                    writer.WriteStartElement("cfvo");
                                    writer.WriteAttributeString("type", "max");
                                    writer.WriteEndElement();
                                    writer.WriteStartElement("color");
                                    writer.WriteAttributeString("rgb", "FF638EC6"); // 蓝色
                                    writer.WriteEndElement();
                                    writer.WriteEndElement();
                                    break;
                                case "iconSet":
                                    // 添加图标集元素
                                    writer.WriteStartElement("iconSet");
                                    writer.WriteAttributeString("iconSet", "3TrafficLights1");
                                    writer.WriteEndElement();
                                    break;
                            }
                            
                            writer.WriteEndElement();
                            
                            writer.WriteEndElement();
                        }
                    }
                }
                
                // P3: pageMargins (页面边距) - 必须在 sheetData 之后
                var ps = worksheet.PageSettings;
                writer.WriteStartElement("pageMargins");
                writer.WriteAttributeString("left", ps.LeftMargin.ToString(System.Globalization.CultureInfo.InvariantCulture));
                writer.WriteAttributeString("right", ps.RightMargin.ToString(System.Globalization.CultureInfo.InvariantCulture));
                writer.WriteAttributeString("top", ps.TopMargin.ToString(System.Globalization.CultureInfo.InvariantCulture));
                writer.WriteAttributeString("bottom", ps.BottomMargin.ToString(System.Globalization.CultureInfo.InvariantCulture));
                writer.WriteAttributeString("header", ps.HeaderMargin.ToString(System.Globalization.CultureInfo.InvariantCulture));
                writer.WriteAttributeString("footer", ps.FooterMargin.ToString(System.Globalization.CultureInfo.InvariantCulture));
                writer.WriteEndElement(); // pageMargins

                // P3: pageSetup (页面设置)
                writer.WriteStartElement("pageSetup");
                writer.WriteAttributeString("paperSize", ps.PaperSize.ToString());
                writer.WriteAttributeString("scale", ps.Scale.ToString());
                if (ps.FitToWidth > 0) writer.WriteAttributeString("fitToWidth", ps.FitToWidth.ToString());
                if (ps.FitToHeight > 0) writer.WriteAttributeString("fitToHeight", ps.FitToHeight.ToString());
                writer.WriteAttributeString("orientation", ps.OrientationLandscape ? "landscape" : "portrait");
                writer.WriteAttributeString("useFirstPageNumber", ps.UsePageNumbers ? "1" : "0");
                writer.WriteEndElement(); // pageSetup

                // P3: headerFooter (页眉页脚)
                if (!string.IsNullOrEmpty(ps.Header) || !string.IsNullOrEmpty(ps.Footer))
                {
                    writer.WriteStartElement("headerFooter");
                    if (!string.IsNullOrEmpty(ps.Header))
                    {
                        writer.WriteElementString("oddHeader", ps.Header);
                    }
                    if (!string.IsNullOrEmpty(ps.Footer))
                    {
                        writer.WriteElementString("oddFooter", ps.Footer);
                    }
                    writer.WriteEndElement(); // headerFooter
                }

                // 生成图表引用 (必须在 pageSetup/headerFooter 之后)
                if (worksheet.Charts.Count > 0)
                {
                    writer.WriteStartElement("drawing");
                    writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "rIdDrawing" + sheetIndex);
                    writer.WriteEndElement();
                }

                // 生成注释引用 (legacyDrawing 必须在 drawing 之后)
                if (worksheet.Comments != null && worksheet.Comments.Count > 0)
                {
                    writer.WriteStartElement("legacyDrawing");
                    writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "rIdComments" + sheetIndex);
                    writer.WriteEndElement();
                }

                writer.WriteEndElement(); // worksheet
                writer.WriteEndDocument();
            }
        }
        
        private void CreateWorksheetRelsXml(ZipArchive archive, Worksheet worksheet, int sheetIndex)
        {
            // 只有在有注释、图表、图片或超链接时才创建关系文件
            bool hasComments = worksheet.Comments != null && worksheet.Comments.Count > 0;
            bool hasDrawings = worksheet.Pictures.Count > 0 || worksheet.EmbeddedObjects.Count > 0 || worksheet.Charts.Count > 0;
            bool hasHyperlinks = worksheet.Hyperlinks != null && worksheet.Hyperlinks.Count > 0;
            
            if (!hasComments && !hasDrawings && !hasHyperlinks)
            {
                return;
            }
            
            var entry = archive.CreateEntry($"xl/worksheets/_rels/sheet{sheetIndex}.xml.rels");
            using (var stream = entry.Open())
            using (var writer = XmlWriter.Create(stream, new XmlWriterSettings { Indent = false, OmitXmlDeclaration = false }))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
                
                // 添加注释关系
                if (hasComments)
                {
                    writer.WriteStartElement("Relationship");
                    writer.WriteAttributeString("Id", "rIdComments" + sheetIndex);
                    writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/legacyDrawing");
                    writer.WriteAttributeString("Target", "../drawings/vmlDrawing" + sheetIndex + ".vml");
                    writer.WriteEndElement();
                }
                
                // 添加图表和图片引用（只引用 drawing.xml）
                if (hasDrawings)
                {
                    writer.WriteStartElement("Relationship");
                    writer.WriteAttributeString("Id", "rIdDrawing" + sheetIndex); // worksheet 统一用 drawing1
                    writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing");
                    writer.WriteAttributeString("Target", "../drawings/drawing" + sheetIndex + ".xml");
                    writer.WriteEndElement();
                }
                
                // 添加超链接关系（超链接属于工作表级别，不是工作簿级别）
                if (hasHyperlinks)
                {
                    for (int i = 0; i < worksheet.Hyperlinks.Count; i++)
                    {
                        var hyperlink = worksheet.Hyperlinks[i];
                        writer.WriteStartElement("Relationship");
                        writer.WriteAttributeString("Id", "rIdHL" + (i + 1));
                        writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink");
                        writer.WriteAttributeString("Target", hyperlink.Target ?? "");
                        writer.WriteAttributeString("TargetMode", "External");
                        writer.WriteEndElement();
                    }
                }
                
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }
        
        private void CreateCommentsXml(ZipArchive archive, Worksheet worksheet, int sheetIndex)
        {
            // 创建VML绘图文件
            var vmlEntry = archive.CreateEntry($"xl/drawings/vmlDrawing{sheetIndex}.vml");
            using (var stream = vmlEntry.Open())
            using (var writer = XmlWriter.Create(stream, new XmlWriterSettings { Indent = false, OmitXmlDeclaration = false }))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("xml", "urn:schemas-microsoft-com:vml");
                writer.WriteAttributeString("xmlns", "o", null, "urn:schemas-microsoft-com:office:office");
                writer.WriteAttributeString("xmlns", "x", null, "urn:schemas-microsoft-com:office:excel");
                writer.WriteAttributeString("xmlns", "v", null, "urn:schemas-microsoft-com:vml");
                
                writer.WriteStartElement("shapelayout", "urn:schemas-microsoft-com:office:office");
                writer.WriteStartElement("idmap", "urn:schemas-microsoft-com:office:office");
                writer.WriteAttributeString("ext", "urn:schemas-microsoft-com:vml", "edit");
                writer.WriteAttributeString("data", "1");
                writer.WriteEndElement();
                writer.WriteEndElement();
                
                writer.WriteStartElement("shapes", "urn:schemas-microsoft-com:vml");
                writer.WriteAttributeString("ext", "urn:schemas-microsoft-com:vml", "edit");
                writer.WriteAttributeString("class", "x:WorksheetComments");
                
                foreach (var comment in worksheet.Comments)
                {
                    string cellRef = GetCellReference(comment.RowIndex, comment.ColumnIndex);
                    
                    writer.WriteStartElement("shape", "urn:schemas-microsoft-com:vml");
                    writer.WriteAttributeString("id", "_x0000_s1025");
                    writer.WriteAttributeString("type", "#_x0000_t202");
                    writer.WriteAttributeString("style", "position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:150pt;height:55.5pt;z-index:1");
                    writer.WriteAttributeString("fillcolor", "#ffffe1");
                    writer.WriteAttributeString("strokecolor", "#000000");
                    writer.WriteAttributeString("insetmode", "urn:schemas-microsoft-com:office:office", "auto");
                    
                    writer.WriteStartElement("fill", "urn:schemas-microsoft-com:vml");
                    writer.WriteAttributeString("on", "true");
                    writer.WriteAttributeString("color", "#ffffe1");
                    writer.WriteEndElement();
                    
                    writer.WriteStartElement("stroke", "urn:schemas-microsoft-com:vml");
                    writer.WriteAttributeString("on", "true");
                    writer.WriteAttributeString("weight", "1pt");
                    writer.WriteEndElement();
                    
                    writer.WriteStartElement("textbox", "urn:schemas-microsoft-com:vml");
                    writer.WriteAttributeString("style", "mso-direction-alt:auto");
                    
                    writer.WriteStartElement("div");
                    writer.WriteAttributeString("style", "text-align:left");
                    
                    if (!string.IsNullOrEmpty(comment.Author))
                    {
                        writer.WriteStartElement("b");
                        writer.WriteString(comment.Author + ": ");
                        writer.WriteEndElement();
                    }
                    
                    if (!string.IsNullOrEmpty(comment.Text))
                    {
                        writer.WriteString(comment.Text);
                    }
                    
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    
                    writer.WriteStartElement("ClientData", "urn:schemas-microsoft-com:office:excel");
                    writer.WriteAttributeString("ObjectType", "Note");
                    writer.WriteElementString("Anchor", "urn:schemas-microsoft-com:office:excel", "1, 15, 0, 2, 3, 15, 2, 2");
                    writer.WriteElementString("AutoFill", "urn:schemas-microsoft-com:office:excel", "False");
                    writer.WriteElementString("Row", "urn:schemas-microsoft-com:office:excel", (comment.RowIndex - 1).ToString());
                    writer.WriteElementString("Column", "urn:schemas-microsoft-com:office:excel", comment.ColumnIndex.ToString());
                    writer.WriteEndElement();
                    
                    writer.WriteEndElement();
                }
                
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        private void CreateStylesXml(ZipArchive archive, Workbook workbook)
        {
            try
            {
                var entry = archive.CreateEntry("xl/styles.xml");
                using (var stream = entry.Open())
                using (var writer = XmlWriter.Create(stream, new XmlWriterSettings { Indent = false, OmitXmlDeclaration = false }))
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("styleSheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                    
                    // 数字格式
                    writer.WriteStartElement("numFmts");
                    // 至少需要一个数字格式（日期时间格式）
                    int numFmtCount = 1;
                    writer.WriteAttributeString("count", numFmtCount.ToString());
                    
                    // 添加日期时间格式
                    writer.WriteStartElement("numFmt");
                    writer.WriteAttributeString("numFmtId", "164");
                    writer.WriteAttributeString("formatCode", "m/d/yyyy");
                    writer.WriteEndElement();
                    
                    writer.WriteEndElement();
                    
                    // 字体
            writer.WriteStartElement("fonts");
            int fontCount = 0;
            List<Font> fonts = new List<Font>();
            
            // 优先添加从XLS文件解析的全局字体
            foreach (var font in workbook.Fonts)
            {
                if (!fonts.Any(f => 
                    f.Name == font.Name && 
                    (f.Size ?? (font.Height / 20.0)) == (font.Size ?? (font.Height / 20.0)) && 
                    (f.Bold ?? font.IsBold) == (font.Bold ?? font.IsBold) &&
                    (f.Italic ?? font.IsItalic) == (font.Italic ?? font.IsItalic) &&
                    (f.Underline ?? font.IsUnderline) == (font.Underline ?? font.IsUnderline) &&
                    f.Color == font.Color))
                {
                    fonts.Add(font);
                    fontCount++;
                }
            }

            // 如果没有字体，添加一个默认字体
            if (fonts.Count == 0)
            {
                fonts.Add(new Font { Name = "Calibri", Size = 11, Bold = false, Italic = false, Underline = false, Color = "00000000" });
                fontCount++;
            }
            
            // 添加工作簿样式中的字体
            foreach (var style in workbook.Styles)
                    {
                        if (style.Font != null && !fonts.Any(f => 
                            f.Name == style.Font.Name && 
                            f.Size == style.Font.Size && 
                            f.Bold == style.Font.Bold &&
                            f.Italic == style.Font.Italic &&
                            f.Underline == style.Font.Underline &&
                            f.Color == style.Font.Color))
                        {
                            fonts.Add(style.Font);
                            fontCount++;
                        }
                    }
                    
                    writer.WriteAttributeString("count", fontCount.ToString());
                    
                    foreach (var font in fonts)
                    {
                        writer.WriteStartElement("font");
                        
                        // OOXML CT_Font sequence: b, i, strike, u, sz, color, name
                        
                        // 粗体 (b)
                        if (font.IsBold || (font.Bold.HasValue && font.Bold.Value))
                        {
                            writer.WriteStartElement("b");
                            writer.WriteEndElement();
                        }
                        
                        // 斜体 (i)
                        if (font.IsItalic || (font.Italic.HasValue && font.Italic.Value))
                        {
                            writer.WriteStartElement("i");
                            writer.WriteEndElement();
                        }
                        
                        // 删除线 (strike)
                        if (font.IsStrikethrough)
                        {
                            writer.WriteStartElement("strike");
                            writer.WriteEndElement();
                        }
                        
                        // 下划线 (u)
                        if (font.IsUnderline || (font.Underline.HasValue && font.Underline.Value))
                        {
                            writer.WriteStartElement("u");
                            writer.WriteAttributeString("val", "single");
                            writer.WriteEndElement();
                        }
                        
                        // 字体大小 (sz)
                        if (font.Size.HasValue)
                        {
                            writer.WriteStartElement("sz");
                            writer.WriteAttributeString("val", font.Size.Value.ToString());
                            writer.WriteEndElement();
                        }
                        else if (font.Height > 0)
                        {
                            writer.WriteStartElement("sz");
                            writer.WriteAttributeString("val", (font.Height / 20.0).ToString());
                            writer.WriteEndElement();
                        }
                        else
                        {
                            writer.WriteStartElement("sz");
                            writer.WriteAttributeString("val", "11");
                            writer.WriteEndElement();
                        }
                        
                        // 字体颜色 (color)
                        if (!string.IsNullOrEmpty(font.Color))
                        {
                            writer.WriteStartElement("color");
                            writer.WriteAttributeString("rgb", font.Color);
                            writer.WriteEndElement();
                        }
                        else
                        {
                            writer.WriteStartElement("color");
                            writer.WriteAttributeString("rgb", "00000000");
                            writer.WriteEndElement();
                        }
                        
                        // 字体名称 (name)
                        if (!string.IsNullOrEmpty(font.Name))
                        {
                            writer.WriteStartElement("name");
                            writer.WriteAttributeString("val", font.Name);
                            writer.WriteEndElement();
                        }
                        else
                        {
                            writer.WriteStartElement("name");
                            writer.WriteAttributeString("val", "Calibri");
                            writer.WriteEndElement();
                        }
                        
                        writer.WriteEndElement();
                    }
                    
                    writer.WriteEndElement();
                    
                    // 填充
                    writer.WriteStartElement("fills");
                    List<Fill> fills = new List<Fill>();
                    // 添加默认填充 (必须)
                    fills.Add(new Fill { PatternType = "none" });
                    fills.Add(new Fill { PatternType = "gray125" });
                    
                    // 添加解析出的全局填充
                    fills.AddRange(workbook.Fills);
                    
                    writer.WriteAttributeString("count", fills.Count.ToString());
                    foreach (var fill in fills)
                    {
                        writer.WriteStartElement("fill");
                        writer.WriteStartElement("patternFill");
                        writer.WriteAttributeString("patternType", !string.IsNullOrEmpty(fill.PatternType) ? fill.PatternType : "none");
                        
                        if (!string.IsNullOrEmpty(fill.ForegroundColor))
                        {
                            writer.WriteStartElement("fgColor");
                            writer.WriteAttributeString("rgb", fill.ForegroundColor);
                            writer.WriteEndElement();
                        }
                        if (!string.IsNullOrEmpty(fill.BackgroundColor))
                        {
                            writer.WriteStartElement("bgColor");
                            writer.WriteAttributeString("rgb", fill.BackgroundColor);
                            writer.WriteEndElement();
                        }
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();
                    
                    // 边框
                    writer.WriteStartElement("borders");
                    List<Border> borders = new List<Border>();
                    // 添加默认边框
                    borders.Add(new Border());
                    
                    // 添加解析出的全局边框
                    borders.AddRange(workbook.Borders);
                    
                    writer.WriteAttributeString("count", borders.Count.ToString());
                    foreach (var border in borders)
                    {
                        writer.WriteStartElement("border");
                        
                        // Left
                        writer.WriteStartElement("left");
                        if (!string.IsNullOrEmpty(border.Left) && border.Left != "none")
                        {
                            writer.WriteAttributeString("style", border.Left);
                            if (!string.IsNullOrEmpty(border.LeftColor))
                            {
                                writer.WriteStartElement("color");
                                writer.WriteAttributeString("rgb", border.LeftColor);
                                writer.WriteEndElement();
                            }
                        }
                        writer.WriteEndElement();
                        
                        // Right
                        writer.WriteStartElement("right");
                        if (!string.IsNullOrEmpty(border.Right) && border.Right != "none")
                        {
                            writer.WriteAttributeString("style", border.Right);
                            if (!string.IsNullOrEmpty(border.RightColor))
                            {
                                writer.WriteStartElement("color");
                                writer.WriteAttributeString("rgb", border.RightColor);
                                writer.WriteEndElement();
                            }
                        }
                        writer.WriteEndElement();
                        
                        // Top
                        writer.WriteStartElement("top");
                        if (!string.IsNullOrEmpty(border.Top) && border.Top != "none")
                        {
                            writer.WriteAttributeString("style", border.Top);
                            if (!string.IsNullOrEmpty(border.TopColor))
                            {
                                writer.WriteStartElement("color");
                                writer.WriteAttributeString("rgb", border.TopColor);
                                writer.WriteEndElement();
                            }
                        }
                        writer.WriteEndElement();
                        
                        // Bottom
                        writer.WriteStartElement("bottom");
                        if (!string.IsNullOrEmpty(border.Bottom) && border.Bottom != "none")
                        {
                            writer.WriteAttributeString("style", border.Bottom);
                            if (!string.IsNullOrEmpty(border.BottomColor))
                            {
                                writer.WriteStartElement("color");
                                writer.WriteAttributeString("rgb", border.BottomColor);
                                writer.WriteEndElement();
                            }
                        }
                        writer.WriteEndElement();
                        
                        // Diagonal
                        writer.WriteStartElement("diagonal");
                        if (!string.IsNullOrEmpty(border.Diagonal) && border.Diagonal != "none")
                        {
                            writer.WriteAttributeString("style", border.Diagonal);
                            if (!string.IsNullOrEmpty(border.DiagonalColor))
                            {
                                writer.WriteStartElement("color");
                                writer.WriteAttributeString("rgb", border.DiagonalColor);
                                writer.WriteEndElement();
                            }
                        }
                        writer.WriteEndElement();
                        
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();
                    
                    // cellStyleXfs (规范要求)
                    writer.WriteStartElement("cellStyleXfs");
                    writer.WriteAttributeString("count", "1");
                    writer.WriteStartElement("xf");
                    writer.WriteAttributeString("fontId", "0");
                    writer.WriteAttributeString("fillId", "0");
                    writer.WriteAttributeString("borderId", "0");
                    writer.WriteAttributeString("numFmtId", "0");
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    
                    // 单元格格式 (cellXfs)
                    writer.WriteStartElement("cellXfs");
                    
                    // 使用解析出的 Xf 列表，或者如果为空则使用默认
                    List<Xf> xfs = workbook.XfList.Count > 0 ? workbook.XfList : new List<Xf> { new Xf() };
                    
                    Logger.Info($"写入 styles.xml: fonts={workbook.Fonts.Count}, fills={fills.Count}, borders={borders.Count}, xfs={xfs.Count}");
                    writer.WriteAttributeString("count", xfs.Count.ToString());
                    foreach (var xf in xfs)
                    {
                        writer.WriteStartElement("xf");
                        writer.WriteAttributeString("numFmtId", xf.NumberFormatIndex.ToString());
                        writer.WriteAttributeString("fontId", xf.FontIndex.ToString());
                        writer.WriteAttributeString("fillId", xf.FillIndex.ToString());
                        writer.WriteAttributeString("borderId", xf.BorderIndex.ToString());
                        writer.WriteAttributeString("xfId", "0");
                        
                        if (xf.NumberFormatIndex > 0)
                            writer.WriteAttributeString("applyNumberFormat", "1");
                        if (xf.FontIndex > 0)
                            writer.WriteAttributeString("applyFont", "1");
                        if (xf.FillIndex > 0)
                            writer.WriteAttributeString("applyFill", "1");
                        if (xf.BorderIndex > 0)
                            writer.WriteAttributeString("applyBorder", "1");
                        
                        if (xf.ApplyAlignment || (!string.IsNullOrEmpty(xf.HorizontalAlignment) || !string.IsNullOrEmpty(xf.VerticalAlignment) || xf.WrapText || xf.Indent > 0))
                        {
                            writer.WriteAttributeString("applyAlignment", "1");
                            writer.WriteStartElement("alignment");
                            if (!string.IsNullOrEmpty(xf.HorizontalAlignment))
                                writer.WriteAttributeString("horizontal", xf.HorizontalAlignment);
                            if (!string.IsNullOrEmpty(xf.VerticalAlignment))
                                writer.WriteAttributeString("vertical", xf.VerticalAlignment);
                            if (xf.WrapText)
                                writer.WriteAttributeString("wrapText", "1");
                            if (xf.Indent > 0)
                                writer.WriteAttributeString("indent", xf.Indent.ToString());
                            writer.WriteEndElement();
                        }
                            
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();
                    
                    // cellStyles (XLSX 规范要求此元素)
                    writer.WriteStartElement("cellStyles");
                    writer.WriteAttributeString("count", "1");
                    writer.WriteStartElement("cellStyle");
                    writer.WriteAttributeString("name", "Normal");
                    writer.WriteAttributeString("xfId", "0");
                    writer.WriteAttributeString("builtinId", "0");
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"创建样式XML时发生错误: {ex.Message}", ex);
                throw new StyleProcessingException($"创建样式XML时发生错误: {ex.Message}", ex);
            }
        }

        private void CreateSharedStringsXml(ZipArchive archive, Workbook workbook)
        {
            // 收集所有共享字符串
            var sharedStrings = new List<string>();
            var stringCount = 0;
            
            // 遍历所有工作表
            foreach (var worksheet in workbook.Worksheets)
            {
                // 遍历所有行和单元格
                foreach (var row in worksheet.Rows)
                {
                    foreach (var cell in row.Cells)
                    {
                        // 只处理文本类型的单元格
                if (cell.DataType == "s" || cell.DataType == "inlineStr" || (cell.DataType == null && cell.Value is string))
                {
                    var textValue = cell.Value?.ToString() ?? "";
                    // 检查字符串是否已经存在
                    if (!sharedStrings.Contains(textValue))
                    {
                        sharedStrings.Add(textValue);
                    }
                    stringCount++;
                }
                    }
                }
            }
            
            // 将共享字符串存储到workbook对象中，供后续使用
            workbook.SharedStrings = sharedStrings;
            
            var entry = archive.CreateEntry("xl/sharedStrings.xml");
            using (var stream = entry.Open())
            using (var writer = XmlWriter.Create(stream, new XmlWriterSettings { Indent = false, OmitXmlDeclaration = false }))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("sst", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                writer.WriteAttributeString("count", stringCount.ToString());
                writer.WriteAttributeString("uniqueCount", sharedStrings.Count.ToString());
                
                // 写入共享字符串
                foreach (var str in sharedStrings)
                {
                    writer.WriteStartElement("si");
                    writer.WriteStartElement("t");
                    writer.WriteString(CleanXmlString(str));
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                }
                
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }
        
        private void CreateDrawings(ZipArchive archive, Workbook workbook)
        {
            // 创建drawings目录
            archive.CreateEntry("xl/drawings/");
            archive.CreateEntry("xl/media/");
            
            // 为每个工作表创建drawing.xml文件
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                var worksheet = workbook.Worksheets[i];
                if (worksheet.Pictures.Count > 0 || worksheet.EmbeddedObjects.Count > 0 || worksheet.Charts.Count > 0)
                {
                    // 为每个图片创建媒体文件
                    for (int j = 0; j < worksheet.Pictures.Count; j++)
                    {
                        var picture = worksheet.Pictures[j];
                        if (picture.Data != null)
                        {
                            try
                            {
                                string extension = picture.Extension ?? "bmp";
                                var entry = archive.CreateEntry($"xl/media/image{j + 1}.{extension}");
                                using (var stream = entry.Open())
                                {
                                    stream.Write(picture.Data, 0, picture.Data.Length);
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new ImageProcessingException($"保存图片时发生错误: {ex.Message}", ex);
                            }
                        }
                    }
                    
                    try
                    {
                        CreateDrawingXml(archive, worksheet, i + 1, workbook);
                    }
                    catch (Exception ex)
                    {
                        throw new ChartProcessingException($"创建绘图XML时发生错误: {ex.Message}", ex);
                    }
                    
                    // 为每个图表创建图表文件
                    for (int j = 0; j < worksheet.Charts.Count; j++)
                    {
                        var chart = worksheet.Charts[j];
                        try
                        {
                            CreateChartXml(archive, chart, i + 1, j + 1);
                        }
                        catch (Exception ex)
                        {
                            throw new ChartProcessingException($"创建图表XML时发生错误: {ex.Message}", ex);
                        }
                    }
                    
                    // 生成 drawing rels
                    CreateDrawingRelsXml(archive, worksheet, i + 1);
                }
            }
        }
        
        private void CreateDrawingRelsXml(ZipArchive archive, Worksheet worksheet, int sheetIndex)
        {
            var entry = archive.CreateEntry($"xl/drawings/_rels/drawing{sheetIndex}.xml.rels");
            using (var stream = entry.Open())
            using (var writer = XmlWriter.Create(stream, new XmlWriterSettings { Indent = false, OmitXmlDeclaration = false }))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
                
                // 为每个图片添加关系
                for (int j = 0; j < worksheet.Pictures.Count; j++)
                {
                    var picture = worksheet.Pictures[j];
                    string extension = picture.Extension ?? "bmp";
                    
                    writer.WriteStartElement("Relationship");
                    writer.WriteAttributeString("Id", "rId" + (j + 1));
                    writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
                    writer.WriteAttributeString("Target", $"../media/image{j + 1}.{extension}");
                    writer.WriteEndElement();
                }
                
                // 为每个图表添加关系
                for (int j = 0; j < worksheet.Charts.Count; j++)
                {
                    writer.WriteStartElement("Relationship");
                    writer.WriteAttributeString("Id", "rIdChart" + (j + 1));
                    writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart");
                    writer.WriteAttributeString("Target", $"../charts/chart{j + 1}.xml");
                    writer.WriteEndElement();
                }
                
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }
        
        private void CreateDrawingXml(ZipArchive archive, Worksheet worksheet, int sheetIndex, Workbook workbook)
        {
            var entry = archive.CreateEntry($"xl/drawings/drawing{sheetIndex}.xml");
            using (var stream = entry.Open())
            using (var writer = XmlWriter.Create(stream, new XmlWriterSettings { Indent = false, OmitXmlDeclaration = false }))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("wsDr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
                writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                writer.WriteAttributeString("xmlns", "a", null, "http://schemas.openxmlformats.org/drawingml/2006/main");
                
                // 写入图片
                for (int i = 0; i < worksheet.Pictures.Count; i++)
                {
                    var picture = worksheet.Pictures[i];
                    writer.WriteStartElement("twoCellAnchor");
                    
                    // 写入图片位置
                    writer.WriteStartElement("from");
                    
                    // 计算列和列偏移（Excel的单位是1/1024英寸，这里转换为EMU单位）
                    double colWidth = 8.43; // 默认列宽（字符）
                    double colPixels = colWidth * 7; // 每个字符约7像素
                    double emuPerPixel = 9525; // 1像素 = 9525 EMU
                    
                    double leftEmu = picture.Left * emuPerPixel;
                    int col = (int)(leftEmu / (colPixels * emuPerPixel));
                    long colOff = (long)(leftEmu % (colPixels * emuPerPixel));
                    
                    // 计算行和行偏移
                    double rowHeight = 15; // 默认行高（像素）
                    double topEmu = picture.Top * emuPerPixel;
                    int row = (int)(topEmu / (rowHeight * emuPerPixel));
                    long rowOff = (long)(topEmu % (rowHeight * emuPerPixel));
                    
                    writer.WriteStartElement("col");
                    writer.WriteString(col.ToString());
                    writer.WriteEndElement();
                    writer.WriteStartElement("colOff");
                    writer.WriteString(colOff.ToString());
                    writer.WriteEndElement();
                    writer.WriteStartElement("row");
                    writer.WriteString(row.ToString());
                    writer.WriteEndElement();
                    writer.WriteStartElement("rowOff");
                    writer.WriteString(rowOff.ToString());
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    
                    writer.WriteStartElement("to");
                    
                    // 计算结束位置
                    double rightEmu = (picture.Left + picture.Width) * emuPerPixel;
                    int endCol = (int)(rightEmu / (colPixels * emuPerPixel));
                    long endColOff = (long)(rightEmu % (colPixels * emuPerPixel));
                    
                    double bottomEmu = (picture.Top + picture.Height) * emuPerPixel;
                    int endRow = (int)(bottomEmu / (rowHeight * emuPerPixel));
                    long endRowOff = (long)(bottomEmu % (rowHeight * emuPerPixel));
                    
                    writer.WriteStartElement("col");
                    writer.WriteString(endCol.ToString());
                    writer.WriteEndElement();
                    writer.WriteStartElement("colOff");
                    writer.WriteString(endColOff.ToString());
                    writer.WriteEndElement();
                    writer.WriteStartElement("row");
                    writer.WriteString(endRow.ToString());
                    writer.WriteEndElement();
                    writer.WriteStartElement("rowOff");
                    writer.WriteString(endRowOff.ToString());
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    
                    // 写入图片数据
                    writer.WriteStartElement("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");
                    writer.WriteStartElement("nvPicPr");
                    writer.WriteStartElement("cNvPr");
                    writer.WriteAttributeString("id", (i + 1).ToString());
                    writer.WriteAttributeString("name", $"Picture {i + 1}");
                    writer.WriteEndElement();
                    writer.WriteStartElement("cNvPicPr");
                    writer.WriteStartElement("picLocks");
                    writer.WriteAttributeString("noChangeAspect", "1");
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    
                    writer.WriteStartElement("blipFill");
                    writer.WriteStartElement("blip", "http://schemas.openxmlformats.org/drawingml/2006/main");
                    writer.WriteAttributeString("r", "embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", $"rId{i + 1}");
                    writer.WriteEndElement();
                    writer.WriteStartElement("stretch", "http://schemas.openxmlformats.org/drawingml/2006/main");
                    writer.WriteStartElement("fillRect", "http://schemas.openxmlformats.org/drawingml/2006/main");
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    
                    writer.WriteStartElement("spPr");
                    writer.WriteStartElement("xfrm");
                    writer.WriteStartElement("off");
                    writer.WriteAttributeString("x", "0");
                    writer.WriteAttributeString("y", "0");
                    writer.WriteEndElement();
                    writer.WriteStartElement("ext");
                    writer.WriteAttributeString("cx", picture.Width.ToString());
                    writer.WriteAttributeString("cy", picture.Height.ToString());
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteStartElement("prstGeom");
                    writer.WriteAttributeString("prst", "rect");
                    writer.WriteStartElement("avLst");
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    
                    writer.WriteEndElement();
                    // clientData 是 twoCellAnchor 的必需子元素
                    writer.WriteStartElement("clientData");
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                }
                
                // 写入图表
                for (int i = 0; i < worksheet.Charts.Count; i++)
                {
                    var chart = worksheet.Charts[i];
                    writer.WriteStartElement("twoCellAnchor");
                    
                    // 写入图表位置
                    writer.WriteStartElement("from");
                    
                    // 计算列和列偏移（Excel的单位是1/1024英寸，这里转换为EMU单位）
                    double colWidth = 8.43; // 默认列宽（字符）
                    double colPixels = colWidth * 7; // 每个字符约7像素
                    double emuPerPixel = 9525; // 1像素 = 9525 EMU
                    
                    double leftEmu = chart.Left * emuPerPixel;
                    int col = (int)(leftEmu / (colPixels * emuPerPixel));
                    long colOff = (long)(leftEmu % (colPixels * emuPerPixel));
                    
                    // 计算行和行偏移
                    double rowHeight = 15; // 默认行高（像素）
                    double topEmu = chart.Top * emuPerPixel;
                    int row = (int)(topEmu / (rowHeight * emuPerPixel));
                    long rowOff = (long)(topEmu % (rowHeight * emuPerPixel));
                    
                    writer.WriteStartElement("col");
                    writer.WriteString(col.ToString());
                    writer.WriteEndElement();
                    writer.WriteStartElement("colOff");
                    writer.WriteString(colOff.ToString());
                    writer.WriteEndElement();
                    writer.WriteStartElement("row");
                    writer.WriteString(row.ToString());
                    writer.WriteEndElement();
                    writer.WriteStartElement("rowOff");
                    writer.WriteString(rowOff.ToString());
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    
                    writer.WriteStartElement("to");
                    
                    // 计算结束位置
                    double rightEmu = (chart.Left + chart.Width) * emuPerPixel;
                    int endCol = (int)(rightEmu / (colPixels * emuPerPixel));
                    long endColOff = (long)(rightEmu % (colPixels * emuPerPixel));
                    
                    double bottomEmu = (chart.Top + chart.Height) * emuPerPixel;
                    int endRow = (int)(bottomEmu / (rowHeight * emuPerPixel));
                    long endRowOff = (long)(bottomEmu % (rowHeight * emuPerPixel));
                    
                    writer.WriteStartElement("col");
                    writer.WriteString(endCol.ToString());
                    writer.WriteEndElement();
                    writer.WriteStartElement("colOff");
                    writer.WriteString(endColOff.ToString());
                    writer.WriteEndElement();
                    writer.WriteStartElement("row");
                    writer.WriteString(endRow.ToString());
                    writer.WriteEndElement();
                    writer.WriteStartElement("rowOff");
                    writer.WriteString(endRowOff.ToString());
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    
                    // 写入图表数据
                    writer.WriteStartElement("graphicFrame");
                    writer.WriteStartElement("nvGraphicFramePr");
                    writer.WriteStartElement("cNvPr");
                    writer.WriteAttributeString("id", (worksheet.Pictures.Count + i + 1).ToString());
                    writer.WriteAttributeString("name", $"Chart {i + 1}");
                    writer.WriteEndElement();
                    writer.WriteStartElement("cNvGraphicFramePr");
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    
                    writer.WriteStartElement("xfrm");
                    writer.WriteStartElement("off", "http://schemas.openxmlformats.org/drawingml/2006/main");
                    writer.WriteAttributeString("x", "0");
                    writer.WriteAttributeString("y", "0");
                    writer.WriteEndElement();
                    writer.WriteStartElement("ext", "http://schemas.openxmlformats.org/drawingml/2006/main");
                    writer.WriteAttributeString("cx", chart.Width.ToString());
                    writer.WriteAttributeString("cy", chart.Height.ToString());
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    
                    writer.WriteStartElement("graphic", "http://schemas.openxmlformats.org/drawingml/2006/main");
                    writer.WriteStartElement("graphicData", "http://schemas.openxmlformats.org/drawingml/2006/main");
                    writer.WriteAttributeString("uri", "http://schemas.openxmlformats.org/drawingml/2006/chart");
                    writer.WriteStartElement("chart", "http://schemas.openxmlformats.org/drawingml/2006/chart");
                    writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", $"rIdChart{i + 1}");
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    
                    writer.WriteEndElement();
                    // clientData 是 twoCellAnchor 的必需子元素
                    writer.WriteStartElement("clientData");
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                }
                
                // 生成超链接
                if (worksheet.Hyperlinks.Count > 0)
                {
                    writer.WriteStartElement("hyperlinks");
                    writer.WriteAttributeString("count", worksheet.Hyperlinks.Count.ToString());
                    
                    foreach (var hyperlink in worksheet.Hyperlinks)
                    {
                        writer.WriteStartElement("hyperlink");
                        writer.WriteAttributeString("ref", hyperlink.Range);
                        // 计算超链接在工作簿中的索引
                        int hyperlinkIndex = workbook.Hyperlinks.IndexOf(hyperlink);
                        if (hyperlinkIndex >= 0)
                        {
                            // 生成唯一的r:id，确保与workbook.xml.rels中的ID对应
                            writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "rId" + (workbook.Worksheets.Count + 3 + hyperlinkIndex));
                        }
                        writer.WriteEndElement();
                    }
                    
                    writer.WriteEndElement();
                }
                
                // 生成注释引用
                if (worksheet.Comments.Count > 0)
                {
                    writer.WriteStartElement("legacyDrawing");
                    writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "rIdComments" + sheetIndex);
                    writer.WriteEndElement();
                }
                
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        private string GetChartElementType(string chartType)
        {
            // 映射图表类型到OpenXML元素名称
            switch (chartType)
            {
                case "barChart": return "barChart";
                case "colChart": return "colChart";
                case "lineChart": return "lineChart";
                case "pieChart": return "pieChart";
                case "scatterChart": return "scatterChart";
                case "areaChart": return "areaChart";
                case "doughnutChart": return "doughnutChart";
                case "radarChart": return "radarChart";
                case "surfaceChart": return "surfaceChart";
                case "bubbleChart": return "bubbleChart";
                case "stockChart": return "stockChart";
                default: return "barChart";
            }
        }
        
        private void CreateChartXml(ZipArchive archive, Chart chart, int sheetIndex, int chartIndex)
        {
            try
            {
                var entry = archive.CreateEntry($"xl/charts/chart{chartIndex}.xml");
                using (var stream = entry.Open())
                using (var writer = XmlWriter.Create(stream, new XmlWriterSettings { Indent = false, OmitXmlDeclaration = false }))
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("chartSpace", "http://schemas.openxmlformats.org/drawingml/2006/chart");
                    writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                    writer.WriteAttributeString("xmlns", "a", null, "http://schemas.openxmlformats.org/drawingml/2006/main");
                    
                    writer.WriteStartElement("chart");
                    
                    // 写入图表标题
                    if (!string.IsNullOrEmpty(chart.Title))
                    {
                        writer.WriteStartElement("title");
                        writer.WriteStartElement("tx");
                        writer.WriteStartElement("rich");
                        writer.WriteStartElement("bodyPr", "http://schemas.openxmlformats.org/drawingml/2006/main");
                        writer.WriteEndElement();
                        writer.WriteStartElement("lstStyle", "http://schemas.openxmlformats.org/drawingml/2006/main");
                        writer.WriteEndElement();
                        writer.WriteStartElement("p", "http://schemas.openxmlformats.org/drawingml/2006/main");
                        writer.WriteStartElement("r", "http://schemas.openxmlformats.org/drawingml/2006/main");
                        writer.WriteStartElement("t", "http://schemas.openxmlformats.org/drawingml/2006/main");
                        writer.WriteString(chart.Title);
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                        writer.WriteStartElement("layout");
                        writer.WriteEndElement();
                        writer.WriteStartElement("overlay");
                        writer.WriteAttributeString("val", "0");
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                    }
                    
                    // OOXML chart structure: chart > (title?) > plotArea > (layout, chartType, catAx, valAx) > legend?
                    
                    // plotArea 包含图表类型和坐标轴
                    writer.WriteStartElement("plotArea");
                    writer.WriteStartElement("layout");
                    writer.WriteEndElement(); // layout
                    
                    // 写入图表类型
                    string chartElementType = GetChartElementType(chart.ChartType);
                    writer.WriteStartElement(chartElementType);
                    
                    // 为不同类型的图表添加适当的元素
                    if (chart.ChartType == "barChart" || chart.ChartType == "colChart")
                    {
                        writer.WriteStartElement("barDir");
                        writer.WriteAttributeString("val", chart.ChartType == "barChart" ? "bar" : "col");
                        writer.WriteEndElement();
                        writer.WriteStartElement("grouping");
                        writer.WriteAttributeString("val", "standard");
                        writer.WriteEndElement();
                    }
                    else if (chart.ChartType == "lineChart" || chart.ChartType == "areaChart")
                    {
                        writer.WriteStartElement("grouping");
                        writer.WriteAttributeString("val", "standard");
                        writer.WriteEndElement();
                    }
                    else if (chart.ChartType == "pieChart" || chart.ChartType == "doughnutChart")
                    {
                        writer.WriteStartElement("firstSliceAng");
                        writer.WriteAttributeString("val", "0");
                        writer.WriteEndElement();
                    }
                    else if (chart.ChartType == "scatterChart")
                    {
                        writer.WriteStartElement("scatterStyle");
                        writer.WriteAttributeString("val", "marker");
                        writer.WriteEndElement();
                    }
                    else if (chart.ChartType == "radarChart")
                    {
                        writer.WriteStartElement("radarStyle");
                        writer.WriteAttributeString("val", "standard");
                        writer.WriteEndElement();
                    }
                    else if (chart.ChartType == "surfaceChart")
                    {
                        writer.WriteStartElement("wireframe");
                        writer.WriteAttributeString("val", "0");
                        writer.WriteEndElement();
                    }
                    else if (chart.ChartType == "bubbleChart")
                    {
                        writer.WriteStartElement("bubbleSizeRepresents");
                        writer.WriteAttributeString("val", "area");
                        writer.WriteEndElement();
                    }
                    else if (chart.ChartType == "stockChart")
                    {
                        writer.WriteStartElement("hiLoLines");
                        writer.WriteStartElement("spPr");
                        writer.WriteStartElement("ln", "http://schemas.openxmlformats.org/drawingml/2006/main");
                        writer.WriteStartElement("prstDash", "http://schemas.openxmlformats.org/drawingml/2006/main");
                        writer.WriteAttributeString("val", "solid");
                        writer.WriteEndElement();
                        writer.WriteStartElement("wid", "http://schemas.openxmlformats.org/drawingml/2006/main");
                        writer.WriteAttributeString("val", "1");
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                    }
                    
                    // 写入系列
                    foreach (var series in chart.Series)
                    {
                        writer.WriteStartElement("ser");
                        writer.WriteStartElement("idx");
                        writer.WriteAttributeString("val", chart.Series.IndexOf(series).ToString());
                        writer.WriteEndElement();
                        writer.WriteStartElement("order");
                        writer.WriteAttributeString("val", chart.Series.IndexOf(series).ToString());
                        writer.WriteEndElement();
                        
                        if (!string.IsNullOrEmpty(series.Name))
                        {
                            writer.WriteStartElement("tx");
                            writer.WriteStartElement("strRef");
                            writer.WriteStartElement("f");
                            writer.WriteString(series.Name);
                            writer.WriteEndElement();
                            writer.WriteStartElement("strCache");
                            writer.WriteStartElement("ptCount");
                            writer.WriteAttributeString("val", "1");
                            writer.WriteEndElement();
                            writer.WriteStartElement("pt");
                            writer.WriteAttributeString("idx", "0");
                            writer.WriteStartElement("v");
                            writer.WriteString(series.Name);
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                        }
                        
                        // 为不同类型的图表添加系列特定元素 (spPr must come before cat/val per OOXML)
                        if (chart.ChartType == "lineChart")
                        {
                            writer.WriteStartElement("spPr");
                            writer.WriteStartElement("ln", "http://schemas.openxmlformats.org/drawingml/2006/main");
                            writer.WriteStartElement("prstDash", "http://schemas.openxmlformats.org/drawingml/2006/main");
                            writer.WriteAttributeString("val", "solid");
                            writer.WriteEndElement();
                            writer.WriteStartElement("wid", "http://schemas.openxmlformats.org/drawingml/2006/main");
                            writer.WriteAttributeString("val", series.LineStyle?.Width.ToString() ?? "2");
                            writer.WriteEndElement();
                            if (!string.IsNullOrEmpty(series.LineStyle?.Color))
                            {
                                writer.WriteStartElement("solidFill", "http://schemas.openxmlformats.org/drawingml/2006/main");
                                writer.WriteStartElement("srgbClr", "http://schemas.openxmlformats.org/drawingml/2006/main");
                                writer.WriteAttributeString("val", series.LineStyle.Color);
                                writer.WriteEndElement();
                                writer.WriteEndElement();
                            }
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                        }
                        else if (chart.ChartType == "barChart" || chart.ChartType == "colChart")
                        {
                            writer.WriteStartElement("spPr");
                            writer.WriteStartElement("solidFill", "http://schemas.openxmlformats.org/drawingml/2006/main");
                            writer.WriteStartElement("srgbClr", "http://schemas.openxmlformats.org/drawingml/2006/main");
                            writer.WriteAttributeString("val", series.FillColor ?? "638EC6");
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                        }
                        else if (chart.ChartType == "pieChart" || chart.ChartType == "doughnutChart")
                        {
                            writer.WriteStartElement("spPr");
                            writer.WriteStartElement("solidFill", "http://schemas.openxmlformats.org/drawingml/2006/main");
                            writer.WriteStartElement("srgbClr", "http://schemas.openxmlformats.org/drawingml/2006/main");
                            writer.WriteAttributeString("val", series.FillColor ?? "638EC6");
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                        }
                        else if (chart.ChartType == "scatterChart")
                        {
                            writer.WriteStartElement("marker");
                            writer.WriteStartElement("symbol");
                            writer.WriteAttributeString("val", "circle");
                            writer.WriteEndElement();
                            writer.WriteStartElement("size");
                            writer.WriteAttributeString("val", "5");
                            writer.WriteEndElement();
                            writer.WriteStartElement("spPr");
                            writer.WriteStartElement("solidFill", "http://schemas.openxmlformats.org/drawingml/2006/main");
                            writer.WriteStartElement("srgbClr", "http://schemas.openxmlformats.org/drawingml/2006/main");
                            writer.WriteAttributeString("val", series.FillColor ?? "638EC6");
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                        }
                        
                        if (!string.IsNullOrEmpty(series.CategoriesRange))
                        {
                            writer.WriteStartElement("cat");
                            writer.WriteStartElement("strRef");
                            writer.WriteStartElement("f");
                            writer.WriteString(series.CategoriesRange);
                            writer.WriteEndElement();
                            writer.WriteStartElement("strCache");
                            writer.WriteStartElement("ptCount");
                            writer.WriteAttributeString("val", "1");
                            writer.WriteEndElement();
                            writer.WriteStartElement("pt");
                            writer.WriteAttributeString("idx", "0");
                            writer.WriteStartElement("v");
                            writer.WriteString("Category");
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                        }
                        
                        if (!string.IsNullOrEmpty(series.ValuesRange))
                        {
                            writer.WriteStartElement("val");
                            writer.WriteStartElement("numRef");
                            writer.WriteStartElement("f");
                            writer.WriteString(series.ValuesRange);
                            writer.WriteEndElement();
                            writer.WriteStartElement("numCache");
                            writer.WriteStartElement("ptCount");
                            writer.WriteAttributeString("val", "1");
                            writer.WriteEndElement();
                            writer.WriteStartElement("pt");
                            writer.WriteAttributeString("idx", "0");
                            writer.WriteStartElement("v");
                            writer.WriteString("0");
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                        }
                        
                        writer.WriteEndElement(); // ser
                    }
                    
                    // barChart/colChart 需要 axId 引用
                    if (chart.ChartType == "barChart" || chart.ChartType == "colChart" || chart.ChartType == "lineChart" || chart.ChartType == "areaChart")
                    {
                        writer.WriteStartElement("axId");
                        writer.WriteAttributeString("val", "1");
                        writer.WriteEndElement();
                        writer.WriteStartElement("axId");
                        writer.WriteAttributeString("val", "2");
                        writer.WriteEndElement();
                    }
                    
                    writer.WriteEndElement(); // chartType element
                    
                    // 写入X轴 (catAx) - inside plotArea
                    if (chart.XAxis != null && chart.XAxis.Visible)
                    {
                        writer.WriteStartElement("catAx");
                        writer.WriteStartElement("axId");
                        writer.WriteAttributeString("val", "1");
                        writer.WriteEndElement();
                        writer.WriteStartElement("scaling");
                        writer.WriteStartElement("orientation");
                        writer.WriteAttributeString("val", "minMax");
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                        writer.WriteStartElement("delete");
                        writer.WriteAttributeString("val", "0");
                        writer.WriteEndElement();
                        writer.WriteStartElement("axPos");
                        writer.WriteAttributeString("val", "b");
                        writer.WriteEndElement();
                        if (!string.IsNullOrEmpty(chart.XAxis.Title))
                        {
                            writer.WriteStartElement("title");
                            writer.WriteStartElement("tx");
                            writer.WriteStartElement("rich");
                            writer.WriteStartElement("bodyPr", "http://schemas.openxmlformats.org/drawingml/2006/main");
                            writer.WriteEndElement();
                            writer.WriteStartElement("lstStyle", "http://schemas.openxmlformats.org/drawingml/2006/main");
                            writer.WriteEndElement();
                            writer.WriteStartElement("p", "http://schemas.openxmlformats.org/drawingml/2006/main");
                            writer.WriteStartElement("r", "http://schemas.openxmlformats.org/drawingml/2006/main");
                            writer.WriteStartElement("t", "http://schemas.openxmlformats.org/drawingml/2006/main");
                            writer.WriteString(chart.XAxis.Title);
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteStartElement("layout");
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                        }
                        writer.WriteStartElement("numFmt");
                        writer.WriteAttributeString("formatCode", chart.XAxis.NumberFormat ?? "General");
                        writer.WriteAttributeString("sourceLinked", "1");
                        writer.WriteEndElement();
                        writer.WriteStartElement("tickLblPos");
                        writer.WriteAttributeString("val", "nextTo");
                        writer.WriteEndElement();
                        writer.WriteStartElement("crossAx");
                        writer.WriteAttributeString("val", "2");
                        writer.WriteEndElement();
                        writer.WriteStartElement("crosses");
                        writer.WriteAttributeString("val", "autoZero");
                        writer.WriteEndElement();
                        writer.WriteStartElement("auto");
                        writer.WriteAttributeString("val", "1");
                        writer.WriteEndElement();
                        writer.WriteStartElement("lblAlgn");
                        writer.WriteAttributeString("val", "ctr");
                        writer.WriteEndElement();
                        writer.WriteStartElement("lblOffset");
                        writer.WriteAttributeString("val", "100");
                        writer.WriteEndElement();
                        
                        writer.WriteEndElement(); // catAx
                    }
                    
                    // 写入Y轴 (valAx) - inside plotArea
                    if (chart.YAxis != null && chart.YAxis.Visible)
                    {
                        writer.WriteStartElement("valAx");
                        writer.WriteStartElement("axId");
                        writer.WriteAttributeString("val", "2");
                        writer.WriteEndElement();
                        writer.WriteStartElement("scaling");
                        writer.WriteStartElement("orientation");
                        writer.WriteAttributeString("val", "minMax");
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                        writer.WriteStartElement("delete");
                        writer.WriteAttributeString("val", "0");
                        writer.WriteEndElement();
                        writer.WriteStartElement("axPos");
                        writer.WriteAttributeString("val", "l");
                        writer.WriteEndElement();
                        writer.WriteStartElement("majorGridlines");
                        writer.WriteEndElement();
                        if (!string.IsNullOrEmpty(chart.YAxis.Title))
                        {
                            writer.WriteStartElement("title");
                            writer.WriteStartElement("tx");
                            writer.WriteStartElement("rich");
                            writer.WriteStartElement("bodyPr", "http://schemas.openxmlformats.org/drawingml/2006/main");
                            writer.WriteEndElement();
                            writer.WriteStartElement("lstStyle", "http://schemas.openxmlformats.org/drawingml/2006/main");
                            writer.WriteEndElement();
                            writer.WriteStartElement("p", "http://schemas.openxmlformats.org/drawingml/2006/main");
                            writer.WriteStartElement("r", "http://schemas.openxmlformats.org/drawingml/2006/main");
                            writer.WriteStartElement("t", "http://schemas.openxmlformats.org/drawingml/2006/main");
                            writer.WriteString(chart.YAxis.Title);
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                            writer.WriteStartElement("layout");
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                        }
                        writer.WriteStartElement("numFmt");
                        writer.WriteAttributeString("formatCode", chart.YAxis.NumberFormat ?? "General");
                        writer.WriteAttributeString("sourceLinked", "1");
                        writer.WriteEndElement();
                        writer.WriteStartElement("tickLblPos");
                        writer.WriteAttributeString("val", "nextTo");
                        writer.WriteEndElement();
                        writer.WriteStartElement("crossAx");
                        writer.WriteAttributeString("val", "1");
                        writer.WriteEndElement();
                        writer.WriteStartElement("crosses");
                        writer.WriteAttributeString("val", "autoZero");
                        writer.WriteEndElement();
                        
                        writer.WriteEndElement(); // valAx
                    }
                    
                    writer.WriteEndElement(); // plotArea
                    
                    // 写入图例 (legend comes AFTER plotArea per OOXML spec)
                    if (chart.Legend != null && chart.Legend.Visible)
                    {
                        writer.WriteStartElement("legend");
                        writer.WriteStartElement("legendPos");
                        // Map position names to OOXML abbreviations
                        string legendPosVal = chart.Legend.Position switch
                        {
                            "right" => "r",
                            "left" => "l",
                            "top" => "t",
                            "bottom" => "b",
                            "topRight" => "tr",
                            _ => chart.Legend.Position ?? "r"
                        };
                        writer.WriteAttributeString("val", legendPosVal);
                        writer.WriteEndElement();
                        writer.WriteStartElement("layout");
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                    }
                    
                    writer.WriteEndElement(); // chart
                    writer.WriteEndElement(); // chartSpace
                    writer.WriteEndDocument();
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"创建图表XML时发生错误: {ex.Message}", ex);
                throw new ChartProcessingException($"创建图表XML时发生错误: {ex.Message}", ex);
            }
        }
        
        private string GetCellReference(int rowIndex, int columnIndex)
        {
            var columnReference = string.Empty;
            int col = columnIndex;
            while (col > 0)
            {
                col--;
                columnReference = (char)('A' + col % 26) + columnReference;
                col /= 26;
            }
            return columnReference + rowIndex;
        }
        
        private double DateTimeToExcelDate(DateTime dateTime)
        {
            // Excel 日期时间是从 1900-01-01 开始的天数
            // 注意：Excel 使用 1900 年 2 月 29 日作为有效日期，即使 1900 年不是闰年
            DateTime excelBaseDate = new DateTime(1900, 1, 1);
            TimeSpan timeSpan = dateTime - excelBaseDate;
            double days = timeSpan.TotalDays;
            
            // 调整 1900 年闰年问题
            if (dateTime >= new DateTime(1900, 3, 1))
            {
                days += 1;
            }
            
            return days;
        }
        
        private void CreateVbaProjectBin(ZipArchive archive, Workbook workbook)
        {
            // 创建VBA项目文件
            var entry = archive.CreateEntry("xl/vbaProject.bin");
            using (var stream = entry.Open())
            {
                if (workbook.VbaProject != null)
                {
                    stream.Write(workbook.VbaProject, 0, workbook.VbaProject.Length);
                }
            }
        }
    }
}