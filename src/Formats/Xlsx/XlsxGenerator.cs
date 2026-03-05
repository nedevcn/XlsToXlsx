using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
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

            // 将一些老式 Wingdings/符号字体中的项目符号等字符规范化为标准 Unicode 符号，避免在不同机器/字体下显示为“乱码”
            char NormalizeChar(char c)
            {
                // 常见的私有区项目符号（例如  等）统一映射为标准实心圆点
                if (c == '\uF0B7' || c == '\uF06C')
                    return '\u2022'; // '•'

                // 非断行空格等统一为普通空格，避免复制/显示异常
                if (c == '\u00A0')
                    return ' ';

                return c;
            }
            
            // 先做字符规范化，再移除 XML 无效字符
            return new string(input
                .Select(NormalizeChar)
                .Where(c =>
                    (c >= 0x0020 && c <= 0xD7FF) ||
                    (c >= 0xE000 && c <= 0xFFFD) ||
                    c == 0x0009 ||
                    c == 0x000A ||
                    c == 0x000D)
                .ToArray());
        }

        /// <summary>
        /// 规范化 sqref / Range 表达式，修复类似 "C65536:65536" 这种缺少列字母的旧 XLS 写法。
        /// </summary>
        private static string SanitizeSqref(string? range)
        {
            if (string.IsNullOrWhiteSpace(range))
                return range ?? string.Empty;

            var parts = range.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < parts.Length; i++)
            {
                var part = parts[i];
                int colonIndex = part.IndexOf(':');
                if (colonIndex <= 0 || colonIndex == part.Length - 1)
                    continue;

                string start = part.Substring(0, colonIndex);
                string end = part.Substring(colonIndex + 1);

                // 如果 end 已经包含列字母，则保持不变
                bool endHasLetter = end.Any(ch => (ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z'));
                if (endHasLetter)
                    continue;

                // 从 start 提取列字母前缀
                int idx = 0;
                while (idx < start.Length && ((start[idx] >= 'A' && start[idx] <= 'Z') || (start[idx] >= 'a' && start[idx] <= 'z')))
                {
                    idx++;
                }
                if (idx == 0)
                    continue;

                string col = start.Substring(0, idx);
                string fixedPart = $"{start}:{col}{end}";
                parts[i] = fixedPart;
            }

            return string.Join(" ", parts);
        }

        /// <summary>
        /// True when the sheet has at least one picture with data or one chart. We do not create a drawing part for
        /// EmbeddedObjects-only or pictures-without-data, to avoid empty/invalid drawing parts and rId/media mismatches.
        /// </summary>
        private static bool SheetHasDrawings(Worksheet ws)
        {
            if (ws == null) return false;
            if (ws.Pictures != null && ws.Pictures.Any(p => p != null && p.Data != null && p.Data.Length > 0))
                return true;
            if (ws.Charts != null && ws.Charts.Count > 0)
                return true;
            return false;
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

                    CreateDocPropsCoreXml(archive, workbook);
                    CreateDocPropsAppXml(archive, workbook);

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
                        CreateWorksheetRelsXml(archive, defaultWorksheet, 1, workbook);
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
                                CreateWorksheetRelsXml(archive, workbook.Worksheets[i], i + 1, workbook);
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
                    
                    // 创建 externalLinks（外部工作簿引用缓存）
                    CreateExternalLinks(archive, workbook);
                    
                    // 数据透视表：pivotCache 与 pivotTable 部件
                    CreatePivotParts(archive, workbook);
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
            using (var writer = XmlWriter.Create(stream, Utf8NoBomXmlSettings))
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
                    if (ws?.Pictures == null) continue;
                    foreach (var pic in ws.Pictures)
                    {
                        if (pic != null && !string.IsNullOrEmpty(pic.Extension))
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
                
                writer.WriteStartElement("Override");
                writer.WriteAttributeString("PartName", "/docProps/core.xml");
                writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-package.core-properties+xml");
                writer.WriteEndElement();
                writer.WriteStartElement("Override");
                writer.WriteAttributeString("PartName", "/docProps/app.xml");
                writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.extended-properties+xml");
                writer.WriteEndElement();
                // 包关系（部分验证器要求显式声明）
                writer.WriteStartElement("Override");
                writer.WriteAttributeString("PartName", "/_rels/.rels");
                writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
                writer.WriteEndElement();
                
                writer.WriteStartElement("Override");
                writer.WriteAttributeString("PartName", "/xl/_rels/workbook.xml.rels");
                writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-package.relationships+xml");
                writer.WriteEndElement();
                
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
                
                // Drawing and chart content types (every part we create must be declared)
                int chartIndex = 1;
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    var ws = workbook.Worksheets[i];
                    bool hasDrawings = SheetHasDrawings(ws);
                    if (hasDrawings)
                    {
                        // Drawing part (required for Pictures, EmbeddedObjects, or Charts)
                        writer.WriteStartElement("Override");
                        writer.WriteAttributeString("PartName", $"/xl/drawings/drawing{i + 1}.xml");
                        writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml");
                        writer.WriteEndElement();
                        
                        for (int j = 0; j < ws.Charts.Count; j++)
                        {
                            writer.WriteStartElement("Override");
                            writer.WriteAttributeString("PartName", $"/xl/charts/chart{chartIndex}.xml");
                            writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml");
                            writer.WriteEndElement();
                            chartIndex++;
                        }
                    }
                    // Do not add Override for /xl/comments{i+1}.xml — we only create VML (vmlDrawing), not the comments part
                }
                
                // VBA项目内容类型
                if (workbook.VbaProject != null)
                {
                    writer.WriteStartElement("Override");
                    writer.WriteAttributeString("PartName", "/xl/vbaProject.bin");
                    writer.WriteAttributeString("ContentType", "application/vnd.ms-office.vbaProject");
                    writer.WriteEndElement();
                }
                
                // externalLinks 内容类型
                int externalLinkCount = workbook.ExternalBooks?.Count(b => b != null && !b.IsSelf && !b.IsAddIn) ?? 0;
                for (int i = 0; i < externalLinkCount; i++)
                {
                    writer.WriteStartElement("Override");
                    writer.WriteAttributeString("PartName", $"/xl/externalLinks/externalLink{i + 1}.xml");
                    writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml");
                    writer.WriteEndElement();
                }
                // 数据透视表缓存与透视表部件
                int totalPivot = GetTotalPivotTableCount(workbook);
                for (int i = 0; i < totalPivot; i++)
                {
                    writer.WriteStartElement("Override");
                    writer.WriteAttributeString("PartName", $"/xl/pivotCache/pivotCacheDefinition{i + 1}.xml");
                    writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml");
                    writer.WriteEndElement();
                    writer.WriteStartElement("Override");
                    writer.WriteAttributeString("PartName", $"/xl/pivotTables/pivotTable{i + 1}.xml");
                    writer.WriteAttributeString("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml");
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
            using (var writer = XmlWriter.Create(stream, Utf8NoBomXmlSettings))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
                
                writer.WriteStartElement("Relationship");
                writer.WriteAttributeString("Id", "rId1");
                writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument");
                writer.WriteAttributeString("Target", "xl/workbook.xml");
                writer.WriteEndElement();
                writer.WriteStartElement("Relationship");
                writer.WriteAttributeString("Id", "rId2");
                writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties");
                writer.WriteAttributeString("Target", "docProps/core.xml");
                writer.WriteEndElement();
                writer.WriteStartElement("Relationship");
                writer.WriteAttributeString("Id", "rId3");
                writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
                writer.WriteAttributeString("Target", "docProps/app.xml");
                writer.WriteEndElement();
                
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        private void CreateDocPropsCoreXml(ZipArchive archive, Workbook workbook)
        {
            var entry = archive.CreateEntry("docProps/core.xml");
            using (var stream = entry.Open())
            using (var writer = XmlWriter.Create(stream, Utf8NoBomXmlSettings))
            {
                const string cpNs = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
                const string dcNs = "http://purl.org/dc/elements/1.1/";
                const string dctermsNs = "http://purl.org/dc/terms/";
                const string dcmitypeNs = "http://purl.org/dc/dcmitype/";
                const string xsiNs = "http://www.w3.org/2001/XMLSchema-instance";

                writer.WriteStartDocument();
                writer.WriteStartElement("cp", "coreProperties", cpNs);
                writer.WriteAttributeString("xmlns", "cp", null, cpNs);
                writer.WriteAttributeString("xmlns", "dc", null, dcNs);
                writer.WriteAttributeString("xmlns", "dcterms", null, dctermsNs);
                writer.WriteAttributeString("xmlns", "dcmitype", null, dcmitypeNs);
                writer.WriteAttributeString("xmlns", "xsi", null, xsiNs);

                string creator = string.IsNullOrEmpty(workbook.Author)
                    ? "Nedev.XlsToXlsx"
                    : workbook.Author!;
                writer.WriteElementString("dc", "creator", dcNs, creator);

                // LastModifiedBy
                string lastModifiedBy = !string.IsNullOrEmpty(workbook.LastAuthor)
                    ? workbook.LastAuthor!
                    : creator;
                writer.WriteElementString("cp", "lastModifiedBy", cpNs, lastModifiedBy);

                // 标题/主题/描述/关键字
                if (!string.IsNullOrEmpty(workbook.Title))
                    writer.WriteElementString("dc", "title", dcNs, workbook.Title);
                if (!string.IsNullOrEmpty(workbook.Subject))
                    writer.WriteElementString("dc", "subject", dcNs, workbook.Subject);
                if (!string.IsNullOrEmpty(workbook.Comments))
                    writer.WriteElementString("dc", "description", dcNs, workbook.Comments);
                if (!string.IsNullOrEmpty(workbook.Keywords))
                    writer.WriteElementString("cp", "keywords", cpNs, workbook.Keywords);

                DateTime created = workbook.CreatedUtc ?? DateTime.UtcNow;
                DateTime modified = workbook.ModifiedUtc ?? created;

                writer.WriteStartElement("dcterms", "created", dctermsNs);
                writer.WriteAttributeString("xsi", "type", xsiNs, "dcterms:W3CDTF");
                writer.WriteString(created.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ"));
                writer.WriteEndElement();

                writer.WriteStartElement("dcterms", "modified", dctermsNs);
                writer.WriteAttributeString("xsi", "type", xsiNs, "dcterms:W3CDTF");
                writer.WriteString(modified.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ"));
                writer.WriteEndElement();

                writer.WriteEndElement(); // cp:coreProperties
                writer.WriteEndDocument();
            }
        }

        private void CreateDocPropsAppXml(ZipArchive archive, Workbook workbook)
        {
            var entry = archive.CreateEntry("docProps/app.xml");
            using (var stream = entry.Open())
            using (var writer = XmlWriter.Create(stream, Utf8NoBomXmlSettings))
            {
                const string propsNs = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
                const string vtNs = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

                writer.WriteStartDocument();
                writer.WriteStartElement("Properties", propsNs);
                writer.WriteAttributeString("xmlns", "vt", null, vtNs);
                writer.WriteElementString("Application", propsNs, "Microsoft Excel");
                writer.WriteElementString("DocSecurity", propsNs, "0");
                writer.WriteElementString("ScaleCrop", propsNs, "false");

                if (!string.IsNullOrEmpty(workbook.Company))
                    writer.WriteElementString("Company", propsNs, workbook.Company);
                if (!string.IsNullOrEmpty(workbook.Manager))
                    writer.WriteElementString("Manager", propsNs, workbook.Manager);
                if (!string.IsNullOrEmpty(workbook.Category))
                    writer.WriteElementString("Category", propsNs, workbook.Category);

                int sheetCount = workbook.Worksheets.Count > 0 ? workbook.Worksheets.Count : 1;

                // HeadingPairs: vector of ("Worksheets", sheetCount)
                writer.WriteStartElement("HeadingPairs", propsNs);
                writer.WriteStartElement("vt", "vector", vtNs);
                writer.WriteAttributeString("size", "2");
                writer.WriteAttributeString("baseType", "variant");

                writer.WriteStartElement("vt", "variant", vtNs);
                writer.WriteElementString("vt", "lpstr", vtNs, "Worksheets");
                writer.WriteEndElement();

                writer.WriteStartElement("vt", "variant", vtNs);
                writer.WriteElementString("vt", "i4", vtNs, sheetCount.ToString());
                writer.WriteEndElement();

                writer.WriteEndElement(); // vt:vector
                writer.WriteEndElement(); // HeadingPairs

                // TitlesOfParts: sheet names
                writer.WriteStartElement("TitlesOfParts", propsNs);
                writer.WriteStartElement("vt", "vector", vtNs);
                writer.WriteAttributeString("size", sheetCount.ToString());
                writer.WriteAttributeString("baseType", "lpstr");
                for (int i = 0; i < sheetCount; i++)
                {
                    string name = workbook.Worksheets.Count > 0 ? (workbook.Worksheets[i].Name ?? "Sheet" + (i + 1)) : "Sheet1";
                    writer.WriteElementString("vt", "lpstr", vtNs, name);
                }
                writer.WriteEndElement(); // vt:vector
                writer.WriteEndElement(); // TitlesOfParts

                writer.WriteEndElement(); // Properties
                writer.WriteEndDocument();
            }
        }

        private static int GetTotalPivotTableCount(Workbook workbook)
        {
            if (workbook?.Worksheets == null) return 0;
            int n = 0;
            foreach (var ws in workbook.Worksheets)
            {
                if (ws?.PivotTables != null)
                    n += ws.PivotTables.Count;
            }
            return n;
        }

        private void CreateExternalLinks(ZipArchive archive, Workbook workbook)
        {
            if (workbook.ExternalBooks == null || workbook.ExternalBooks.Count == 0) return;
            var externalBooks = workbook.ExternalBooks.Where(b => b != null && !b.IsSelf && !b.IsAddIn).ToList();
            if (externalBooks.Count == 0) return;

            for (int i = 0; i < externalBooks.Count; i++)
            {
                var extBook = externalBooks[i];
                int linkIndex = i + 1;

                // externalLinkN.xml
                var entry = archive.CreateEntry($"xl/externalLinks/externalLink{linkIndex}.xml");
                using (var stream = entry.Open())
                using (var writer = XmlWriter.Create(stream, Utf8NoBomXmlSettings))
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("externalLink", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                    writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

                    writer.WriteStartElement("externalBook");
                    writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "rId1");

                    // sheetNames
                    if (extBook.SheetNames != null && extBook.SheetNames.Count > 0)
                    {
                        writer.WriteStartElement("sheetNames");
                        foreach (var name in extBook.SheetNames)
                        {
                            writer.WriteStartElement("sheetName");
                            writer.WriteAttributeString("val", name ?? string.Empty);
                            writer.WriteEndElement();
                        }
                        writer.WriteEndElement();
                    }

                    writer.WriteEndElement(); // externalBook
                    writer.WriteEndElement(); // externalLink
                    writer.WriteEndDocument();
                }

                // xl/externalLinks/_rels/externalLinkN.xml.rels
                var relEntry = archive.CreateEntry($"xl/externalLinks/_rels/externalLink{linkIndex}.xml.rels");
                using (var stream = relEntry.Open())
                using (var writer = XmlWriter.Create(stream, Utf8NoBomXmlSettings))
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
                    writer.WriteStartElement("Relationship");
                    writer.WriteAttributeString("Id", "rId1");
                    writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath");
                    string target = !string.IsNullOrEmpty(extBook.FileName) ? extBook.FileName! : string.Empty;
                    writer.WriteAttributeString("Target", target);
                    writer.WriteAttributeString("TargetMode", "External");
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }
            }
        }

        private void CreateWorkbookXml(ZipArchive archive, Workbook workbook)
        {
            var entry = archive.CreateEntry("xl/workbook.xml");
            using (var stream = entry.Open())
            using (var writer = XmlWriter.Create(stream, Utf8NoBomXmlSettings))
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
                // 工作簿保护（结构/窗口 + 密码哈希）
                if (workbook.IsStructureProtected || workbook.IsWindowsProtected || !string.IsNullOrEmpty(workbook.WorkbookPasswordHash))
                {
                    writer.WriteStartElement("workbookProtection");
                    if (workbook.IsStructureProtected)
                        writer.WriteAttributeString("lockStructure", "1");
                    if (workbook.IsWindowsProtected)
                        writer.WriteAttributeString("lockWindows", "1");
                    if (!string.IsNullOrEmpty(workbook.WorkbookPasswordHash))
                        writer.WriteAttributeString("password", workbook.WorkbookPasswordHash);
                    writer.WriteEndElement();
                }

                if (workbook.DefinedNames != null && workbook.DefinedNames.Count > 0)
                {
                    writer.WriteStartElement("definedNames");
                    int sheetCountForNames = Math.Max(1, workbook.Worksheets?.Count ?? 0);
                    foreach (var dn in workbook.DefinedNames)
                    {
                        if (string.IsNullOrEmpty(dn?.Name)) continue;
                        writer.WriteStartElement("definedName");
                        writer.WriteAttributeString("name", GetOoxmlBuiltInName(dn.Name!));
                        if (dn.LocalSheetId.HasValue && sheetCountForNames > 0)
                        {
                            int localId = Math.Clamp(dn.LocalSheetId.Value, 0, sheetCountForNames - 1);
                            writer.WriteAttributeString("localSheetId", localId.ToString());
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

                // 数据透视表缓存引用（每个透视表一个 cache）
                int totalPivotTables = GetTotalPivotTableCount(workbook);
                if (totalPivotTables > 0)
                {
                    int relIdBase = workbook.Worksheets.Count + 3;
                    writer.WriteStartElement("pivotCaches");
                    for (int i = 0; i < totalPivotTables; i++)
                    {
                        writer.WriteStartElement("pivotCache");
                        writer.WriteAttributeString("cacheId", (i + 1).ToString());
                        writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "rId" + (relIdBase + i));
                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement(); // pivotCaches
                }

                writer.WriteEndElement(); // workbook
                writer.WriteEndDocument();
            }
        }

        private void CreateWorkbookRelsXml(ZipArchive archive, Workbook workbook)
        {
            var entry = archive.CreateEntry("xl/_rels/workbook.xml.rels");
            using (var stream = entry.Open())
            using (var writer = XmlWriter.Create(stream, Utf8NoBomXmlSettings))
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
                
                int nextRelId = worksheetCount + 3;
                
                // Pivot cache relationships
                int totalPivotTables = GetTotalPivotTableCount(workbook);
                for (int i = 0; i < totalPivotTables; i++)
                {
                    writer.WriteStartElement("Relationship");
                    writer.WriteAttributeString("Id", "rId" + nextRelId);
                    writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition");
                    writer.WriteAttributeString("Target", "pivotCache/pivotCacheDefinition" + (i + 1) + ".xml");
                    writer.WriteEndElement();
                    nextRelId++;
                }
                
                // externalLinks relationships
                int externalLinkCount = workbook.ExternalBooks?.Count(b => b != null && !b.IsSelf && !b.IsAddIn) ?? 0;
                for (int i = 0; i < externalLinkCount; i++)
                {
                    writer.WriteStartElement("Relationship");
                    writer.WriteAttributeString("Id", "rId" + nextRelId);
                    writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink");
                    writer.WriteAttributeString("Target", "externalLinks/externalLink" + (i + 1) + ".xml");
                    writer.WriteEndElement();
                    nextRelId++;
                }
                
                // VBA项目关系（ID 在 pivot/externalLinks 之后）
                if (workbook.VbaProject != null)
                {
                    writer.WriteStartElement("Relationship");
                    writer.WriteAttributeString("Id", "rId" + nextRelId);
                    writer.WriteAttributeString("Type", "http://schemas.microsoft.com/office/2006/relationships/vbaProject");
                    writer.WriteAttributeString("Target", "vbaProject.bin");
                    writer.WriteEndElement();
                }
                
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        private static readonly XmlWriterSettings WorksheetXmlSettings = new XmlWriterSettings
        {
            Indent = false,
            OmitXmlDeclaration = false,
            Encoding = new UTF8Encoding(false)
        };

        private static readonly XmlWriterSettings Utf8NoBomXmlSettings = new XmlWriterSettings
        {
            Indent = false,
            OmitXmlDeclaration = false,
            Encoding = new UTF8Encoding(false)
        };

        private void CreateWorksheetXml(ZipArchive archive, Worksheet worksheet, int sheetIndex, Workbook workbook)
        {
            var entry = archive.CreateEntry($"xl/worksheets/sheet{sheetIndex}.xml");
            using (var stream = entry.Open())
            using (var writer = XmlWriter.Create(stream, WorksheetXmlSettings))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("worksheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                
                int styleCount = (workbook.XfList != null && workbook.XfList.Count > 0) ? workbook.XfList.Count : 1;
                int maxStyleId = Math.Max(0, styleCount - 1);
                
                // sheetPr（含 pageSetUpPr fitToPage，使 fitToWidth/fitToHeight 生效）
                var ps = worksheet.PageSettings;
                if (ps != null && (ps.FitToWidth > 0 || ps.FitToHeight > 0))
                {
                    writer.WriteStartElement("sheetPr");
                    writer.WriteStartElement("pageSetUpPr");
                    writer.WriteAttributeString("fitToPage", "1");
                    writer.WriteEndElement();
                    writer.WriteEndElement(); // sheetPr
                }
                
                // dimension（工作表范围；Excel 最大行 1048576，最大列 16384）
                writer.WriteStartElement("dimension");
                int maxRow = Math.Clamp(Math.Max(1, worksheet.MaxRow), 1, 1048576);
                int maxCol = Math.Clamp(Math.Max(1, worksheet.MaxColumn), 1, 16384);
                string maxRef = GetCellReference(maxRow, maxCol);
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
                // 工作表保护
                if (worksheet.IsProtected || !string.IsNullOrEmpty(worksheet.SheetPasswordHash))
                {
                    writer.WriteStartElement("sheetProtection");
                    if (worksheet.IsProtected)
                        writer.WriteAttributeString("sheet", "1");
                    if (!string.IsNullOrEmpty(worksheet.SheetPasswordHash))
                        writer.WriteAttributeString("password", worksheet.SheetPasswordHash);
                    // 常见默认：保护对象和方案
                    writer.WriteAttributeString("objects", "1");
                    writer.WriteAttributeString("scenarios", "1");
                    writer.WriteEndElement();
                }

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
                        if (colInfo == null) continue;
                        writer.WriteStartElement("col");
                        writer.WriteAttributeString("min", (colInfo.FirstColumn + 1).ToString()); // 1-based
                        writer.WriteAttributeString("max", (colInfo.LastColumn + 1).ToString());
                        // 将 1/256 字符宽度转为 XLSX 的字符宽度（除以 256）
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
                
                // sheetData：OOXML 要求 row 元素按 r 属性升序排列
                var rows = (worksheet.Rows ?? Enumerable.Empty<Row>()).OrderBy(r => r.RowIndex).ToList();
                writer.WriteStartElement("sheetData");
                foreach (var row in rows)
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
                    
                    foreach (var cell in row.Cells ?? Enumerable.Empty<Cell>())
                    {
                        if (cell == null) continue;
                        if (cell.ColumnIndex > 0 && cell.ColumnIndex <= 16384 && !processedColumns.Contains(cell.ColumnIndex))
                        {
                            writer.WriteStartElement("c");
                            writer.WriteAttributeString("r", GetCellReference(row.RowIndex, cell.ColumnIndex));

                            bool isDateCell = cell.Value is DateTime;
                            
                            // 数组公式：t="array" 且 ref 为范围
                            if (cell.IsArrayFormula && !string.IsNullOrEmpty(cell.ArrayRef))
                            {
                                writer.WriteAttributeString("t", "array");
                                writer.WriteAttributeString("ref", cell.ArrayRef);
                            }
                            // t 属性：日期单元格使用数值序列（不使用 t=\"d\"），避免 Sem_CellValue 错误
                            else if (!string.IsNullOrEmpty(cell.DataType) && !isDateCell)
                            {
                                writer.WriteAttributeString("t", cell.DataType);
                            }
                            
                            // 样式ID：优先用单元格的 StyleId，否则用行的默认 XF（如整行背景）
                            int? styleToUse = null;
                            if (!string.IsNullOrEmpty(cell.StyleId) && int.TryParse(cell.StyleId, out int styleIdVal))
                                styleToUse = styleIdVal;
                            else if (row.DefaultXfIndex.HasValue)
                                styleToUse = row.DefaultXfIndex.Value;
                            else if (isDateCell)
                                styleToUse = Math.Min(1, maxStyleId);
                            if (styleToUse.HasValue)
                                writer.WriteAttributeString("s", Math.Clamp(styleToUse.Value, 0, maxStyleId).ToString());
                            
                            // 处理公式（含数组公式）
                            if (cell.DataType == "f" || cell.IsArrayFormula)
                            {
                                writer.WriteStartElement("f");
                                writer.WriteString(cell.Formula ?? "");
                                writer.WriteEndElement();
                                writer.WriteStartElement("v");
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
                                if (isDateCell)
                                {
                                    // Excel 日期时间是从 1900-01-01 开始的天数，使用数值序列存储
                                    var dateTime = (DateTime)cell.Value!;
                                    double excelDate = DateTimeToExcelDate(dateTime);
                                    writer.WriteStartElement("v");
                                    writer.WriteString(excelDate.ToString());
                                    writer.WriteEndElement();
                                }
                                else if (cell.DataType == "s")
                                {
                                    // 共享字符串类型，写入索引（必须落在 [0, SharedStrings.Count-1] 否则会报错）
                                    var sharedList = workbook.SharedStrings ?? Enumerable.Empty<string>().ToList();
                                    var textValue = cell.Value?.ToString();
                                    writer.WriteStartElement("v");
                                    int ssIndex = 0;
                                    if (!string.IsNullOrEmpty(textValue))
                                    {
                                        int idx = sharedList.IndexOf(textValue);
                                        ssIndex = idx >= 0 ? Math.Min(idx, Math.Max(0, sharedList.Count - 1)) : 0;
                                    }
                                    writer.WriteString(ssIndex.ToString());
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
                        if (mergeCell == null) continue;
                        writer.WriteStartElement("mergeCell");
                        string refValue = GetCellReference(mergeCell.StartRow, mergeCell.StartColumn) + ":" + 
                                         GetCellReference(mergeCell.EndRow, mergeCell.EndColumn);
                        writer.WriteAttributeString("ref", refValue);
                        writer.WriteEndElement();
                    }
                    
                    writer.WriteEndElement();
                }
                
                // 自动筛选
                if (!string.IsNullOrEmpty(worksheet.AutoFilterRange))
                {
                    writer.WriteStartElement("autoFilter");
                    writer.WriteAttributeString("ref", worksheet.AutoFilterRange);
                    if (worksheet.AutoFilterColumnIndices != null && worksheet.AutoFilterColumnIndices.Count > 0)
                    {
                        foreach (int colId in worksheet.AutoFilterColumnIndices)
                        {
                            writer.WriteStartElement("filterColumn");
                            writer.WriteAttributeString("colId", colId.ToString());
                            writer.WriteStartElement("filters");
                            writer.WriteEndElement();
                            writer.WriteEndElement();
                        }
                    }
                    writer.WriteEndElement();
                }
                
                // 写入条件格式信息 (OOXML 要求 conditionalFormatting 在 dataValidations 之前)
                if (worksheet.ConditionalFormats != null && worksheet.ConditionalFormats.Count > 0)
                {
                    int cfPriority = 1; // 每个 cfRule 必须有唯一的 priority
                    foreach (var cf in worksheet.ConditionalFormats)
                    {
                        if (cf == null) continue;
                        if (!string.IsNullOrEmpty(cf.Range))
                        {
                            writer.WriteStartElement("conditionalFormatting");
                            writer.WriteAttributeString("sqref", SanitizeSqref(cf.Range));
                            
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
                            // between / notBetween 需要两个公式
                            if ((cf.Operator == "between" || cf.Operator == "notBetween") && !string.IsNullOrEmpty(cf.Formula2))
                            {
                                writer.WriteElementString("formula", cf.Formula2);
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
                
                // 写入数据验证信息 (必须在 conditionalFormatting 之后，pageMargins 之前)
                if (worksheet.DataValidations != null && worksheet.DataValidations.Count > 0)
                {
                    writer.WriteStartElement("dataValidations");
                    writer.WriteAttributeString("count", worksheet.DataValidations.Count.ToString());
                    
                    foreach (var dv in worksheet.DataValidations)
                    {
                        if (dv == null) continue;
                        writer.WriteStartElement("dataValidation");
                        if (!string.IsNullOrEmpty(dv.Range))
                        {
                            writer.WriteAttributeString("sqref", SanitizeSqref(dv.Range));
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
                
                // P3: pageMargins (页面边距) - 必须在 sheetData 之后
                if (ps != null)
                {
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
                }

                // 生成 drawing 引用 (仅当有图片数据或图表时；必须在 pageSetup/headerFooter 之后)
                bool hasDrawings = SheetHasDrawings(worksheet);
                if (hasDrawings)
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
        
        private void CreateWorksheetRelsXml(ZipArchive archive, Worksheet worksheet, int sheetIndex, Workbook workbook)
        {
            bool hasComments = worksheet.Comments != null && worksheet.Comments.Count > 0;
            bool hasDrawings = SheetHasDrawings(worksheet);
            bool hasHyperlinks = worksheet.Hyperlinks != null && worksheet.Hyperlinks.Count > 0;
            bool hasPivotTables = worksheet.PivotTables != null && worksheet.PivotTables.Count > 0;
            if (!hasComments && !hasDrawings && !hasHyperlinks && !hasPivotTables)
                return;

            var entry = archive.CreateEntry($"xl/worksheets/_rels/sheet{sheetIndex}.xml.rels");
            using (var stream = entry.Open())
            using (var writer = XmlWriter.Create(stream, Utf8NoBomXmlSettings))
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
                        if (hyperlink == null) continue;
                        writer.WriteStartElement("Relationship");
                        writer.WriteAttributeString("Id", "rIdHL" + (i + 1));
                        writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink");
                        writer.WriteAttributeString("Target", hyperlink.Target ?? "");
                        writer.WriteAttributeString("TargetMode", "External");
                        writer.WriteEndElement();
                    }
                }
                // 数据透视表关系
                if (hasPivotTables && workbook != null)
                {
                    int globalPivotIndex = 0;
                    for (int s = 0; s < sheetIndex - 1 && s < workbook.Worksheets.Count; s++)
                        globalPivotIndex += workbook.Worksheets[s].PivotTables?.Count ?? 0;
                    for (int i = 0; i < worksheet.PivotTables!.Count; i++)
                    {
                        writer.WriteStartElement("Relationship");
                        writer.WriteAttributeString("Id", "rIdPivot" + (i + 1));
                        writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable");
                        writer.WriteAttributeString("Target", "../pivotTables/pivotTable" + (globalPivotIndex + i + 1) + ".xml");
                        writer.WriteEndElement();
                    }
                }
                
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        private void CreatePivotParts(ZipArchive archive, Workbook workbook)
        {
            int total = GetTotalPivotTableCount(workbook);
            if (total == 0) return;
            int idx = 0;
            for (int s = 0; s < workbook.Worksheets.Count; s++)
            {
                var ws = workbook.Worksheets[s];
                if (ws.PivotTables == null) continue;
                foreach (var pt in ws.PivotTables)
                {
                    if (pt == null) continue;
                    idx++;
                    CreatePivotCacheDefinitionXml(archive, idx, pt, ws.Name ?? "Sheet" + (s + 1));
                    CreatePivotTableXml(archive, idx, pt);
                    CreatePivotTableRelsXml(archive, idx);
                }
            }
        }

        private static readonly string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        private void CreatePivotCacheDefinitionXml(ZipArchive archive, int cacheIndex, PivotTable pivotTable, string sheetName)
        {
            var entry = archive.CreateEntry($"xl/pivotCache/pivotCacheDefinition{cacheIndex}.xml");
            using (var stream = entry.Open())
            using (var writer = XmlWriter.Create(stream, Utf8NoBomXmlSettings))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("pivotCacheDefinition", MainNs);
                writer.WriteAttributeString("xmlns", "r", null, RelNs);
                writer.WriteAttributeString("r", "id", RelNs, "rId1");
                writer.WriteAttributeString("refreshedVersion", "3");
                writer.WriteAttributeString("createdVersion", "3");
                writer.WriteAttributeString("recordCount", "0");
                writer.WriteStartElement("cacheSource");
                writer.WriteAttributeString("type", "worksheet");
                writer.WriteStartElement("worksheetSource");
                writer.WriteAttributeString("ref", pivotTable.DataSource ?? "A1:Z100");
                writer.WriteAttributeString("sheet", sheetName);
                writer.WriteEndElement();
                writer.WriteEndElement();
                var allFields = (pivotTable.RowFields ?? new List<PivotField>())
                    .Concat(pivotTable.ColumnFields ?? new List<PivotField>())
                    .Concat(pivotTable.PageFields ?? new List<PivotField>())
                    .Concat(pivotTable.DataFields ?? new List<PivotField>())
                    .ToList();
                writer.WriteStartElement("cacheFields");
                writer.WriteAttributeString("count", allFields.Count.ToString());
                foreach (var f in allFields)
                {
                    writer.WriteStartElement("cacheField");
                    writer.WriteAttributeString("name", CleanXmlString(f?.Name ?? "Field"));
                    writer.WriteStartElement("sharedItems");
                    var items = f?.Items;
                    bool hasStr = items != null && items.Any(s => !string.IsNullOrEmpty(s));
                    writer.WriteAttributeString("containsString", hasStr ? "1" : "0");
                    writer.WriteAttributeString("containsNumber", "0");
                    if (hasStr && items != null && items.Count > 0)
                    {
                        writer.WriteAttributeString("count", items.Count.ToString());
                        foreach (var it in items)
                            writer.WriteElementString("s", CleanXmlString(it ?? ""));
                    }
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                }
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        private void CreatePivotTableXml(ZipArchive archive, int pivotIndex, PivotTable pivotTable)
        {
            var entry = archive.CreateEntry($"xl/pivotTables/pivotTable{pivotIndex}.xml");
            using (var stream = entry.Open())
            using (var writer = XmlWriter.Create(stream, Utf8NoBomXmlSettings))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("pivotTableDefinition", MainNs);
                writer.WriteAttributeString("xmlns", "r", null, RelNs);
                writer.WriteAttributeString("name", CleanXmlString(pivotTable.Name ?? "PivotTable" + pivotIndex));
                writer.WriteAttributeString("cacheId", pivotIndex.ToString());
                writer.WriteAttributeString("applyNumberFormats", "0");
                writer.WriteAttributeString("applyBorderFormats", "0");
                writer.WriteAttributeString("applyFontFormats", "0");
                writer.WriteAttributeString("applyPatternFormats", "0");
                writer.WriteAttributeString("applyAlignmentFormats", "0");
                writer.WriteAttributeString("applyWidthHeightFormats", "1");
                writer.WriteAttributeString("dataCaption", "Values");
                writer.WriteAttributeString("updatedVersion", "3");
                writer.WriteAttributeString("minRefreshableVersion", "3");
                writer.WriteAttributeString("useAutoFormatting", "1");
                writer.WriteAttributeString("itemPrintTitles", "1");
                writer.WriteAttributeString("createdVersion", "3");
                writer.WriteAttributeString("indent", "0");
                writer.WriteAttributeString("outline", "1");
                writer.WriteAttributeString("outlineData", "1");
                string locationRef = pivotTable.Range ?? pivotTable.DataSource ?? "A1:Z100";
                writer.WriteStartElement("location");
                writer.WriteAttributeString("ref", locationRef);
                writer.WriteAttributeString("firstHeaderRow", "1");
                writer.WriteAttributeString("firstDataRow", "1");
                writer.WriteAttributeString("firstDataCol", "0");
                writer.WriteEndElement();
                var allFields = (pivotTable.RowFields ?? new List<PivotField>())
                    .Concat(pivotTable.ColumnFields ?? new List<PivotField>())
                    .Concat(pivotTable.PageFields ?? new List<PivotField>())
                    .Concat(pivotTable.DataFields ?? new List<PivotField>())
                    .ToList();
                writer.WriteStartElement("pivotFields");
                writer.WriteAttributeString("count", allFields.Count.ToString());
                foreach (var f in allFields)
                {
                    bool isData = f?.Type == "data";
                    writer.WriteStartElement("pivotField");
                    if (isData)
                        writer.WriteAttributeString("dataField", "1");
                    writer.WriteAttributeString("showAll", "0");
                    writer.WriteAttributeString("sortType", (f?.SortType ?? "manual").ToLowerInvariant());
                    writer.WriteEndElement();
                }
                writer.WriteEndElement();
                writer.WriteStartElement("rowItems");
                writer.WriteAttributeString("count", "1");
                writer.WriteStartElement("i");
                writer.WriteAttributeString("t", "data");
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteStartElement("colItems");
                writer.WriteAttributeString("count", "1");
                writer.WriteStartElement("i");
                writer.WriteAttributeString("t", "data");
                writer.WriteEndElement();
                writer.WriteEndElement();
                int rowColPageCount = (pivotTable.RowFields?.Count ?? 0) + (pivotTable.ColumnFields?.Count ?? 0) + (pivotTable.PageFields?.Count ?? 0);
                int dataCount = pivotTable.DataFields?.Count ?? 0;
                writer.WriteStartElement("dataFields");
                writer.WriteAttributeString("count", dataCount.ToString());
                for (int i = 0; i < dataCount; i++)
                {
                    var df = pivotTable.DataFields![i];
                    string name = string.IsNullOrEmpty(df.Name) ? "Sum of " + (df.SourceName ?? "Value") : df.Name;
                    string subtotal = (df.Function ?? "sum").ToLowerInvariant();
                    writer.WriteStartElement("dataField");
                    writer.WriteAttributeString("name", CleanXmlString(name));
                    writer.WriteAttributeString("fld", (rowColPageCount + i).ToString());
                    writer.WriteAttributeString("subtotal", subtotal);
                    writer.WriteAttributeString("baseField", "0");
                    writer.WriteAttributeString("baseItem", "0");
                    writer.WriteAttributeString("numFmtId", "0");
                    writer.WriteEndElement();
                }
                writer.WriteEndElement();
                writer.WriteStartElement("tableStyle");
                writer.WriteAttributeString("name", "TableStyleLight16");
                writer.WriteAttributeString("showRowHeaders", "1");
                writer.WriteAttributeString("showColHeaders", "1");
                writer.WriteAttributeString("showRowStripes", "0");
                writer.WriteAttributeString("showColStripes", "0");
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        private void CreatePivotTableRelsXml(ZipArchive archive, int pivotIndex)
        {
            var entry = archive.CreateEntry($"xl/pivotTables/_rels/pivotTable{pivotIndex}.xml.rels");
            using (var stream = entry.Open())
            using (var writer = XmlWriter.Create(stream, Utf8NoBomXmlSettings))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
                writer.WriteStartElement("Relationship");
                writer.WriteAttributeString("Id", "rId1");
                writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition");
                writer.WriteAttributeString("Target", "../pivotCache/pivotCacheDefinition" + pivotIndex + ".xml");
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }
        
        private void CreateCommentsXml(ZipArchive archive, Worksheet worksheet, int sheetIndex)
        {
            // 创建VML绘图文件
            var vmlEntry = archive.CreateEntry($"xl/drawings/vmlDrawing{sheetIndex}.vml");
            using (var stream = vmlEntry.Open())
            using (var writer = XmlWriter.Create(stream, Utf8NoBomXmlSettings))
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

                // 定义批注使用的 shapetype（_x0000_t202），否则引用 type=\"#_x0000_t202\" 的 shape 在某些 Excel 版本中会被当作损坏的 drawing 修复/移除
                writer.WriteStartElement("shapetype", "urn:schemas-microsoft-com:vml");
                writer.WriteAttributeString("id", "_x0000_t202");
                writer.WriteAttributeString("coordsize", "21600,21600");
                writer.WriteAttributeString("o", "spt", "urn:schemas-microsoft-com:office:office", "202");
                writer.WriteAttributeString("path", "m,l,21600r21600,l21600,xe");
                writer.WriteStartElement("stroke", "urn:schemas-microsoft-com:vml");
                writer.WriteAttributeString("joinstyle", "miter");
                writer.WriteEndElement();
                writer.WriteStartElement("path", "urn:schemas-microsoft-com:vml");
                writer.WriteAttributeString("gradientshapeok", "t");
                writer.WriteAttributeString("o", "connecttype", "urn:schemas-microsoft-com:office:office", "rect");
                writer.WriteEndElement();
                writer.WriteEndElement();
                
                writer.WriteStartElement("shapes", "urn:schemas-microsoft-com:vml");
                writer.WriteAttributeString("ext", "urn:schemas-microsoft-com:vml", "edit");
                writer.WriteAttributeString("class", "x:WorksheetComments");
                
                // 为每个批注生成唯一的 shape id，避免所有批注共用 _x0000_s1025 导致 Excel 报错/修复 drawing
                int shapeIndex = 1025;
                foreach (var comment in worksheet.Comments)
                {
                    if (comment == null) continue;
                    string cellRef = GetCellReference(comment.RowIndex, comment.ColumnIndex);
                    
                    writer.WriteStartElement("shape", "urn:schemas-microsoft-com:vml");
                    writer.WriteAttributeString("id", "_x0000_s" + shapeIndex.ToString());
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
                    shapeIndex++;
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
                using (var writer = XmlWriter.Create(stream, Utf8NoBomXmlSettings))
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("styleSheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                    
                    // 数字格式：内置 0–163 由 Excel 定义，自定义从 164 起
                    writer.WriteStartElement("numFmts");
                    var customFormatIndices = (workbook.NumberFormats?.Keys ?? Enumerable.Empty<ushort>()).OrderBy(x => x).ToList();
                    int numFmtCount = 1 + customFormatIndices.Count;
                    writer.WriteAttributeString("count", numFmtCount.ToString());
                    writer.WriteStartElement("numFmt");
                    writer.WriteAttributeString("numFmtId", "164");
                    writer.WriteAttributeString("formatCode", "m/d/yyyy");
                    writer.WriteEndElement();
                    for (int i = 0; i < customFormatIndices.Count; i++)
                    {
                        ushort idx = customFormatIndices[i];
                        if (workbook.NumberFormats!.TryGetValue(idx, out string? formatCode) && !string.IsNullOrEmpty(formatCode))
                        {
                            writer.WriteStartElement("numFmt");
                            writer.WriteAttributeString("numFmtId", (165 + i).ToString());
                            writer.WriteAttributeString("formatCode", CleanXmlString(formatCode));
                            writer.WriteEndElement();
                        }
                    }
                    writer.WriteEndElement();
                    
                    // 字体
            writer.WriteStartElement("fonts");
            int fontCount = 0;
            List<Font> fonts = new List<Font>();
            
            // 优先添加从XLS文件解析的全局字体
            foreach (var font in workbook.Fonts ?? Enumerable.Empty<Font>())
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
            foreach (var style in workbook.Styles ?? Enumerable.Empty<Style>())
                    {
                        if (style == null) continue;
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
                    
                    writer.WriteAttributeString("count", fonts.Count.ToString());
                    
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
                        
                        // 字体颜色 (color)：OOXML rgb 为 6 或 8 位十六进制，不含 #
                        writer.WriteStartElement("color");
                        string fontRgb = (font.Color ?? "").Replace("#", "").Trim();
                        if (fontRgb.Length < 6) fontRgb = "000000";
                        writer.WriteAttributeString("rgb", fontRgb.Length <= 6 ? fontRgb : fontRgb.Substring(0, 6));
                        writer.WriteEndElement();
                        
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
                            writer.WriteAttributeString("rgb", fill.ForegroundColor.TrimStart('#'));
                            writer.WriteEndElement();
                        }
                        if (!string.IsNullOrEmpty(fill.BackgroundColor))
                        {
                            writer.WriteStartElement("bgColor");
                            writer.WriteAttributeString("rgb", fill.BackgroundColor.TrimStart('#'));
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
                    List<Xf> xfs = (workbook.XfList != null && workbook.XfList.Count > 0) ? workbook.XfList : new List<Xf> { new Xf() };
                    
                    Logger.Info($"写入 styles.xml: fonts={workbook.Fonts.Count}, fills={fills.Count}, borders={borders.Count}, xfs={xfs.Count}");
                    writer.WriteAttributeString("count", xfs.Count.ToString());
                    int maxFontId = Math.Max(0, fonts.Count - 1);
                    int maxFillId = Math.Max(0, fills.Count - 1);
                    int maxBorderId = Math.Max(0, borders.Count - 1);
                    int customFormatBaseId = 165;
                    foreach (var xf in xfs)
                    {
                        writer.WriteStartElement("xf");
                        int numFmtId;
                        if (xf.NumberFormatIndex > 0 && customFormatIndices.Count > 0)
                        {
                            int customIdx = customFormatIndices.IndexOf(xf.NumberFormatIndex);
                            if (customIdx >= 0)
                                numFmtId = customFormatBaseId + customIdx;
                            else
                                numFmtId = Math.Min((int)xf.NumberFormatIndex, 163);
                        }
                        else
                            numFmtId = xf.NumberFormatIndex > 0 ? Math.Min((int)xf.NumberFormatIndex, 163) : 0;
                        int fontId = xf.FontIndex > 0 ? Math.Min((int)xf.FontIndex, maxFontId) : 0;
                        int fillId = xf.FillIndex > 0 ? Math.Min(xf.FillIndex, maxFillId) : 0;
                        int borderId = xf.BorderIndex > 0 ? Math.Min(xf.BorderIndex, maxBorderId) : 0;
                        writer.WriteAttributeString("numFmtId", numFmtId.ToString());
                        writer.WriteAttributeString("fontId", fontId.ToString());
                        writer.WriteAttributeString("fillId", fillId.ToString());
                        writer.WriteAttributeString("borderId", borderId.ToString());
                        writer.WriteAttributeString("xfId", "0");
                        
                        if (numFmtId > 0)
                            writer.WriteAttributeString("applyNumberFormat", "1");
                        if (fontId > 0)
                            writer.WriteAttributeString("applyFont", "1");
                        if (fillId > 0)
                            writer.WriteAttributeString("applyFill", "1");
                        if (borderId > 0)
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
            var sharedStrings = new List<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            int stringCount = 0;

            foreach (var worksheet in workbook.Worksheets)
            {
                foreach (var row in worksheet.Rows ?? Enumerable.Empty<Row>())
                {
                    foreach (var cell in row.Cells ?? Enumerable.Empty<Cell>())
                    {
                        if (cell == null) continue;
                        if (cell.DataType != "s" && cell.DataType != "inlineStr" && (cell.DataType != null || !(cell.Value is string)))
                            continue;
                        var textValue = cell.Value?.ToString() ?? "";
                        if (!seen.Contains(textValue))
                        {
                            seen.Add(textValue);
                            sharedStrings.Add(textValue);
                        }
                        stringCount++;
                    }
                }
            }
            
            // 将共享字符串存储到workbook对象中，供后续使用
            workbook.SharedStrings = sharedStrings;
            
            var entry = archive.CreateEntry("xl/sharedStrings.xml");
            using (var stream = entry.Open())
            using (var writer = XmlWriter.Create(stream, Utf8NoBomXmlSettings))
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
            // Do not CreateEntry("xl/drawings/") or ("xl/media/") — directory-only entries are not OOXML parts
            // and cause "file corrupted" because they are not declared in [Content_Types].xml.
            int globalImageIndex = 1;
            int globalChartIndex = 1;
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                var worksheet = workbook.Worksheets[i];
                if (SheetHasDrawings(worksheet))
                {
                    int sheetStartImageIndex = globalImageIndex;
                    for (int j = 0; j < worksheet.Pictures.Count; j++)
                    {
                        var picture = worksheet.Pictures[j];
                        if (picture.Data != null)
                        {
                            try
                            {
                                string extension = picture.Extension ?? "bmp";
                                var entry = archive.CreateEntry($"xl/media/image{globalImageIndex}.{extension}");
                                using (var stream = entry.Open())
                                {
                                    stream.Write(picture.Data, 0, picture.Data.Length);
                                }
                                globalImageIndex++;
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
                    
                    int sheetStartChartIndex = globalChartIndex;
                    for (int j = 0; j < worksheet.Charts.Count; j++)
                    {
                        var chart = worksheet.Charts[j];
                        try
                        {
                            CreateChartXml(archive, chart, i + 1, globalChartIndex);
                            globalChartIndex++;
                        }
                        catch (Exception ex)
                        {
                            throw new ChartProcessingException($"创建图表XML时发生错误: {ex.Message}", ex);
                        }
                    }
                    
                    CreateDrawingRelsXml(archive, worksheet, i + 1, sheetStartImageIndex, sheetStartChartIndex);
                }
            }
        }
        
        private void CreateDrawingRelsXml(ZipArchive archive, Worksheet worksheet, int sheetIndex, int firstImageIndex, int firstChartIndex)
        {
            var entry = archive.CreateEntry($"xl/drawings/_rels/drawing{sheetIndex}.xml.rels");
            using (var stream = entry.Open())
            using (var writer = XmlWriter.Create(stream, Utf8NoBomXmlSettings))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("Relationships", "http://schemas.openxmlformats.org/package/2006/relationships");
                
                int imageOffset = 0;
                foreach (var picture in worksheet.Pictures ?? Enumerable.Empty<Picture>())
                {
                    if (picture?.Data == null || picture.Data.Length == 0) continue;
                    string extension = picture.Extension ?? "bmp";
                    writer.WriteStartElement("Relationship");
                    writer.WriteAttributeString("Id", "rId" + (imageOffset + 1));
                    writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
                    writer.WriteAttributeString("Target", $"../media/image{firstImageIndex + imageOffset}.{extension}");
                    writer.WriteEndElement();
                    imageOffset++;
                }
                
                for (int j = 0; j < (worksheet.Charts?.Count ?? 0); j++)
                {
                    writer.WriteStartElement("Relationship");
                    writer.WriteAttributeString("Id", "rIdChart" + (j + 1));
                    writer.WriteAttributeString("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart");
                    writer.WriteAttributeString("Target", $"../charts/chart{firstChartIndex + j}.xml");
                    writer.WriteEndElement();
                }
                
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }
        
        private const long EmuPerPixel = 9525L; // DrawingML: 1 pixel = 9525 EMU

        private void CreateDrawingXml(ZipArchive archive, Worksheet worksheet, int sheetIndex, Workbook workbook)
        {
            var entry = archive.CreateEntry($"xl/drawings/drawing{sheetIndex}.xml");
            string xdrNs = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
            using (var stream = entry.Open())
            using (var writer = XmlWriter.Create(stream, Utf8NoBomXmlSettings))
            {
                writer.WriteStartDocument();
                writer.WriteStartElement("wsDr", xdrNs);
                writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                writer.WriteAttributeString("xmlns", "a", null, "http://schemas.openxmlformats.org/drawingml/2006/main");
                
                int picIndex = 0;
                foreach (var picture in worksheet.Pictures ?? Enumerable.Empty<Picture>())
                {
                    if (picture?.Data == null || picture.Data.Length == 0) continue;
                    long leftEmu = (long)picture.Left * EmuPerPixel;
                    long topEmu = (long)picture.Top * EmuPerPixel;
                    long wEmu = (long)Math.Max(1, picture.Width) * EmuPerPixel;
                    long hEmu = (long)Math.Max(1, picture.Height) * EmuPerPixel;
                    long colWidthEmu = (long)(8.43 * 7 * EmuPerPixel);
                    long rowHeightEmu = (long)(15 * EmuPerPixel);
                    int col = (int)(leftEmu / colWidthEmu);
                    long colOff = leftEmu % colWidthEmu;
                    int row = (int)(topEmu / rowHeightEmu);
                    long rowOff = topEmu % rowHeightEmu;
                    int endCol = (int)((leftEmu + wEmu) / colWidthEmu);
                    long endColOff = (leftEmu + wEmu) % colWidthEmu;
                    int endRow = (int)((topEmu + hEmu) / rowHeightEmu);
                    long endRowOff = (topEmu + hEmu) % rowHeightEmu;
                    col = Math.Max(0, col);
                    row = Math.Max(0, row);
                    endCol = Math.Max(col + 1, endCol);
                    endRow = Math.Max(row + 1, endRow);
                    
                    writer.WriteStartElement("twoCellAnchor", xdrNs);
                    writer.WriteAttributeString("editAs", "twoCell");
                    writer.WriteStartElement("from", xdrNs);
                    writer.WriteElementString("col", xdrNs, col.ToString());
                    writer.WriteElementString("colOff", xdrNs, colOff.ToString());
                    writer.WriteElementString("row", xdrNs, row.ToString());
                    writer.WriteElementString("rowOff", xdrNs, rowOff.ToString());
                    writer.WriteEndElement();
                    writer.WriteStartElement("to", xdrNs);
                    writer.WriteElementString("col", xdrNs, endCol.ToString());
                    writer.WriteElementString("colOff", xdrNs, endColOff.ToString());
                    writer.WriteElementString("row", xdrNs, endRow.ToString());
                    writer.WriteElementString("rowOff", xdrNs, endRowOff.ToString());
                    writer.WriteEndElement();
                    
                    writer.WriteStartElement("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");
                    writer.WriteStartElement("nvPicPr");
                    writer.WriteStartElement("cNvPr");
                    writer.WriteAttributeString("id", (picIndex + 1).ToString());
                    writer.WriteAttributeString("name", $"Picture {picIndex + 1}");
                    writer.WriteEndElement();
                    writer.WriteStartElement("cNvPicPr");
                    writer.WriteStartElement("picLocks");
                    writer.WriteAttributeString("noChangeAspect", "1");
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteStartElement("blipFill");
                    writer.WriteStartElement("blip", "http://schemas.openxmlformats.org/drawingml/2006/main");
                    writer.WriteAttributeString("r", "embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", $"rId{picIndex + 1}");
                    writer.WriteEndElement();
                    writer.WriteStartElement("stretch", "http://schemas.openxmlformats.org/drawingml/2006/main");
                    writer.WriteStartElement("fillRect", "http://schemas.openxmlformats.org/drawingml/2006/main");
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteStartElement("spPr");
                    writer.WriteStartElement("xfrm");
                    writer.WriteStartElement("off", "http://schemas.openxmlformats.org/drawingml/2006/main");
                    writer.WriteAttributeString("x", "0");
                    writer.WriteAttributeString("y", "0");
                    writer.WriteEndElement();
                    writer.WriteStartElement("ext", "http://schemas.openxmlformats.org/drawingml/2006/main");
                    writer.WriteAttributeString("cx", wEmu.ToString());
                    writer.WriteAttributeString("cy", hEmu.ToString());
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteStartElement("prstGeom", "http://schemas.openxmlformats.org/drawingml/2006/main");
                    writer.WriteAttributeString("prst", "rect");
                    writer.WriteStartElement("avLst", "http://schemas.openxmlformats.org/drawingml/2006/main");
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    writer.WriteStartElement("clientData", xdrNs);
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                    picIndex++;
                }
                
                int pictureCountWithData = picIndex;
                for (int i = 0; i < (worksheet.Charts?.Count ?? 0); i++)
                {
                    var chart = worksheet.Charts[i];
                    long leftEmu = (long)chart.Left * EmuPerPixel;
                    long topEmu = (long)chart.Top * EmuPerPixel;
                    long wEmu = (long)Math.Max(1, chart.Width) * EmuPerPixel;
                    long hEmu = (long)Math.Max(1, chart.Height) * EmuPerPixel;
                    long colWidthEmu = (long)(8.43 * 7 * EmuPerPixel);
                    long rowHeightEmu = (long)(15 * EmuPerPixel);
                    int col = Math.Max(0, (int)(leftEmu / colWidthEmu));
                    long colOff = leftEmu % colWidthEmu;
                    int row = Math.Max(0, (int)(topEmu / rowHeightEmu));
                    long rowOff = topEmu % rowHeightEmu;
                    int endCol = Math.Max(col + 1, (int)((leftEmu + wEmu) / colWidthEmu));
                    long endColOff = (leftEmu + wEmu) % colWidthEmu;
                    int endRow = Math.Max(row + 1, (int)((topEmu + hEmu) / rowHeightEmu));
                    long endRowOff = (topEmu + hEmu) % rowHeightEmu;
                    
                    writer.WriteStartElement("twoCellAnchor", xdrNs);
                    writer.WriteAttributeString("editAs", "twoCell");
                    writer.WriteStartElement("from", xdrNs);
                    writer.WriteElementString("col", xdrNs, col.ToString());
                    writer.WriteElementString("colOff", xdrNs, colOff.ToString());
                    writer.WriteElementString("row", xdrNs, row.ToString());
                    writer.WriteElementString("rowOff", xdrNs, rowOff.ToString());
                    writer.WriteEndElement();
                    writer.WriteStartElement("to", xdrNs);
                    writer.WriteElementString("col", xdrNs, endCol.ToString());
                    writer.WriteElementString("colOff", xdrNs, endColOff.ToString());
                    writer.WriteElementString("row", xdrNs, endRow.ToString());
                    writer.WriteElementString("rowOff", xdrNs, endRowOff.ToString());
                    writer.WriteEndElement();
                    
                    writer.WriteStartElement("graphicFrame");
                    writer.WriteStartElement("nvGraphicFramePr");
                    writer.WriteStartElement("cNvPr");
                    writer.WriteAttributeString("id", (pictureCountWithData + i + 1).ToString());
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
                    writer.WriteAttributeString("cx", wEmu.ToString());
                    writer.WriteAttributeString("cy", hEmu.ToString());
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
                    writer.WriteStartElement("clientData", xdrNs);
                    writer.WriteEndElement();
                    writer.WriteEndElement();
                }
                
                // Drawing XML must contain only anchors (pic/graphicFrame). Do not write hyperlinks or legacyDrawing here —
                // those belong in the worksheet XML and cause "Removed Part: Drawing shape" corruption.
                
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
        }

        private string GetChartElementType(Chart chart)
        {
            string chartType = chart.ChartType ?? "barChart";
            bool is3D = chart.Is3D || string.Equals(chartType, "surfaceChart", StringComparison.OrdinalIgnoreCase);

            // 映射图表类型到OpenXML元素名称
            switch (chartType)
            {
                case "barChart": return is3D ? "bar3DChart" : "barChart";
                case "colChart": return is3D ? "bar3DChart" : "colChart";
                case "lineChart": return is3D ? "line3DChart" : "lineChart";
                case "pieChart": return is3D ? "pie3DChart" : "pieChart";
                case "scatterChart": return "scatterChart";
                case "areaChart": return is3D ? "area3DChart" : "areaChart";
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
                using (var writer = XmlWriter.Create(stream, Utf8NoBomXmlSettings))
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
                    
                    // OOXML chart structure: chart > (title?) > view3D? > plotArea > (layout, chartType, catAx, valAx) > legend?
                    
                    // 3D 视图（如果由 CHART3D 标记为 3D）
                    if (chart.Is3D)
                    {
                        writer.WriteStartElement("view3D");
                        writer.WriteStartElement("rotX");
                        writer.WriteAttributeString("val", "20");
                        writer.WriteEndElement();
                        writer.WriteStartElement("rotY");
                        writer.WriteAttributeString("val", "20");
                        writer.WriteEndElement();
                        writer.WriteStartElement("perspective");
                        writer.WriteAttributeString("val", "30");
                        writer.WriteEndElement();
                        writer.WriteEndElement();
                    }
                    
                    // plotArea 包含图表类型和坐标轴
                    writer.WriteStartElement("plotArea");
                    writer.WriteStartElement("layout");
                    writer.WriteEndElement(); // layout
                    
                    // 写入图表类型
                    string chartElementType = GetChartElementType(chart);
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
        
        /// <summary>将 BIFF 内置名称映射为 OOXML 保留名称，便于 Excel 识别打印区域等。</summary>
        private static string GetOoxmlBuiltInName(string name)
        {
            return name switch
            {
                "Print_Area" => "_xlnm.Print_Area",
                "Print_Titles" => "_xlnm.Print_Titles",
                "Consolidate_Area" => "_xlnm.Consolidate_Area",
                "Database" => "_xlnm.Database",
                "Criteria" => "_xlnm.Criteria",
                "FilterDatabase" => "_xlnm._FilterDatabase",
                "Sheet_Title" => "_xlnm.Sheet_Title",
                _ => name
            };
        }
        
        private string GetCellReference(int rowIndex, int columnIndex)
        {
            rowIndex = Math.Clamp(rowIndex, 1, 1048576);
            columnIndex = Math.Clamp(columnIndex, 1, 16384);
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
            
            // Excel 将 1900 当作闰年：序列 60 = 1900-02-29（虚构），61 = 1900-03-01。.NET 中 1900-03-01 距 1900-01-01 为 59 天，故需 +2 得 61。
            if (dateTime >= new DateTime(1900, 3, 1))
                days += 2;

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