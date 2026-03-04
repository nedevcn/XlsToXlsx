using System.IO;
using System.Collections.Generic;
using System;
using System.Linq;
using Nedev.XlsToXlsx;
using Nedev.XlsToXlsx.Exceptions;

namespace Nedev.XlsToXlsx.Formats.Xls
{
    public class XlsParser
    {
        private Stream _stream;
        private BinaryReader _reader;
        private List<string> _sharedStrings;
        private List<Font> _fonts = new List<Font>();
        private List<Xf> _xfList = new List<Xf>();
        private Dictionary<ushort, string> _formats = new Dictionary<ushort, string>();
        private Dictionary<int, string> _palette = new Dictionary<int, string>();
        private Workbook _workbook;
        private const long MAX_FILE_SIZE = 100 * 1024 * 1024; // 100MB文件大小限制

        public XlsParser(Stream stream)
        {
            // 验证流是否可读
            if (!stream.CanRead)
            {
                throw new XlsToXlsxException("Stream must be readable", 1000, "StreamError");
            }

            // 检查文件大小限制
            if (stream.CanSeek)
            {
                long fileSize = stream.Length;
                if (fileSize > MAX_FILE_SIZE)
                {
                    throw new XlsToXlsxException($"File size exceeds limit of {MAX_FILE_SIZE / (1024 * 1024)}MB", 1002, "FileSizeError");
                }
            }

            // 使用BufferedStream提高读取速度
            _stream = new BufferedStream(stream);
            _reader = new BinaryReader(_stream);
            _sharedStrings = new List<string>();
        }

        public Workbook Parse()
        {
            _workbook = new Workbook();
            var workbook = _workbook;

            try
            {
                Logger.Info("开始解析XLS文件");

                // 解析XLS文件头
                ParseHeader();
                Logger.Info("文件头解析完成");

                // 解析Compound File Binary格式
                ParseCompoundFile();
                Logger.Info("Compound File格式解析完成");

                // 解析Workbook流
                ParseWorkbookStream(workbook);
                Logger.Info("Workbook流解析完成");

                // 解析Worksheet流
                ParseWorksheetStreams(workbook);
                Logger.Info("Worksheet流解析完成");

                // 将解析到的全局数据转移到工作簿对象
                workbook.SharedStrings = _sharedStrings;
                workbook.Fonts = _fonts;
                workbook.XfList = _xfList;
                workbook.NumberFormats = _formats;
                workbook.Palette = _palette;

                // 解析VBA流
                ParseVbaStream(workbook);
                Logger.Info("VBA流解析完成");

                Logger.Info("XLS文件解析成功");
                return workbook;
            }
            catch (XlsToXlsxException)
            {
                throw;
            }
            catch (System.IO.InvalidDataException ex)
            {
                Logger.Error("解析XLS文件时发生数据格式错误", ex);
                throw new XlsParseException($"解析XLS文件时发生数据格式错误: {ex.Message}", ex);
            }
            catch (System.IO.IOException ex)
            {
                Logger.Error("解析XLS文件时发生IO错误", ex);
                throw new XlsParseException($"解析XLS文件时发生IO错误: {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                Logger.Error("解析XLS文件时发生未知错误", ex);
                throw new XlsParseException($"解析XLS文件时发生未知错误: {ex.Message}", ex);
            }
        }

        public async Task<Workbook> ParseAsync()
        {
            // 使用Task.Run在后台线程中执行解析，避免阻塞主线程
            return await Task.Run(() => Parse());
        }

        private void ParseHeader()
        {
            // 验证文件头
            var header = _reader.ReadBytes(8);
            // 检查是否读取到足够的数据
            if (header.Length < 8)
            {
                throw new System.IO.InvalidDataException("Not a valid XLS file: insufficient data for file header");
            }
            // 检查是否为BIFF8格式
            if (header[0] != 0xD0 || header[1] != 0xCF || header[2] != 0x11 || header[3] != 0xE0 ||
                header[4] != 0xA1 || header[5] != 0xB1 || header[6] != 0x1A || header[7] != 0xE1)
            {
                throw new System.IO.InvalidDataException("Not a valid XLS file: invalid file header signature");
            }
        }

        private int _sectorSize;
        
        private void ParseCompoundFile()
        {
            // 读取Compound File Binary格式的头部信息
            _stream.Seek(0x00, SeekOrigin.Begin);
            var header = _reader.ReadBytes(512);
            
            // 验证文件头
            if (header[0] != 0xD0 || header[1] != 0xCF || header[2] != 0x11 || header[3] != 0xE0 ||
                header[4] != 0xA1 || header[5] != 0xB1 || header[6] != 0x1A || header[7] != 0xE1)
            {
                throw new System.IO.InvalidDataException("Not a valid Compound File Binary format");
            }
            
            // 读取DIFAT和FAT相关信息
            _sectorSize = 1 << (header[0x0C] & 0xFF); // 扇区大小
            int directorySectorStart = BitConverter.ToInt32(header, 0x30); // 目录区起始扇区
            
            // 定位到目录区
            _stream.Seek(directorySectorStart * _sectorSize, SeekOrigin.Begin);
            
            // 解析目录条目
            ParseDirectoryEntries();
        }
        
        private void ParseDirectoryEntries()
        {
            // 每个目录条目大小为128字节
            byte[] entryBuffer = new byte[128];
            
            // 持续解析直到遇到空条目或无法读取完整条目
            int entryCount = 0;
            const int MAX_ENTRY_COUNT = 1000; // 增加最大条目数限制
            
            // 存储所有流类型的条目，用于备用查找
            List<WorksheetStreamInfo> allStreams = new List<WorksheetStreamInfo>();
            
            while (entryCount < MAX_ENTRY_COUNT)
            {
                int bytesRead = _reader.Read(entryBuffer, 0, 128);
                if (bytesRead < 128)
                    break;
                
                // 检查是否为空条目（名称全为0）
                bool isEmpty = true;
                for (int j = 0; j < 64; j++)
                {
                    if (entryBuffer[j] != 0)
                    {
                        isEmpty = false;
                        break;
                    }
                }
                if (isEmpty)
                    break;
                
                // 读取条目名称（UTF-16LE编码）
                string name = System.Text.Encoding.Unicode.GetString(entryBuffer, 0, 64).TrimEnd('\0');
                // 清理名称，移除非打印字符
                name = new string(name.Where(c => char.IsLetterOrDigit(c) || char.IsPunctuation(c) || char.IsWhiteSpace(c)).ToArray());
                
                // 读取条目类型
                byte entryType = entryBuffer[0x40];
                
                // 读取子类型
                byte entrySubType = entryBuffer[0x41];
                
                // 读取起始扇区（无符号整数）
                uint startSectorUInt = BitConverter.ToUInt32(entryBuffer, 0x44);
                int startSector = (int)startSectorUInt;
                
                // 读取流大小（无符号整数）
                ulong streamSizeUInt = BitConverter.ToUInt64(entryBuffer, 0x48);
                long streamSize = (long)streamSizeUInt;
                
                // 合理性检查
                bool isValidStream = true;
                // 只检查起始扇区是否为负数，不检查大小，因为损坏的文件可能会有异常大的流大小值
                if (startSector < 0)
                {
                    isValidStream = false;
                }
                // 对于损坏的文件，我们需要尽可能尝试恢复数据，所以不设置流大小限制
                // 即使流大小看起来不合理，我们也尝试解析它
                
                // 处理流类型和存储类型的条目
                if (entryType == 1 || entryType == 2) // 1 = Storage, 2 = Stream
                {
                    // 记录所有找到的条目
                    Logger.Info($"找到条目: 名称='{name}', 类型={entryType}, 起始扇区={startSector}, 大小={streamSize}");
                    
                    // 检查是否为Workbook
                    if (name == "Workbook" && isValidStream)
                    {
                        // 记录Workbook的位置
                        _workbookStreamStart = startSector;
                        _workbookStreamSize = streamSize;
                        Logger.Info($"找到Workbook流: 起始扇区={startSector}, 大小={streamSize}");
                    }
                    
                    // 检查是否为Worksheet
                    if ((name.StartsWith("Sheet") || name.StartsWith("Worksheet")) && isValidStream)
                    {
                        // 记录Worksheet的位置
                        _worksheetStreams.Add(new WorksheetStreamInfo
                        {
                            Name = name,
                            StartSector = startSector,
                            Size = streamSize
                        });
                        Logger.Info($"找到Worksheet流: 名称='{name}', 起始扇区={startSector}, 大小={streamSize}");
                    }
                    
                    // 检查是否为其他可能的工作表名称
                    if (name.Length > 0 && char.IsDigit(name[0]) && isValidStream)
                    {
                        // 记录Worksheet的位置
                        _worksheetStreams.Add(new WorksheetStreamInfo
                        {
                            Name = name,
                            StartSector = startSector,
                            Size = streamSize
                        });
                        Logger.Info($"找到Worksheet流: 名称='{name}', 起始扇区={startSector}, 大小={streamSize}");
                    }
                    
                    // 检查是否为VBA
                    if ((name == "VBA" || name.StartsWith("VBA/")) && isValidStream)
                    {
                        // 记录VBA的位置
                        _vbaStreamStart = startSector;
                        _vbaStreamSize = streamSize;
                        Logger.Info($"找到VBA流: 起始扇区={startSector}, 大小={streamSize}");
                    }
                    
                    // 记录所有流类型的条目，用于备用查找
                    if (entryType == 2) // 只记录流类型
                    {
                        // 过滤掉起始扇区为负数的流，以及大小不合理的流
                        if (isValidStream)
                        {
                            allStreams.Add(new WorksheetStreamInfo
                            {
                                Name = name,
                                StartSector = startSector,
                                Size = streamSize
                            });
                        }
                    }
                }
                
                entryCount++;
            }
            
            // 如果没有找到Workbook流，尝试从所有流中找到最可能的Workbook流
            if (_workbookStreamStart == 0 && _workbookStreamSize == 0 && allStreams.Count > 0)
            {
                Logger.Info("未找到命名为Workbook的流，尝试从所有流中查找可能的Workbook流");
                
                // 找到大小合理的流（Workbook流通常在1KB到10MB之间）
                var possibleWorkbookStreams = allStreams.Where(s => s.Size > 1024 && s.Size < 10 * 1024 * 1024).ToList();
                
                if (possibleWorkbookStreams.Count > 0)
                {
                    // 按大小排序，选择最大的一个作为Workbook流
                    var largestStream = possibleWorkbookStreams.OrderByDescending(s => s.Size).First();
                    _workbookStreamStart = largestStream.StartSector;
                    _workbookStreamSize = largestStream.Size;
                    Logger.Info($"选择最大的流作为Workbook流: 起始扇区={largestStream.StartSector}, 大小={largestStream.Size}");
                    
                    // 从剩余的流中查找可能的工作表流
                    var possibleWorksheetStreams = possibleWorkbookStreams.Where(s => s != largestStream).ToList();
                    foreach (var stream in possibleWorksheetStreams)
                    {
                        _worksheetStreams.Add(stream);
                        Logger.Info($"添加可能的工作表流: 名称='{stream.Name}', 起始扇区={stream.StartSector}, 大小={stream.Size}");
                    }
                }
                else
                {
                    // 尝试从所有流中找到可能的Workbook流，即使大小不合理
                    Logger.Info("尝试从所有流中找到可能的Workbook流，即使大小不合理");
                    
                    // 从所有流中选择一个作为Workbook流
                    if (allStreams.Count > 0)
                    {
                        // 按起始扇区排序，选择起始扇区较小的流作为Workbook流
                        var sortedStreams = allStreams.OrderBy(s => s.StartSector).ToList();
                        var selectedStream = sortedStreams.First();
                        _workbookStreamStart = selectedStream.StartSector;
                        _workbookStreamSize = selectedStream.Size;
                        Logger.Info($"选择起始扇区最小的流作为Workbook流: 起始扇区={selectedStream.StartSector}, 大小={selectedStream.Size}");
                        
                        // 从剩余的流中查找可能的工作表流
                        var possibleWorksheetStreams = sortedStreams.Where(s => s != selectedStream).ToList();
                        foreach (var stream in possibleWorksheetStreams)
                        {
                            _worksheetStreams.Add(stream);
                            Logger.Info($"添加可能的工作表流: 名称='{stream.Name}', 起始扇区={stream.StartSector}, 大小={stream.Size}");
                        }
                    }
                }
            }
            
            // 确保至少有一个工作表
            if (_worksheetStreams.Count == 0 && allStreams.Count > 0)
            {
                Logger.Info("未找到工作表流，尝试从所有流中查找可能的工作表流");
                
                // 从所有流中选择一些作为工作表流
                foreach (var stream in allStreams.Take(3)) // 最多添加3个工作表
                {
                    _worksheetStreams.Add(stream);
                    Logger.Info($"添加可能的工作表流: 名称='{stream.Name}', 起始扇区={stream.StartSector}, 大小={stream.Size}");
                }
            }
        }
        
        private long _workbookStreamStart;
        private long _workbookStreamSize;
        private List<WorksheetStreamInfo> _worksheetStreams = new List<WorksheetStreamInfo>();
        private long _vbaStreamStart;
        private long _vbaStreamSize;
        
        /// <summary>
        /// VBA项目大小限制（字节）
        /// </summary>
        public long VbaSizeLimit { get; set; } = 50 * 1024 * 1024;
        private string _currentCFRange = string.Empty; // 当前条件格式的范围
        
        private class WorksheetStreamInfo
        {
            public string? Name { get; set; }
            public int StartSector { get; set; }
            public long Size { get; set; }
        }

        private void ParseWorkbookStream(Workbook workbook)
        {
            // 定位到Workbook流
            _stream.Seek(_workbookStreamStart * _sectorSize, SeekOrigin.Begin);

            // 读取BIFF记录
            long workbookStreamEnd = _workbookStreamStart * _sectorSize + _workbookStreamSize;
            while (_stream.Position < workbookStreamEnd && _stream.Position >= 0 && workbookStreamEnd >= 0)
                {
                    try
                    {
                        var record = BiffRecord.Read(_reader);
                        
                        switch (record.Id)
                        {
                            case (ushort)BiffRecordType.BOF:
                                // 开始记录
                                break;
                            case (ushort)BiffRecordType.EOF:
                                // 结束记录
                                return;
                            case (ushort)BiffRecordType.SHEET:
                                // 工作表定义
                                ParseSheetRecord(record, workbook);
                                break;
                            case (ushort)BiffRecordType.SST:
                                // 共享字符串表（可能包含 CONTINUE 记录）
                                ParseSstInfo(record, workbookStreamEnd);
                                break;
                            case (ushort)BiffRecordType.FONT:
                                ParseFontRecordToGlobal(record);
                                break;
                            case (ushort)BiffRecordType.XF:
                                ParseXfRecordToGlobal(record);
                                break;
                            case (ushort)BiffRecordType.FORMAT:
                                ParseFormatRecordGlobal(record);
                                break;
                            case (ushort)BiffRecordType.PALETTE:
                                ParsePaletteRecordGlobal(record);
                                break;
                            case (ushort)BiffRecordType.NAME:
                                ParseNameRecord(record, workbook);
                                break;
                            default:
                                // 跳过其他记录
                                break;
                        }
                    }
                    catch (EndOfStreamException)
                    {
                        // 遇到流结束，正常退出循环
                        break;
                    }
                    catch (XlsToXlsxException)
                    {
                        // 重新抛出自定义异常
                        throw;
                    }
                    catch (Exception ex)
                    {
                        // 记录详细的错误信息
                        Logger.Error($"解析Workbook流时发生错误: {ex.Message}", ex);
                        // 继续处理下一条记录
                        continue;
                    }

                }
        }

        private object _streamLock = new object();
        private const int CACHE_SIZE = 65536; // 64KB缓存
        private const int MAX_ROW_CELLS = 16384; // 最大行单元格数
        
        private void ParseWorksheetStreams(Workbook workbook)
        {
            // 为每个工作表创建Worksheet对象
            
            // 确保工作表数量与流数量匹配
            while (workbook.Worksheets.Count < _worksheetStreams.Count)
            {
                var worksheet = new Worksheet();
                worksheet.Name = "Sheet" + (workbook.Worksheets.Count + 1);
                workbook.Worksheets.Add(worksheet);
            }
            
            // 为每个流设置工作表名称
            for (int i = 0; i < _worksheetStreams.Count; i++)
            {
                if (i < workbook.Worksheets.Count)
                {
                    var streamInfo = _worksheetStreams[i];
                    if (!string.IsNullOrEmpty(streamInfo.Name))
                    {
                        // 尝试从流名称中提取有意义的工作表名称
                        string name = streamInfo.Name.Trim();
                        if (!string.IsNullOrEmpty(name))
                        {
                            // 移除特殊字符和控制字符
                            name = new string(name.Where(c => char.IsLetterOrDigit(c) || c == ' ' || c == '_').ToArray());
                            if (!string.IsNullOrEmpty(name))
                            {
                                workbook.Worksheets[i].Name = name.Length > 31 ? name.Substring(0, 31) : name;
                            }
                            else
                            {
                                workbook.Worksheets[i].Name = "Sheet" + (i + 1);
                            }
                        }
                        else
                        {
                            workbook.Worksheets[i].Name = "Sheet" + (i + 1);
                        }
                    }
                    else
                    {
                        workbook.Worksheets[i].Name = "Sheet" + (i + 1);
                    }
                }
            }
            
            // 尝试直接从文件中搜索BIFF记录，不依赖于目录条目的信息
            Logger.Info("尝试直接从文件中搜索BIFF记录");
            
            // 限制搜索范围，避免处理整个文件
            long searchLimit = 500 * 1024 * 1024; // 500MB
            long fileSize = _stream.Length;
            long searchEnd = Math.Min(searchLimit, fileSize);
            
            // 第一遍扫描: 找到工作簿全局流(BOF type 0x0005)来解析BOUNDSHEET/SST等全局记录
            // 第二遍扫描: 找到工作表流(BOF type 0x0010)来解析数据
            long position = 0;
            bool foundGlobals = false;
            
            // 第一遍: 查找并解析工作簿全局流
            while (position < searchEnd)
            {
                try
                {
                    position = SearchForBofRecord(position, searchEnd);
                    if (position < 0)
                        break;
                    
                    // 读取BOF记录获取子流类型
                    ushort subStreamType = 0;
                    lock (_streamLock)
                    {
                        _stream.Seek(position, SeekOrigin.Begin);
                        var bofRecord = BiffRecord.Read(_reader);
                        if (bofRecord.Data != null && bofRecord.Data.Length >= 4)
                        {
                            subStreamType = BitConverter.ToUInt16(bofRecord.Data, 2);
                        }
                    }
                    
                    if (subStreamType == 0x0005) // Workbook Globals
                    {
                        Logger.Info($"在位置 {position} 找到工作簿全局流");
                        foundGlobals = true;
                        
                        // 清除之前从_worksheetStreams创建的默认工作表
                        workbook.Worksheets.Clear();
                        
                        // 解析全局记录 (BOUNDSHEET, SST, FONT, XF, etc.)
                        lock (_streamLock)
                        {
                            _stream.Seek(position, SeekOrigin.Begin);
                            BiffRecord.Read(_reader); // skip BOF
                            long globalsEnd = searchEnd;
                            while (_stream.Position < globalsEnd)
                            {
                                try
                                {
                                    var record = BiffRecord.Read(_reader);
                                    switch (record.Id)
                                    {
                                        case (ushort)BiffRecordType.EOF:
                                            goto doneGlobals;
                                        case (ushort)BiffRecordType.SHEET:
                                            ParseSheetRecord(record, workbook);
                                            break;
                                        case (ushort)BiffRecordType.SST:
                                            ParseSstInfo(record, globalsEnd);
                                            break;
                                        case (ushort)BiffRecordType.FONT:
                                            ParseFontRecordToGlobal(record);
                                            break;
                                        case (ushort)BiffRecordType.XF:
                                            ParseXfRecordToGlobal(record);
                                            break;
                                        case (ushort)BiffRecordType.FORMAT:
                                            ParseFormatRecordGlobal(record);
                                            break;
                                        case (ushort)BiffRecordType.PALETTE:
                                            ParsePaletteRecordGlobal(record);
                                            break;
                                        case (ushort)BiffRecordType.NAME:
                                            ParseNameRecord(record, workbook);
                                            break;
                                    }
                                }
                                catch (EndOfStreamException)
                                {
                                    break;
                                }
                            }
                        }
                        doneGlobals:
                        Logger.Info($"工作簿全局流解析完成，找到 {workbook.Worksheets.Count} 个工作表定义");
                        break; // 只需要找第一个全局流
                    }
                    
                    // 不是全局流，跳到下一个位置继续搜索
                    position += 4;
                }
                catch (Exception ex)
                {
                    Logger.Error($"搜索工作簿全局流时发生错误", ex);
                    position += 1024;
                }
            }
            
            // 第二遍: 查找并解析工作表子流 (BOF type 0x0010)
            position = 0;
            int worksheetIndex = 0;
            
            while (position < searchEnd && worksheetIndex < Math.Max(workbook.Worksheets.Count, 3))
            {
                try
                {
                    position = SearchForBofRecord(position, searchEnd);
                    if (position < 0)
                        break;
                    
                    // 读取BOF记录获取子流类型
                    ushort subStreamType = 0;
                    lock (_streamLock)
                    {
                        _stream.Seek(position, SeekOrigin.Begin);
                        var bofRecord = BiffRecord.Read(_reader);
                        if (bofRecord.Data != null && bofRecord.Data.Length >= 4)
                        {
                            subStreamType = BitConverter.ToUInt16(bofRecord.Data, 2);
                        }
                    }
                    
                    if (subStreamType != 0x0010) // 不是工作表子流
                    {
                        // 跳过非工作表BOF（全局流、图表等）
                        position += 4;
                        continue;
                    }
                    
                    // 确保有对应的工作表对象
                    if (worksheetIndex >= workbook.Worksheets.Count)
                    {
                        var newWorksheet = new Worksheet();
                        newWorksheet.Name = "Sheet" + (workbook.Worksheets.Count + 1);
                        workbook.Worksheets.Add(newWorksheet);
                    }
                    
                    var worksheet = workbook.Worksheets[worksheetIndex];
                    Logger.Info($"在位置 {position} 找到工作表子流 {worksheet.Name}");
                    
                    // 解析从当前位置开始的记录
                    long endPosition = ParseWorksheetFromPosition(worksheet, position, searchEnd);
                    if (endPosition > position)
                    {
                        position = endPosition;
                        worksheetIndex++;
                        Logger.Info($"工作表 {worksheet.Name} 解析完成，共 {worksheet.Rows.Count} 行");
                    }
                    else
                    {
                        position += 1024;
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error($"搜索BIFF记录时发生错误", ex);
                    position += 1024;
                }
            }
            
            // 如果没有找到任何工作表数据，尝试从目录条目指定的流中解析
            if (worksheetIndex == 0)
            {
                Logger.Info("未找到BIFF记录，尝试从目录条目指定的流中解析");
                
                // 串行处理工作表流，避免并发问题
                for (int i = 0; i < _worksheetStreams.Count; i++)
                {
                    var streamInfo = _worksheetStreams[i];
                    // 确保索引不越界
                    if (i < workbook.Worksheets.Count)
                    {
                        var worksheet = workbook.Worksheets[i];
                    
                    // 直接使用原始流的位置，避免一次性加载整个工作表数据到内存
                    long startPosition = streamInfo.StartSector * _sectorSize;
                    
                    // 限制流大小，避免处理过大的流导致内存问题
                    long maxStreamSize = 100 * 1024 * 1024; // 100MB
                    long streamSize = Math.Min(streamInfo.Size, maxStreamSize);
                    long endPosition = startPosition + streamSize;
                    
                    // 读取BIFF记录
                    var currentRow = new Row();
                    currentRow.Cells.Capacity = 100; // 预分配容量
                    long currentPosition = startPosition;
                    
                    // 缓存机制
                    byte[] buffer = new byte[CACHE_SIZE];
                    int bufferLength = 0;
                    int bufferPosition = 0;
                    
                    Logger.Info($"开始解析工作表 {worksheet.Name}，起始位置: {startPosition}, 大小: {streamSize}");
                    
                    while (currentPosition < endPosition)
                    {
                        long remainingBytes = endPosition - currentPosition;
                        try
                        {
                            // 计算剩余数据量
                            if (remainingBytes < 4) // 至少需要4字节来读取记录ID和长度
                                break;
                            
                            // 读取记录，使用锁确保并发安全
                            BiffRecord record;
                            
                            // 检查缓存是否足够
                            if (bufferPosition + 4 > bufferLength)
                            {
                                // 填充缓存
                                lock (_streamLock)
                                {
                                    _stream.Seek(currentPosition, SeekOrigin.Begin);
                                    bufferLength = _reader.Read(buffer, 0, Math.Min(CACHE_SIZE, (int)remainingBytes));
                                    bufferPosition = 0;
                                }
                            }
                            
                            // 从缓存中读取记录ID和长度
                            ushort recordId = BitConverter.ToUInt16(buffer, bufferPosition);
                            ushort recordLength = BitConverter.ToUInt16(buffer, bufferPosition + 2);
                            bufferPosition += 4;
                            
                            // 检查缓存是否足够读取整个记录
                            if (bufferPosition + recordLength > bufferLength)
                            {
                                // 直接从流中读取完整记录
                                lock (_streamLock)
                                {
                                    _stream.Seek(currentPosition, SeekOrigin.Begin);
                                    record = BiffRecord.Read(_reader);
                                    currentPosition = _stream.Position;
                                }
                            }
                            else
                            {
                                // 从缓存中读取记录数据
                                byte[] recordData = new byte[recordLength];
                                Array.Copy(buffer, bufferPosition, recordData, 0, recordLength);
                                bufferPosition += recordLength;
                                record = new BiffRecord();
                                record.Id = recordId;
                                record.Length = recordLength;
                                record.Data = recordData;
                                currentPosition += 4 + recordLength;
                            }
                            
                            switch (record.Id)
                            {
                                case (ushort)BiffRecordType.BOF:
                                    // 开始记录
                                    break;
                                case (ushort)BiffRecordType.SST:
                                    // 共享字符串表
                                    ParseSstInfo(record, searchEnd);
                                    break;
                                case (ushort)BiffRecordType.EOF:
                                    // 结束记录
                                    if (currentRow.Cells.Count > 0)
                                    {
                                        worksheet.Rows.Add(currentRow);
                                        if (currentRow.RowIndex > worksheet.MaxRow) worksheet.MaxRow = (int)currentRow.RowIndex;
                                    }
                                    goto nextWorksheet;
                                case (ushort)BiffRecordType.ROW:
                                    // 行记录
                                    var parsedRow = ParseRowRecord(record);
                                    var existingRow = GetOrCreateRow(worksheet, ref currentRow, parsedRow.RowIndex);
                                    existingRow.Height = parsedRow.Height;
                                    existingRow.CustomHeight = parsedRow.CustomHeight;
                                    if (existingRow.RowIndex > worksheet.MaxRow) worksheet.MaxRow = (int)existingRow.RowIndex;
                                    break;
                                case (ushort)BiffRecordType.CELL_BLANK:
                                case (ushort)BiffRecordType.CELL_BOOLERR:
                                case (ushort)BiffRecordType.CELL_LABEL:
                                case (ushort)BiffRecordType.CELL_LABELSST:
                                case (ushort)BiffRecordType.CELL_NUMBER:
                                case (ushort)BiffRecordType.CELL_RK:
                                    // 单元格记录
                                    var cell = ParseCellRecord(record);
                                    if (cell.ColumnIndex > 0 && cell.ColumnIndex <= 16384)
                                    {
                                        var targetRow = GetOrCreateRow(worksheet, ref currentRow, cell.RowIndex);
                                        targetRow.Cells.Add(cell);
                                        if (cell.ColumnIndex > worksheet.MaxColumn) worksheet.MaxColumn = cell.ColumnIndex;
                                        if (targetRow.Cells.Count > MAX_ROW_CELLS) Logger.Warn($"行 {targetRow.RowIndex} 单元格数超过限制，可能导致内存问题");
                                    }
                                    break;
                                case (ushort)BiffRecordType.CELL_FORMULA:
                                    // 公式记录
                                    var formulaCell = ParseCellRecord(record);
                                    if (formulaCell.ColumnIndex > 0 && formulaCell.ColumnIndex <= 16384)
                                    {
                                        var targetRow = GetOrCreateRow(worksheet, ref currentRow, formulaCell.RowIndex);
                                        targetRow.Cells.Add(formulaCell);
                                        if (formulaCell.ColumnIndex > worksheet.MaxColumn) worksheet.MaxColumn = formulaCell.ColumnIndex;
                                    }
                                    break;
                                case (ushort)BiffRecordType.STRING:
                                    // 公式产生的字符串结果记录
                                    if (currentRow != null && currentRow.Cells.Count > 0)
                                    {
                                        var lastCell = currentRow.Cells[currentRow.Cells.Count - 1];
                                        if (record.Data != null && record.Data.Length > 0)
                                        {
                                            int strOffset = 0;
                                            lastCell.Value = ReadBiffString(record.Data, ref strOffset);
                                            lastCell.DataType = "inlineStr";
                                        }
                                    }
                                    break;
                                case (ushort)BiffRecordType.MULRK:
                                    // 多值RK记录
                                    ParseMulRkRecord(record, ref currentRow, worksheet);
                                    break;
                                case (ushort)BiffRecordType.MULBLANK:
                                    // 多空白单元格记录
                                    ParseMulBlankRecord(record, ref currentRow, worksheet);
                                    break;
                                case (ushort)BiffRecordType.MERGECELLS:
                                    // 合并单元格记录
                                    ParseMergeCellsRecord(record, worksheet);
                                    break;
                                case (ushort)BiffRecordType.PALETTE:
                                    // 调色板记录已经通过全局解析处理，此处跳过
                                    break;
                                case (ushort)BiffRecordType.CHART:
                                    // 图表记录
                                    ParseChartRecord(record, worksheet);
                                    break;
                                case (ushort)BiffRecordType.CHARTTITLE:
                                    // 图表标题记录
                                    ParseChartTitleRecord(record, worksheet);
                                    break;
                                case (ushort)BiffRecordType.SERIES:
                                    // 数据系列记录
                                    ParseSeriesRecord(record, worksheet);
                                    break;
                                case (ushort)BiffRecordType.MSODRAWING:
                                    // 图片和绘图对象
                                    ParseMSODrawingRecord(record, worksheet);
                                    break;
                                case (ushort)BiffRecordType.PICTURE:
                                    // 图片记录
                                    ParsePictureRecord(record, worksheet);
                                    break;
                                case (ushort)BiffRecordType.OBJ:
                                    // 嵌入对象记录
                                    ParseObjRecord(record, worksheet);
                                    break;
                                case (ushort)BiffRecordType.DV:
                                    // 数据验证记录
                                    ParseDVRecord(record, worksheet);
                                    break;
                                case (ushort)BiffRecordType.CF:
                                    // 条件格式记录
                                    ParseCFRecord(record, worksheet);
                                    break;
                                case (ushort)BiffRecordType.CFHEADER:
                                    // 条件格式头部记录
                                    ParseCFHeaderRecord(record, worksheet);
                                    break;
                                case (ushort)BiffRecordType.HYPERLINK:
                                    // 超链接记录
                                    ParseHyperlinkRecord(record, worksheet);
                                    break;
                                case (ushort)BiffRecordType.COLINFO:
                                    // 列宽信息
                                    ParseColInfoRecord(record, worksheet);
                                    break;
                                case (ushort)BiffRecordType.DEFCOLWIDTH:
                                    // 默认列宽
                                    if (record.Data != null && record.Data.Length >= 2)
                                        worksheet.DefaultColumnWidth = BitConverter.ToUInt16(record.Data, 0);
                                    break;
                                case (ushort)BiffRecordType.DEFAULTROWHEIGHT:
                                    // 默认行高
                                    if (record.Data != null && record.Data.Length >= 4)
                                        worksheet.DefaultRowHeight = BitConverter.ToUInt16(record.Data, 2) / 20.0;
                                    break;
                                case (ushort)BiffRecordType.NOTE:
                                    // 注释记录 (如果需要保留)
                                    ParseCommentRecord(record, worksheet);
                                    break;
                                case (ushort)BiffRecordType.HEADER:
                                    if (record.Data != null && record.Data.Length >= 1)
                                    {
                                        int pos = 0;
                                        ushort len = record.Data[0];
                                        pos = 1;
                                        worksheet.PageSettings.Header = ReadBiffStringFromBytes(record.Data, ref pos, len);
                                    }
                                    break;
                                case (ushort)BiffRecordType.FOOTER:
                                    if (record.Data != null && record.Data.Length >= 1)
                                    {
                                        int pos = 0;
                                        ushort len = record.Data[0];
                                        pos = 1;
                                        worksheet.PageSettings.Footer = ReadBiffStringFromBytes(record.Data, ref pos, len);
                                    }
                                    break;
                                case (ushort)BiffRecordType.LEFTMARGIN:
                                    if (record.Data != null && record.Data.Length >= 8)
                                        worksheet.PageSettings.LeftMargin = BitConverter.ToDouble(record.Data, 0);
                                    break;
                                case (ushort)BiffRecordType.RIGHTMARGIN:
                                    if (record.Data != null && record.Data.Length >= 8)
                                        worksheet.PageSettings.RightMargin = BitConverter.ToDouble(record.Data, 0);
                                    break;
                                case (ushort)BiffRecordType.TOPMARGIN:
                                    if (record.Data != null && record.Data.Length >= 8)
                                        worksheet.PageSettings.TopMargin = BitConverter.ToDouble(record.Data, 0);
                                    break;
                                case (ushort)BiffRecordType.BOTTOMMARGIN:
                                    if (record.Data != null && record.Data.Length >= 8)
                                        worksheet.PageSettings.BottomMargin = BitConverter.ToDouble(record.Data, 0);
                                    break;
                                case (ushort)BiffRecordType.HCENTER:
                                    if (record.Data != null && record.Data.Length >= 2)
                                        worksheet.PageSettings.HorizontalCenter = BitConverter.ToUInt16(record.Data, 0) != 0;
                                    break;
                                case (ushort)BiffRecordType.VCENTER:
                                    if (record.Data != null && record.Data.Length >= 2)
                                        worksheet.PageSettings.VerticalCenter = BitConverter.ToUInt16(record.Data, 0) != 0;
                                    break;
                                case (ushort)BiffRecordType.PAGESETUP:
                                    ParsePageSetupRecord(record, worksheet);
                                    break;
                                case (ushort)BiffRecordType.DIMENSION:
                                    // 工作表范围 (BIFF8: rwMic(4), rwMac(4), colMic(2), colMac(2))
                                    if (record.Data != null && record.Data.Length >= 12)
                                    {
                                        worksheet.MaxRow = BitConverter.ToInt32(record.Data, 4);
                                        worksheet.MaxColumn = BitConverter.ToUInt16(record.Data, 10);
                                    }
                                    break;
                                case (ushort)BiffRecordType.WINDOW2:
                                    // 工作表窗口设置（包含冻结窗格标志）
                                    ParseWindow2Record(record, worksheet);
                                    break;
                                case 0x0041: // PANE 记录
                                    ParsePaneRecord(record, worksheet);
                                    break;
                                default:
                                    // 跳过其他记录
                                    break;
                            }
                        }
                        catch (EndOfStreamException)
                        {
                            // 遇到流结束，正常退出循环
                            break;
                        }
                        catch (Exception ex)
                        {
                            // 记录详细的错误信息
                            Logger.Error($"解析工作表 {worksheet.Name} 在位置 {currentPosition} 时发生错误", ex);
                            // 继续处理下一条记录，而不是中断整个工作表的解析
                            // 计算下一条记录的位置
                            try
                            {
                                lock (_streamLock)
                                {
                                    // 尝试跳过当前损坏的记录
                                    _stream.Seek(currentPosition, SeekOrigin.Begin);
                                    // 读取记录长度
                                    if (remainingBytes >= 2)
                                    {
                                        short recordLength = _reader.ReadInt16();
                                        // 跳过记录数据
                                        currentPosition += 4 + recordLength; // 2字节ID + 2字节长度 + 数据长度
                                    }
                                    else
                                    {
                                        // 无法读取记录长度，直接退出
                                        break;
                                    }
                                }
                            }
                            catch
                            {
                                // 如果无法跳过，直接退出
                                break;
                        }
                        }
                    } // Ends while (currentPosition < endPosition)
                    
                    nextWorksheet:
                        // 处理完一个工作表
                        Logger.Info($"解析工作表 {worksheet.Name} 完成，共 {worksheet.Rows.Count} 行");
                    }
                }
            }
        }

        
        /// <summary>
        /// 搜索BOF记录
        /// </summary>
        /// <param name="startPosition">开始位置</param>
        /// <param name="endPosition">结束位置</param>
        /// <returns>找到的BOF记录位置，未找到返回-1</returns>
        private long SearchForBofRecord(long startPosition, long endPosition)
        {
            long position = startPosition;
            byte[] buffer = new byte[4096];
            
            while (position < endPosition)
            {
                try
                {
                    lock (_streamLock)
                    {
                        _stream.Seek(position, SeekOrigin.Begin);
                        int bytesRead = _reader.Read(buffer, 0, Math.Min(buffer.Length, (int)(endPosition - position)));
                        if (bytesRead < 4)
                            break;
                        
                        // 搜索BOF记录（0x0009）
                        for (int i = 0; i < bytesRead - 3; i++)
                        {
                            ushort recordId = BitConverter.ToUInt16(buffer, i);
                            if (recordId == (ushort)BiffRecordType.BOF)
                            {
                                return position + i;
                            }
                        }
                    }
                    
                    position += buffer.Length - 3;
                }
                catch (Exception ex)
                {
                    Logger.Error($"搜索BOF记录时发生错误", ex);
                    position += 1024;
                }
            }
            
            return -1;
        }
        
        private Row GetOrCreateRow(Worksheet worksheet, ref Row currentRow, int rowIndex)
    {
        if (currentRow != null && currentRow.RowIndex == rowIndex)
            return currentRow;
            
        // 倒序查找，因为行通常是顺序添加的，从后往前找最快
        for (int i = worksheet.Rows.Count - 1; i >= 0; i--)
        {
            if (worksheet.Rows[i].RowIndex == rowIndex)
            {
                currentRow = worksheet.Rows[i];
                return currentRow;
            }
        }
        
        // 找不到则创建新行
        var newRow = new Row { RowIndex = rowIndex };
        newRow.Cells.Capacity = 20;
        worksheet.Rows.Add(newRow);
        currentRow = newRow;
        return newRow;
    }

    /// <summary>
    /// 从指定位置解析工作表
        /// </summary>
        /// <param name="worksheet">工作表对象</param>
        /// <param name="startPosition">开始位置</param>
        /// <param name="endPosition">结束位置</param>
        /// <returns>解析结束的位置</returns>
        private long ParseWorksheetFromPosition(Worksheet worksheet, long startPosition, long endPosition)
        {
            long currentPosition = startPosition;
            var currentRow = (Row)null; // Initialize to null, GetOrCreateRow will manage it
            
            // 缓存机制
            byte[] buffer = new byte[CACHE_SIZE];
            int bufferLength = 0;
            int bufferPosition = 0;
            
            while (currentPosition < endPosition)
            {
                long remainingBytes = endPosition - currentPosition;
                try
                {
                    // 计算剩余数据量
                    if (remainingBytes < 4) // 至少需要4字节来读取记录ID和长度
                        break;
                    
                    // 读取记录，使用锁确保并发安全
                    BiffRecord record;
                    
                    // 检查缓存是否足够
                    if (bufferPosition + 4 > bufferLength)
                    {
                        // 填充缓存
                        lock (_streamLock)
                        {
                            _stream.Seek(currentPosition, SeekOrigin.Begin);
                            bufferLength = _reader.Read(buffer, 0, Math.Min(CACHE_SIZE, (int)remainingBytes));
                            bufferPosition = 0;
                        }
                    }
                    
                    // 从缓存中读取记录ID和长度
                    ushort recordId = BitConverter.ToUInt16(buffer, bufferPosition);
                    ushort recordLength = BitConverter.ToUInt16(buffer, bufferPosition + 2);
                    bufferPosition += 4;
                    
                    // 检查缓存是否足够读取整个记录
                    if (bufferPosition + recordLength > bufferLength)
                    {
                        // 直接从流中读取完整记录
                        lock (_streamLock)
                        {
                            _stream.Seek(currentPosition, SeekOrigin.Begin);
                            record = BiffRecord.Read(_reader);
                            currentPosition = _stream.Position;
                        }
                    }
                    else
                    {
                        // 从缓存中读取记录数据
                        byte[] recordData = new byte[recordLength];
                        Array.Copy(buffer, bufferPosition, recordData, 0, recordLength);
                        bufferPosition += recordLength;
                        record = new BiffRecord();
                        record.Id = recordId;
                        record.Length = recordLength;
                        record.Data = recordData;
                        currentPosition += 4 + recordLength;
                    }
                    
                    switch (record.Id)
                    {
                        case (ushort)BiffRecordType.EOF:
                            // 结束记录
                            return currentPosition;
                        case (ushort)BiffRecordType.SST:
                            // 共享字符串表
                            ParseSstInfo(record, endPosition);
                            break;
                    // 行记录
                    var parsedRow = ParseRowRecord(record);
                    var existingRow = GetOrCreateRow(worksheet, ref currentRow, parsedRow.RowIndex);
                    // ROW记录提供了行高信息，更新当前行属性
                    existingRow.Height = parsedRow.Height;
                    existingRow.CustomHeight = parsedRow.CustomHeight;
                    if (existingRow.RowIndex > worksheet.MaxRow) worksheet.MaxRow = (int)existingRow.RowIndex;
                    break;
                case (ushort)BiffRecordType.CELL_BLANK:
                case (ushort)BiffRecordType.CELL_BOOLERR:
                case (ushort)BiffRecordType.CELL_LABEL:
                case (ushort)BiffRecordType.CELL_LABELSST:
                case (ushort)BiffRecordType.CELL_NUMBER:
                case (ushort)BiffRecordType.CELL_RK:
                    // 单元格记录
                    var cell = ParseCellRecord(record);
                    // 确保列索引有效
                    if (cell.ColumnIndex >= 1 && cell.ColumnIndex <= 16384) // Excel最大列数
                    {
                        var targetRow = GetOrCreateRow(worksheet, ref currentRow, cell.RowIndex);
                        targetRow.Cells.Add(cell);
                        if (cell.ColumnIndex > worksheet.MaxColumn)
                        {
                            worksheet.MaxColumn = cell.ColumnIndex;
                        }
                        if (cell.RowIndex > worksheet.MaxRow)
                        {
                            worksheet.MaxRow = cell.RowIndex;
                        }
                        // 防止行单元格过多导致内存溢出
                        if (targetRow.Cells.Count > MAX_ROW_CELLS)
                        {
                            Logger.Warn($"行 {targetRow.RowIndex} 单元格数超过限制，可能导致内存问题");
                        }
                    }
                    break;
                case (ushort)BiffRecordType.CELL_FORMULA:
                    // 公式记录
                    var formulaCell = ParseCellRecord(record);
                    // 确保列索引有效
                    if (formulaCell.ColumnIndex >= 1 && formulaCell.ColumnIndex <= 16384) // Excel最大列数
                    {
                        var targetRow = GetOrCreateRow(worksheet, ref currentRow, formulaCell.RowIndex);
                        targetRow.Cells.Add(formulaCell);
                        if (formulaCell.ColumnIndex > worksheet.MaxColumn)
                        {
                            worksheet.MaxColumn = formulaCell.ColumnIndex;
                        }
                        if (formulaCell.RowIndex > worksheet.MaxRow)
                        {
                            worksheet.MaxRow = formulaCell.RowIndex;
                        }
                    }
                    break;
                case (ushort)BiffRecordType.STRING:
                    // 公式产生的字符串结果记录
                    if (currentRow != null && currentRow.Cells.Count > 0)
                    {
                        var lastCell = currentRow.Cells[currentRow.Cells.Count - 1];
                        if (record.Data != null && record.Data.Length > 0)
                        {
                            int strOffset = 0;
                            lastCell.Value = ReadBiffString(record.Data, ref strOffset);
                            lastCell.DataType = "inlineStr";
                        }
                    }
                    break;
                case (ushort)BiffRecordType.MULRK:
                    // 多值RK记录
                    ParseMulRkRecord(record, ref currentRow, worksheet);
                    break;
                case (ushort)BiffRecordType.MULBLANK:
                    // 多空白单元格记录
                    ParseMulBlankRecord(record, ref currentRow, worksheet);
                    break;
                        case (ushort)BiffRecordType.MERGECELLS:
                            // 合并单元格记录
                            ParseMergeCellsRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.FORMAT:
                            // 格式记录
                            ParseFormatRecord(record);
                            break;
                        case (ushort)BiffRecordType.PALETTE:
                            // 全局记录已在Workbook流中处理
                            break;
                        case (ushort)BiffRecordType.CHART:
                            // 图表记录
                            ParseChartRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.CHARTTITLE:
                            // 图表标题记录
                            ParseChartTitleRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.SERIES:
                            // 数据系列记录
                            ParseSeriesRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.MSODRAWING:
                            // 图片和绘图对象
                            ParseMSODrawingRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.PICTURE:
                            // 图片记录
                            ParsePictureRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.OBJ:
                            // 嵌入对象记录
                            ParseObjRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.DV:
                            // 数据验证记录
                            ParseDVRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.CF:
                            // 条件格式记录
                            ParseCFRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.CFHEADER:
                            // 条件格式头部记录
                            ParseCFHeaderRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.HYPERLINK:
                            // 超链接记录
                            ParseHyperlinkRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.COLINFO:
                            // 列宽信息
                            ParseColInfoRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.DEFCOLWIDTH:
                            // 默认列宽
                            if (record.Data != null && record.Data.Length >= 2)
                                worksheet.DefaultColumnWidth = BitConverter.ToUInt16(record.Data, 0);
                            break;
                        case (ushort)BiffRecordType.DEFAULTROWHEIGHT:
                            // 默认行高
                            if (record.Data != null && record.Data.Length >= 4)
                                worksheet.DefaultRowHeight = BitConverter.ToUInt16(record.Data, 2) / 20.0;
                            break;
                        case (ushort)BiffRecordType.DIMENSION:
                            // 工作表范围
                            if (record.Data != null && record.Data.Length >= 12)
                            {
                                worksheet.MaxRow = BitConverter.ToInt32(record.Data, 4);
                                worksheet.MaxColumn = BitConverter.ToUInt16(record.Data, 10);
                            }
                            break;
                        case (ushort)BiffRecordType.WINDOW2:
                            // 工作表窗口设置（包含冻结窗格标志）
                            ParseWindow2Record(record, worksheet);
                            break;
                        case 0x0041: // PANE 记录
                            ParsePaneRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.NOTE:
                            // 注释记录 (如果需要保留)
                            ParseCommentRecord(record, worksheet);
                            break;
                        case (ushort)BiffRecordType.HEADER:
                            if (record.Data != null && record.Data.Length >= 1)
                            {
                                int pos = 0;
                                ushort len = record.Data[0];
                                pos = 1;
                                worksheet.PageSettings.Header = ReadBiffStringFromBytes(record.Data, ref pos, len);
                            }
                            break;
                        case (ushort)BiffRecordType.FOOTER:
                            if (record.Data != null && record.Data.Length >= 1)
                            {
                                int pos = 0;
                                ushort len = record.Data[0];
                                pos = 1;
                                worksheet.PageSettings.Footer = ReadBiffStringFromBytes(record.Data, ref pos, len);
                            }
                            break;
                        case (ushort)BiffRecordType.LEFTMARGIN:
                            if (record.Data != null && record.Data.Length >= 8)
                                worksheet.PageSettings.LeftMargin = BitConverter.ToDouble(record.Data, 0);
                            break;
                        case (ushort)BiffRecordType.RIGHTMARGIN:
                            if (record.Data != null && record.Data.Length >= 8)
                                worksheet.PageSettings.RightMargin = BitConverter.ToDouble(record.Data, 0);
                            break;
                        case (ushort)BiffRecordType.TOPMARGIN:
                            if (record.Data != null && record.Data.Length >= 8)
                                worksheet.PageSettings.TopMargin = BitConverter.ToDouble(record.Data, 0);
                            break;
                        case (ushort)BiffRecordType.BOTTOMMARGIN:
                            if (record.Data != null && record.Data.Length >= 8)
                                worksheet.PageSettings.BottomMargin = BitConverter.ToDouble(record.Data, 0);
                            break;
                        case (ushort)BiffRecordType.HCENTER:
                            if (record.Data != null && record.Data.Length >= 2)
                                worksheet.PageSettings.HorizontalCenter = BitConverter.ToUInt16(record.Data, 0) != 0;
                            break;
                        case (ushort)BiffRecordType.VCENTER:
                            if (record.Data != null && record.Data.Length >= 2)
                                worksheet.PageSettings.VerticalCenter = BitConverter.ToUInt16(record.Data, 0) != 0;
                            break;
                        case (ushort)BiffRecordType.PAGESETUP:
                            ParsePageSetupRecord(record, worksheet);
                            break;
                        default:
                            // 跳过其他记录
                            break;
                    }
                }
                catch (EndOfStreamException)
                {
                    // 遇到流结束，正常退出循环
                    break;
                }
                catch (Exception ex)
                {
                    // 记录详细的错误信息
                    Logger.Error($"解析工作表在位置 {currentPosition} 时发生错误", ex);
                    // 继续处理下一条记录，而不是中断整个工作表的解析
                    // 计算下一条记录的位置
                    try
                    {
                        lock (_streamLock)
                        {
                            // 尝试跳过当前损坏的记录
                            _stream.Seek(currentPosition, SeekOrigin.Begin);
                            // 读取记录长度
                            if (remainingBytes >= 2)
                            {
                                short recordLength = _reader.ReadInt16();
                                // 跳过记录数据
                                currentPosition += 4 + recordLength; // 2字节ID + 2字节长度 + 数据长度
                            }
                            else
                            {
                                // 无法读取记录长度，直接退出
                                break;
                            }
                        }
                    }
                    catch
                    {
                        // 如果无法跳过，直接退出
                        break;
                    }
                }
            }
            
            return currentPosition;
        }
        
        private void ParseFormatRecord(BiffRecord record)
        {
            // 解析格式记录
            if (record.Data != null && record.Data.Length >= 18)
            {
                // 读取格式索引
                ushort formatIndex = BitConverter.ToUInt16(record.Data, 0);
                
                // 读取格式字符串长度
                byte formatLength = record.Data[2];
                
                // 读取格式字符串
                if (record.Data.Length >= 3 + formatLength)
                {
                    string formatString = System.Text.Encoding.ASCII.GetString(record.Data, 3, formatLength);
                    _formats[formatIndex] = formatString;
                }
            }
        }
        
        private void ParseMulRkRecord(BiffRecord record, ref Row currentRow, Worksheet worksheet)
{
    // MULRK 记录：一条记录包含多个连续单元格的 RK 值
    // 格式: row(2) + firstCol(2) + [xfIndex(2) + rkValue(4)] * N + lastCol(2)
    if (record.Data != null && record.Data.Length >= 6)
    {
        ushort row = BitConverter.ToUInt16(record.Data, 0);
        ushort firstCol = BitConverter.ToUInt16(record.Data, 2);
        // 最后2字节是lastCol
        int dataLength = record.Data.Length - 6; // 去掉row(2) + firstCol(2) + lastCol(2)
        int numCells = dataLength / 6; // 每个单元格占6字节: xfIndex(2) + rkValue(4)
        
        var targetRow = GetOrCreateRow(worksheet, ref currentRow, row + 1);
        
        for (int j = 0; j < numCells; j++)
        {
            int offset = 4 + j * 6;
            if (offset + 6 > record.Data.Length - 2) break; // 边界检查
            
            ushort xfIndex = BitConverter.ToUInt16(record.Data, offset);
            int rkValue = BitConverter.ToInt32(record.Data, offset + 2);
            double value = DecodeRKValue(rkValue);
            
            var cell = new Cell();
            cell.RowIndex = row + 1; // 转为1-based
            cell.ColumnIndex = firstCol + j + 1; // 转为1-based
            cell.Value = value;
            cell.DataType = "n";
            cell.StyleId = xfIndex.ToString();
            targetRow.Cells.Add(cell);
            
            if (cell.ColumnIndex > worksheet.MaxColumn)
            {
                worksheet.MaxColumn = cell.ColumnIndex;
            }
        }
    }
}
        
        private void ParseMulBlankRecord(BiffRecord record, ref Row currentRow, Worksheet worksheet)
{
    // MULBLANK 记录：一条记录包含多个连续空白单元格的样式信息
    // 格式: row(2) + firstCol(2) + [xfIndex(2)] * N + lastCol(2)
    if (record.Data != null && record.Data.Length >= 6)
    {
        ushort row = BitConverter.ToUInt16(record.Data, 0);
        ushort firstCol = BitConverter.ToUInt16(record.Data, 2);
        int dataLength = record.Data.Length - 6; // 去掉row(2) + firstCol(2) + lastCol(2)
        int numCells = dataLength / 2; // 每个单元格占2字节: xfIndex(2)
        
        var targetRow = GetOrCreateRow(worksheet, ref currentRow, row + 1);
        
        for (int j = 0; j < numCells; j++)
        {
            int offset = 4 + j * 2;
            if (offset + 2 > record.Data.Length - 2) break;
            
            ushort xfIndex = BitConverter.ToUInt16(record.Data, offset);
            
            var cell = new Cell();
            cell.RowIndex = row + 1; // 转为1-based
            cell.ColumnIndex = firstCol + j + 1; // 转为1-based
            cell.Value = null;
            cell.StyleId = xfIndex.ToString();
            targetRow.Cells.Add(cell);
            
            if (cell.ColumnIndex > worksheet.MaxColumn)
            {
                worksheet.MaxColumn = cell.ColumnIndex;
            }
        }
    }
}
        
        private void ParseColInfoRecord(BiffRecord record, Worksheet worksheet)
        {
            // COLINFO 记录格式: firstCol(2) + lastCol(2) + width(2) + xfIndex(2) + options(2) + reserved(2)
            if (record.Data != null && record.Data.Length >= 10)
            {
                ushort firstCol = BitConverter.ToUInt16(record.Data, 0);
                ushort lastCol = BitConverter.ToUInt16(record.Data, 2);
                ushort width = BitConverter.ToUInt16(record.Data, 4);
                ushort xfIndex = BitConverter.ToUInt16(record.Data, 6);
                ushort options = BitConverter.ToUInt16(record.Data, 8);
                bool hidden = (options & 0x0001) != 0;
                
                var colInfo = new ColumnInfo
                {
                    FirstColumn = firstCol,
                    LastColumn = lastCol,
                    Width = width,
                    XfIndex = xfIndex,
                    Hidden = hidden
                };
                worksheet.ColumnInfos.Add(colInfo);
            }
        }
        
        private void ParseWindow2Record(BiffRecord record, Worksheet worksheet)
        {
            // WINDOW2 记录: options(2) + ...
            // options bit 3: 是否冻结窗格 (fFrozen)
            if (record.Data != null && record.Data.Length >= 2)
            {
                ushort options = BitConverter.ToUInt16(record.Data, 0);
                bool isFrozen = (options & 0x0008) != 0;
                
                if (isFrozen && worksheet.FreezePane == null)
                {
                    // 标记为冻结，具体位置由PANE记录设置
                    worksheet.FreezePane = new FreezePane();
                }
            }
        }
        
        private void ParsePaneRecord(BiffRecord record, Worksheet worksheet)
        {
            // PANE 记录: x(2) + y(2) + topRow(2) + leftCol(2) + activePane(1)
            if (record.Data != null && record.Data.Length >= 8)
            {
                ushort x = BitConverter.ToUInt16(record.Data, 0); // 水平分割位置（列数或像素）
                ushort y = BitConverter.ToUInt16(record.Data, 2); // 垂直分割位置（行数或像素）
                ushort topRow = BitConverter.ToUInt16(record.Data, 4);
                ushort leftCol = BitConverter.ToUInt16(record.Data, 6);
                
                if (worksheet.FreezePane != null)
                {
                    worksheet.FreezePane.ColSplit = x;
                    worksheet.FreezePane.RowSplit = y;
                    worksheet.FreezePane.TopRow = topRow + 1; // 转为1-based
                    worksheet.FreezePane.LeftCol = leftCol + 1; // 转为1-based
                }
                else
                {
                    worksheet.FreezePane = new FreezePane
                    {
                        ColSplit = x,
                        RowSplit = y,
                        TopRow = topRow + 1, // 转为1-based
                        LeftCol = leftCol + 1 // 转为1-based
                    };
                }
            }
        }
        
        private void ParseFontRecord(BiffRecord record, Worksheet worksheet)
        {
            // 解析字体记录
            if (record.Data != null && record.Data.Length >= 48)
            {
                var font = new Font();
                font.Height = BitConverter.ToInt16(record.Data, 0);
                font.IsBold = (BitConverter.ToUInt16(record.Data, 2) & 0x0001) != 0;
                font.IsItalic = (BitConverter.ToUInt16(record.Data, 2) & 0x0002) != 0;
                font.IsUnderline = (BitConverter.ToUInt16(record.Data, 2) & 0x0004) != 0;
                font.IsStrikethrough = (BitConverter.ToUInt16(record.Data, 2) & 0x0008) != 0;
                font.ColorIndex = BitConverter.ToUInt16(record.Data, 6);
                font.Name = System.Text.Encoding.ASCII.GetString(record.Data, 40, record.Data.Length - 40).TrimEnd('\0');
                
                worksheet.Fonts.Add(font);
            }
        }
        
        private void ParseXfRecord(BiffRecord record, Worksheet worksheet)
        {
            // 解析扩展格式记录
            if (record.Data != null && record.Data.Length >= 28)
            {
                var xf = new Xf();
                xf.FontIndex = BitConverter.ToUInt16(record.Data, 0);
                xf.NumberFormatIndex = BitConverter.ToUInt16(record.Data, 2);
                xf.CellFormatIndex = BitConverter.ToUInt16(record.Data, 4);
                
                // 解析对齐方式
                ushort alignment = BitConverter.ToUInt16(record.Data, 6);
                byte horizontalAlign = (byte)((alignment & 0x000F) >> 0);
                byte verticalAlign = (byte)((alignment & 0x00F0) >> 4);
                
                switch (horizontalAlign)
                {
                    case 0: xf.HorizontalAlignment = "general";
                        break;
                    case 1: xf.HorizontalAlignment = "left";
                        break;
                    case 2: xf.HorizontalAlignment = "center";
                        break;
                    case 3: xf.HorizontalAlignment = "right";
                        break;
                    case 4: xf.HorizontalAlignment = "fill";
                        break;
                    case 5: xf.HorizontalAlignment = "justify";
                        break;
                    case 6: xf.HorizontalAlignment = "centerContinuous";
                        break;
                    case 7: xf.HorizontalAlignment = "distributed";
                        break;
                }
                
                switch (verticalAlign)
                {
                    case 0: xf.VerticalAlignment = "top";
                        break;
                    case 1: xf.VerticalAlignment = "center";
                        break;
                    case 2: xf.VerticalAlignment = "bottom";
                        break;
                    case 3: xf.VerticalAlignment = "justify";
                        break;
                    case 4: xf.VerticalAlignment = "distributed";
                        break;
                }
                
                // 解析缩进
                xf.Indent = (byte)((alignment & 0x0F00) >> 8);
                
                // 解析文本换行
                xf.WrapText = (alignment & 0x1000) != 0;

                // 解析边框 (偏移10-17)
                if (record.Data.Length >= 18)
                {
                    uint border1 = BitConverter.ToUInt32(record.Data, 10);
                    uint border2 = BitConverter.ToUInt32(record.Data, 14);

                    var border = new Border();
                    border.Left = GetBorderLineStyle((byte)(border1 & 0x0F));
                    border.Right = GetBorderLineStyle((byte)((border1 >> 4) & 0x0F));
                    border.Top = GetBorderLineStyle((byte)((border1 >> 8) & 0x0F));
                    border.Bottom = GetBorderLineStyle((byte)((border1 >> 12) & 0x0F));

                    border.LeftColor = GetColorFromPalette((int)((border1 >> 16) & 0x7F));
                    border.RightColor = GetColorFromPalette((int)((border1 >> 23) & 0x7F));

                    border.TopColor = GetColorFromPalette((int)(border2 & 0x7F));
                    border.BottomColor = GetColorFromPalette((int)((border2 >> 7) & 0x7F));
                    border.DiagonalColor = GetColorFromPalette((int)((border2 >> 14) & 0x7F));
                    border.Diagonal = GetBorderLineStyle((byte)((border2 >> 21) & 0x0F));

                    // 添加到全局列表并分配索引
                    _workbook.Borders.Add(border);
                    xf.BorderIndex = _workbook.Borders.Count - 1;
                }

                // 解析填充 (偏移18-21)
                if (record.Data.Length >= 22)
                {
                    ushort fillData = BitConverter.ToUInt16(record.Data, 18);
                    byte pattern = (byte)(fillData & 0x3F);
                    
                    var fill = new Fill();
                    fill.PatternType = GetPatternType(pattern);
                    // 填充颜色处理通常更复杂，涉及前景色和背景色
                    
                    _workbook.Fills.Add(fill);
                    xf.FillIndex = _workbook.Fills.Count - 1;
                }
                
                // 解析锁定和隐藏状态
                xf.IsLocked = (BitConverter.ToUInt16(record.Data, 26) & 0x0001) != 0;
                xf.IsHidden = (BitConverter.ToUInt16(record.Data, 26) & 0x0002) != 0;
                
                worksheet.Xfs.Add(xf);
            }
        }
        
        private void ParsePaletteRecord(BiffRecord record, Worksheet worksheet)
        {
            // 解析调色板记录
            if (record.Data != null && record.Data.Length >= 10)
            {
                int startIndex = BitConverter.ToUInt16(record.Data, 0);
                int colorCount = (record.Data.Length - 2) / 10;
                
                for (int i = 0; i < colorCount; i++)
                {
                    int offset = 2 + i * 10;
                    if (offset + 10 <= record.Data.Length)
                    {
                        byte red = record.Data[offset + 2];
                        byte green = record.Data[offset + 4];
                        byte blue = record.Data[offset + 6];
                        string color = $"#{red:X2}{green:X2}{blue:X2}";
                        worksheet.Palette[startIndex + i] = color;
                    }
                }
            }
        }
        
        private void ParseChartRecord(BiffRecord record, Worksheet worksheet)
        {
            // 解析图表记录
            var chart = new Chart();
            // 默认为柱状图
            chart.ChartType = "barChart";
            
            // 解析图表类型
            if (record.Data != null && record.Data.Length >= 2)
            {
                ushort chartType = BitConverter.ToUInt16(record.Data, 0);
                chart.ChartType = GetChartType(chartType);
            }
            
            // 解析图表位置和大小
            if (record.Data != null && record.Data.Length >= 18)
            {
                // BIFF8 uses 4 byte integers (2+16 offset bytes typically for this sub-record)
                // Simplified fallback extraction trying to avoid breaking existing layout assumptions
                chart.Left = BitConverter.ToInt16(record.Data, 2);
                chart.Top = BitConverter.ToInt16(record.Data, 6);
                chart.Width = BitConverter.ToInt16(record.Data, 10);
                chart.Height = BitConverter.ToInt16(record.Data, 14);
            }
            
            // 添加默认图例
            chart.Legend = new Legend
            {
                Visible = true,
                Position = "right"
            };
            
            // 添加默认坐标轴
            chart.XAxis = new Axis
            {
                Visible = true,
                Title = "X轴",
                NumberFormat = "General"
            };
            
            chart.YAxis = new Axis
            {
                Visible = true,
                Title = "Y轴",
                NumberFormat = "General"
            };
            
            worksheet.Charts.Add(chart);
        }
        
        private void ParseChartTitleRecord(BiffRecord record, Worksheet worksheet)
        {
            // 解析图表标题记录
            if (worksheet.Charts.Count > 0 && record.Data != null)
            {
                var chart = worksheet.Charts[worksheet.Charts.Count - 1];
                // 读取标题字符串
                string title = System.Text.Encoding.ASCII.GetString(record.Data);
                chart.Title = title.Trim();
            }
        }
        
        private void ParseSeriesRecord(BiffRecord record, Worksheet worksheet)
        {
            // 解析数据系列记录
            if (worksheet.Charts.Count > 0 && record.Data != null)
            {
                var chart = worksheet.Charts[worksheet.Charts.Count - 1];
                var series = new Series();
                
                // 读取系列名称
                if (record.Data.Length >= 2)
                {
                    ushort nameLength = BitConverter.ToUInt16(record.Data, 0);
                    if (nameLength > 0 && nameLength + 2 <= record.Data.Length)
                    {
                        string seriesName = System.Text.Encoding.ASCII.GetString(record.Data, 2, nameLength);
                        series.Name = seriesName.Trim();
                    }
                }
                
                // 解析值范围和类别范围
                if (record.Data.Length >= 6)
                {
                    // 读取值范围
                    ushort valuesOffset = BitConverter.ToUInt16(record.Data, 2 + (record.Data[2] > 0 ? record.Data[2] : 0));
                    if (valuesOffset > 0 && valuesOffset < record.Data.Length)
                    {
                        string valuesRange = ParseRange(record.Data, valuesOffset);
                        if (!string.IsNullOrEmpty(valuesRange))
                        {
                            series.ValuesRange = valuesRange;
                        }
                    }
                    
                    // 读取类别范围
                    ushort categoriesOffset = BitConverter.ToUInt16(record.Data, 4 + (record.Data[2] > 0 ? record.Data[2] : 0));
                    if (categoriesOffset > 0 && categoriesOffset < record.Data.Length)
                    {
                        string categoriesRange = ParseRange(record.Data, categoriesOffset);
                        if (!string.IsNullOrEmpty(categoriesRange))
                        {
                            series.CategoriesRange = categoriesRange;
                        }
                    }
                }
                
                // 如果没有解析到范围，使用默认值
                if (string.IsNullOrEmpty(series.ValuesRange))
                {
                    series.ValuesRange = $"{worksheet.Name}!$B$2:$B$6";
                }
                if (string.IsNullOrEmpty(series.CategoriesRange))
                {
                    series.CategoriesRange = $"{worksheet.Name}!$A$2:$A$6";
                }
                
                // 添加默认系列样式
                series.LineStyle = new LineStyle
                {
                    Width = 2
                };
                
                chart.Series.Add(series);
            }
        }
        
        private string ParseRange(byte[] data, int offset)
        {
            // 解析范围字符串
            if (offset + 2 <= data.Length)
            {
                ushort length = BitConverter.ToUInt16(data, offset);
                if (length > 0 && offset + 2 + length <= data.Length)
                {
                    string range = System.Text.Encoding.ASCII.GetString(data, offset + 2, length).Trim();
                    return range;
                }
            }
            return string.Empty;
        }
        
        private string GetChartType(ushort chartType)
        {
            // 映射BIFF8图表类型到OpenXML图表类型
            switch (chartType)
            {
                case 0: return "barChart";
                case 1: return "colChart";
                case 2: return "lineChart";
                case 3: return "pieChart";
                case 4: return "scatterChart";
                case 5: return "areaChart";
                case 6: return "doughnutChart";
                case 7: return "radarChart";
                case 8: return "surfaceChart";
                case 9: return "bubbleChart";
                case 10: return "stockChart";
                default: return "colChart";
            }
        }
        
        private void ParseMSODrawingRecord(BiffRecord record, Worksheet worksheet)
        {
            // 解析图片和绘图对象记录
            // 这里只是简单地识别记录，实际解析需要更复杂的逻辑
        }
        
        private void ParsePictureRecord(BiffRecord record, Worksheet worksheet)
        {
            // 解析图片记录
            if (record.Data != null)
            {
                try
                {
                    var picture = new Picture();
                    // 读取图片数据
                    picture.Data = new byte[record.Data.Length];
                    Array.Copy(record.Data, picture.Data, record.Data.Length);
                    
                    // 识别图片格式
                    if (record.Data.Length >= 4)
                    {
                        // 检查图片格式
                        if (record.Data[0] == 0x89 && record.Data[1] == 0x50 && record.Data[2] == 0x4E && record.Data[3] == 0x47)
                        {
                            // PNG格式
                            picture.MimeType = "image/png";
                            picture.Extension = "png";
                        }
                        else if (record.Data[0] == 0xFF && record.Data[1] == 0xD8)
                        {
                            // JPEG格式
                            picture.MimeType = "image/jpeg";
                            picture.Extension = "jpg";
                        }
                        else if (record.Data[0] == 0x47 && record.Data[1] == 0x49 && record.Data[2] == 0x46)
                        {
                            // GIF格式
                            picture.MimeType = "image/gif";
                            picture.Extension = "gif";
                        }
                        else if (record.Data[0] == 0x42 && record.Data[1] == 0x4D)
                        {
                            // BMP格式
                            picture.MimeType = "image/bmp";
                            picture.Extension = "bmp";
                        }
                        else if (record.Data[0] == 0x52 && record.Data[1] == 0x49 && record.Data[2] == 0x46 && record.Data[3] == 0x46)
                        {
                            // WebP格式
                            if (record.Data.Length >= 12 && record.Data[8] == 0x57 && record.Data[9] == 0x45 && record.Data[10] == 0x42 && record.Data[11] == 0x50)
                            {
                                picture.MimeType = "image/webp";
                                picture.Extension = "webp";
                            }
                            else
                            {
                                // 默认为BMP格式
                                picture.MimeType = "image/bmp";
                                picture.Extension = "bmp";
                            }
                        }
                        else if (record.Data[0] == 0x49 && record.Data[1] == 0x49 && record.Data[2] == 0x2A && record.Data[3] == 0x00)
                        {
                            // TIFF格式 (小端序)
                            picture.MimeType = "image/tiff";
                            picture.Extension = "tiff";
                        }
                        else if (record.Data[0] == 0x4D && record.Data[1] == 0x4D && record.Data[2] == 0x00 && record.Data[3] == 0x2A)
                        {
                            // TIFF格式 (大端序)
                            picture.MimeType = "image/tiff";
                            picture.Extension = "tiff";
                        }
                        else if (record.Data[0] == 0x38 && record.Data[1] == 0x42 && record.Data[2] == 0x50 && record.Data[3] == 0x53)
                        {
                            // PSD格式
                            picture.MimeType = "image/vnd.adobe.photoshop";
                            picture.Extension = "psd";
                        }
                        else if (record.Data[0] == 0x52 && record.Data[1] == 0x49 && record.Data[2] == 0x46 && record.Data[3] == 0x46 && record.Data.Length >= 10)
                        {
                            // RTF格式
                            picture.MimeType = "image/rtf";
                            picture.Extension = "rtf";
                        }
                        else if (record.Data[0] == 0x49 && record.Data[1] == 0x4D && record.Data[2] == 0x47)
                        {
                            // IMG格式
                            picture.MimeType = "image/x-ms-bmp";
                            picture.Extension = "img";
                        }
                        else if (record.Data[0] == 0x43 && record.Data[1] == 0x57 && record.Data[2] == 0x53)
                        {
                            // CWS格式 (Compressed Web Archive)
                            picture.MimeType = "application/x-cws";
                            picture.Extension = "cws";
                        }
                        else
                        {
                            // 默认为BMP格式
                            picture.MimeType = "image/bmp";
                            picture.Extension = "bmp";
                        }
                    }
                    else
                    {
                        // 默认为BMP格式
                        picture.MimeType = "image/bmp";
                        picture.Extension = "bmp";
                    }
                    
                    worksheet.Pictures.Add(picture);
                }
                catch (Exception ex)
                {
                    throw new ImageProcessingException($"处理图片时发生错误: {ex.Message}", ex);
                }
            }
        }
        
        private void ParseObjRecord(BiffRecord record, Worksheet worksheet)
        {
            // 解析嵌入对象记录
            if (record.Data != null)
            {
                var embeddedObject = new EmbeddedObject();
                // 读取对象数据
                embeddedObject.Data = new byte[record.Data.Length];
                Array.Copy(record.Data, embeddedObject.Data, record.Data.Length);
                embeddedObject.MimeType = "application/octet-stream"; // 默认为二进制流
                worksheet.EmbeddedObjects.Add(embeddedObject);
            }
        }
        
        private void ParseDVRecord(BiffRecord record, Worksheet worksheet)
        {
            // 解析数据验证记录
            if (record.Data != null && record.Data.Length >= 16)
            {
                var dataValidation = new DataValidation();
                
                // 解析选项标志
                ushort options = BitConverter.ToUInt16(record.Data, 0);
                dataValidation.AllowBlank = (options & 0x01) != 0;
                
                // 解析条件类型
                ushort validationType = BitConverter.ToUInt16(record.Data, 2);
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
                
                // 解析操作符
                ushort operatorType = BitConverter.ToUInt16(record.Data, 4);
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
                
                // 解析公式偏移量 (不再是简单的字符串偏移，而是用于二进制 Ptg 读取)
        int currentOffset = 6;
        ushort formula1Size = BitConverter.ToUInt16(record.Data, currentOffset); currentOffset += 2;
        ushort formula2Size = BitConverter.ToUInt16(record.Data, currentOffset); currentOffset += 2;
        
        // 解析公式1
        if (formula1Size > 0 && currentOffset + formula1Size <= record.Data.Length)
        {
            byte[] formula1Bytes = new byte[formula1Size];
            Array.Copy(record.Data, currentOffset, formula1Bytes, 0, formula1Size);
            dataValidation.Formula1 = FormulaDecompiler.Decompile(formula1Bytes);
            currentOffset += formula1Size;
        }
        
        // 解析公式2
        if (formula2Size > 0 && currentOffset + formula2Size <= record.Data.Length)
        {
            byte[] formula2Bytes = new byte[formula2Size];
            Array.Copy(record.Data, currentOffset, formula2Bytes, 0, formula2Size);
            dataValidation.Formula2 = FormulaDecompiler.Decompile(formula2Bytes);
            currentOffset += formula2Size;
        }
        
        // 解析范围
        if (currentOffset + 8 <= record.Data.Length)
        {
            // 解析范围的起始和结束单元格
            ushort firstRow = BitConverter.ToUInt16(record.Data, currentOffset); currentOffset += 2;
            ushort lastRow = BitConverter.ToUInt16(record.Data, currentOffset); currentOffset += 2;
            ushort firstCol = BitConverter.ToUInt16(record.Data, currentOffset); currentOffset += 2;
            ushort lastCol = BitConverter.ToUInt16(record.Data, currentOffset); currentOffset += 2;
                    
                    // 生成范围字符串
                    dataValidation.Range = $"{GetColumnLetter(firstCol)}{firstRow + 1}:{GetColumnLetter(lastCol)}{lastRow + 1}";
                }
                else
                {
                    // 如果没有范围信息，使用默认值
                    dataValidation.Range = "A1:A10";
                }
                
                worksheet.DataValidations.Add(dataValidation);
            }
        }
        
        private void ParseCFRecord(BiffRecord record, Worksheet worksheet)
        {
            // 解析条件格式记录
            if (record.Data != null && record.Data.Length >= 8)
            {
                var conditionalFormat = new ConditionalFormat();
                
                // 解析条件类型
                ushort conditionType = BitConverter.ToUInt16(record.Data, 0);
                switch (conditionType)
                {
                    case 0: conditionalFormat.Type = "cellIs"; break;
                    case 1: conditionalFormat.Type = "expression"; break;
                    case 2: conditionalFormat.Type = "colorScale"; break;
                    case 3: conditionalFormat.Type = "dataBar"; break;
                    case 4: conditionalFormat.Type = "iconSet"; break;
                }
                
                // 解析操作符
                ushort operatorType = BitConverter.ToUInt16(record.Data, 2);
                switch (operatorType)
                {
                    case 0: conditionalFormat.Operator = "between"; break;
                    case 1: conditionalFormat.Operator = "notBetween"; break;
                    case 2: conditionalFormat.Operator = "equal"; break;
                    case 3: conditionalFormat.Operator = "notEqual"; break;
                    case 4: conditionalFormat.Operator = "greaterThan"; break;
                    case 5: conditionalFormat.Operator = "lessThan"; break;
                    case 6: conditionalFormat.Operator = "greaterThanOrEqual"; break;
                    case 7: conditionalFormat.Operator = "lessThanOrEqual"; break;
                    case 8: conditionalFormat.Operator = "containsText"; break;
                    case 9: conditionalFormat.Operator = "notContainsText"; break;
                    case 10: conditionalFormat.Operator = "beginsWith"; break;
                    case 11: conditionalFormat.Operator = "endsWith"; break;
                }
                
                // 解析公式
                if (record.Data.Length >= 12)
                {
                    int currentOffset = 4;
                    ushort formula1Size = BitConverter.ToUInt16(record.Data, currentOffset); 
                    currentOffset += 2;
                    ushort formula2Size = BitConverter.ToUInt16(record.Data, currentOffset);
                    currentOffset += 2;
                    
                    // Skip optional parts we don't handle yet block
                    // Formula strings start further down (in BIFF8: size1, size2, 4 bytes padding usually, then formulas)
                    if (currentOffset + 4 <= record.Data.Length)
                    {
                        currentOffset += 4;
                        
                        if (formula1Size > 0 && currentOffset + formula1Size <= record.Data.Length)
                        {
                            byte[] ptg1 = new byte[formula1Size];
                            Array.Copy(record.Data, currentOffset, ptg1, 0, formula1Size);
                            conditionalFormat.Formula = FormulaDecompiler.Decompile(ptg1);
                            currentOffset += formula1Size;
                        }

                        if (formula2Size > 0 && currentOffset + formula2Size <= record.Data.Length)
                        {
                            byte[] ptg2 = new byte[formula2Size];
                            Array.Copy(record.Data, currentOffset, ptg2, 0, formula2Size);
                            // Not stored in POCO model: conditionalFormat.Formula2 = FormulaDecompiler.Decompile(ptg2);
                        }
                    }
                }
                
                // 解析范围（从CFHeader记录中获取）
                conditionalFormat.Range = !string.IsNullOrEmpty(_currentCFRange) ? _currentCFRange : "A1:A10";
                
                worksheet.ConditionalFormats.Add(conditionalFormat);
            }
        }
        
        private void ParseCFHeaderRecord(BiffRecord record, Worksheet worksheet)
        {
            // 解析条件格式头部记录
            if (record.Data != null && record.Data.Length >= 12)
            {
                // 解析条件格式数量
                ushort conditionCount = BitConverter.ToUInt16(record.Data, 0);
                
                // 解析范围
                ushort firstRow = BitConverter.ToUInt16(record.Data, 2);
                ushort lastRow = BitConverter.ToUInt16(record.Data, 4);
                ushort firstCol = BitConverter.ToUInt16(record.Data, 6);
                ushort lastCol = BitConverter.ToUInt16(record.Data, 8);
                
                // 生成范围字符串
                _currentCFRange = $"{GetColumnLetter(firstCol)}{firstRow + 1}:{GetColumnLetter(lastCol)}{lastRow + 1}";
                
                // 存储范围信息，供后续的CF记录使用
            }
        }

        private void ParseHyperlinkRecord(BiffRecord record, Worksheet worksheet)
        {
            // 解析超链接记录
            if (record.Data != null && record.Data.Length >= 20)
            {
                var hyperlink = new Hyperlink();
                
                // 解析超链接范围
                ushort firstRow = BitConverter.ToUInt16(record.Data, 0);
                ushort lastRow = BitConverter.ToUInt16(record.Data, 2);
                ushort firstCol = BitConverter.ToUInt16(record.Data, 4);
                ushort lastCol = BitConverter.ToUInt16(record.Data, 6);
                hyperlink.Range = $"{GetColumnLetter(firstCol)}{firstRow + 1}:{GetColumnLetter(lastCol)}{lastRow + 1}";
                
                // 解析目标URL
                int urlLength = BitConverter.ToInt16(record.Data, 18);
                if (urlLength > 0 && record.Data.Length >= 20 + urlLength)
                {
                    hyperlink.Target = System.Text.Encoding.ASCII.GetString(record.Data, 20, urlLength);
                }
                
                worksheet.Hyperlinks.Add(hyperlink);
            }
        }

        private void ParseCommentRecord(BiffRecord record, Worksheet worksheet)
        {
            // 解析注释记录
            if (record.Data != null && record.Data.Length >= 12)
            {
                var comment = new Comment();
                
                // 解析行和列索引
                ushort row = BitConverter.ToUInt16(record.Data, 0);
                ushort col = BitConverter.ToUInt16(record.Data, 2);
                comment.RowIndex = row + 1; // 转换为1-based索引
                comment.ColumnIndex = col + 1; // 转换为1-based索引
                
                // 解析作者和注释文本
                if (record.Data.Length >= 14)
                {
                    // 读取作者长度
                    byte authorLength = record.Data[12];
                    if (authorLength > 0 && record.Data.Length >= 13 + authorLength)
                    {
                        // 读取作者名称
                        comment.Author = System.Text.Encoding.ASCII.GetString(record.Data, 13, authorLength);
                        
                        // 读取注释文本
                        int textOffset = 13 + authorLength;
                        if (record.Data.Length > textOffset)
                        {
                            // 简单实现，实际需要更复杂的解析逻辑
                            comment.Text = System.Text.Encoding.ASCII.GetString(record.Data, textOffset, record.Data.Length - textOffset);
                        }
                    }
                }
                
                worksheet.Comments.Add(comment);
            }
        }

        private string GetColumnLetter(int columnIndex)
        {
            var columnReference = string.Empty;
            int col = columnIndex;
            while (col > 0)
            {
                col--;
                columnReference = (char)('A' + col % 26) + columnReference;
                col /= 26;
            }
            return columnReference;
        }

        private List<RichTextRun> ParseRichText(byte[] data, int offset)
        {
            var richTextRuns = new List<RichTextRun>();
            int currentOffset = offset;
            
            try
            {
                // 读取富文本记录的结构
                // 根据BIFF8格式规范解析富文本数据
                while (currentOffset < data.Length)
                {
                    // 读取文本长度
                    if (currentOffset + 2 <= data.Length)
                    {
                        short textLength = BitConverter.ToInt16(data, currentOffset);
                        currentOffset += 2;
                        
                        // 读取文本类型（0=ASCII, 1=Unicode）
                        byte textType = 0;
                        if (currentOffset < data.Length)
                        {
                            textType = data[currentOffset];
                            currentOffset += 1;
                        }
                        
                        // 读取文本内容
                        string text;
                        if (textType == 0)
                        {
                            // ASCII字符串
                            if (currentOffset + textLength <= data.Length)
                            {
                                text = System.Text.Encoding.ASCII.GetString(data, currentOffset, textLength);
                                currentOffset += textLength;
                            }
                            else
                            {
                                break;
                            }
                        }
                        else
                        {
                            // Unicode字符串
                            if (currentOffset + textLength * 2 <= data.Length)
                            {
                                text = System.Text.Encoding.Unicode.GetString(data, currentOffset, textLength * 2);
                                currentOffset += textLength * 2;
                            }
                            else
                            {
                                break;
                            }
                        }
                        
                        // 读取字体索引
                        short fontIndex = 0;
                        if (currentOffset + 2 <= data.Length)
                        {
                            fontIndex = BitConverter.ToInt16(data, currentOffset);
                            currentOffset += 2;
                        }
                        
                        // 创建富文本运行
                        var run = new RichTextRun
                        {
                            Text = text,
                            // 根据fontIndex获取对应的字体信息
                            Font = GetFontByIndex(fontIndex)
                        };
                        richTextRuns.Add(run);
                    }
                    else
                    {
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error("解析富文本时发生错误", ex);
                // 继续处理，返回已解析的部分
            }
            
            return richTextRuns;
        }

        private Font GetFontByIndex(short fontIndex)
        {
            // 根据字体索引获取字体信息
            try
            {
                // 检查字体索引是否有效
                if (fontIndex >= 0 && fontIndex < _fonts.Count)
                {
                    return _fonts[fontIndex];
                }
                else
                {
                    // 返回默认字体
                    return new Font
                    {
                        Name = "Arial",
                        Size = 11,
                        Bold = false,
                        Italic = false,
                        Underline = false,
                        Color = "000000"
                    };
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"获取字体信息时发生错误，字体索引: {fontIndex}", ex);
                // 返回默认字体
                return new Font
                {
                    Name = "Arial",
                    Size = 11,
                    Bold = false,
                    Italic = false,
                    Underline = false,
                    Color = "000000"
                };
            }
        }

        private void ParseSheetRecord(BiffRecord record, Workbook workbook)
        {
            // 解析工作表记录
            var worksheet = new Worksheet();
            
            // 从记录数据中提取工作表名称
            if (record.Data != null && record.Data.Length >= 8)
            {
                // BIFF8 BoundSheet 记录包含流位置(4) + 类型(1) + 隐藏标志(1) + 名称
                // 名称是 ShortXLUnicodeString (1 byte len + 1 byte option + data)
                int nameOffset = 6;
                if (record.Data.Length > nameOffset)
                {
                    byte len = record.Data[nameOffset];
                    int pos = nameOffset + 1;
                    worksheet.Name = ReadBiffStringFromBytes(record.Data, ref pos, len);
                }
            }
            
            // 如果没有提取到名称，使用默认名称
            if (string.IsNullOrEmpty(worksheet.Name))
            {
                worksheet.Name = "Sheet" + (workbook.Worksheets.Count + 1);
            }
            
            workbook.Worksheets.Add(worksheet);
        }

        private void ParseSstInfo(BiffRecord record, long workbookStreamEnd)
        {
            // SST 记录后面可能紧跟多个 CONTINUE 记录
            using (var ms = new MemoryStream())
            {
                if (record.Data != null)
                {
                    ms.Write(record.Data, 0, record.Data.Length);
                }

                // 连续读取后续的 CONTINUE (0x003C) 记录
                while (_stream.Position + 4 <= workbookStreamEnd)
                {
                    long currentPos = _stream.Position;
                    ushort nextId = _reader.ReadUInt16();
                    ushort nextLen = _reader.ReadUInt16();

                    if (nextId == 0x003C) // CONTINUE
                    {
                        byte[] contData = _reader.ReadBytes(nextLen);
                        ms.Write(contData, 0, contData.Length);
                    }
                    else
                    {
                        // 不是 CONTINUE，回退流位置
                        _stream.Seek(currentPos, SeekOrigin.Begin);
                        break;
                    }
                }

                byte[] fullData = ms.ToArray();
                if (fullData.Length < 8) return;

                int uniqueCount = BitConverter.ToInt32(fullData, 4);
                _sharedStrings.Capacity = Math.Max(_sharedStrings.Capacity, uniqueCount);

                int offset = 8;
                for (int i = 0; i < uniqueCount && offset < fullData.Length; i++)
                {
                    _sharedStrings.Add(ReadBiffString(fullData, ref offset));
                }
            }
        }

        private string ReadBiffString(byte[] data, ref int offset)
        {
            if (offset + 2 > data.Length) return string.Empty;
            ushort charCount = BitConverter.ToUInt16(data, offset);
            offset += 2;
            if (offset >= data.Length) return string.Empty;
            
            byte option = data[offset];
            offset += 1;
            
            bool isUnicode = (option & 0x01) != 0;
            bool hasRichText = (option & 0x08) != 0;
            bool hasExtended = (option & 0x04) != 0;
            
            int runs = 0;
            if (hasRichText)
            {
                if (offset + 2 > data.Length) return string.Empty;
                runs = BitConverter.ToUInt16(data, offset);
                offset += 2;
            }
            
            int extendedSize = 0;
            if (hasExtended)
            {
                if (offset + 4 > data.Length) return string.Empty;
                extendedSize = BitConverter.ToInt32(data, offset);
                offset += 4; // Skip the size header
            }

            string result;
            if (isUnicode)
            {
                int byteCount = charCount * 2;
                if (offset + byteCount > data.Length) byteCount = data.Length - offset;
                result = System.Text.Encoding.Unicode.GetString(data, offset, byteCount);
                offset += byteCount;
            }
            else
            {
                int byteCount = charCount;
                if (offset + byteCount > data.Length) byteCount = data.Length - offset;
                result = System.Text.Encoding.ASCII.GetString(data, offset, byteCount);
                offset += byteCount;
            }

            if (hasRichText)
            {
                offset += runs * 4; // Skip formatting runs
            }
            
            if (hasExtended)
            {
                offset += extendedSize; // Skip the phonetic string data payload
            }
            
            return result;
        }

        private string ReadBiffStringFromBytes(byte[] data, ref int offset, int charCount)
        {
            if (offset >= data.Length) return string.Empty;
            byte option = data[offset];
            offset += 1;
            bool isUnicode = (option & 0x01) != 0;
            string result;
            if (isUnicode)
            {
                int byteCount = charCount * 2;
                if (offset + byteCount > data.Length) byteCount = data.Length - offset;
                result = System.Text.Encoding.Unicode.GetString(data, offset, byteCount);
                offset += byteCount;
            }
            else
            {
                int byteCount = charCount;
                if (offset + byteCount > data.Length) byteCount = data.Length - offset;
                result = System.Text.Encoding.ASCII.GetString(data, offset, byteCount);
                offset += byteCount;
            }
            return result;
        }

        private void ParseFontRecordToGlobal(BiffRecord record)
        {
            if (record.Data != null && record.Data.Length >= 14)
            {
                var font = new Font();
                font.Height = BitConverter.ToInt16(record.Data, 0);
                ushort grbit = BitConverter.ToUInt16(record.Data, 2);
                font.IsBold = BitConverter.ToUInt16(record.Data, 6) >= 700;
                font.IsItalic = (grbit & 0x0002) != 0;
                font.IsUnderline = (record.Data[10]) != 0;
                font.IsStrikethrough = (grbit & 0x0008) != 0;
                font.ColorIndex = BitConverter.ToUInt16(record.Data, 4);
                
                int nameOffset = 14;
                if (record.Data.Length > nameOffset)
                {
                    // BIFF8 Font record uses ShortXLUnicodeString at offset 14 (1 byte len + 1 byte option)
                    byte len = record.Data[nameOffset];
                    if (record.Data.Length > nameOffset + 1)
                    {
                        byte opt = record.Data[nameOffset + 1];
                        bool isUni = (opt & 0x01) != 0;
                        if (isUni)
                        {
                           font.Name = System.Text.Encoding.Unicode.GetString(record.Data, nameOffset + 2, Math.Min(len * 2, record.Data.Length - nameOffset - 2));
                        }
                        else
                        {
                           font.Name = System.Text.Encoding.ASCII.GetString(record.Data, nameOffset + 2, Math.Min(len, record.Data.Length - nameOffset - 2));
                        }
                    }
                }
                _fonts.Add(font);
            }
        }

        private void ParseXfRecordToGlobal(BiffRecord record)
        {
            if (record.Data != null && record.Data.Length >= 20)
            {
                var xf = new Xf();
                xf.FontIndex = BitConverter.ToUInt16(record.Data, 0);
                xf.NumberFormatIndex = BitConverter.ToUInt16(record.Data, 2);
                
                // 解析对齐方式 (offset 6-9)
                ushort alignment = BitConverter.ToUInt16(record.Data, 6);
                byte horizontalAlign = (byte)(alignment & 0x07);
                byte verticalAlign = (byte)((alignment & 0x70) >> 4);
                
                xf.HorizontalAlignment = horizontalAlign switch {
                    1 => "left", 2 => "center", 3 => "right", 4 => "fill", 5 => "justify", 6 => "centerContinuous", 7 => "distributed", _ => "general"
                };
                xf.VerticalAlignment = verticalAlign switch {
                    1 => "center", 2 => "bottom", 3 => "justify", 4 => "distributed", _ => "top"
                };
                
                xf.WrapText = (alignment & 0x08) != 0;
                xf.Indent = (byte)((alignment >> 8) & 0x0F);

                // 解析边框 (偏移10-17)
                if (record.Data.Length >= 18)
                {
                    uint border1 = BitConverter.ToUInt32(record.Data, 10);
                    uint border2 = BitConverter.ToUInt32(record.Data, 14);

                    var border = new Border {
                        Left = GetBorderLineStyle((byte)(border1 & 0x0F)),
                        Right = GetBorderLineStyle((byte)((border1 >> 4) & 0x0F)),
                        Top = GetBorderLineStyle((byte)((border1 >> 8) & 0x0F)),
                        Bottom = GetBorderLineStyle((byte)((border1 >> 12) & 0x0F)),
                        LeftColor = GetColorFromPalette((int)((border1 >> 16) & 0x7F)),
                        RightColor = GetColorFromPalette((int)((border1 >> 23) & 0x7F)),
                        TopColor = GetColorFromPalette((int)(border2 & 0x7F)),
                        BottomColor = GetColorFromPalette((int)((border2 >> 7) & 0x7F)),
                        DiagonalColor = GetColorFromPalette((int)((border2 >> 14) & 0x7F)),
                        Diagonal = GetBorderLineStyle((byte)((border2 >> 21) & 0x0F))
                    };

                    // 只有当边框不是全部为 none 时才添加或查找
                    if (border.Left != "none" || border.Right != "none" || border.Top != "none" || border.Bottom != "none" || border.Diagonal != "none")
                    {
                        int existingBorderIdx = _workbook.Borders.FindIndex(b => 
                            b.Left == border.Left && b.Right == border.Right && b.Top == border.Top && b.Bottom == border.Bottom && 
                            b.LeftColor == border.LeftColor && b.RightColor == border.RightColor && b.TopColor == border.TopColor && b.BottomColor == border.BottomColor);
                        
                        if (existingBorderIdx >= 0)
                        {
                            xf.BorderIndex = existingBorderIdx + 1; // 0 是默认
                        }
                        else
                        {
                            _workbook.Borders.Add(border);
                            xf.BorderIndex = _workbook.Borders.Count;
                        }
                    }
                    else
                    {
                        xf.BorderIndex = 0;
                    }
                }

                // 解析填充 (偏移18-21)
                if (record.Data.Length >= 20)
                {
                    ushort fillData = BitConverter.ToUInt16(record.Data, 18);
                    byte pattern = (byte)(fillData & 0x3F);
                    if (pattern > 0)
                    {
                        var fill = new Fill { PatternType = GetPatternType(pattern) };
                        int existingFillIdx = _workbook.Fills.FindIndex(f => f.PatternType == fill.PatternType);
                        
                        if (existingFillIdx >= 0)
                        {
                            xf.FillIndex = existingFillIdx + 2; // 0 和 1 是默认
                        }
                        else
                        {
                            _workbook.Fills.Add(fill);
                            xf.FillIndex = _workbook.Fills.Count + 1;
                        }
                    }
                    else
                    {
                        xf.FillIndex = 0;
                    }
                }
                
                // 解析锁定和隐藏状态 (偏移26)
                if (record.Data.Length >= 28)
                {
                    ushort options = BitConverter.ToUInt16(record.Data, 26);
                    xf.IsLocked = (options & 0x0001) != 0;
                    xf.IsHidden = (options & 0x0002) != 0;
                }


                _xfList.Add(xf);
            }
        }

        private void ParseFormatRecordGlobal(BiffRecord record)
        {
            if (record.Data != null && record.Data.Length >= 5)
            {
                ushort index = BitConverter.ToUInt16(record.Data, 0);
                int offset = 2;
                _formats[index] = ReadBiffString(record.Data, ref offset);
            }
        }

        private void ParsePaletteRecordGlobal(BiffRecord record)
        {
            if (record.Data != null && record.Data.Length >= 4)
            {
                int count = BitConverter.ToUInt16(record.Data, 0);
                for (int i = 0; i < count && (2 + i * 4 + 4 <= record.Data.Length); i++)
                {
                    byte r = record.Data[2 + i * 4];
                    byte g = record.Data[2 + i * 4 + 1];
                    byte b = record.Data[2 + i * 4 + 2];
                    _palette[8 + i] = $"#{r:X2}{g:X2}{b:X2}";
                }
            }
        }

        private Row ParseRowRecord(BiffRecord record)
        {
            var row = new Row();
            row.Cells.Capacity = 100;
            
            if (record.Data != null && record.Data.Length >= 2)
            {
                row.RowIndex = BitConverter.ToUInt16(record.Data, 0) + 1; // 转为1-based
                
                // BIFF8 ROW 记录格式: row(2) + firstCol(2) + lastCol(2) + height(2) + ...
                if (record.Data.Length >= 8)
                {
                    ushort rawHeight = BitConverter.ToUInt16(record.Data, 6);
                    // BIFF8 spec: bit 15 of miHeight is fGhost (1 = default height, 0 = custom height)
                    row.Height = (ushort)(rawHeight & 0x7FFF);
                    row.CustomHeight = (rawHeight & 0x8000) == 0;
                    
                    // Option flags at offset 12
                    if (record.Data.Length >= 14)
                    {
                        ushort options = BitConverter.ToUInt16(record.Data, 12);
                        // bit 6: fDyZero (hidden row)
                        if ((options & 0x0040) != 0) 
                        {
                            row.CustomHeight = true;
                            row.Height = 0;
                        }
                    }
                }
            }
            else
            {
                row.RowIndex = 1;
            }
            
            return row;
        }

        private Cell ParseCellRecord(BiffRecord record)
        {
            // 解析单元格记录
            var cell = new Cell();
            
            if (record.Data != null && record.Data.Length >= 6)
            {
                // 读取行索引
                ushort rowIndex = BitConverter.ToUInt16(record.Data, 0);
                cell.RowIndex = rowIndex + 1; // 转为1-based
                
                // 读取列索引（从1开始）
                ushort colIndex = BitConverter.ToUInt16(record.Data, 2);
                cell.ColumnIndex = colIndex + 1;
                
                // 读取样式索引 (XF index)
                ushort styleIndex = BitConverter.ToUInt16(record.Data, 4);
                if (styleIndex > 0)
                {
                    cell.StyleId = styleIndex.ToString();
                }
                
                // 根据记录类型解析单元格值
                switch (record.Id)
                {
                    case (ushort)BiffRecordType.CELL_BLANK:
                        // 空单元格
                        cell.Value = null;
                        break;
                    case (ushort)BiffRecordType.CELL_BOOLERR:
                        // BOOLERR 记录：根据标志字节区分布尔值和错误值
                        if (record.Data.Length >= 8)
                        {
                            byte valueOrError = record.Data[6];
                            byte isError = record.Data[7]; // 0 = 布尔值, 1 = 错误值
                            if (isError == 0)
                            {
                                cell.Value = valueOrError != 0;
                                cell.DataType = "b";
                            }
                            else
                            {
                                cell.Value = GetErrorString(valueOrError);
                                cell.DataType = "e";
                            }
                        }
                        break;
                    case (ushort)BiffRecordType.CELL_LABEL:
                        // 文本值 (BIFF8 LABEL record uses XLUnicodeString)
                        if (record.Data.Length > 6)
                        {
                            int offset = 6;
                            cell.Value = ReadBiffString(record.Data, ref offset);
                            cell.DataType = "inlineStr";
                        }
                        break;
                    case (ushort)BiffRecordType.CELL_RICH_TEXT:
                        // 富文本值
                        if (record.Data.Length > 6)
                        {
                            // 解析富文本格式
                            cell.RichText = ParseRichText(record.Data, 6);
                            cell.DataType = "inlineStr";
                            // 同时设置Value为纯文本，确保兼容性
                            cell.Value = string.Join("", cell.RichText.Select(r => r.Text));
                        }
                        break;
                    case (ushort)BiffRecordType.CELL_LABELSST:
                        // 共享字符串表中的索引
                        if (record.Data.Length >= 8)
                        {
                            int sstIndex = BitConverter.ToInt32(record.Data, 6);
                            if (sstIndex >= 0 && sstIndex < _sharedStrings.Count)
                            {
                                cell.Value = _sharedStrings[sstIndex];
                            }
                            cell.DataType = "s";
                        }
                        break;
                    case (ushort)BiffRecordType.CELL_NUMBER:
                        // 数值
                        if (record.Data.Length >= 14)
                        {
                            double value = BitConverter.ToDouble(record.Data, 6);
                            // 检查是否为日期时间值（Excel 日期时间是从 1900-01-01 开始的天数）
                            if (IsDateTimeValue(value))
                            {
                                cell.Value = ExcelDateToDateTime(value);
                                cell.DataType = "d";
                            }
                            else
                            {
                                cell.Value = value;
                                cell.DataType = "n";
                            }
                        }
                        break;
                    case (ushort)BiffRecordType.CELL_RK:
                        // 压缩数值
                        if (record.Data.Length >= 10)
                        {
                            int rkValue = BitConverter.ToInt32(record.Data, 6);
                            double value = DecodeRKValue(rkValue);
                            // 检查是否为日期时间值
                            if (IsDateTimeValue(value))
                            {
                                cell.Value = ExcelDateToDateTime(value);
                                cell.DataType = "d";
                            }
                            else
                            {
                                cell.Value = value;
                                cell.DataType = "n";
                            }
                        }
                        break;
                    case (ushort)BiffRecordType.CELL_FORMULA:
                        // 公式
                        try
                        {
                            if (record.Data.Length >= 20)
                            {
                                // 读取公式结果
                                double result = BitConverter.ToDouble(record.Data, 6);
                                
                                // 读取公式字符串长度
                                int formulaLength = BitConverter.ToUInt16(record.Data, 20); // Length of Ptgs is at offset 20 usually
                                
                                // 读取公式 Ptgs
                                if (record.Data.Length >= 22 + formulaLength)
                                {
                                    byte[] ptgs = new byte[formulaLength];
                                    Array.Copy(record.Data, 22, ptgs, 0, formulaLength);
                                    string formula = FormulaDecompiler.Decompile(ptgs);
                                    
                                    cell.Formula = formula;
                                    
                                    // 处理公式结果
                                    if (IsDateTimeValue(result))
                                    {
                                        cell.Value = ExcelDateToDateTime(result);
                                        cell.DataType = "d";
                                    }
                                    else if (IsErrorValue(result))
                                    {
                                        cell.Value = GetErrorString((byte)result);
                                        cell.DataType = "e";
                                    }
                                    else
                                    {
                                        cell.Value = result;
                                        cell.DataType = "f";
                                    }
                                }
                                else
                                {
                                    // 如果无法读取公式字符串，使用结果值
                                    if (IsDateTimeValue(result))
                                    {
                                        cell.Value = ExcelDateToDateTime(result);
                                        cell.DataType = "d";
                                    }
                                    else if (IsErrorValue(result))
                                    {
                                        cell.Value = GetErrorString((byte)result);
                                        cell.DataType = "e";
                                    }
                                    else
                                    {
                                        cell.Value = result;
                                        cell.DataType = "n";
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Error("解析公式时发生错误", ex);
                            // 继续处理，设置默认值
                            cell.Value = "#ERROR!";
                            cell.DataType = "e";
                        }
                        break;
                }
            }
            else
            {
                cell.ColumnIndex = 1;
                cell.Value = "";
            }
            
            return cell;
        }
        
        private string GetErrorString(byte errorCode)
        {
            // 映射错误代码到错误字符串
            switch (errorCode)
            {
                case 0x00: return "#NULL!";
                case 0x07: return "#DIV/0!";
                case 0x0F: return "#VALUE!";
                case 0x17: return "#REF!";
                case 0x1D: return "#NAME?";
                case 0x24: return "#NUM!";
                case 0x2A: return "#N/A";
                default: return "#ERROR!";
            }
        }
        
        private double DecodeRKValue(int rkValue)
        {
            // 解码RK值（压缩数值）
            // bit 0: 1 = 除以100, 0 = 不变
            // bit 1: 1 = 30位有符号整数, 0 = IEEE 754 double的前30位
            double value;
            if ((rkValue & 0x02) != 0)
            {
                // 30位整数
                value = (double)(rkValue >> 2);
            }
            else
            {
                // IEEE double 的前30位
                long bits = (long)(rkValue & 0xFFFFFFFC) << 32;
                value = BitConverter.Int64BitsToDouble(bits);
            }

            if ((rkValue & 0x01) != 0)
            {
                value /= 100.0;
            }

            return value;
        }
        
        private bool IsDateTimeValue(double value)
        {
            // 检查是否为日期时间值
            // Excel 日期时间范围通常在 25569（1970-01-01）到 44197（2020-12-31）之间
            // 扩展范围以覆盖更多可能的日期
            return value >= 25569 && value <= 730485; // 从1970-01-01到9999-12-31
        }

        private bool IsErrorValue(double value)
        {
            // 判断是否为错误值（Excel错误值通常是特殊的整数值）
            int intValue = (int)value;
            return intValue >= 0 && intValue <= 0x2A && 
                   (intValue == 0x00 || // #NULL!
                    intValue == 0x07 || // #DIV/0!
                    intValue == 0x0F || // #VALUE!
                    intValue == 0x17 || // #REF!
                    intValue == 0x1D || // #NAME?
                    intValue == 0x24 || // #NUM!
                    intValue == 0x2A);  // #N/A
        }
        
        private DateTime ExcelDateToDateTime(double excelDate)
        {
            // 将 Excel 日期时间值转换为 .NET DateTime
            // Excel 日期时间是从 1900-01-01 开始的天数
            // 注意：Excel 使用 1900 年 2 月 29 日作为有效日期，即使 1900 年不是闰年
            DateTime excelBaseDate = new DateTime(1900, 1, 1);
            
            // 调整 1900 年闰年问题
            if (excelDate >= 60)
            {
                excelDate -= 1;
            }
            
            return excelBaseDate.AddDays(excelDate);
        }
        
        private void ParseMergeCellsRecord(BiffRecord record, Worksheet worksheet)
        {
            // 解析合并单元格记录
            if (record.Data != null && record.Data.Length >= 8)
            {
                // 读取合并单元格的范围
                ushort startRow = BitConverter.ToUInt16(record.Data, 0);
                ushort startCol = BitConverter.ToUInt16(record.Data, 2);
                ushort endRow = BitConverter.ToUInt16(record.Data, 4);
                ushort endCol = BitConverter.ToUInt16(record.Data, 6);
                
                // 验证合并单元格范围是否有效
                if (startRow > 0 && endRow > 0 && startRow <= endRow && startCol <= endCol)
                {
                    // 创建合并单元格对象
                    var mergeCell = new MergeCell
                    {
                        StartRow = startRow + 1, // 转为1-based
                        StartColumn = startCol + 1, // 转换为从1开始的索引
                        EndRow = endRow + 1, // 转为1-based
                        EndColumn = endCol + 1 // 转换为从1开始的索引
                    };
                    
                    // 添加到工作表的合并单元格列表
                    worksheet.MergeCells.Add(mergeCell);
                }
            }
        }
        
        private void ParseVbaStream(Workbook workbook)
        {
            // 解析VBA流
            if (_vbaStreamStart > 0 && _vbaStreamSize > 0)
            {
                try
                {
                    // 验证VBA流大小
                    if (_vbaStreamSize > VbaSizeLimit)
                    {
                        Logger.Warn($"VBA项目大小超过{VbaSizeLimit / (1024 * 1024)}MB限制，可能导致转换失败");
                    }
                    
                    // 定位到VBA流
                    _stream.Seek(_vbaStreamStart * _sectorSize, SeekOrigin.Begin);
                    
                    // 读取VBA流数据
                    byte[] vbaData = new byte[_vbaStreamSize];
                    int bytesRead = _reader.Read(vbaData, 0, (int)_vbaStreamSize);
                    if (bytesRead == _vbaStreamSize)
                    {
                        // 验证VBA流头部
                        if (vbaData.Length >= 8)
                        {
                            // 检查VBA流头部签名
                            uint signature = BitConverter.ToUInt32(vbaData, 0);
                            if (signature == 0x61CC61CC) // VBA流签名
                            {
                                workbook.VbaProject = vbaData;
                                Logger.Info($"成功解析VBA项目，大小: {vbaData.Length} 字节");
                            }
                            else
                            {
                                Logger.Warn("VBA流头部签名无效，跳过VBA项目");
                            }
                        }
                        else
                        {
                            Logger.Warn("VBA流数据太短，跳过VBA项目");
                        }
                    }
                    else
                    {
                        Logger.Warn($"VBA流读取不完整，预期: {_vbaStreamSize} 字节，实际: {bytesRead} 字节");
                    }
                }
                catch (Exception ex)
                {
                    Logger.Error("解析VBA流时发生错误", ex);
                    // 继续执行，不影响其他部分的解析
                }
            }
        }

        private void ParseNameRecord(BiffRecord record, Workbook workbook)
        {
            // NAME 记录 (0x0018) - BIFF8
            if (record.Data != null && record.Data.Length >= 14)
            {
                ushort options = BitConverter.ToUInt16(record.Data, 0);
                byte nameLen = record.Data[3];
                ushort formulaLen = BitConverter.ToUInt16(record.Data, 4);
                
                bool hidden = (options & 0x0001) != 0;
                int localSheetId = (options >> 5) & 0x0FFF;
                
                int offset = 14;
                string name = ReadBiffStringFromBytes(record.Data, ref offset, nameLen);
                
                byte[] formulaData = new byte[formulaLen];
                if (offset + formulaLen <= record.Data.Length)
                {
                    Array.Copy(record.Data, offset, formulaData, 0, formulaLen);
                }
                
                string formula = FormulaDecompiler.Decompile(formulaData);
                
                var definedName = new DefinedName
                {
                    Name = name,
                    Formula = formula,
                    Hidden = hidden,
                    LocalSheetId = localSheetId > 0 ? (int?)(localSheetId - 1) : null
                };
                
                workbook.DefinedNames.Add(definedName);
            }
        }

        private void ParsePageSetupRecord(BiffRecord record, Worksheet worksheet)
        {
            // PAGESETUP (0x00A1) - BIFF8
            if (record.Data != null && record.Data.Length >= 34)
            {
                var ps = worksheet.PageSettings;
                ps.PaperSize = BitConverter.ToUInt16(record.Data, 0);
                ps.Scale = BitConverter.ToUInt16(record.Data, 2);
                ps.FitToWidth = BitConverter.ToUInt16(record.Data, 6);
                ps.FitToHeight = BitConverter.ToUInt16(record.Data, 8);
                
                ushort options = BitConverter.ToUInt16(record.Data, 10);
                ps.OrientationLandscape = (options & 0x0002) == 0;
                ps.UsePageNumbers = (options & 0x0001) != 0;
            }
        }

        private string GetBorderLineStyle(byte styleId)
        {
            switch (styleId)
            {
                case 0: return "none";
                case 1: return "thin";
                case 2: return "medium";
                case 3: return "dashed";
                case 4: return "dotted";
                case 5: return "thick";
                case 6: return "double";
                case 7: return "hair";
                case 8: return "mediumDashed";
                case 9: return "dashDot";
                case 10: return "mediumDashDot";
                case 11: return "dashDotDot";
                case 12: return "mediumDashDotDot";
                case 13: return "slantDashDot";
                default: return "none";
            }
        }

        private string? GetColorFromPalette(int colorIndex)
        {
            if (colorIndex == 64) return null; // System Foreground
            if (colorIndex == 65) return null; // System Background
            
            if (_workbook.Palette.TryGetValue(colorIndex, out string? color))
                return color.Replace("#", "");
            
            if (_palette.TryGetValue(colorIndex, out string? palColor))
                return palColor.Replace("#", "");
                
            return null;
        }

        private string GetPatternType(byte patternId)
        {
            switch (patternId)
            {
                case 0: return "none";
                case 1: return "solid";
                case 2: return "mediumGray";
                case 3: return "darkGray";
                case 4: return "lightGray";
                case 5: return "darkHorizontal";
                case 6: return "darkVertical";
                case 7: return "darkDown";
                case 8: return "darkUp";
                case 9: return "darkGrid";
                case 10: return "darkTrellis";
                case 11: return "lightHorizontal";
                case 12: return "lightVertical";
                case 13: return "lightDown";
                case 14: return "lightUp";
                case 15: return "lightGrid";
                case 16: return "lightTrellis";
                case 17: return "gray125";
                case 18: return "gray0625";
                default: return "none";
            }
        }
    }
}