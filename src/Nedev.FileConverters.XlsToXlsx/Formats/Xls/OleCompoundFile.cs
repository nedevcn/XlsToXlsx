using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Nedev.FileConverters.XlsToXlsx.Formats.Xls
{
    /// <summary>
    /// OLE Compound Binary File 格式解析器。
    /// 实现完整的 FAT 链追踪、DIFAT、Mini FAT / Mini Stream，以及目录树解析。
    /// 参考规范: [MS-CFB] Compound File Binary File Format
    /// </summary>
    public class OleCompoundFile
    {
        // ===== 常量 =====
        private static readonly byte[] Signature = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
        private const uint ENDOFCHAIN = 0xFFFFFFFE;
        private const uint FREESECT = 0xFFFFFFFF;
        private const uint FATSECT = 0xFFFFFFFD;
        private const uint DIFSECT = 0xFFFFFFFC;
        private const int HEADER_SIZE = 512;
        private const int DIRECTORY_ENTRY_SIZE = 128;
        private const int HEADER_DIFAT_COUNT = 109;

        // ===== 头部字段 =====
        private int _sectorSize;
        private int _miniSectorSize;
        private int _miniStreamCutoffSize;

        private int _fatSectorCount;
        private int _firstDirectorySector;
        private int _firstMiniFatSector;
        private int _miniFatSectorCount;
        private int _firstDifatSector;
        private int _difatSectorCount;

        // 头部中的 109 个 DIFAT 条目
        private int[] _headerDifat = new int[HEADER_DIFAT_COUNT];

        // ===== 运行时数据 =====
        private Stream _stream;
        private int[] _fat = Array.Empty<int>();             // 完整 FAT
        private int[] _miniFat = Array.Empty<int>();         // Mini FAT
        private byte[] _miniStream = Array.Empty<byte>();    // Mini Stream (Root Entry 的流数据)
        private List<DirectoryEntry> _directoryEntries = new List<DirectoryEntry>();

        /// <summary>
        /// 所有目录条目
        /// </summary>
        public IReadOnlyList<DirectoryEntry> DirectoryEntries => _directoryEntries;

        /// <summary>
        /// 扇区大小
        /// </summary>
        public int SectorSize => _sectorSize;

        public OleCompoundFile(Stream stream)
        {
            _stream = stream ?? throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead || !stream.CanSeek)
                throw new ArgumentException("Stream must be readable and seekable");

            ParseHeader();
            BuildFat();
            ParseDirectoryEntries();
            BuildMiniStream();
            BuildMiniFat();
        }

        // ===== 头部解析 =====

        private void ParseHeader()
        {
            _stream.Seek(0, SeekOrigin.Begin);
            byte[] header = new byte[HEADER_SIZE];
            int read = _stream.Read(header, 0, HEADER_SIZE);
            if (read < HEADER_SIZE)
                throw new InvalidDataException("文件太小，不是有效的 OLE 复合文件");

            // 验证签名
            for (int i = 0; i < 8; i++)
            {
                if (header[i] != Signature[i])
                    throw new InvalidDataException("无效的 OLE 复合文件签名");
            }

            // 扇区大小: 2^(header[0x1E:0x1F])，仅支持 512 (2^9) 或 4096 (2^12)
            ushort sectorShift = BitConverter.ToUInt16(header, 0x1E);
            if (sectorShift != 9 && sectorShift != 12)
                sectorShift = 9;
            _sectorSize = 1 << sectorShift;

            // Mini 扇区大小: 2^(header[0x20:0x21])，通常 64，防止异常值
            ushort miniSectorShift = BitConverter.ToUInt16(header, 0x20);
            if (miniSectorShift > 31) miniSectorShift = 6;
            _miniSectorSize = 1 << miniSectorShift;

            // FAT 扇区总数
            _fatSectorCount = BitConverter.ToInt32(header, 0x2C);

            // 目录流起始扇区
            _firstDirectorySector = BitConverter.ToInt32(header, 0x30);

            // Mini Stream 截断大小（小于此值的流存储在 Mini Stream 中）
            _miniStreamCutoffSize = BitConverter.ToInt32(header, 0x38);

            // Mini FAT 起始扇区
            _firstMiniFatSector = BitConverter.ToInt32(header, 0x3C);

            // Mini FAT 扇区数
            _miniFatSectorCount = BitConverter.ToInt32(header, 0x40);

            // DIFAT 起始扇区
            _firstDifatSector = BitConverter.ToInt32(header, 0x44);

            // DIFAT 扇区数
            _difatSectorCount = BitConverter.ToInt32(header, 0x48);

            // 读取头部中的 109 个 DIFAT 条目 (偏移 0x4C 开始，每个 4 字节)
            for (int i = 0; i < HEADER_DIFAT_COUNT; i++)
            {
                _headerDifat[i] = BitConverter.ToInt32(header, 0x4C + i * 4);
            }

            Logger.Debug($"OLE头部: sectorSize={_sectorSize}, miniSectorSize={_miniSectorSize}, " +
                         $"fatSectors={_fatSectorCount}, dirSector={_firstDirectorySector}, " +
                         $"miniFatSector={_firstMiniFatSector}, difatSector={_firstDifatSector}");
        }

        // ===== FAT 构建 =====

        private void BuildFat()
        {
            // 1. 收集所有 FAT 扇区编号
            var fatSectorIds = new List<int>();

            // 从头部 DIFAT 取前 _fatSectorCount 个（最多 109 个）
            int fromHeader = Math.Min(_fatSectorCount, HEADER_DIFAT_COUNT);
            for (int i = 0; i < fromHeader; i++)
            {
                int sid = _headerDifat[i];
                if (sid >= 0 && (uint)sid < ENDOFCHAIN)
                    fatSectorIds.Add(sid);
            }

            // 如果 FAT 扇区数 > 109，使用 DIFAT 链获取剩余的 FAT 扇区编号
            if (_fatSectorCount > HEADER_DIFAT_COUNT && _firstDifatSector >= 0 && (uint)_firstDifatSector < ENDOFCHAIN)
            {
                int difatSector = _firstDifatSector;
                int remaining = _fatSectorCount - HEADER_DIFAT_COUNT;
                int maxDifatChain = _difatSectorCount > 0 ? _difatSectorCount : 1000; // 安全上限
                int difatCount = 0;

                while (remaining > 0 && difatSector >= 0 && (uint)difatSector < ENDOFCHAIN && difatCount < maxDifatChain)
                {
                    byte[] difatData = ReadSectorRaw(difatSector);
                    int entriesPerSector = _sectorSize / 4 - 1; // 最后 4 字节是下一个 DIFAT 扇区

                    for (int i = 0; i < entriesPerSector && remaining > 0; i++)
                    {
                        int sid = BitConverter.ToInt32(difatData, i * 4);
                        if (sid >= 0 && (uint)sid < ENDOFCHAIN)
                        {
                            fatSectorIds.Add(sid);
                            remaining--;
                        }
                    }

                    // 下一个 DIFAT 扇区
                    difatSector = BitConverter.ToInt32(difatData, _sectorSize - 4);
                    difatCount++;
                }
            }

            // 2. 读取所有 FAT 扇区并拼接为完整 FAT
            int entriesTotal = fatSectorIds.Count * (_sectorSize / 4);
            _fat = new int[entriesTotal];

            for (int i = 0; i < fatSectorIds.Count; i++)
            {
                byte[] sectorData = ReadSectorRaw(fatSectorIds[i]);
                int entriesInSector = _sectorSize / 4;

                for (int j = 0; j < entriesInSector; j++)
                {
                    _fat[i * entriesInSector + j] = BitConverter.ToInt32(sectorData, j * 4);
                }
            }

            Logger.Debug($"FAT 构建完成: {_fat.Length} 个条目, 来自 {fatSectorIds.Count} 个 FAT 扇区");
        }

        // ===== 目录解析 =====

        private void ParseDirectoryEntries()
        {
            // 沿 FAT 链读取目录流
            byte[] directoryData = ReadStreamFromFat(_firstDirectorySector);
            int entryCount = directoryData.Length / DIRECTORY_ENTRY_SIZE;

            _directoryEntries.Clear();

            for (int i = 0; i < entryCount; i++)
            {
                int offset = i * DIRECTORY_ENTRY_SIZE;

                // 条目类型: 0=Unknown/Empty, 1=Storage, 2=Stream, 5=Root
                byte objectType = directoryData[offset + 0x42];
                if (objectType == 0)
                {
                    // 空条目，仍记录以维持索引一致
                    _directoryEntries.Add(new DirectoryEntry
                    {
                        Index = i,
                        ObjectType = DirectoryEntryType.Empty
                    });
                    continue;
                }

                // 名称长度（字节数，包含终止符）；MS-CFB 目录名最多 32 个 UTF-16 码元 = 64 字节，防止异常值越界
                ushort nameLen = BitConverter.ToUInt16(directoryData, offset + 0x40);
                int nameBytesToDecode = Math.Max(0, nameLen <= 2 ? 0 : Math.Min(nameLen - 2, 62));
                if (offset + nameBytesToDecode > directoryData.Length)
                    nameBytesToDecode = Math.Max(0, directoryData.Length - offset);
                string name = nameBytesToDecode > 0
                    ? Encoding.Unicode.GetString(directoryData, offset, nameBytesToDecode)
                    : string.Empty;

                int startSector = BitConverter.ToInt32(directoryData, offset + 0x74);
                long streamSize;

                // v4 文件使用 8 字节大小，v3 文件只用前 4 字节
                if (_sectorSize == 4096) // v4
                {
                    streamSize = BitConverter.ToInt64(directoryData, offset + 0x78);
                }
                else // v3
                {
                    streamSize = BitConverter.ToUInt32(directoryData, offset + 0x78);
                }

                var entryType = objectType switch
                {
                    1 => DirectoryEntryType.Storage,
                    2 => DirectoryEntryType.Stream,
                    5 => DirectoryEntryType.RootEntry,
                    _ => DirectoryEntryType.Empty
                };

                // 子节点索引（红黑树）
                int childId = BitConverter.ToInt32(directoryData, offset + 0x4C);
                int leftSiblingId = BitConverter.ToInt32(directoryData, offset + 0x44);
                int rightSiblingId = BitConverter.ToInt32(directoryData, offset + 0x48);

                var entry = new DirectoryEntry
                {
                    Index = i,
                    Name = name,
                    ObjectType = entryType,
                    StartSector = startSector,
                    StreamSize = streamSize,
                    ChildId = childId,
                    LeftSiblingId = leftSiblingId,
                    RightSiblingId = rightSiblingId,
                };

                _directoryEntries.Add(entry);

                Logger.Debug($"目录条目[{i}]: 名称='{name}', 类型={entryType}, 起始扇区={startSector}, 大小={streamSize}");
            }

            Logger.Info($"目录解析完成: {_directoryEntries.Count} 个条目");
        }

        // ===== Mini Stream / Mini FAT =====

        private void BuildMiniStream()
        {
            // Root Entry (index 0) 的流数据就是 Mini Stream 容器
            if (_directoryEntries.Count == 0) return;
            var rootEntry = _directoryEntries[0];
            if (rootEntry.ObjectType != DirectoryEntryType.RootEntry) return;
            if (rootEntry.StartSector < 0 || (uint)rootEntry.StartSector >= ENDOFCHAIN) return;

            // 用普通 FAT 链读取 Root Entry 的流（即 Mini Stream）
            _miniStream = ReadStreamFromFat(rootEntry.StartSector, rootEntry.StreamSize);
            Logger.Debug($"Mini Stream 构建完成: {_miniStream.Length} 字节");
        }

        private void BuildMiniFat()
        {
            if (_firstMiniFatSector < 0 || (uint)_firstMiniFatSector >= ENDOFCHAIN || _miniFatSectorCount <= 0)
            {
                _miniFat = Array.Empty<int>();
                return;
            }

            // 用普通 FAT 链读取 Mini FAT
            byte[] miniFatData = ReadStreamFromFat(_firstMiniFatSector);
            int count = miniFatData.Length / 4;
            _miniFat = new int[count];

            for (int i = 0; i < count; i++)
            {
                _miniFat[i] = BitConverter.ToInt32(miniFatData, i * 4);
            }

            Logger.Debug($"Mini FAT 构建完成: {_miniFat.Length} 个条目");
        }

        // ===== 公共 API =====

        /// <summary>
        /// 按名称查找并读取流数据。
        /// 名称匹配不区分大小写。
        /// </summary>
        public byte[]? ReadStreamByName(string name)
        {
            var entry = _directoryEntries.FirstOrDefault(
                e => e.ObjectType == DirectoryEntryType.Stream &&
                     string.Equals(e.Name, name, StringComparison.OrdinalIgnoreCase));

            if (entry == null) return null;
            return ReadStream(entry);
        }

        /// <summary>
        /// 读取指定目录条目的流数据。
        /// 自动判断使用普通 FAT 还是 Mini FAT。
        /// </summary>
        public byte[] ReadStream(DirectoryEntry entry)
        {
            if (entry.ObjectType == DirectoryEntryType.Empty)
                return Array.Empty<byte>();

            if (entry.StartSector < 0 || (uint)entry.StartSector >= ENDOFCHAIN)
                return Array.Empty<byte>();

            long size = entry.StreamSize;
            if (size <= 0)
                return Array.Empty<byte>();

            // Root Entry 始终使用普通 FAT
            if (entry.ObjectType == DirectoryEntryType.RootEntry)
            {
                return ReadStreamFromFat(entry.StartSector, size);
            }

            // 小流使用 Mini FAT，大流使用普通 FAT
            if (size < _miniStreamCutoffSize && _miniFat.Length > 0 && _miniStream.Length > 0)
            {
                return ReadStreamFromMiniFat(entry.StartSector, size);
            }
            else
            {
                return ReadStreamFromFat(entry.StartSector, size);
            }
        }

        /// <summary>
        /// 查找指定存储下的所有子条目。
        /// </summary>
        public List<DirectoryEntry> GetChildEntries(DirectoryEntry storageEntry)
        {
            var result = new List<DirectoryEntry>();
            if (storageEntry.ChildId < 0 || storageEntry.ChildId >= _directoryEntries.Count)
                return result;

            // 遍历红黑树（中序遍历收集所有节点）
            CollectSiblings(_directoryEntries[storageEntry.ChildId], result);
            return result;
        }

        /// <summary>
        /// 在目录中按路径查找条目，如 "VBA/dir"。
        /// </summary>
        public DirectoryEntry? FindEntry(string path)
        {
            var parts = path.Split('/', '\\');
            DirectoryEntry? current = _directoryEntries.Count > 0 ? _directoryEntries[0] : null; // Root

            foreach (var part in parts)
            {
                if (current == null) return null;
                var children = GetChildEntries(current);
                current = children.FirstOrDefault(
                    c => string.Equals(c.Name, part, StringComparison.OrdinalIgnoreCase));
            }

            return current;
        }

        // ===== 内部流读取方法 =====

        /// <summary>
        /// 从物理扇区直接读取原始数据（不经过 FAT 链）
        /// </summary>
        private byte[] ReadSectorRaw(int sectorId)
        {
            if (sectorId < 0)
                return new byte[_sectorSize];
            long offset = HEADER_SIZE + (long)sectorId * _sectorSize;
            _stream.Seek(offset, SeekOrigin.Begin);
            byte[] data = new byte[_sectorSize];
            int read = _stream.Read(data, 0, _sectorSize);
            if (read < _sectorSize)
            {
                // 文件末尾，返回已读取的部分
                Array.Resize(ref data, read);
            }
            return data;
        }

        /// <summary>
        /// 沿 FAT 链读取流数据，不限制大小（用于目录流等内部流）。
        /// </summary>
        private byte[] ReadStreamFromFat(int startSector)
        {
            return ReadStreamFromFat(startSector, long.MaxValue);
        }

        /// <summary>
        /// 沿 FAT 链读取流数据，限制为指定大小。
        /// </summary>
        private byte[] ReadStreamFromFat(int startSector, long maxSize)
        {
            if (startSector < 0 || (uint)startSector >= ENDOFCHAIN)
                return Array.Empty<byte>();
            if (_fat == null || _fat.Length == 0)
                return Array.Empty<byte>();

            // 收集扇区链
            var sectorChain = new List<int>();
            int current = startSector;
            int maxChainLength = _fat.Length + 1; // 安全上限

            while (current >= 0 && (uint)current < ENDOFCHAIN && sectorChain.Count < maxChainLength)
            {
                sectorChain.Add(current);

                if (current >= _fat.Length)
                {
                    Logger.Warn($"FAT 链中扇区 {current} 超出 FAT 范围 ({_fat.Length})");
                    break;
                }

                current = _fat[current];
            }

            // 计算实际大小
            long totalAvailable = (long)sectorChain.Count * _sectorSize;
            long actualSize = Math.Min(totalAvailable, maxSize);

            byte[] result = new byte[actualSize];
            long written = 0;

            foreach (int sectorId in sectorChain)
            {
                if (written >= actualSize) break;

                long offset = HEADER_SIZE + (long)sectorId * _sectorSize;
                _stream.Seek(offset, SeekOrigin.Begin);

                int toRead = (int)Math.Min(_sectorSize, actualSize - written);
                int read = _stream.Read(result, (int)written, toRead);
                written += read;

                if (read < toRead) break; // 文件末尾
            }

            // 如果实际读取不足，截断
            if (written < actualSize)
            {
                Array.Resize(ref result, (int)written);
            }

            return result;
        }

        /// <summary>
        /// 沿 Mini FAT 链从 Mini Stream 中读取流数据。
        /// </summary>
        private byte[] ReadStreamFromMiniFat(int startSector, long size)
        {
            if (startSector < 0 || (uint)startSector >= ENDOFCHAIN)
                return Array.Empty<byte>();

            byte[] result = new byte[size];
            long written = 0;
            int current = startSector;
            int maxChainLength = _miniFat.Length + 1;
            int chainCount = 0;

            while (current >= 0 && (uint)current < ENDOFCHAIN && written < size && chainCount < maxChainLength)
            {
                long miniOffset = (long)current * _miniSectorSize;
                int toRead = (int)Math.Min(_miniSectorSize, size - written);

                if (miniOffset + toRead <= _miniStream.Length)
                {
                    Array.Copy(_miniStream, miniOffset, result, written, toRead);
                }
                else if (miniOffset < _miniStream.Length)
                {
                    int available = (int)(_miniStream.Length - miniOffset);
                    Array.Copy(_miniStream, miniOffset, result, written, available);
                    toRead = available;
                }
                else
                {
                    break; // Mini Stream 越界
                }

                written += toRead;

                if (current >= _miniFat.Length)
                {
                    Logger.Warn($"Mini FAT 链中扇区 {current} 超出 Mini FAT 范围 ({_miniFat.Length})");
                    break;
                }

                current = _miniFat[current];
                chainCount++;
            }

            if (written < size)
            {
                Array.Resize(ref result, (int)written);
            }

            return result;
        }

        /// <summary>
        /// 递归遍历红黑树的兄弟节点
        /// </summary>
        private void CollectSiblings(DirectoryEntry node, List<DirectoryEntry> result)
        {
            if (node == null) return;

            // 左兄弟
            if (node.LeftSiblingId >= 0 && node.LeftSiblingId < _directoryEntries.Count)
            {
                CollectSiblings(_directoryEntries[node.LeftSiblingId], result);
            }

            result.Add(node);

            // 右兄弟
            if (node.RightSiblingId >= 0 && node.RightSiblingId < _directoryEntries.Count)
            {
                CollectSiblings(_directoryEntries[node.RightSiblingId], result);
            }
        }
    }

    /// <summary>
    /// OLE 目录条目类型
    /// </summary>
    public enum DirectoryEntryType
    {
        Empty = 0,
        Storage = 1,
        Stream = 2,
        RootEntry = 5
    }

    /// <summary>
    /// OLE 目录条目
    /// </summary>
    public class DirectoryEntry
    {
        /// <summary>目录条目索引</summary>
        public int Index { get; set; }
        /// <summary>条目名称</summary>
        public string Name { get; set; } = string.Empty;
        /// <summary>条目类型</summary>
        public DirectoryEntryType ObjectType { get; set; }
        /// <summary>流起始扇区</summary>
        public int StartSector { get; set; }
        /// <summary>流大小</summary>
        public long StreamSize { get; set; }
        /// <summary>子条目 ID（红黑树根）</summary>
        public int ChildId { get; set; } = -1;
        /// <summary>左兄弟 ID</summary>
        public int LeftSiblingId { get; set; } = -1;
        /// <summary>右兄弟 ID</summary>
        public int RightSiblingId { get; set; } = -1;

        public override string ToString() => $"{ObjectType}: '{Name}' (sector={StartSector}, size={StreamSize})";
    }
}
