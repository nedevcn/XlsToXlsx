using System;
using System.Collections.Generic;
using System.IO;
using Nedev.XlsToXlsx;

namespace Nedev.XlsToXlsx.Formats.Xls.Escher
{
    public class EscherRecord
    {
        public int Version { get; set; }
        public int Instance { get; set; }
        public int Type { get; set; }
        public int Length { get; set; }
        public byte[]? Data { get; set; }
        public List<EscherRecord> Children { get; set; } = new List<EscherRecord>();

        public bool IsContainer => Version == 0x0F;
    }

    /// <summary>
    /// Lightweight scanner for Escher (OfficeArt) streams to extract Blip (Image) Data and Anchors.
    /// </summary>
    public class EscherParser
    {
        public const int DggContainer = 0xF000;
        public const int BstoreContainer = 0xF001;
        public const int DgContainer = 0xF002;
        public const int SpgrContainer = 0xF003;
        public const int SpContainer = 0xF004;
        
        public const int BSE = 0xF007;          // Blip Store Entry (Image Info)
        public const int ClientAnchor = 0xF010; // Pixel coordinate mapping
        public const int ClientData = 0xF011;   // Binds shape to Obj record

        /// <summary>
        /// Reads an Escher structure recursively from a concatenated byte array.
        /// </summary>
        public static List<EscherRecord> ParseStream(byte[] data)
        {
            var records = new List<EscherRecord>();
            if (data == null || data.Length == 0) return records;

            using (var stream = new MemoryStream(data))
            using (var reader = new BinaryReader(stream))
            {
                ReadRecords(reader, data.Length, records);
            }
            return records;
        }

        private static void ReadRecords(BinaryReader reader, long endPosition, List<EscherRecord> records)
        {
            while (reader.BaseStream.Position + 8 <= endPosition)
            {
                var record = new EscherRecord();
                
                // 1. Read Header
                ushort temp = reader.ReadUInt16();
                record.Version = temp & 0x0F;
                record.Instance = temp >> 4;
                record.Type = reader.ReadUInt16();
                record.Length = reader.ReadInt32();

                // Validation safeguard against malformed lengths
                if (record.Length < 0 || reader.BaseStream.Position + record.Length > reader.BaseStream.Length)
                {
                    break;
                }

                long dataStartPos = reader.BaseStream.Position;

                // 2. Container vs Atomic
                if (record.IsContainer)
                {
                    // Recursively parse children within this length bound
                    ReadRecords(reader, dataStartPos + record.Length, record.Children);
                }
                else
                {
                    // Read atomic payload
                    record.Data = reader.ReadBytes(record.Length);
                }

                // Ensure stream pointer correctly advances even if children reading failed
                reader.BaseStream.Position = dataStartPos + record.Length;
                
                records.Add(record);
            }
        }
    }
}
