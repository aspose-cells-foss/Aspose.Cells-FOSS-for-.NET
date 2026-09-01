using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace Aspose.Cells_FOSS.Rendering
{
    internal sealed class TrueTypeSubsetFontProgram
    {
        private const ushort CompoundGlyphFlagArg1And2AreWords = 0x0001;
        private const ushort CompoundGlyphFlagMoreComponents = 0x0020;
        private const ushort CompoundGlyphFlagWeHaveAScale = 0x0008;
        private const ushort CompoundGlyphFlagWeHaveAnXAndYScale = 0x0040;
        private const ushort CompoundGlyphFlagWeHaveATwoByTwo = 0x0080;
        private const ushort CompoundGlyphFlagWeHaveInstructions = 0x0100;

        private readonly byte[] _fontBytes;
        private readonly Dictionary<string, TableRecord> _tables;
        private readonly int _unitsPerEm;
        private readonly short _ascent;
        private readonly short _descent;
        private readonly short _lineGap;
        private readonly short _xMin;
        private readonly short _yMin;
        private readonly short _xMax;
        private readonly short _yMax;
        private readonly ushort _numberOfHMetrics;
        private readonly int _numGlyphs;
        private readonly int _indexToLocFormat;
        private readonly int[] _glyphOffsets;
        private readonly HorizontalMetric[] _metrics;

        private struct TableRecord
        {
            public string Tag;
            public uint Checksum;
            public uint Offset;
            public uint Length;
        }

        private struct HorizontalMetric
        {
            public ushort AdvanceWidth;
            public short LeftSideBearing;
        }

        public TrueTypeSubsetFontProgram(byte[] fontBytes)
        {
            if (fontBytes == null || fontBytes.Length < 12)
            {
                throw new ArgumentException("TrueType subset source bytes were invalid.", "fontBytes");
            }

            _fontBytes = fontBytes;
            _tables = ParseTableDirectory(fontBytes);
            _unitsPerEm = ReadUShort(TableBytes("head"), 18);
            _xMin = ReadShort(TableBytes("head"), 36);
            _yMin = ReadShort(TableBytes("head"), 38);
            _xMax = ReadShort(TableBytes("head"), 40);
            _yMax = ReadShort(TableBytes("head"), 42);
            _indexToLocFormat = ReadShort(TableBytes("head"), 50);
            _ascent = ReadShort(TableBytes("hhea"), 4);
            _descent = ReadShort(TableBytes("hhea"), 6);
            _lineGap = ReadShort(TableBytes("hhea"), 8);
            _numberOfHMetrics = ReadUShort(TableBytes("hhea"), 34);
            _numGlyphs = ReadUShort(TableBytes("maxp"), 4);
            _glyphOffsets = ParseGlyphOffsets();
            _metrics = ParseHorizontalMetrics();
        }

        public int UnitsPerEm
        {
            get { return _unitsPerEm; }
        }

        public short Ascent
        {
            get { return _ascent; }
        }

        public short Descent
        {
            get { return _descent; }
        }

        public short XMin
        {
            get { return _xMin; }
        }

        public short YMin
        {
            get { return _yMin; }
        }

        public short XMax
        {
            get { return _xMax; }
        }

        public short YMax
        {
            get { return _yMax; }
        }

        public byte[] BuildSubset(IList<ushort> initialGlyphIds, IDictionary<byte, ushort> codeToGlyphId, IDictionary<ushort, int> glyphIdToUnicode, string subsetFontName)
        {
            if (initialGlyphIds == null || initialGlyphIds.Count == 0)
            {
                throw new InvalidOperationException("TrueType subset requires at least one glyph.");
            }

            var closure = BuildGlyphClosure(initialGlyphIds);
            var oldToNew = new Dictionary<ushort, ushort>();
            for (var i = 0; i < closure.Count; i++)
            {
                oldToNew[closure[i]] = (ushort)i;
            }

            var glyfTable = BuildGlyfTable(closure, oldToNew);
            var locaTable = BuildLocaTable(glyfTable.LocaOffsets);
            var hmtxTable = BuildHmtxTable(closure);
            var maxpTable = PatchMaxpTable((ushort)closure.Count);
            var hheaTable = PatchHheaTable((ushort)closure.Count);
            var headTable = PatchHeadTable();
            var cmapTable = BuildCmapTable(codeToGlyphId, glyphIdToUnicode, oldToNew);

            var outputTables = new Dictionary<string, byte[]>();
            CopyTableIfPresent(outputTables, "gasp");
            outputTables["OS/2"] = BuildOs2Table();
            outputTables["cmap"] = cmapTable;
            outputTables["glyf"] = glyfTable.Bytes;
            outputTables["head"] = headTable;
            outputTables["hhea"] = hheaTable;
            outputTables["hmtx"] = hmtxTable;
            outputTables["loca"] = locaTable;
            outputTables["maxp"] = maxpTable;
            outputTables["name"] = BuildNameTable(subsetFontName);
            outputTables["post"] = BuildPostTable();

            return BuildFontFile(outputTables);
        }

        private void CopyTableIfPresent(IDictionary<string, byte[]> outputTables, string tag)
        {
            TableRecord record;
            if (_tables.TryGetValue(tag, out record))
            {
                outputTables[tag] = SliceBytes(_fontBytes, (int)record.Offset, (int)record.Length);
            }
        }

        private List<ushort> BuildGlyphClosure(IList<ushort> initialGlyphIds)
        {
            var seen = new HashSet<ushort>();
            var queue = new Queue<ushort>();
            queue.Enqueue(0);
            for (var i = 0; i < initialGlyphIds.Count; i++)
            {
                queue.Enqueue(initialGlyphIds[i]);
            }

            while (queue.Count > 0)
            {
                var glyphId = queue.Dequeue();
                if (glyphId >= _numGlyphs || !seen.Add(glyphId))
                {
                    continue;
                }

                var bytes = GetGlyphBytes(glyphId);
                if (bytes == null || bytes.Length < 10)
                {
                    continue;
                }

                var contourCount = ReadShort(bytes, 0);
                if (contourCount >= 0)
                {
                    continue;
                }

                var offset = 10;
                while (offset + 4 <= bytes.Length)
                {
                    var flags = ReadUShort(bytes, offset);
                    var componentGlyphId = ReadUShort(bytes, offset + 2);
                    queue.Enqueue(componentGlyphId);
                    offset += 4;
                    offset += (flags & CompoundGlyphFlagArg1And2AreWords) != 0 ? 4 : 2;
                    if ((flags & CompoundGlyphFlagWeHaveAScale) != 0)
                    {
                        offset += 2;
                    }
                    else if ((flags & CompoundGlyphFlagWeHaveAnXAndYScale) != 0)
                    {
                        offset += 4;
                    }
                    else if ((flags & CompoundGlyphFlagWeHaveATwoByTwo) != 0)
                    {
                        offset += 8;
                    }

                    if ((flags & CompoundGlyphFlagMoreComponents) == 0)
                    {
                        break;
                    }
                }
            }

            var ordered = new List<ushort>(seen);
            ordered.Sort();
            if (ordered.Count == 0 || ordered[0] != 0)
            {
                ordered.Insert(0, 0);
            }

            return ordered;
        }

        private GlyfBuildResult BuildGlyfTable(IList<ushort> closure, IDictionary<ushort, ushort> oldToNew)
        {
            var locaOffsets = new List<uint>();
            using (var output = new MemoryStream())
            {
                for (var i = 0; i < closure.Count; i++)
                {
                    locaOffsets.Add((uint)output.Position);
                    var originalGlyphId = closure[i];
                    var bytes = GetGlyphBytes(originalGlyphId);
                    if (bytes != null && bytes.Length > 0)
                    {
                        bytes = RewriteCompoundGlyphReferences(bytes, oldToNew);
                        output.Write(bytes, 0, bytes.Length);
                        while ((output.Position & 3) != 0)
                        {
                            output.WriteByte(0);
                        }
                    }
                }

                locaOffsets.Add((uint)output.Position);
                return new GlyfBuildResult(output.ToArray(), locaOffsets);
            }
        }

        private byte[] RewriteCompoundGlyphReferences(byte[] glyphBytes, IDictionary<ushort, ushort> oldToNew)
        {
            var copy = new byte[glyphBytes.Length];
            Buffer.BlockCopy(glyphBytes, 0, copy, 0, glyphBytes.Length);
            if (copy.Length < 10)
            {
                return copy;
            }

            var contourCount = ReadShort(copy, 0);
            if (contourCount >= 0)
            {
                return StripSimpleGlyphInstructions(copy, contourCount);
            }

            var output = new MemoryStream();
            output.Write(copy, 0, 10);

            var offset = 10;
            while (offset + 4 <= copy.Length)
            {
                var flags = ReadUShort(copy, offset);
                flags = (ushort)(flags & ~CompoundGlyphFlagWeHaveInstructions);
                var oldGlyphId = ReadUShort(copy, offset + 2);
                ushort newGlyphId;
                if (!oldToNew.TryGetValue(oldGlyphId, out newGlyphId))
                {
                    throw new InvalidOperationException("Compound glyph referenced an unmapped component glyph.");
                }

                WriteUShort(output, flags);
                WriteUShort(output, newGlyphId);
                offset += 4;
                var argumentBytes = (flags & CompoundGlyphFlagArg1And2AreWords) != 0 ? 4 : 2;
                output.Write(copy, offset, argumentBytes);
                offset += argumentBytes;
                if ((flags & CompoundGlyphFlagWeHaveAScale) != 0)
                {
                    output.Write(copy, offset, 2);
                    offset += 2;
                }
                else if ((flags & CompoundGlyphFlagWeHaveAnXAndYScale) != 0)
                {
                    output.Write(copy, offset, 4);
                    offset += 4;
                }
                else if ((flags & CompoundGlyphFlagWeHaveATwoByTwo) != 0)
                {
                    output.Write(copy, offset, 8);
                    offset += 8;
                }

                if ((flags & CompoundGlyphFlagMoreComponents) == 0)
                {
                    break;
                }
            }

            return output.ToArray();
        }

        private byte[] BuildLocaTable(IList<uint> offsets)
        {
            using (var output = new MemoryStream())
            {
                for (var i = 0; i < offsets.Count; i++)
                {
                    WriteUInt(output, offsets[i]);
                }

                return output.ToArray();
            }
        }

        private byte[] BuildHmtxTable(IList<ushort> closure)
        {
            using (var output = new MemoryStream())
            {
                for (var i = 0; i < closure.Count; i++)
                {
                    var metric = MetricForGlyph(closure[i]);
                    WriteUShort(output, metric.AdvanceWidth);
                    WriteShort(output, metric.LeftSideBearing);
                }

                return output.ToArray();
            }
        }

        private byte[] PatchMaxpTable(ushort numGlyphs)
        {
            var bytes = TableBytes("maxp");
            var copy = new byte[bytes.Length];
            Buffer.BlockCopy(bytes, 0, copy, 0, bytes.Length);
            WriteUShort(copy, 4, numGlyphs);
            return copy;
        }

        private byte[] PatchHheaTable(ushort numberOfHMetrics)
        {
            var bytes = TableBytes("hhea");
            var copy = new byte[bytes.Length];
            Buffer.BlockCopy(bytes, 0, copy, 0, bytes.Length);
            WriteUShort(copy, 34, numberOfHMetrics);
            return copy;
        }

        private byte[] PatchHeadTable()
        {
            var bytes = TableBytes("head");
            var copy = new byte[bytes.Length];
            Buffer.BlockCopy(bytes, 0, copy, 0, bytes.Length);
            WriteUInt(copy, 8, 0);
            WriteShort(copy, 50, 1);
            return copy;
        }

        private byte[] BuildCmapTable(IDictionary<byte, ushort> codeToGlyphId, IDictionary<ushort, int> glyphIdToUnicode, IDictionary<ushort, ushort> oldToNew)
        {
            var mappings = new List<CharMapping>();
            var seen = new HashSet<int>();
            foreach (var pair in codeToGlyphId)
            {
                int unicode;
                if (!glyphIdToUnicode.TryGetValue(pair.Value, out unicode))
                {
                    unicode = pair.Key;
                }

                ushort newGlyphId;
                if (!oldToNew.TryGetValue(pair.Value, out newGlyphId))
                {
                    continue;
                }

                if (seen.Add(unicode))
                {
                    mappings.Add(new CharMapping(unicode, newGlyphId));
                }
            }

            mappings.Sort(delegate (CharMapping a, CharMapping b) { return a.Unicode.CompareTo(b.Unicode); });
            return BuildFormat4Cmap(mappings);
        }

        private byte[] BuildFormat4Cmap(IList<CharMapping> mappings)
        {
            var segmentCount = mappings.Count + 1;
            var segCountX2 = (ushort)(segmentCount * 2);
            ushort searchRange = 2;
            ushort entrySelector = 0;
            while (searchRange * 2 <= segCountX2)
            {
                searchRange *= 2;
                entrySelector++;
            }

            var rangeShift = (ushort)(segCountX2 - searchRange);
            using (var format4 = new MemoryStream())
            {
                WriteUShort(format4, 4);
                WriteUShort(format4, 0);
                WriteUShort(format4, 0);
                WriteUShort(format4, segCountX2);
                WriteUShort(format4, searchRange);
                WriteUShort(format4, entrySelector);
                WriteUShort(format4, rangeShift);

                for (var i = 0; i < mappings.Count; i++)
                {
                    WriteUShort(format4, (ushort)mappings[i].Unicode);
                }

                WriteUShort(format4, 0xFFFF);
                WriteUShort(format4, 0);

                for (var i = 0; i < mappings.Count; i++)
                {
                    WriteUShort(format4, (ushort)mappings[i].Unicode);
                }

                WriteUShort(format4, 0xFFFF);
                for (var i = 0; i < mappings.Count; i++)
                {
                    int delta = mappings[i].GlyphId - mappings[i].Unicode;
                    while (delta < 0)
                    {
                        delta += 65536;
                    }

                    WriteUShort(format4, (ushort)delta);
                }

                WriteUShort(format4, 1);
                for (var i = 0; i < segmentCount; i++)
                {
                    WriteUShort(format4, 0);
                }

                var bytes = format4.ToArray();
                WriteUShort(bytes, 2, (ushort)bytes.Length);
                return BuildCmapTableHeader(bytes);
            }
        }

        private byte[] BuildCmapTableHeader(byte[] format4)
        {
            using (var output = new MemoryStream())
            {
                WriteUShort(output, 0);
                WriteUShort(output, 1);
                WriteUShort(output, 3);
                WriteUShort(output, 1);
                WriteUInt(output, 12);
                output.Write(format4, 0, format4.Length);
                return output.ToArray();
            }
        }

        private byte[] BuildPostTable()
        {
            var original = TableBytes("post");
            using (var output = new MemoryStream())
            {
                WriteUInt(output, 0x00030000);
                WriteUInt(output, ReadUInt(original, 4));
                WriteShort(output, ReadShort(original, 8));
                WriteShort(output, ReadShort(original, 10));
                WriteUInt(output, ReadUInt(original, 12));
                WriteUInt(output, ReadUInt(original, 16));
                WriteUInt(output, ReadUInt(original, 20));
                WriteUInt(output, ReadUInt(original, 24));
                WriteUInt(output, ReadUInt(original, 28));
                return output.ToArray();
            }
        }

        private byte[] BuildOs2Table()
        {
            var original = TableBytes("OS/2");
            var copy = new byte[original.Length];
            Buffer.BlockCopy(original, 0, copy, 0, original.Length);
            if (copy.Length >= 62)
            {
                WriteUShort(copy, 62, 0);
            }

            return copy;
        }

        private byte[] BuildNameTable(string subsetFontName)
        {
            var records = new List<NameRecord>();
            AddNameRecord(records, 3, 1, 1033, 1, "Calibri");
            AddNameRecord(records, 3, 1, 1033, 2, "Regular");
            AddNameRecord(records, 3, 1, 1033, 4, subsetFontName);
            AddNameRecord(records, 3, 1, 1033, 6, subsetFontName);
            using (var output = new MemoryStream())
            {
                WriteUShort(output, 0);
                WriteUShort(output, (ushort)records.Count);
                WriteUShort(output, (ushort)(6 + records.Count * 12));
                var runningOffset = 0;
                for (var i = 0; i < records.Count; i++)
                {
                    var record = records[i];
                    WriteUShort(output, record.PlatformId);
                    WriteUShort(output, record.EncodingId);
                    WriteUShort(output, record.LanguageId);
                    WriteUShort(output, record.NameId);
                    WriteUShort(output, (ushort)record.Bytes.Length);
                    WriteUShort(output, (ushort)runningOffset);
                    runningOffset += record.Bytes.Length;
                }

                for (var i = 0; i < records.Count; i++)
                {
                    var bytes = records[i].Bytes;
                    output.Write(bytes, 0, bytes.Length);
                }

                return output.ToArray();
            }
        }

        private static void AddNameRecord(IList<NameRecord> records, ushort platformId, ushort encodingId, ushort languageId, ushort nameId, string text)
        {
            records.Add(new NameRecord
            {
                PlatformId = platformId,
                EncodingId = encodingId,
                LanguageId = languageId,
                NameId = nameId,
                Bytes = Encoding.BigEndianUnicode.GetBytes(text)
            });
        }

        private byte[] BuildFontFile(IDictionary<string, byte[]> tables)
        {
            var tags = new List<string>(tables.Keys);
            tags.Sort(StringComparer.Ordinal);
            var numTables = (ushort)tags.Count;
            ushort entrySelector = 0;
            ushort searchRange = 1;
            while (searchRange * 2 <= numTables)
            {
                searchRange *= 2;
                entrySelector++;
            }

            searchRange = (ushort)(searchRange * 16);
            var rangeShift = (ushort)(numTables * 16 - searchRange);
            using (var output = new MemoryStream())
            {
                WriteUInt(output, 0x00010000);
                WriteUShort(output, numTables);
                WriteUShort(output, searchRange);
                WriteUShort(output, entrySelector);
                WriteUShort(output, rangeShift);

                var recordPosition = output.Position;
                for (var i = 0; i < numTables; i++)
                {
                    output.Write(new byte[16], 0, 16);
                }

                var records = new List<OutputRecord>();
                for (var i = 0; i < tags.Count; i++)
                {
                    var tag = tags[i];
                    var bytes = tables[tag];
                    while ((output.Position & 3) != 0)
                    {
                        output.WriteByte(0);
                    }

                    var offset = (uint)output.Position;
                    output.Write(bytes, 0, bytes.Length);
                    while ((output.Position & 3) != 0)
                    {
                        output.WriteByte(0);
                    }

                    records.Add(new OutputRecord(tag, CalculateChecksum(bytes), offset, (uint)bytes.Length));
                }

                var fontBytes = output.ToArray();
                for (var i = 0; i < records.Count; i++)
                {
                    var entryOffset = (int)recordPosition + i * 16;
                    WriteTag(fontBytes, entryOffset, records[i].Tag);
                    WriteUInt(fontBytes, entryOffset + 4, records[i].Checksum);
                    WriteUInt(fontBytes, entryOffset + 8, records[i].Offset);
                    WriteUInt(fontBytes, entryOffset + 12, records[i].Length);
                }

                var headRecord = records.Find(delegate (OutputRecord item) { return item.Tag == "head"; });
                WriteUInt(fontBytes, (int)headRecord.Offset + 8, 0);
                var checksum = CalculateChecksum(fontBytes);
                var adjustment = 0xB1B0AFBAu - checksum;
                WriteUInt(fontBytes, (int)headRecord.Offset + 8, adjustment);
                return fontBytes;
            }
        }

        private byte[] TableBytes(string tag)
        {
            TableRecord record;
            if (!_tables.TryGetValue(tag, out record))
            {
                throw new InvalidOperationException("TrueType subset source was missing required table " + tag + ".");
            }

            return SliceBytes(_fontBytes, (int)record.Offset, (int)record.Length);
        }

        private HorizontalMetric MetricForGlyph(ushort glyphId)
        {
            if (glyphId >= _metrics.Length)
            {
                return _metrics[_metrics.Length - 1];
            }

            return _metrics[glyphId];
        }

        private byte[] GetGlyphBytes(ushort glyphId)
        {
            if (glyphId >= _numGlyphs)
            {
                return null;
            }

            var start = _glyphOffsets[glyphId];
            var end = _glyphOffsets[glyphId + 1];
            if (end <= start)
            {
                return new byte[0];
            }

            var glyf = TableBytes("glyf");
            return SliceBytes(glyf, start, end - start);
        }

        private int[] ParseGlyphOffsets()
        {
            var loca = TableBytes("loca");
            var offsets = new int[_numGlyphs + 1];
            if (_indexToLocFormat == 0)
            {
                for (var i = 0; i <= _numGlyphs; i++)
                {
                    offsets[i] = ReadUShort(loca, i * 2) * 2;
                }
            }
            else
            {
                for (var i = 0; i <= _numGlyphs; i++)
                {
                    offsets[i] = (int)ReadUInt(loca, i * 4);
                }
            }

            return offsets;
        }

        private HorizontalMetric[] ParseHorizontalMetrics()
        {
            var bytes = TableBytes("hmtx");
            var metrics = new HorizontalMetric[_numGlyphs];
            ushort lastAdvance = 0;
            short lastLsb = 0;
            for (var i = 0; i < _numGlyphs; i++)
            {
                if (i < _numberOfHMetrics)
                {
                    var metricOffset = i * 4;
                    lastAdvance = ReadUShort(bytes, metricOffset);
                    lastLsb = ReadShort(bytes, metricOffset + 2);
                }
                else
                {
                    var lsbOffset = _numberOfHMetrics * 4 + (i - _numberOfHMetrics) * 2;
                    lastLsb = ReadShort(bytes, lsbOffset);
                }

                metrics[i] = new HorizontalMetric { AdvanceWidth = lastAdvance, LeftSideBearing = lastLsb };
            }

            return metrics;
        }

        private static Dictionary<string, TableRecord> ParseTableDirectory(byte[] fontBytes)
        {
            var numTables = ReadUShort(fontBytes, 4);
            var tables = new Dictionary<string, TableRecord>(StringComparer.Ordinal);
            for (var i = 0; i < numTables; i++)
            {
                var offset = 12 + i * 16;
                var record = new TableRecord();
                record.Tag = Encoding.ASCII.GetString(fontBytes, offset, 4);
                record.Checksum = ReadUInt(fontBytes, offset + 4);
                record.Offset = ReadUInt(fontBytes, offset + 8);
                record.Length = ReadUInt(fontBytes, offset + 12);
                tables[record.Tag] = record;
            }

            return tables;
        }

        private static byte[] SliceBytes(byte[] source, int offset, int length)
        {
            var bytes = new byte[length];
            Buffer.BlockCopy(source, offset, bytes, 0, length);
            return bytes;
        }

        private static byte[] StripSimpleGlyphInstructions(byte[] glyphBytes, short contourCount)
        {
            if (contourCount < 0)
            {
                return glyphBytes;
            }

            var contourArrayBytes = contourCount * 2;
            var instructionLengthOffset = 10 + contourArrayBytes;
            if (instructionLengthOffset + 2 > glyphBytes.Length)
            {
                return glyphBytes;
            }

            var instructionLength = ReadUShort(glyphBytes, instructionLengthOffset);
            var flagsOffset = instructionLengthOffset + 2 + instructionLength;
            if (flagsOffset > glyphBytes.Length)
            {
                return glyphBytes;
            }

            using (var output = new MemoryStream())
            {
                output.Write(glyphBytes, 0, instructionLengthOffset);
                WriteUShort(output, 0);
                output.Write(glyphBytes, flagsOffset, glyphBytes.Length - flagsOffset);
                return output.ToArray();
            }
        }

        private static ushort ReadUShort(byte[] bytes, int offset)
        {
            return (ushort)((bytes[offset] << 8) | bytes[offset + 1]);
        }

        private static short ReadShort(byte[] bytes, int offset)
        {
            return unchecked((short)ReadUShort(bytes, offset));
        }

        private static uint ReadUInt(byte[] bytes, int offset)
        {
            return (uint)((bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3]);
        }

        private static void WriteUShort(byte[] bytes, int offset, ushort value)
        {
            bytes[offset] = (byte)(value >> 8);
            bytes[offset + 1] = (byte)value;
        }

        private static void WriteShort(byte[] bytes, int offset, short value)
        {
            WriteUShort(bytes, offset, unchecked((ushort)value));
        }

        private static void WriteUInt(byte[] bytes, int offset, uint value)
        {
            bytes[offset] = (byte)(value >> 24);
            bytes[offset + 1] = (byte)(value >> 16);
            bytes[offset + 2] = (byte)(value >> 8);
            bytes[offset + 3] = (byte)value;
        }

        private static void WriteTag(byte[] bytes, int offset, string tag)
        {
            bytes[offset] = (byte)tag[0];
            bytes[offset + 1] = (byte)tag[1];
            bytes[offset + 2] = (byte)tag[2];
            bytes[offset + 3] = (byte)tag[3];
        }

        private static void WriteUShort(Stream stream, ushort value)
        {
            stream.WriteByte((byte)(value >> 8));
            stream.WriteByte((byte)value);
        }

        private static void WriteShort(Stream stream, short value)
        {
            WriteUShort(stream, unchecked((ushort)value));
        }

        private static void WriteUInt(Stream stream, uint value)
        {
            stream.WriteByte((byte)(value >> 24));
            stream.WriteByte((byte)(value >> 16));
            stream.WriteByte((byte)(value >> 8));
            stream.WriteByte((byte)value);
        }

        private static uint CalculateChecksum(byte[] bytes)
        {
            uint sum = 0;
            var paddedLength = (bytes.Length + 3) & ~3;
            for (var i = 0; i < paddedLength; i += 4)
            {
                uint value = 0;
                if (i < bytes.Length) value |= (uint)bytes[i] << 24;
                if (i + 1 < bytes.Length) value |= (uint)bytes[i + 1] << 16;
                if (i + 2 < bytes.Length) value |= (uint)bytes[i + 2] << 8;
                if (i + 3 < bytes.Length) value |= bytes[i + 3];
                sum += value;
            }

            return sum;
        }

        private sealed class GlyfBuildResult
        {
            public GlyfBuildResult(byte[] bytes, IList<uint> locaOffsets)
            {
                Bytes = bytes;
                LocaOffsets = locaOffsets;
            }

            public byte[] Bytes { get; private set; }

            public IList<uint> LocaOffsets { get; private set; }
        }

        private struct CharMapping
        {
            public CharMapping(int unicode, ushort glyphId)
            {
                Unicode = unicode;
                GlyphId = glyphId;
            }

            public int Unicode;
            public ushort GlyphId;
        }

        private struct OutputRecord
        {
            public OutputRecord(string tag, uint checksum, uint offset, uint length)
            {
                Tag = tag;
                Checksum = checksum;
                Offset = offset;
                Length = length;
            }

            public string Tag;
            public uint Checksum;
            public uint Offset;
            public uint Length;
        }

        private struct NameRecord
        {
            public ushort PlatformId;
            public ushort EncodingId;
            public ushort LanguageId;
            public ushort NameId;
            public byte[] Bytes;
        }
    }
}
