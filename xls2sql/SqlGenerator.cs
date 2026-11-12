using Sylvan.Data.Csv;
using Sylvan.Data.Excel;
using System;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace xls2sql
{
    /// <summary>All raw cell strings of one worksheet. Row 0 is the header row.</summary>
    internal sealed class SheetData
    {
        public SheetData(List<string[]> rows) => Rows = rows;

        /// <summary>Each row holds <c>FieldCount</c> cells as reported by the reader for that row.</summary>
        public List<string[]> Rows { get; }
    }

    /// <summary>
    /// Generated script as ordered UTF-8 chunks (concatenation = full script, identical to what
    /// File.WriteAllText would write). Kept chunked: half the memory of a UTF-16 string, no
    /// 2 GB single-object limit, and no full-size copy is ever needed for display or saving.
    /// Chunk boundaries never split a character.
    /// </summary>
    internal sealed class SqlOutput
    {
        public static readonly SqlOutput Empty = new SqlOutput(new byte[0][]);

        public SqlOutput(byte[][] chunks)
        {
            Chunks = chunks;
            foreach (byte[] c in chunks) Length += c.Length;
        }

        public byte[][] Chunks { get; }

        /// <summary>Total size in UTF-8 bytes.</summary>
        public long Length { get; }

        public void WriteTo(Stream stream)
        {
            foreach (byte[] c in Chunks) stream.Write(c, 0, c.Length);
        }

        /// <summary>Decodes the whole script into one string (only for clipboard).</summary>
        public string ToText()
        {
            long total = 0;
            foreach (byte[] c in Chunks) total += Encoding.UTF8.GetCharCount(c);
            var sb = new StringBuilder(checked((int)total));
            foreach (byte[] c in Chunks) sb.Append(Encoding.UTF8.GetString(c));
            return sb.ToString();
        }
    }

    internal static class SheetReader
    {
        private static readonly ExcelDataReaderOptions _excelOptions = new ExcelDataReaderOptions
        {
            Schema = ExcelSchema.NoHeaders,
        };
        private static readonly CsvDataReaderOptions _csvOptions = new CsvDataReaderOptions
        {
            HasHeaders = false,
        };

        public static bool IsCsv(string filePath) =>
            filePath.EndsWith(".csv", StringComparison.OrdinalIgnoreCase);

        public static string[] ReadSheetNames(string filePath)
        {
            if (IsCsv(filePath))
                return new[] { Path.GetFileNameWithoutExtension(filePath) };

            var names = new List<string>();
            using (var reader = ExcelDataReader.Create(filePath, _excelOptions))
            {
                do { names.Add(reader.WorksheetName); } while (reader.NextResult());
            }
            return names.ToArray();
        }

        /// <summary>Reads every cell of the sheet as a string (single pass, forward-only).</summary>
        public static SheetData Load(string filePath, string sheetName)
        {
            var rows = new List<string[]>();
            using (DbDataReader reader = OpenSheet(filePath, sheetName))
            {
                while (reader.Read())
                {
                    int fieldCount = reader.FieldCount;
                    int rowFieldCount = Math.Min(RowFieldCount(reader), fieldCount);
                    var cells = new string[fieldCount];
                    for (int i = 0; i < rowFieldCount; i++)
                        cells[i] = reader.GetString(i) ?? string.Empty;
                    for (int i = rowFieldCount; i < fieldCount; i++)
                        cells[i] = string.Empty;
                    rows.Add(cells);
                }
            }
            return new SheetData(rows);
        }

        // Opens a forward-only reader positioned before the first row of the given sheet.
        private static DbDataReader OpenSheet(string filePath, string sheetName)
        {
            if (IsCsv(filePath))
            {
                // Windows-1250 fallback when no BOM, matching previous ExcelDataReader config.
                var textReader = new StreamReader(filePath, Encoding.GetEncoding(1250), detectEncodingFromByteOrderMarks: true);
                return CsvDataReader.Create(textReader, _csvOptions);
            }

            var excelReader = ExcelDataReader.Create(filePath, _excelOptions);
            if (!excelReader.TryOpenWorksheet(sheetName))
            {
                excelReader.Dispose();
                throw new InvalidOperationException($"Worksheet '{sheetName}' not found.");
            }
            return excelReader;
        }

        private static int RowFieldCount(DbDataReader reader)
        {
            switch (reader)
            {
                case ExcelDataReader e: return e.RowFieldCount;
                case CsvDataReader c: return c.RowFieldCount;
                default: return reader.FieldCount;
            }
        }
    }

    internal static class SqlGenerator
    {
        private static readonly Regex _lineBreaks = new Regex(@"\t|\n|\r", RegexOptions.Compiled);

        // Rows per parallel work item. Large enough to amortize scheduling, small enough to balance.
        private const int ChunkRows = 1024;

        public static SqlOutput Generate(SqlSettings s, SheetData data)
        {
            string nl = Environment.NewLine;
            List<string[]> rows = data.Rows;

            List<string> headers = GetColumnNames(rows, s.TrimWhiteSpaces);
            string columnNames = GenerateColumnNames(headers);
            int headerCount = headers.Count;
            string insertHead = $"INSERT INTO {s.TableName} ({columnNames}) {nl}VALUES ";

            var prologue = new StringBuilder();
            prologue.Append("USE ").Append(s.DatabaseName).Append(';').Append(nl);
            prologue.Append("SET ANSI_NULLS ON;").Append(nl);
            prologue.Append("SET QUOTED_IDENTIFIER ON;").Append(nl);
            prologue.Append(nl);
            if (s.IsCreateTable)
            {
                prologue.Append("CREATE TABLE ").Append(s.TableName).Append(" (")
                        .Append(GenerateColumnNamesForCreateTable(headers, s.FirstColumnId)).Append(");").Append(nl);
                prologue.Append(nl);
            }
            prologue.Append(insertHead);

            int dataRows = Math.Max(0, rows.Count - 1);
            int chunkCount = (dataRows + ChunkRows - 1) / ChunkRows;
            // [prologue, chunk 0..n-1, ";"]
            var parts = new byte[chunkCount + 2][];
            parts[0] = Encoding.UTF8.GetBytes(prologue.ToString());
            parts[chunkCount + 1] = new[] { (byte)';' };

            // Each chunk is formatted independently; the batch separator only depends on the
            // global data-row index, so chunks can be concatenated in order afterwards.
            // Workers reuse their builder/buffers, so the only per-chunk allocation is the result.
            Parallel.For(0, chunkCount,
                () => new ChunkWorker(ChunkRows * Math.Max(16, headerCount * 12)),
                (c, _, w) =>
                {
                    int from = c * ChunkRows;
                    int to = Math.Min(from + ChunkRows, dataRows);
                    StringBuilder sb = w.Builder.Clear();
                    var f = new RowFormatter(sb, s.TrimWhiteSpaces, s.PreferNulls);
                    for (int d = from; d < to; d++)
                    {
                        if (d > 0)
                        {
                            if (d % s.Separator == 0)
                                sb.Append(';').Append(nl).Append(nl).Append(insertHead);
                            else
                                sb.Append(", ");
                        }
                        string[] cells = rows[d + 1];
                        f.AppendRow(cells, Math.Min(headerCount, cells.Length));
                    }
                    parts[c + 1] = w.ToUtf8();
                    return w;
                },
                _ => { });

            return new SqlOutput(parts);
        }

        /// <summary>Per-thread reusable formatting state.</summary>
        private sealed class ChunkWorker
        {
            public readonly StringBuilder Builder;
#if NETCOREAPP
            // Same replacement fallback as Encoding.UTF8.GetBytes(string) (lone surrogate -> EF BF BD).
            private readonly Encoder _encoder = Encoding.UTF8.GetEncoder();
            private byte[] _buffer = new byte[0];
#endif

            public ChunkWorker(int capacity) => Builder = new StringBuilder(capacity);

            /// <summary>
            /// UTF-8 bytes of <see cref="Builder"/>, identical to Encoding.UTF8.GetBytes(Builder.ToString())
            /// but without materializing the UTF-16 string (which would land on the LOH as garbage).
            /// </summary>
            public byte[] ToUtf8()
            {
#if NETCOREAPP
                _encoder.Reset();
                int pos = 0;
                foreach (ReadOnlyMemory<char> m in Builder.GetChunks())
                {
                    ReadOnlySpan<char> src = m.Span;
                    EnsureCapacity(pos + Encoding.UTF8.GetMaxByteCount(src.Length));
                    // flush: false keeps a surrogate pair split across builder chunks intact.
                    pos += _encoder.GetBytes(src, _buffer.AsSpan(pos), flush: false);
                }
                EnsureCapacity(pos + Encoding.UTF8.GetMaxByteCount(0));
                pos += _encoder.GetBytes(ReadOnlySpan<char>.Empty, _buffer.AsSpan(pos), flush: true);
                return _buffer.AsSpan(0, pos).ToArray();
#else
                return Encoding.UTF8.GetBytes(Builder.ToString());
#endif
            }

#if NETCOREAPP
            private void EnsureCapacity(int size)
            {
                if (_buffer.Length < size)
                    Array.Resize(ref _buffer, Math.Max(size, _buffer.Length * 2));
            }
#endif
        }

        /// <summary>
        /// Allocation-free equivalent of the original per-row logic:
        /// <code>
        /// value = cell.Replace("'", "''"); if (trim) value = value.Trim();
        /// if (value.Length == 0 &amp;&amp; preferNulls) value = "NULL";
        /// row = $"('{string.Join("', N'", values)}')";
        /// if (preferNulls) row = row.Replace("N'NULL'", "NULL");
        /// </code>
        /// Because quotes inside values are always doubled, "N'NULL'" can only match at the
        /// opening quote of column 1+ whose escaped value is exactly "NULL" or starts with "NULL'".
        /// </summary>
        private readonly struct RowFormatter
        {
            private readonly StringBuilder _sb;
            private readonly bool _trim;
            private readonly bool _preferNulls;

            public RowFormatter(StringBuilder sb, bool trim, bool preferNulls)
            {
                _sb = sb;
                _trim = trim;
                _preferNulls = preferNulls;
            }

            public void AppendRow(string[] cells, int colCount)
            {
                StringBuilder sb = _sb;
                sb.Append('(');
                if (colCount == 0)
                {
                    sb.Append("''");
                }
                for (int j = 0; j < colCount; j++)
                {
                    string v = cells[j];
                    int start = 0, end = v.Length;
                    if (_trim)
                    {
                        while (start < end && char.IsWhiteSpace(v[start])) start++;
                        while (end > start && char.IsWhiteSpace(v[end - 1])) end--;
                    }
                    int len = end - start;

                    if (len == 0 && _preferNulls)
                    {
                        sb.Append(j == 0 ? "'NULL'" : ", NULL");
                        continue;
                    }

                    if (j == 0)
                    {
                        sb.Append('\'');
                    }
                    else if (_preferNulls && len >= 4 && IsNullPrefix(v, start)
                             && (len == 4 || v[start + 4] == '\''))
                    {
                        if (len == 4)
                        {
                            sb.Append(", NULL");
                            continue;
                        }
                        // "N'NULL''rest'" -> "NULL'rest'"
                        sb.Append(", NULL'");
                        AppendEscaped(sb, v, start + 5, len - 5);
                        sb.Append('\'');
                        continue;
                    }
                    else
                    {
                        sb.Append(", N'");
                    }

                    AppendEscaped(sb, v, start, len);
                    sb.Append('\'');
                }
                sb.Append(')');
            }

            private static bool IsNullPrefix(string v, int i) =>
                v[i] == 'N' && v[i + 1] == 'U' && v[i + 2] == 'L' && v[i + 3] == 'L';

            private static void AppendEscaped(StringBuilder sb, string v, int start, int len)
            {
                int end = start + len;
                while (start < end)
                {
                    int q = v.IndexOf('\'', start, end - start);
                    if (q < 0)
                    {
                        sb.Append(v, start, end - start);
                        return;
                    }
                    sb.Append(v, start, q - start).Append("''");
                    start = q + 1;
                }
            }
        }

        private static List<string> GetColumnNames(List<string[]> rows, bool trimWhiteSpaces)
        {
            var headers = new List<string>();
            if (rows.Count == 0) return headers;

            foreach (string cell in rows[0])
            {
                string header = cell;
                if (trimWhiteSpaces) header = header.Trim();
                header = _lineBreaks.Replace(header, "");
                if (!string.IsNullOrEmpty(header))
                    headers.Add(header);
            }
            return headers;
        }

        private static string GenerateColumnNames(List<string> headers)
        {
            if (headers == null || headers.Count == 0) return string.Empty;
            return $"[{string.Join("], [", headers)}]";
        }

        private static string GenerateColumnNamesForCreateTable(List<string> headers, int firstColumnId = 0)
        {
            if (headers == null || headers.Count == 0) return string.Empty;

            var sb = new StringBuilder();
            if (firstColumnId == 1)
                sb.Append("Id INT IDENTITY (1, 1) NOT NULL, ");
            else if (firstColumnId == 2)
                sb.Append("Id UNIQUEIDENTIFIER DEFAULT NEWSEQUENTIALID() NOT NULL, ");

            sb.Append($"[{string.Join("] varchar(max) NULL, [", headers)}] varchar(max) NULL");
            return sb.ToString();
        }
    }

    internal class SqlSettings
    {
        public string DatabaseName { get; set; }
        public string TableName { get; set; }
        public bool IsCreateTable { get; set; }
        public int Separator { get; set; }
        public bool PreferNulls { get; set; }
        public bool TrimWhiteSpaces { get; set; }
        public int FirstColumnId { get; set; }
        public string FilePath { get; set; }
        public string SheetName { get; set; }
    }
}
