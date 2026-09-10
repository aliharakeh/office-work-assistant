using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using ClosedXML.Excel;

namespace WorkAssistant.Features.ExcelProcessing
{
    public sealed class ExcelProcessingLoadResult
    {
        public DataTable Table { get; set; }
        public string[] Sheets { get; set; }
        public string Sheet { get; set; }
    }

    public sealed class ExcelProcessingCompareResult
    {
        public DataTable Table { get; set; }
        public char[] Sources { get; set; }
        public string[] Originals { get; set; }
    }

    public sealed class ExcelProcessingMatch
    {
        public string ColumnA { get; set; }
        public string ColumnB { get; set; }
        public string Value { get; set; }
    }

    public static class ExcelProcessingWork
    {
        public static ExcelProcessingLoadResult Load(string path, string sheetName)
        {
            using (var wb = new XLWorkbook(path))
            {
                var sheets = new string[wb.Worksheets.Count];
                var i = 0;
                foreach (var w in wb.Worksheets)
                    sheets[i++] = w.Name;

                var ws = string.IsNullOrEmpty(sheetName)
                    ? wb.Worksheet(1)
                    : wb.Worksheet(sheetName);

                return new ExcelProcessingLoadResult
                {
                    Table = ToTable(ws),
                    Sheets = sheets,
                    Sheet = ws.Name
                };
            }
        }

        public static void Save(DataTable table, string path)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Result");
                for (var c = 0; c < table.Columns.Count; c++)
                    ws.Cell(1, c + 1).Value = table.Columns[c].ColumnName;

                for (var r = 0; r < table.Rows.Count; r++)
                {
                    for (var c = 0; c < table.Columns.Count; c++)
                    {
                        var v = table.Rows[r][c];
                        if (v != null && v != DBNull.Value)
                            ws.Cell(r + 2, c + 1).Value = Convert.ToString(v, CultureInfo.InvariantCulture);
                    }
                }

                wb.SaveAs(path);
            }
        }

        public static ExcelProcessingCompareResult OnlyInA(DataTable a, DataTable b, string keyA, string keyB, IList<ExcelProcessingMatch> extra)
        {
            return FromOne(CopyUnmatched(a, keyA, extra, true, KeySet(b, keyB, extra, false)), 'A');
        }

        public static ExcelProcessingCompareResult OnlyInB(DataTable a, DataTable b, string keyA, string keyB, IList<ExcelProcessingMatch> extra)
        {
            return FromOne(CopyUnmatched(b, keyB, extra, false, KeySet(a, keyA, extra, true)), 'B');
        }

        public static ExcelProcessingCompareResult Common(DataTable a, DataTable b, string keyA, string keyB, IList<ExcelProcessingMatch> extra)
        {
            string[] originals;
            char[] sources;
            var result = NewJoinTable(a, b, keyB, out originals, out sources);
            var lookup = BuildLookup(b, keyB, extra, false);

            foreach (DataRow rowA in a.Rows)
            {
                if (!RowFilter(rowA, extra, true))
                    continue;
                var k = FullKey(rowA, keyA, extra, true);
                List<DataRow> hits;
                if (k.Length == 0 || !lookup.TryGetValue(k, out hits))
                    continue;
                foreach (var rowB in hits)
                    result.Rows.Add(FillJoin(result, a, b, keyB, rowA, rowB));
            }

            return new ExcelProcessingCompareResult { Table = result, Sources = sources, Originals = originals };
        }

        public static DataTable Project(DataTable source, IList<string> internals, IList<string> outputs)
        {
            var table = new DataTable();
            if (outputs != null)
            {
                for (var i = 0; i < outputs.Count; i++)
                    table.Columns.Add(UniqueName(table, outputs[i]), typeof(object));
            }

            if (internals == null || internals.Count == 0)
                return table;

            foreach (DataRow row in source.Rows)
            {
                var dest = table.NewRow();
                for (var i = 0; i < internals.Count; i++)
                    dest[i] = row[internals[i]];
                table.Rows.Add(dest);
            }
            return table;
        }

        public static string OutputName(string original, char source, IList<string> selectedOriginals, IList<char> selectedSources)
        {
            var mixed = false;
            var hasA = false;
            var hasB = false;
            if (selectedSources != null)
            {
                for (var i = 0; i < selectedSources.Count; i++)
                {
                    if (selectedSources[i] == 'A')
                        hasA = true;
                    else if (selectedSources[i] == 'B')
                        hasB = true;
                }
                mixed = hasA && hasB;
            }

            if (!mixed)
                return original;

            var hits = 0;
            if (selectedOriginals != null)
            {
                for (var i = 0; i < selectedOriginals.Count; i++)
                {
                    if (string.Equals(selectedOriginals[i], original, StringComparison.OrdinalIgnoreCase))
                        hits++;
                }
            }

            if (hits > 1)
                return original + "_" + source;
            return original;
        }

        static ExcelProcessingCompareResult FromOne(DataTable table, char source)
        {
            var sources = new char[table.Columns.Count];
            var originals = new string[table.Columns.Count];
            for (var i = 0; i < sources.Length; i++)
            {
                sources[i] = source;
                originals[i] = table.Columns[i].ColumnName;
            }
            return new ExcelProcessingCompareResult { Table = table, Sources = sources, Originals = originals };
        }

        static DataTable ToTable(IXLWorksheet ws)
        {
            var table = new DataTable();
            var range = ws.RangeUsed();
            if (range == null)
                return table;

            var header = range.FirstRow();
            var colCount = range.ColumnCount();
            for (var c = 1; c <= colCount; c++)
            {
                var name = header.Cell(c).GetString();
                if (string.IsNullOrWhiteSpace(name))
                    name = "Column" + c;
                table.Columns.Add(UniqueName(table, name), typeof(object));
            }

            foreach (var row in range.RowsUsed())
            {
                if (row.RowNumber() == header.RowNumber())
                    continue;
                var dr = table.NewRow();
                for (var c = 1; c <= colCount; c++)
                {
                    var cell = row.Cell(c);
                    dr[c - 1] = cell.IsEmpty() ? (object)DBNull.Value : cell.Value;
                }
                table.Rows.Add(dr);
            }

            return table;
        }

        static string UniqueName(DataTable table, string name)
        {
            if (!table.Columns.Contains(name))
                return name;
            var n = 2;
            while (table.Columns.Contains(name + "_" + n))
                n++;
            return name + "_" + n;
        }

        static string NormKey(object value)
        {
            if (value == null || value == DBNull.Value)
                return "";
            return Convert.ToString(value, CultureInfo.InvariantCulture).Trim().ToLowerInvariant();
        }

        static Dictionary<string, List<DataRow>> BuildLookup(DataTable table, string keyCol, IList<ExcelProcessingMatch> extra, bool isA)
        {
            var map = new Dictionary<string, List<DataRow>>(StringComparer.Ordinal);
            foreach (DataRow row in table.Rows)
            {
                if (!RowFilter(row, extra, isA))
                    continue;
                var k = FullKey(row, keyCol, extra, isA);
                if (k.Length == 0)
                    continue;
                List<DataRow> list;
                if (!map.TryGetValue(k, out list))
                {
                    list = new List<DataRow>();
                    map[k] = list;
                }
                list.Add(row);
            }
            return map;
        }

        static HashSet<string> KeySet(DataTable table, string keyCol, IList<ExcelProcessingMatch> extra, bool isA)
        {
            var set = new HashSet<string>(StringComparer.Ordinal);
            foreach (DataRow row in table.Rows)
            {
                if (!RowFilter(row, extra, isA))
                    continue;
                var k = FullKey(row, keyCol, extra, isA);
                if (k.Length > 0)
                    set.Add(k);
            }
            return set;
        }

        static string FullKey(DataRow row, string keyCol, IList<ExcelProcessingMatch> extra, bool isA)
        {
            var k = NormKey(row[keyCol]);
            if (k.Length == 0)
                return "";
            if (extra == null)
                return k;
            for (var i = 0; i < extra.Count; i++)
            {
                var m = extra[i];
                if (string.IsNullOrEmpty(m.ColumnA) || string.IsNullOrEmpty(m.ColumnB))
                    continue;
                k += "\x1f" + NormKey(row[isA ? m.ColumnA : m.ColumnB]);
            }
            return k;
        }

        static bool RowFilter(DataRow row, IList<ExcelProcessingMatch> extra, bool isA)
        {
            if (extra == null)
                return true;
            for (var i = 0; i < extra.Count; i++)
            {
                var m = extra[i];
                if (string.IsNullOrWhiteSpace(m.Value))
                    continue;
                var col = isA ? m.ColumnA : m.ColumnB;
                if (string.IsNullOrEmpty(col))
                    continue;
                if (NormKey(row[col]) != NormKey(m.Value))
                    return false;
            }
            return true;
        }

        static DataTable NewJoinTable(DataTable a, DataTable b, string keyB, out string[] originals, out char[] sources)
        {
            var result = new DataTable();
            var orig = new List<string>();
            var src = new List<char>();

            foreach (DataColumn col in a.Columns)
            {
                result.Columns.Add("c" + result.Columns.Count, typeof(object));
                orig.Add(col.ColumnName);
                src.Add('A');
            }

            foreach (DataColumn col in b.Columns)
            {
                if (string.Equals(col.ColumnName, keyB, StringComparison.OrdinalIgnoreCase))
                    continue;
                result.Columns.Add("c" + result.Columns.Count, typeof(object));
                orig.Add(col.ColumnName);
                src.Add('B');
            }

            originals = orig.ToArray();
            sources = src.ToArray();
            return result;
        }

        static object[] FillJoin(DataTable result, DataTable a, DataTable b, string keyB, DataRow rowA, DataRow rowB)
        {
            var values = new object[result.Columns.Count];
            var i = 0;
            foreach (DataColumn col in a.Columns)
                values[i++] = rowA[col.ColumnName];

            foreach (DataColumn col in b.Columns)
            {
                if (string.Equals(col.ColumnName, keyB, StringComparison.OrdinalIgnoreCase))
                    continue;
                values[i++] = rowB[col.ColumnName];
            }

            return values;
        }

        static DataTable CopyUnmatched(DataTable source, string keyCol, IList<ExcelProcessingMatch> extra, bool isA, HashSet<string> otherKeys)
        {
            var result = source.Clone();
            foreach (DataColumn col in result.Columns)
                col.DataType = typeof(object);

            foreach (DataRow row in source.Rows)
            {
                if (!RowFilter(row, extra, isA))
                    continue;
                var k = FullKey(row, keyCol, extra, isA);
                if (k.Length > 0 && otherKeys.Contains(k))
                    continue;
                result.ImportRow(row);
            }
            return result;
        }
    }
}
