using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using ClosedXML.Excel;
using WorkAssistant.Expressions;

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
        public string ExpressionA { get; set; }
        public string ExpressionB { get; set; }
        public string Operator { get; set; }
        public string SplitA { get; set; }
        public string SplitB { get; set; }
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

        public static ExcelProcessingCompareResult OnlyInA(DataTable a, DataTable b, IList<ExcelProcessingMatch> extra)
        {
            return FromOne(CopyUnmatched(a, b, extra, true), 'A');
        }

        public static ExcelProcessingCompareResult OnlyInB(DataTable a, DataTable b, IList<ExcelProcessingMatch> extra)
        {
            return FromOne(CopyUnmatched(b, a, extra, false), 'B');
        }

        public static ExcelProcessingCompareResult Common(DataTable a, DataTable b, IList<ExcelProcessingMatch> extra)
        {
            string[] originals;
            char[] sources;
            var result = NewJoinTable(a, b, out originals, out sources);
            var lookup = BuildLookup(b, extra, false);

            foreach (DataRow rowA in a.Rows)
            {
                if (!RowFilter(a, rowA, extra, true))
                    continue;
                var k = FullKey(a, rowA, extra, true);
                List<DataRow> hits;
                if (k.Length == 0 || !lookup.TryGetValue(k, out hits))
                    continue;
                foreach (var rowB in hits)
                {
                    if (!PairFilter(a, rowA, b, rowB, extra))
                        continue;
                    result.Rows.Add(FillJoin(result, a, b, rowA, rowB));
                }
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

        static Dictionary<string, List<DataRow>> BuildLookup(DataTable table, IList<ExcelProcessingMatch> extra, bool isA)
        {
            var map = new Dictionary<string, List<DataRow>>(StringComparer.Ordinal);
            foreach (DataRow row in table.Rows)
            {
                if (!RowFilter(table, row, extra, isA))
                    continue;
                var k = FullKey(table, row, extra, isA);
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

        static string FullKey(DataTable table, DataRow row, IList<ExcelProcessingMatch> extra, bool isA)
        {
            var k = "";
            if (extra != null)
            {
                for (var i = 0; i < extra.Count; i++)
                {
                    var m = extra[i];
                    if (!IsJoin(m) || Op(m) != "==")
                        continue;
                    var part = NormKey(SideValue(table, row, m, isA));
                    if (part.Length == 0)
                        return "";
                    if (k.Length > 0)
                        k += "\x1f";
                    k += part;
                }
            }
            if (k.Length == 0)
                return HasEquality(extra) ? "" : "\x1e"; // ponytail: one bucket, nested PairFilter if no == key
            return k;
        }

        static bool RowFilter(DataTable table, DataRow row, IList<ExcelProcessingMatch> extra, bool isA)
        {
            if (extra == null)
                return true;
            for (var i = 0; i < extra.Count; i++)
            {
                var m = extra[i];
                if (IsJoin(m))
                    continue;
                if (isA ? !HasA(m) : !HasB(m))
                    continue;
                if (!PassSide(table, row, m, isA))
                    return false;
            }
            return true;
        }

        static bool PairFilter(DataTable a, DataRow rowA, DataTable b, DataRow rowB, IList<ExcelProcessingMatch> extra)
        {
            if (extra == null)
                return true;
            for (var i = 0; i < extra.Count; i++)
            {
                var m = extra[i];
                if (!IsJoin(m) || Op(m) == "==")
                    continue;
                if (!ExpressionEngine.Compare(Op(m), SideValue(a, rowA, m, true), SideValue(b, rowB, m, false)))
                    return false;
            }
            return true;
        }

        static DataTable NewJoinTable(DataTable a, DataTable b, out string[] originals, out char[] sources)
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
                result.Columns.Add("c" + result.Columns.Count, typeof(object));
                orig.Add(col.ColumnName);
                src.Add('B');
            }

            originals = orig.ToArray();
            sources = src.ToArray();
            return result;
        }

        static object[] FillJoin(DataTable result, DataTable a, DataTable b, DataRow rowA, DataRow rowB)
        {
            var values = new object[result.Columns.Count];
            var i = 0;
            foreach (DataColumn col in a.Columns)
                values[i++] = rowA[col.ColumnName];
            foreach (DataColumn col in b.Columns)
                values[i++] = rowB[col.ColumnName];
            return values;
        }

        static DataTable CopyUnmatched(DataTable source, DataTable other, IList<ExcelProcessingMatch> extra, bool isA)
        {
            var result = source.Clone();
            foreach (DataColumn col in result.Columns)
                col.DataType = typeof(object);

            var lookup = BuildLookup(other, extra, !isA);
            foreach (DataRow row in source.Rows)
            {
                if (!RowFilter(source, row, extra, isA))
                    continue;
                if (HasPair(source, row, other, lookup, extra, isA))
                    continue;
                result.ImportRow(row);
            }
            return result;
        }

        static bool HasPair(DataTable source, DataRow row, DataTable other, Dictionary<string, List<DataRow>> lookup, IList<ExcelProcessingMatch> extra, bool isA)
        {
            var k = FullKey(source, row, extra, isA);
            if (k.Length == 0)
                return false;
            List<DataRow> hits;
            if (!lookup.TryGetValue(k, out hits))
                return false;
            for (var i = 0; i < hits.Count; i++)
            {
                if (isA)
                {
                    if (PairFilter(source, row, other, hits[i], extra))
                        return true;
                }
                else if (PairFilter(other, hits[i], source, row, extra))
                    return true;
            }
            return false;
        }

        static bool HasEquality(IList<ExcelProcessingMatch> extra)
        {
            if (extra == null)
                return false;
            for (var i = 0; i < extra.Count; i++)
            {
                if (IsJoin(extra[i]) && Op(extra[i]) == "==")
                    return true;
            }
            return false;
        }

        static string Op(ExcelProcessingMatch m)
        {
            if (m == null || string.IsNullOrWhiteSpace(m.Operator))
                return "==";
            return m.Operator.Trim();
        }

        static bool HasA(ExcelProcessingMatch m)
        {
            return m != null && !string.IsNullOrWhiteSpace(m.ExpressionA);
        }

        static bool HasB(ExcelProcessingMatch m)
        {
            return m != null && !string.IsNullOrWhiteSpace(m.ExpressionB);
        }

        static bool IsJoin(ExcelProcessingMatch m)
        {
            return HasA(m) && HasB(m);
        }

        static bool PassSide(DataTable table, DataRow row, ExcelProcessingMatch m, bool isA)
        {
            return ExpressionEngine.ToBool(SideValue(table, row, m, isA));
        }

        static object SideValue(DataTable table, DataRow row, ExcelProcessingMatch m, bool isA)
        {
            return EvalExpr(table, row, isA ? m.ExpressionA : m.ExpressionB, PartsFor(table, row, m, isA));
        }

        static string[] PartsFor(DataTable table, DataRow row, ExcelProcessingMatch m, bool isA)
        {
            var sep = isA ? m.SplitA : m.SplitB;
            var col = isA ? m.ColumnA : m.ColumnB;
            if (string.IsNullOrEmpty(sep) || string.IsNullOrEmpty(col) || table == null || !table.Columns.Contains(col))
                return new string[0];
            return SplitParts(ExpressionEngine.ToText(row[col]), sep[0]);
        }

        static string[] SplitParts(string text, char separator)
        {
            if (string.IsNullOrEmpty(text))
                return new string[0];
            var raw = text.Split(new[] { separator }, StringSplitOptions.RemoveEmptyEntries);
            var parts = new List<string>(raw.Length);
            for (var i = 0; i < raw.Length; i++)
            {
                var t = raw[i].Trim();
                if (t.Length > 0)
                    parts.Add(t);
            }
            return parts.ToArray();
        }

        static object LookupPart(string[] parts, string token)
        {
            if (parts == null || string.IsNullOrWhiteSpace(token))
                return "";
            var t = token.Trim();
            if (t.Length > 0 && t[0] == '$')
                t = t.Substring(1);
            int n;
            if (!int.TryParse(t, NumberStyles.Integer, CultureInfo.InvariantCulture, out n))
                return "";
            if (n < 1 || n > parts.Length)
                return "";
            return parts[n - 1];
        }

        static object EvalExpr(DataTable table, DataRow row, string expression, string[] parts)
        {
            return ExpressionEngine.Evaluate(expression, delegate(string name) { return FindValue(table, row, parts, name); });
        }

        static object FindValue(DataTable table, DataRow row, string[] parts, string name)
        {
            if (string.IsNullOrEmpty(name))
                return null;
            if (name[0] == '$')
                return LookupPart(parts, name);
            object builtin;
            if (ExpressionEngine.TryGetVariable(name, out builtin))
                return builtin;
            if (table == null)
                return null;
            int letter;
            if (TryParseColumnLetter(name, out letter) && letter >= 0 && letter < table.Columns.Count)
                return row[letter];
            if (table.Columns.Contains(name))
                return row[name];
            foreach (DataColumn c in table.Columns)
            {
                if (string.Equals(c.ColumnName, name, StringComparison.OrdinalIgnoreCase))
                    return row[c.ColumnName];
            }
            return null;
        }

        static bool TryParseColumnLetter(string text, out int index)
        {
            index = -1;
            if (string.IsNullOrWhiteSpace(text))
                return false;
            var t = text.Trim();
            var n = 0;
            for (var i = 0; i < t.Length; i++)
            {
                var c = t[i];
                if (c >= 'a' && c <= 'z')
                    c = (char)(c - 'a' + 'A');
                if (c < 'A' || c > 'Z')
                    return false;
                n = n * 26 + (c - 'A' + 1);
            }
            if (n < 1 || n > 16384)
                return false;
            index = n - 1;
            return true;
        }
    }
}
