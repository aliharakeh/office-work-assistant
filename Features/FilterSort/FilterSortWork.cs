using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using ClosedXML.Excel;
using WorkAssistant.Expressions;

namespace WorkAssistant.Features.FilterSort
{
    public enum FilterConditionKind
    {
        Column,
        Formula
    }

    public sealed class FilterCondition
    {
        public FilterConditionKind Kind { get; set; }
        public string Column { get; set; }
        public string Operator { get; set; }
        public string Value { get; set; }
        public string Expression { get; set; }

        public FilterCondition()
        {
            Kind = FilterConditionKind.Column;
            Column = "";
            Operator = "==";
            Value = "";
            Expression = "";
        }

        public override string ToString()
        {
            if (Kind == FilterConditionKind.Formula)
                return "Formula: " + (Expression != null ? Expression : "");
            string col = Column != null ? Column : "";
            if (Operator == "isEmpty")
                return col + " is empty";
            if (Operator == "isNotEmpty")
                return col + " is not empty";
            return col + " " + (Operator != null ? Operator : "") + " " + (Value != null ? Value : "");
        }
    }

    public sealed class SortKey
    {
        public string Column { get; set; }
        public bool Descending { get; set; }

        public SortKey()
        {
            Column = "";
        }

        public override string ToString()
        {
            string name = Column != null ? Column : "";
            return Descending ? name + " (Z to A)" : name + " (A to Z)";
        }
    }

    public sealed class FilterSortSourceResult
    {
        public DataTable Table { get; set; }
        public string[] Sheets { get; set; }
        public string Sheet { get; set; }
        public int TotalRows { get; set; }
        public bool Truncated { get; set; }
    }

    // Pure logic: load, filter, sort and save. No WPF, no page code.
    public static class FilterSortWork
    {
        // Index-aligned with the Op combo in the page: same order, same count.
        public static readonly string[] Operators =
        {
            "==", "!=", ">", ">=", "<", "<=", "contains", "startsWith", "endsWith", "isEmpty", "isNotEmpty"
        };

        public static FilterSortSourceResult Load(string path, string sheetName)
        {
            return Load(path, sheetName, 0);
        }

        // maxRows <= 0 loads everything; a positive value keeps the UI preview small.
        public static FilterSortSourceResult Load(string path, string sheetName, int maxRows)
        {
            try
            {
                // Shared read: succeeds even while Excel has the file open.
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var wb = new XLWorkbook(stream))
                {
                    if (wb.Worksheets.Count == 0)
                        throw new InvalidOperationException("The workbook has no sheets.");

                    var sheets = new string[wb.Worksheets.Count];
                    var i = 0;
                    foreach (var w in wb.Worksheets)
                        sheets[i++] = w.Name;

                    var ws = string.IsNullOrEmpty(sheetName)
                        ? wb.Worksheet(1)
                        : wb.Worksheet(sheetName);

                    var range = ws.RangeUsed();
                    var total = range == null ? 0 : range.RowCount() - 1;
                    if (total < 0)
                        total = 0;

                    var table = ToTable(ws, maxRows);

                    return new FilterSortSourceResult
                    {
                        Table = table,
                        Sheets = sheets,
                        Sheet = ws.Name,
                        TotalRows = total,
                        Truncated = maxRows > 0 && total > table.Rows.Count
                    };
                }
            }
            catch (Exception ex)
            {
                throw new IOException("Could not read \"" + Path.GetFileName(path) + "\". " +
                    "Make sure it is a valid .xlsx file (a .xls renamed to .xlsx will not open). " + ex.Message, ex);
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
                        if (v == null || v == DBNull.Value)
                            continue;
                        if (v is DateTime)
                        {
                            DateTime dt = (DateTime)v;
                            ws.Cell(r + 2, c + 1).Value = dt;
                            ws.Cell(r + 2, c + 1).Style.DateFormat.Format =
                                dt.TimeOfDay == TimeSpan.Zero ? "dd/MM/yyyy" : "dd/MM/yyyy HH:mm:ss";
                        }
                        else if (v is bool)
                            ws.Cell(r + 2, c + 1).Value = (bool)v;
                        else if (v is double)
                            ws.Cell(r + 2, c + 1).Value = (double)v;
                        else if (v is float)
                            ws.Cell(r + 2, c + 1).Value = (double)(float)v;
                        else if (v is int)
                            ws.Cell(r + 2, c + 1).Value = (double)(int)v;
                        else if (v is long)
                            ws.Cell(r + 2, c + 1).Value = (double)(long)v;
                        else if (v is decimal)
                            ws.Cell(r + 2, c + 1).Value = (double)(decimal)v;
                        else
                            ws.Cell(r + 2, c + 1).Value = Convert.ToString(v, CultureInfo.InvariantCulture);
                    }
                }

                wb.SaveAs(path);
            }
        }

        public static string Validate(IList<FilterCondition> conditions)
        {
            if (conditions == null)
                return null;
            for (var i = 0; i < conditions.Count; i++)
            {
                FilterCondition cond = conditions[i];
                if (cond == null)
                    continue;
                if (cond.Kind == FilterConditionKind.Column)
                {
                    if (string.IsNullOrWhiteSpace(cond.Column))
                        return "Condition #" + (i + 1) + ": choose a column.";
                }
                else
                {
                    string err = ExpressionEngine.Validate(cond.Expression);
                    if (err != null)
                        return "Condition #" + (i + 1) + ": " + err;
                }
            }
            return null;
        }

        // Keeps the source column order, filters rows, then sorts.
        // keepColumns: null keeps every column; letters like "A" pick by position.
        public static DataTable Apply(DataTable source, IList<string> keepColumns, IList<FilterCondition> conditions, bool matchAll, IList<SortKey> sorts)
        {
            if (source == null)
                throw new InvalidOperationException("Choose a data file first.");
            string error = Validate(conditions);
            if (error != null)
                throw new InvalidOperationException(error);

            List<int> keep = ResolveColumns(source, keepColumns);
            if (keep.Count == 0)
                throw new InvalidOperationException("Keep at least one column.");

            var rows = new List<int>();
            for (var r = 0; r < source.Rows.Count; r++)
            {
                if (Passes(source, source.Rows[r], conditions, matchAll))
                    rows.Add(r);
            }

            var order = rows.ToArray();
            if (sorts != null && sorts.Count > 0)
            {
                var keys = new int[sorts.Count];
                var desc = new bool[sorts.Count];
                for (var i = 0; i < sorts.Count; i++)
                {
                    if (sorts[i] == null || string.IsNullOrWhiteSpace(sorts[i].Column))
                        throw new InvalidOperationException("Sort #" + (i + 1) + ": choose a column.");
                    keys[i] = OneColumn(source, sorts[i].Column);
                    desc[i] = sorts[i].Descending;
                }
                Array.Sort(order, delegate(int x, int y)
                {
                    int c = CompareRows(source.Rows[x], source.Rows[y], keys, desc);
                    return c != 0 ? c : x.CompareTo(y); // ties keep the original row order
                });
            }

            var output = new DataTable();
            output.Locale = CultureInfo.InvariantCulture;
            for (var c = 0; c < keep.Count; c++)
                output.Columns.Add(source.Columns[keep[c]].ColumnName, typeof(object));

            for (var i = 0; i < order.Length; i++)
            {
                DataRow row = source.Rows[order[i]];
                DataRow dest = output.NewRow();
                for (var c = 0; c < keep.Count; c++)
                    dest[c] = row[keep[c]];
                output.Rows.Add(dest);
            }
            return output;
        }

        static bool Passes(DataTable source, DataRow row, IList<FilterCondition> conditions, bool matchAll)
        {
            if (conditions == null || conditions.Count == 0)
                return true;
            if (matchAll)
            {
                for (var i = 0; i < conditions.Count; i++)
                {
                    if (!Pass(source, row, conditions[i]))
                        return false;
                }
                return true;
            }
            for (var i = 0; i < conditions.Count; i++)
            {
                if (Pass(source, row, conditions[i]))
                    return true;
            }
            return false;
        }

        static bool Pass(DataTable source, DataRow row, FilterCondition cond)
        {
            if (cond == null)
                return true;

            if (cond.Kind == FilterConditionKind.Formula)
            {
                if (string.IsNullOrWhiteSpace(cond.Expression))
                    return true;
                try
                {
                    return ExpressionEngine.ToBool(
                        ExpressionEngine.Evaluate(cond.Expression, delegate(string name) { return FindValue(source, row, name); }));
                }
                catch
                {
                    return false;
                }
            }

            object cell = FindValue(source, row, cond.Column);
            string op = string.IsNullOrWhiteSpace(cond.Operator) ? "==" : cond.Operator;
            if (op == "isEmpty")
                return IsEmpty(cell);
            if (op == "isNotEmpty")
                return !IsEmpty(cell);

            // $Today and the other built-ins expand; plain values pass through.
            string value = ExpressionEngine.ExpandVariables(cond.Value);
            string text = ExpressionEngine.ToText(cell);
            if (op == "contains")
                return text.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0;
            if (op == "startsWith")
                return text.StartsWith(value, StringComparison.OrdinalIgnoreCase);
            if (op == "endsWith")
                return text.EndsWith(value, StringComparison.OrdinalIgnoreCase);
            return ExpressionEngine.Compare(op, cell, value);
        }

        static bool IsEmpty(object value)
        {
            if (value == null || value == DBNull.Value)
                return true;
            return ExpressionEngine.ToText(value).Trim().Length == 0;
        }

        static object FindValue(DataTable source, DataRow row, string name)
        {
            if (string.IsNullOrEmpty(name))
                return null;
            if (source != null)
            {
                // Excel letters win: $A = first column, $B = second, etc.
                int letterIndex;
                if (TryParseColumnLetter(ExtractLetter(name), out letterIndex) &&
                    letterIndex >= 0 && letterIndex < source.Columns.Count)
                    return row[letterIndex];
                if (source.Columns.Contains(name))
                    return row[name];
                foreach (DataColumn c in source.Columns)
                {
                    if (string.Equals(c.ColumnName, name, StringComparison.OrdinalIgnoreCase))
                        return row[c.ColumnName];
                }
            }
            object v;
            if (ExpressionEngine.TryGetVariable(name, out v))
                return v;
            return null;
        }

        static int CompareRows(DataRow a, DataRow b, int[] keys, bool[] desc)
        {
            for (var k = 0; k < keys.Length; k++)
            {
                object va = a[keys[k]];
                object vb = b[keys[k]];
                bool ea = IsEmpty(va);
                bool eb = IsEmpty(vb);
                if (ea || eb)
                {
                    if (ea && eb)
                        continue; // both empty: the next key decides
                    return ea ? 1 : -1; // empty cells sink to the bottom either way
                }
                int c = CompareValues(va, vb);
                if (c != 0)
                    return desc[k] ? -c : c;
            }
            return 0;
        }

        static int CompareValues(object a, object b)
        {
            double x;
            double y;
            if (ExpressionEngine.TryToNumber(a, out x) && ExpressionEngine.TryToNumber(b, out y))
                return x.CompareTo(y);

            DateTime da;
            DateTime db;
            if ((ExpressionEngine.IsDateLike(a) || ExpressionEngine.IsDateLike(b)) &&
                ExpressionEngine.TryToDate(a, out da) && ExpressionEngine.TryToDate(b, out db))
                return DateTime.Compare(da, db);

            return string.Compare(ExpressionEngine.ToText(a), ExpressionEngine.ToText(b), StringComparison.OrdinalIgnoreCase);
        }

        // null keeps every column; otherwise letters ("A") or header names.
        static List<int> ResolveColumns(DataTable source, IList<string> columns)
        {
            var seen = new HashSet<int>();
            if (columns == null)
            {
                for (var i = 0; i < source.Columns.Count; i++)
                    seen.Add(i);
            }
            else
            {
                for (var i = 0; i < columns.Count; i++)
                {
                    if (string.IsNullOrWhiteSpace(columns[i]))
                        continue;
                    seen.Add(OneColumn(source, columns[i]));
                }
            }
            var list = new List<int>(seen);
            list.Sort();
            return list;
        }

        static int OneColumn(DataTable source, string column)
        {
            string name = column != null ? column.Trim() : "";
            int index;
            if (TryParseColumnLetter(ExtractLetter(name), out index) &&
                index >= 0 && index < source.Columns.Count)
                return index;
            if (source.Columns.Contains(name))
                return source.Columns.IndexOf(name);
            foreach (DataColumn c in source.Columns)
            {
                if (string.Equals(c.ColumnName, name, StringComparison.OrdinalIgnoreCase))
                    return c.Ordinal;
            }
            throw new InvalidOperationException("Column '" + name + "' is not in the file.");
        }

        // Excel-style column letters: 0->A, 1->B, ... 25->Z, 26->AA, etc.
        public static string ColumnLetter(int index)
        {
            string s = "";
            int n = index + 1;
            while (n > 0)
            {
                n--;
                s = (char)('A' + (n % 26)) + s;
                n /= 26;
            }
            return s;
        }

        public static bool TryParseColumnLetter(string s, out int index)
        {
            index = -1;
            if (string.IsNullOrEmpty(s))
                return false;
            string t = s.Trim();
            if (t.Length == 0 || t.Length > 3)
                return false;
            int n = 0;
            for (var i = 0; i < t.Length; i++)
            {
                char c = t[i];
                if (c >= 'a' && c <= 'z')
                    c = (char)(c - 'a' + 'A');
                if (c < 'A' || c > 'Z')
                    return false;
                n = n * 26 + (c - 'A' + 1);
            }
            // Excel max is XFD (16384 columns).
            if (n < 1 || n > 16384)
                return false;
            index = n - 1;
            return true;
        }

        // "A - Qty" style labels so users see the letter plus the real header.
        public static string SourceLabel(DataTable table, int index)
        {
            return ColumnLetter(index) + " - " + table.Columns[index].ColumnName;
        }

        public static string[] SourceLabels(DataTable table)
        {
            if (table == null)
                return new string[0];
            var labels = new string[table.Columns.Count];
            for (var i = 0; i < table.Columns.Count; i++)
                labels[i] = SourceLabel(table, i);
            return labels;
        }

        // Accepts "A", "a", or a UI label like "A - Qty" and returns "A".
        public static string ExtractLetter(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";
            string t = text.Trim();
            // Take the first token before space, '-', or ':' ("A - Qty" -> "A").
            int end = t.Length;
            for (var i = 0; i < t.Length; i++)
            {
                if (t[i] == ' ' || t[i] == '-' || t[i] == ':')
                {
                    end = i;
                    break;
                }
            }
            string head = t.Substring(0, end).Trim();
            int dummy;
            if (TryParseColumnLetter(head, out dummy))
                return head.ToUpperInvariant();
            return t;
        }

        static DataTable ToTable(IXLWorksheet ws, int maxRows)
        {
            var table = new DataTable();
            table.Locale = CultureInfo.InvariantCulture;
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

            var loaded = 0;
            foreach (var row in range.RowsUsed())
            {
                if (row.RowNumber() == header.RowNumber())
                    continue;
                if (maxRows > 0 && loaded >= maxRows)
                    break;
                var dr = table.NewRow();
                for (var c = 1; c <= colCount; c++)
                    dr[c - 1] = CellValue(row.Cell(c));
                table.Rows.Add(dr);
                loaded++;
            }

            return table;
        }

        // One odd cell must not kill the whole sheet.
        static object CellValue(IXLCell cell)
        {
            if (cell == null || cell.IsEmpty())
                return DBNull.Value;
            try
            {
                return cell.Value;
            }
            catch
            {
                try { return cell.GetFormattedString(); }
                catch { return DBNull.Value; }
            }
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
    }
}
