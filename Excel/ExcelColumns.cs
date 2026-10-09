using System;
using System.Collections.Generic;
using System.Data;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Excel
{
    // Excel column letters (A = 1st column) and "A - Qty" labels, shared by the step editors.
    public static class ExcelColumns
    {
        // 0 -> A, 1 -> B, ... 25 -> Z, 26 -> AA.
        public static string Letter(int index)
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

        public static bool TryParseLetter(string s, out int index)
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

        // "A - Qty": the letter plus the real header.
        public static string Label(DataTable table, int index)
        {
            return Letter(index) + " - " + ExcelFile.Header(table.Columns[index]);
        }

        public static string[] Labels(DataTable table)
        {
            if (table == null)
                return new string[0];
            var labels = new string[table.Columns.Count];
            for (var i = 0; i < table.Columns.Count; i++)
                labels[i] = Label(table, i);
            return labels;
        }

        // "A", "a" or a label like "A - Qty" gives "A"; anything else comes back trimmed.
        public static string ExtractLetter(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";
            string t = text.Trim();
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
            if (TryParseLetter(head, out dummy))
                return head.ToUpperInvariant();
            return t;
        }

        // A letter that fits the table wins, then a header. -1 when neither is there.
        public static int IndexOf(DataTable table, string column)
        {
            if (table == null || string.IsNullOrWhiteSpace(column))
                return -1;
            int index;
            if (TryParseLetter(ExtractLetter(column), out index) && index < table.Columns.Count)
                return index;
            var named = ExcelFile.FindColumn(table, column.Trim());
            return named != null ? named.Ordinal : -1;
        }

        // Like IndexOf, but a missing column is an error.
        public static int Index(DataTable table, string column)
        {
            int index = IndexOf(table, column);
            if (index < 0)
                throw new InvalidOperationException("Column '" + (column ?? "").Trim() + "' is not in this step's input.");
            return index;
        }

        public static List<int> Indexes(DataTable table, IList<string> columns)
        {
            var list = new List<int>();
            if (columns == null)
                return list;
            foreach (var c in columns)
            {
                if (string.IsNullOrWhiteSpace(c))
                    continue;
                int i = Index(table, c);
                if (!list.Contains(i))
                    list.Add(i);
            }
            return list;
        }

        // A name that a formula can use after $: the header when it is one word, otherwise the letter.
        public static string FormulaName(DataTable table, int index)
        {
            string header = ExcelFile.Header(table.Columns[index]);
            if (header.Length == 0 || !(char.IsLetter(header[0]) || header[0] == '_'))
                return Letter(index);
            for (var i = 0; i < header.Length; i++)
            {
                if (!char.IsLetterOrDigit(header[i]) && header[i] != '_')
                    return Letter(index);
            }
            // A one-word header that reads as a letter ("ID", "Qty") would hit that column instead.
            int asLetter;
            if (TryParseLetter(header, out asLetter) && asLetter < table.Columns.Count && asLetter != index)
                return Letter(index);
            return header;
        }
    }

    // Resolves $names in a formula against one row: $A by letter, $Qty by header, $Today and
    // the other built-ins, and $Value for the current cell when one is set. Set Row (and Value)
    // before each evaluation; names are looked up once per table.
    public sealed class RowLookup
    {
        readonly DataTable _table;
        readonly Dictionary<string, int> _columns = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        public DataRow Row;
        public bool HasValue;
        public object Value;

        public RowLookup(DataTable table)
        {
            _table = table;
        }

        public object Find(string name)
        {
            if (string.IsNullOrEmpty(name))
                return null;
            if (HasValue && string.Equals(name, "Value", StringComparison.OrdinalIgnoreCase))
                return Value;
            int index;
            if (!_columns.TryGetValue(name, out index))
            {
                index = ExcelColumns.IndexOf(_table, name);
                _columns[name] = index;
            }
            if (index >= 0 && Row != null)
                return Row[index];
            object v;
            if (ExpressionEngine.TryGetVariable(name, out v))
                return v;
            return null;
        }
    }
}
