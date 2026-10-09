using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.SplitCombine
{
    public enum SplitMode
    {
        Delimiter,
        FixedWidths,
        Pattern
    }

    public sealed class SplitColumnSettings
    {
        // Excel letter.
        public string Column { get; set; }
        public SplitMode Mode { get; set; }
        // \t is a tab.
        public string Delimiter { get; set; }
        // Character counts such as "3,5,2"; anything after them is one more part.
        public string Widths { get; set; }
        // A regular expression. With ( ) groups the parts are the groups of the first match;
        // without, the text is cut wherever it matches.
        public string Pattern { get; set; }
        // 0 = as many parts as the text has; otherwise the last part keeps the rest.
        public int MaxParts { get; set; }
        // Headers for the parts, in order; blank ones become "<column> 1", "<column> 2", ...
        public List<string> Headers { get; set; }
        public bool TrimParts { get; set; }
        public bool KeepSource { get; set; }

        public SplitColumnSettings()
        {
            Column = "";
            Delimiter = ",";
            Widths = "";
            Pattern = "";
            Headers = new List<string>();
            TrimParts = true;
            KeepSource = true;
        }
    }

    public sealed class CombineColumnsSettings
    {
        // Excel letters, in the order they are joined.
        public List<string> Columns { get; set; }
        // \t is a tab, \n a line break.
        public string Separator { get; set; }
        public bool SkipBlanks { get; set; }
        public string Header { get; set; }
        public bool RemoveSources { get; set; }

        public CombineColumnsSettings()
        {
            Columns = new List<string>();
            Separator = " ";
            SkipBlanks = true;
            Header = "Combined";
        }
    }

    // Pure logic: one column into several, several into one. No WPF.
    public static class SplitCombineWork
    {
        const int MaxColumns = 200;

        public static string Unescape(string text)
        {
            return (text ?? "").Replace("\\t", "\t").Replace("\\n", "\n");
        }

        // ---------- split ----------

        public static DataTable Split(DataTable input, SplitColumnSettings s)
        {
            if (input == null)
                throw new InvalidOperationException("Connect an input first.");
            if (s == null)
                throw new InvalidOperationException("Set up the split first.");
            int source = ExcelColumns.Index(input, s.Column);

            int[] widths = null;
            Regex regex = null;
            string delimiter = null;
            if (s.Mode == SplitMode.FixedWidths)
                widths = ParseWidths(s.Widths);
            else if (s.Mode == SplitMode.Pattern)
            {
                if (string.IsNullOrEmpty(s.Pattern))
                    throw new InvalidOperationException("Enter a pattern to split on.");
                try
                {
                    regex = new Regex(s.Pattern, RegexOptions.IgnoreCase, TimeSpan.FromSeconds(2));
                }
                catch (ArgumentException ex)
                {
                    throw new InvalidOperationException("The pattern is not valid: " + ex.Message);
                }
            }
            else
            {
                delimiter = Unescape(s.Delimiter);
                if (delimiter.Length == 0)
                    throw new InvalidOperationException("Enter the text to split on, e.g. a comma.");
            }

            var parts = new List<string[]>(input.Rows.Count);
            int count = 0;
            foreach (DataRow row in input.Rows)
            {
                var v = row[source];
                string text = v == null || v == DBNull.Value ? "" : ExpressionEngine.ToText(v);
                string[] p;
                if (widths != null)
                    p = ByWidths(text, widths);
                else if (regex != null)
                    p = ByPattern(text, regex, s.MaxParts);
                else
                    p = s.MaxParts > 0 ? text.Split(new[] { delimiter }, s.MaxParts, StringSplitOptions.None) : text.Split(new[] { delimiter }, StringSplitOptions.None);
                parts.Add(p);
                count = Math.Max(count, p.Length);
            }
            if (s.MaxParts > 0)
                count = Math.Min(count, s.MaxParts);
            count = Math.Max(count, Math.Max(1, CountHeaders(s.Headers)));
            if (count > MaxColumns)
                throw new InvalidOperationException("The split would make " + count + " columns. Set a maximum number of parts.");

            string name = ExcelFile.Header(input.Columns[source]);
            var output = new DataTable();
            output.Locale = CultureInfo.InvariantCulture;
            var from = new List<int>(); // input column, or -1 - part
            for (var c = 0; c < input.Columns.Count; c++)
            {
                if (c != source || s.KeepSource)
                {
                    ExcelFile.AddColumn(output, ExcelFile.Header(input.Columns[c]));
                    from.Add(c);
                }
                if (c != source)
                    continue;
                for (var i = 0; i < count; i++)
                {
                    string h = i < s.Headers.Count && !string.IsNullOrWhiteSpace(s.Headers[i])
                        ? s.Headers[i].Trim()
                        : name + " " + (i + 1).ToString(CultureInfo.InvariantCulture);
                    ExcelFile.AddColumn(output, h);
                    from.Add(-1 - i);
                }
            }

            for (var r = 0; r < input.Rows.Count; r++)
            {
                var src = input.Rows[r];
                var p = parts[r];
                var dest = output.NewRow();
                for (var c = 0; c < from.Count; c++)
                {
                    if (from[c] >= 0)
                    {
                        dest[c] = src[from[c]];
                        continue;
                    }
                    int i = -1 - from[c];
                    string part = i < p.Length ? p[i] : null;
                    if (part != null && s.TrimParts)
                        part = part.Trim();
                    dest[c] = string.IsNullOrEmpty(part) ? (object)DBNull.Value : part;
                }
                output.Rows.Add(dest);
            }
            return output;
        }

        static int CountHeaders(List<string> headers)
        {
            int n = 0;
            for (var i = 0; i < headers.Count; i++)
            {
                if (!string.IsNullOrWhiteSpace(headers[i]))
                    n = i + 1;
            }
            return n;
        }

        public static int[] ParseWidths(string text)
        {
            var list = new List<int>();
            foreach (var piece in (text ?? "").Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries))
            {
                int w;
                if (!int.TryParse(piece, NumberStyles.Integer, CultureInfo.InvariantCulture, out w) || w <= 0)
                    throw new InvalidOperationException("Widths must be whole numbers above 0, such as 3,5,2.");
                list.Add(w);
            }
            if (list.Count == 0)
                throw new InvalidOperationException("Enter the widths of the parts, such as 3,5,2.");
            return list.ToArray();
        }

        static string[] ByWidths(string text, int[] widths)
        {
            var parts = new List<string>();
            int at = 0;
            foreach (var w in widths)
            {
                if (at >= text.Length)
                    break;
                int n = Math.Min(w, text.Length - at);
                parts.Add(text.Substring(at, n));
                at += n;
            }
            if (at < text.Length)
                parts.Add(text.Substring(at));
            return parts.ToArray();
        }

        static string[] ByPattern(string text, Regex regex, int max)
        {
            if (regex.GetGroupNumbers().Length > 1)
            {
                var m = regex.Match(text);
                if (!m.Success)
                    return new string[0];
                var groups = new string[m.Groups.Count - 1];
                for (var i = 1; i < m.Groups.Count; i++)
                    groups[i - 1] = m.Groups[i].Value;
                return groups;
            }
            return max > 0 ? regex.Split(text, max) : regex.Split(text);
        }

        // ---------- combine ----------

        public static DataTable Combine(DataTable input, CombineColumnsSettings s)
        {
            if (input == null)
                throw new InvalidOperationException("Connect an input first.");
            if (s == null)
                throw new InvalidOperationException("Set up the step first.");
            var sources = ExcelColumns.Indexes(input, s.Columns);
            if (sources.Count < 2)
                throw new InvalidOperationException("Pick at least two columns to combine.");
            if (string.IsNullOrWhiteSpace(s.Header))
                throw new InvalidOperationException("Enter a header for the combined column.");

            // Removed sources: the result takes the leftmost one's place. Kept: it goes after the rightmost.
            int left = int.MaxValue;
            int right = -1;
            foreach (var i in sources)
            {
                left = Math.Min(left, i);
                right = Math.Max(right, i);
            }
            int at = s.RemoveSources ? left : right;
            var remove = new HashSet<int>(s.RemoveSources ? sources : new List<int>());
            string separator = Unescape(s.Separator);

            var output = new DataTable();
            output.Locale = CultureInfo.InvariantCulture;
            var from = new List<int>(); // input column, or -1 for the combined one
            for (var c = 0; c < input.Columns.Count; c++)
            {
                if (s.RemoveSources && c == at)
                {
                    ExcelFile.AddColumn(output, s.Header.Trim());
                    from.Add(-1);
                }
                if (!remove.Contains(c))
                {
                    ExcelFile.AddColumn(output, ExcelFile.Header(input.Columns[c]));
                    from.Add(c);
                }
                if (!s.RemoveSources && c == at)
                {
                    ExcelFile.AddColumn(output, s.Header.Trim());
                    from.Add(-1);
                }
            }

            var sb = new StringBuilder();
            foreach (DataRow row in input.Rows)
            {
                sb.Length = 0;
                bool any = false;
                foreach (var i in sources)
                {
                    var v = row[i];
                    string text = v == null || v == DBNull.Value ? "" : ExpressionEngine.ToText(v);
                    if (s.SkipBlanks && text.Trim().Length == 0)
                        continue;
                    if (any)
                        sb.Append(separator);
                    sb.Append(text);
                    any = true;
                }
                var dest = output.NewRow();
                for (var c = 0; c < from.Count; c++)
                    dest[c] = from[c] >= 0 ? row[from[c]] : (sb.Length == 0 ? (object)DBNull.Value : sb.ToString());
                output.Rows.Add(dest);
            }
            return output;
        }
    }
}
