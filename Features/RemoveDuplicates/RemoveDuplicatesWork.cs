using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Text;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.RemoveDuplicates
{
    public enum DuplicateKeep
    {
        First,
        Last,
        UniqueOnly,
        DuplicatesOnly
    }

    public sealed class RemoveDuplicatesSettings
    {
        // true: two rows are the same when KeyFormula gives the same result for both,
        // e.g. LOWER(TRIM($Name)) & "|" & MONTH($Date). false: when Keys match.
        public bool UseFormula { get; set; }
        public string KeyFormula { get; set; }
        // Excel letters of the columns that make two rows the same; empty = every column.
        public List<string> Keys { get; set; }
        public bool IgnoreCase { get; set; }
        public bool TrimSpaces { get; set; }
        public DuplicateKeep Keep { get; set; }
        // Blank = no count column; otherwise a column with how many rows share the key.
        public string CountHeader { get; set; }

        public RemoveDuplicatesSettings()
        {
            KeyFormula = "";
            Keys = new List<string>();
            IgnoreCase = true;
            TrimSpaces = true;
            Keep = DuplicateKeep.First;
            CountHeader = "";
        }
    }

    // Pure logic: rows that repeat a key. The kept rows stay in their original order. No WPF.
    public static class RemoveDuplicatesWork
    {
        public static DataTable Run(DataTable input, RemoveDuplicatesSettings s)
        {
            if (input == null)
                throw new InvalidOperationException("Connect an input first.");
            if (s == null)
                throw new InvalidOperationException("Set up the step first.");

            List<int> keys = null;
            Func<Func<string, object>, object> formula = null;
            var lookup = new RowLookup(input);
            Func<string, object> find = lookup.Find;
            if (s.UseFormula)
            {
                string err = ExpressionEngine.Validate(s.KeyFormula);
                if (err != null)
                    throw new InvalidOperationException("Key formula: " + err);
                formula = ExpressionEngine.Compile(s.KeyFormula);
            }
            else
            {
                keys = ExcelColumns.Indexes(input, s.Keys);
                if (keys.Count == 0)
                {
                    for (var i = 0; i < input.Columns.Count; i++)
                        keys.Add(i);
                }
            }

            int n = input.Rows.Count;
            var rowKeys = new string[n];
            var count = new Dictionary<string, int>(StringComparer.Ordinal);
            var first = new Dictionary<string, int>(StringComparer.Ordinal);
            var last = new Dictionary<string, int>(StringComparer.Ordinal);
            for (var r = 0; r < n; r++)
            {
                string key;
                if (formula != null)
                {
                    lookup.Row = input.Rows[r];
                    // A row the formula cannot handle never matches another row.
                    try
                    {
                        key = Normalize(ExpressionEngine.ToText(formula(find)), s);
                    }
                    catch
                    {
                        key = "\u001Eerror" + r.ToString(CultureInfo.InvariantCulture);
                    }
                }
                else
                    key = KeyOf(input.Rows[r], keys, s);
                rowKeys[r] = key;
                int c;
                count[key] = count.TryGetValue(key, out c) ? c + 1 : 1;
                if (!first.ContainsKey(key))
                    first[key] = r;
                last[key] = r;
            }

            var output = new DataTable();
            output.Locale = CultureInfo.InvariantCulture;
            foreach (DataColumn col in input.Columns)
                ExcelFile.AddColumn(output, ExcelFile.Header(col));
            bool withCount = !string.IsNullOrWhiteSpace(s.CountHeader);
            if (withCount)
                ExcelFile.AddColumn(output, s.CountHeader.Trim());

            int width = input.Columns.Count;
            for (var r = 0; r < n; r++)
            {
                string key = rowKeys[r];
                bool keep;
                switch (s.Keep)
                {
                    case DuplicateKeep.Last: keep = last[key] == r; break;
                    case DuplicateKeep.UniqueOnly: keep = count[key] == 1; break;
                    case DuplicateKeep.DuplicatesOnly: keep = count[key] > 1; break;
                    default: keep = first[key] == r; break;
                }
                if (!keep)
                    continue;
                var dest = output.NewRow();
                var src = input.Rows[r];
                for (var c = 0; c < width; c++)
                    dest[c] = src[c];
                if (withCount)
                    dest[width] = (double)count[key];
                output.Rows.Add(dest);
            }
            return output;
        }

        static string KeyOf(DataRow row, List<int> keys, RemoveDuplicatesSettings s)
        {
            var sb = new StringBuilder();
            foreach (var k in keys)
                sb.Append(Normalize(ExpressionEngine.ToText(row[k]), s)).Append('\u001F');
            return sb.ToString();
        }

        static string Normalize(string text, RemoveDuplicatesSettings s)
        {
            if (s.TrimSpaces)
                text = text.Trim();
            if (s.IgnoreCase)
                text = text.ToUpperInvariant();
            return text;
        }

        public static string KeepLabel(DuplicateKeep keep)
        {
            switch (keep)
            {
                case DuplicateKeep.Last: return "Keep the last of each";
                case DuplicateKeep.UniqueOnly: return "Keep only rows with no repeat";
                case DuplicateKeep.DuplicatesOnly: return "Keep only repeated rows";
                default: return "Keep the first of each";
            }
        }
    }
}
