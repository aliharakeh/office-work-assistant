using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using OfficeWorkAssistant.Excel;

namespace OfficeWorkAssistant.Features.Append
{
    public enum AppendMatch
    {
        ByHeader,
        ByPosition
    }

    public sealed class AppendSettings
    {
        public AppendMatch Match { get; set; }
        // Only columns every input has; otherwise every column of every input.
        public bool CommonOnly { get; set; }
        // Blank = none; otherwise a first column with the name of the step each row came from.
        public string SourceHeader { get; set; }

        public AppendSettings()
        {
            SourceHeader = "";
        }
    }

    // Pure logic: rows of several tables, one after another. No WPF.
    public static class AppendWork
    {
        // names: one per input, for the source column.
        public static DataTable Run(IList<DataTable> inputs, IList<string> names, AppendSettings s)
        {
            if (inputs == null || inputs.Count == 0)
                throw new InvalidOperationException("Connect the tables to append.");
            if (s == null)
                throw new InvalidOperationException("Set up the step first.");

            // For every output column, the input column of each table (-1 = that table has none).
            var headers = new List<string>();
            var map = new List<int[]>();
            if (s.Match == AppendMatch.ByPosition)
                ByPosition(inputs, s.CommonOnly, headers, map);
            else
                ByHeader(inputs, s.CommonOnly, headers, map);
            if (headers.Count == 0)
                throw new InvalidOperationException("The tables have no column in common.");

            var output = new DataTable();
            output.Locale = CultureInfo.InvariantCulture;
            bool withSource = !string.IsNullOrWhiteSpace(s.SourceHeader);
            if (withSource)
                ExcelFile.AddColumn(output, s.SourceHeader.Trim());
            foreach (var h in headers)
                ExcelFile.AddColumn(output, h);
            int shift = withSource ? 1 : 0;

            for (var t = 0; t < inputs.Count; t++)
            {
                string name = names != null && t < names.Count ? names[t] : (t + 1).ToString(CultureInfo.InvariantCulture);
                foreach (DataRow row in inputs[t].Rows)
                {
                    var dest = output.NewRow();
                    if (withSource)
                        dest[0] = name;
                    for (var c = 0; c < map.Count; c++)
                    {
                        int from = map[c][t];
                        dest[c + shift] = from >= 0 ? row[from] : DBNull.Value;
                    }
                    output.Rows.Add(dest);
                }
            }
            return output;
        }

        // The n-th "Qty" of one table matches the n-th "Qty" of another; case and spaces do not matter.
        static void ByHeader(IList<DataTable> inputs, bool commonOnly, List<string> headers, List<int[]> map)
        {
            var keyed = new Dictionary<string, int>(StringComparer.Ordinal);
            for (var t = 0; t < inputs.Count; t++)
            {
                var seen = new Dictionary<string, int>(StringComparer.Ordinal);
                for (var c = 0; c < inputs[t].Columns.Count; c++)
                {
                    string header = ExcelFile.Header(inputs[t].Columns[c]);
                    string norm = header.Trim().ToUpperInvariant();
                    int n;
                    seen[norm] = seen.TryGetValue(norm, out n) ? n + 1 : 1;
                    string key = norm + "#" + seen[norm].ToString(CultureInfo.InvariantCulture);
                    int at;
                    if (!keyed.TryGetValue(key, out at))
                    {
                        at = headers.Count;
                        keyed[key] = at;
                        headers.Add(header);
                        map.Add(Missing(inputs.Count));
                    }
                    map[at][t] = c;
                }
            }
            if (commonOnly)
                DropPartial(headers, map);
        }

        static void ByPosition(IList<DataTable> inputs, bool commonOnly, List<string> headers, List<int[]> map)
        {
            int most = 0;
            foreach (var t in inputs)
                most = Math.Max(most, t.Columns.Count);
            for (var c = 0; c < most; c++)
            {
                var cols = Missing(inputs.Count);
                string header = null;
                for (var t = 0; t < inputs.Count; t++)
                {
                    if (c >= inputs[t].Columns.Count)
                        continue;
                    cols[t] = c;
                    if (header == null)
                        header = ExcelFile.Header(inputs[t].Columns[c]);
                }
                headers.Add(header);
                map.Add(cols);
            }
            if (commonOnly)
                DropPartial(headers, map);
        }

        static int[] Missing(int count)
        {
            var cols = new int[count];
            for (var i = 0; i < count; i++)
                cols[i] = -1;
            return cols;
        }

        static void DropPartial(List<string> headers, List<int[]> map)
        {
            for (var c = map.Count - 1; c >= 0; c--)
            {
                if (Array.IndexOf(map[c], -1) >= 0)
                {
                    map.RemoveAt(c);
                    headers.RemoveAt(c);
                }
            }
        }
    }
}
