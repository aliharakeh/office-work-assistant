using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using OfficeWorkAssistant.Excel;

namespace OfficeWorkAssistant.Features.ArrangeColumns
{
    public sealed class ArrangeColumn
    {
        // Excel letter of the input column.
        public string Source { get; set; }
        // New header; blank keeps the old one.
        public string Header { get; set; }
        public bool Keep { get; set; }

        public ArrangeColumn()
        {
            Source = "";
            Header = "";
            Keep = true;
        }
    }

    // The listed columns, in list order, kept or dropped and maybe renamed.
    public sealed class ArrangeColumnsSettings
    {
        public List<ArrangeColumn> Columns { get; set; }
        // Columns not in the list (say, added by an earlier step later on) go at the end.
        public bool KeepNew { get; set; }

        public ArrangeColumnsSettings()
        {
            Columns = new List<ArrangeColumn>();
            KeepNew = true;
        }
    }

    // Pure logic: pick, order and rename columns. No WPF.
    public static class ArrangeColumnsWork
    {
        public static DataTable Run(DataTable input, ArrangeColumnsSettings s)
        {
            if (input == null)
                throw new InvalidOperationException("Connect an input first.");
            if (s == null)
                throw new InvalidOperationException("Set up the columns first.");

            var sources = new List<int>();
            var headers = new List<string>();
            var listed = new HashSet<int>();
            foreach (var c in s.Columns)
            {
                int index = ExcelColumns.Index(input, c.Source);
                listed.Add(index);
                if (!c.Keep)
                    continue;
                sources.Add(index);
                headers.Add(string.IsNullOrWhiteSpace(c.Header) ? ExcelFile.Header(input.Columns[index]) : c.Header.Trim());
            }
            if (s.KeepNew)
            {
                for (var i = 0; i < input.Columns.Count; i++)
                {
                    if (listed.Contains(i))
                        continue;
                    sources.Add(i);
                    headers.Add(ExcelFile.Header(input.Columns[i]));
                }
            }
            if (sources.Count == 0)
                throw new InvalidOperationException("Keep at least one column.");

            var output = new DataTable();
            output.Locale = CultureInfo.InvariantCulture;
            foreach (var h in headers)
                ExcelFile.AddColumn(output, h);
            foreach (DataRow row in input.Rows)
            {
                var dest = output.NewRow();
                for (var i = 0; i < sources.Count; i++)
                    dest[i] = row[sources[i]];
                output.Rows.Add(dest);
            }
            return output;
        }

        public static string Summary(ArrangeColumnsSettings s)
        {
            int kept = 0;
            int renamed = 0;
            int dropped = 0;
            foreach (var c in s.Columns)
            {
                if (!c.Keep)
                    dropped++;
                else
                {
                    kept++;
                    if (!string.IsNullOrWhiteSpace(c.Header))
                        renamed++;
                }
            }
            return kept + " kept, " + dropped + " dropped, " + renamed + " renamed" + (s.KeepNew ? ", new ones kept" : "");
        }
    }
}
