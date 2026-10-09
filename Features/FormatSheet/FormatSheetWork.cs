using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using OfficeWorkAssistant.Excel;

namespace OfficeWorkAssistant.Features.FormatSheet
{
    public sealed class ColumnFormat
    {
        // Excel letter.
        public string Column { get; set; }
        // In Excel's width units (about one character each); 0 = leave it.
        public double Width { get; set; }
        public string NumberFormat { get; set; }
        public CellAlign Align { get; set; }
        public bool Wrap { get; set; }

        public ColumnFormat()
        {
            Column = "";
            NumberFormat = "";
        }

        public override string ToString()
        {
            var parts = new List<string>();
            if (Width > 0)
                parts.Add("width " + Width.ToString("0.#", CultureInfo.InvariantCulture));
            if (!string.IsNullOrWhiteSpace(NumberFormat))
                parts.Add("format " + NumberFormat);
            if (Align != CellAlign.General)
                parts.Add(Align.ToString().ToLowerInvariant());
            if (Wrap)
                parts.Add("wrap");
            return Column + ": " + (parts.Count == 0 ? "nothing" : string.Join(", ", parts.ToArray()));
        }
    }

    public sealed class FormatSheetSettings
    {
        public bool StyleHeader { get; set; }
        public CellStyle Header { get; set; }
        public bool FreezeHeader { get; set; }
        public int FreezeColumns { get; set; }
        public bool AutoFilter { get; set; }
        public bool AutoFit { get; set; }
        public bool BandRows { get; set; }
        public string BandFill { get; set; }
        public bool Borders { get; set; }
        public List<ColumnFormat> Columns { get; set; }

        public FormatSheetSettings()
        {
            StyleHeader = true;
            Header = new CellStyle { Bold = true, Fill = "#D9E1F2", Border = CellBorder.Thin };
            FreezeHeader = true;
            AutoFilter = true;
            AutoFit = true;
            BandFill = "#F2F2F2";
            Columns = new List<ColumnFormat>();
        }
    }

    // Pure logic: the sheet-wide look Save writes. The data is not changed. No WPF.
    public static class FormatSheetWork
    {
        public static DataTable Run(DataTable input, FormatSheetSettings s)
        {
            if (input == null)
                throw new InvalidOperationException("Connect an input first.");
            if (s == null)
                throw new InvalidOperationException("Set up the step first.");
            if (s.FreezeColumns < 0)
                throw new InvalidOperationException("The number of columns to freeze cannot be below 0.");
            if (s.StyleHeader)
            {
                string err = s.Header == null ? null : s.Header.Check();
                if (err != null)
                    throw new InvalidOperationException("Header: " + err);
            }
            if (s.BandRows && string.IsNullOrEmpty(NamedColors.ToHex(s.BandFill)))
                throw new InvalidOperationException("The band colour '" + s.BandFill + "' is not a colour name or #RRGGBB.");

            var old = ExcelFile.GetStyle(input);
            var style = old != null ? old.Copy() : new SheetStyle();
            if (s.StyleHeader && s.Header != null)
                style.Header = style.Header == null ? s.Header.Copy() : style.Header.Over(s.Header);
            if (s.FreezeHeader)
                style.FreezeRows = 1;
            if (s.FreezeColumns > 0)
                style.FreezeColumns = s.FreezeColumns;
            if (s.AutoFilter)
                style.AutoFilter = true;
            if (s.AutoFit)
                style.AutoFit = true;
            if (s.Borders)
            {
                var all = new CellStyle { Border = CellBorder.Thin };
                style.AllCells = style.AllCells == null ? all : style.AllCells.Over(all);
            }
            if (s.BandRows)
                style.Band = new CellStyle { Fill = NamedColors.ToHex(s.BandFill) };

            foreach (var f in s.Columns)
            {
                int col = ExcelColumns.Index(input, f.Column);
                style.SetWidth(col, f.Width);
                var look = new CellStyle { NumberFormat = (f.NumberFormat ?? "").Trim(), Align = f.Align, Wrap = f.Wrap };
                if (!look.IsEmpty)
                    style.StyleColumn(col, look, true);
            }
            return ExcelFile.WithStyle(input, style);
        }

        public static string Summary(FormatSheetSettings s)
        {
            var parts = new List<string>();
            if (s.StyleHeader) parts.Add("header");
            if (s.FreezeHeader || s.FreezeColumns > 0) parts.Add("freeze");
            if (s.AutoFilter) parts.Add("filter");
            if (s.AutoFit) parts.Add("fit widths");
            if (s.BandRows) parts.Add("bands");
            if (s.Borders) parts.Add("borders");
            if (s.Columns.Count > 0) parts.Add(s.Columns.Count + " column format" + (s.Columns.Count == 1 ? "" : "s"));
            return parts.Count == 0 ? "Nothing to format" : string.Join(", ", parts.ToArray());
        }
    }
}
