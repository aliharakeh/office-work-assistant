using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using OfficeWorkAssistant.Excel;

namespace OfficeWorkAssistant.Features.FormatSheet
{
    public sealed class ColumnFormat
    {
        // Excel letter in the editor, header or binding ID in a saved pipeline.
        public string Column { get; set; }
        public bool ColumnIsId { get; set; }
        // In Excel's width units (about one character each). Zero inherits the default.
        public double Width { get; set; }
        public CellStyle Style { get; set; }
        // Older saved formats store these directly on the column.
        public string NumberFormat { get; set; }
        public CellAlign Align { get; set; }
        public bool Wrap { get; set; }

        public ColumnFormat()
        {
            Column = "";
            NumberFormat = "";
        }

        public CellStyle GetStyle()
        {
            return new CellStyle { NumberFormat = NumberFormat, Align = Align, Wrap = Wrap }.Over(Style);
        }

        public string Describe()
        {
            var parts = new List<string>();
            if (Width > 0)
                parts.Add("width " + Width.ToString("0.#", CultureInfo.InvariantCulture));
            var look = GetStyle();
            if (!look.IsEmpty)
                parts.Add(look.Describe());
            return parts.Count == 0 ? "uses defaults" : string.Join(", ", parts.ToArray());
        }

        public static bool TryWidth(string text, out double width)
        {
            width = 0;
            return string.IsNullOrWhiteSpace(text) ||
                (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out width) && IsValidWidth(width));
        }

        public static bool IsValidWidth(double width)
        {
            return width >= 0 && width <= 255;
        }
    }

    public sealed class FormatSheetSettings
    {
        public bool StyleHeader { get; set; }
        public CellStyle Header { get; set; }
        public bool RightToLeft { get; set; }
        public bool FreezeHeader { get; set; }
        public int FreezeColumns { get; set; }
        public bool AutoFilter { get; set; }
        public bool AutoFit { get; set; }
        public bool BandRows { get; set; }
        public string BandFill { get; set; }
        public CellStyle SheetCells { get; set; }
        public bool BorderValuesOnly { get; set; }
        public CellStyle DataRows { get; set; }
        public CellStyle DefaultColumn { get; set; }
        public double DefaultColumnWidth { get; set; }
        // Older saved pipelines use this for thin borders on all data cells.
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

        public CellStyle GetDataRowStyle()
        {
            return Borders ? new CellStyle { Border = CellBorder.Thin }.Over(DataRows) : DataRows;
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
            if (s.BandRows && string.IsNullOrEmpty(NamedColors.ToHex(s.BandFill)))
                throw new InvalidOperationException("The band colour '" + s.BandFill + "' is not a colour name or #RRGGBB.");
            if (!ColumnFormat.IsValidWidth(s.DefaultColumnWidth))
                throw new InvalidOperationException("Default column width must be a number from 0 to 255.");

            var old = ExcelFile.GetStyle(input);
            var style = old != null ? old.Copy() : new SheetStyle();
            if (s.StyleHeader)
                style.Header = Apply(style.Header, s.Header, "Header row");
            style.AllCells = Apply(style.AllCells, s.SheetCells, "Sheet");
            if (s.SheetCells != null && s.SheetCells.Border != CellBorder.None)
                style.BorderValuesOnly = s.BorderValuesOnly;
            style.DataRows = Apply(style.DataRows, s.GetDataRowStyle(), "Data rows");
            style.DefaultColumn = Apply(style.DefaultColumn, s.DefaultColumn, "Column defaults");
            if (s.DefaultColumnWidth > 0)
                style.DefaultColumnWidth = s.DefaultColumnWidth;
            if (s.RightToLeft)
                style.RightToLeft = true;
            if (s.FreezeHeader)
                style.FreezeRows = 1;
            if (s.FreezeColumns > 0)
                style.FreezeColumns = s.FreezeColumns;
            if (s.AutoFilter)
                style.AutoFilter = true;
            if (s.AutoFit)
                style.AutoFit = true;
            if (s.BandRows)
                style.Band = new CellStyle { Fill = NamedColors.ToHex(s.BandFill) };

            foreach (var f in s.Columns)
            {
                int col = ExcelColumns.Index(input, f.Column);
                if (!ColumnFormat.IsValidWidth(f.Width))
                    throw new InvalidOperationException(ExcelColumns.Label(input, col) + " width must be a number from 0 to 255.");
                style.SetWidth(col, f.Width);
                var look = Apply(null, f.GetStyle(), ExcelColumns.Label(input, col));
                if (look != null)
                    style.StyleColumn(col, look, true);
            }
            return ExcelFile.WithStyle(input, style);
        }

        static CellStyle Apply(CellStyle current, CellStyle changes, string section)
        {
            if (changes == null || changes.IsEmpty)
                return current;
            string error = changes.Check();
            if (error != null)
                throw new InvalidOperationException(section + ": " + error);
            return current == null ? changes.Copy() : current.Over(changes);
        }

        public static string Summary(FormatSheetSettings s)
        {
            var parts = new List<string>();
            if (s.StyleHeader) parts.Add("header");
            if (s.RightToLeft) parts.Add("right-to-left");
            if (s.FreezeHeader || s.FreezeColumns > 0) parts.Add("freeze");
            if (s.AutoFilter) parts.Add("filter");
            if (s.AutoFit) parts.Add("fit widths");
            if (s.BandRows) parts.Add("bands");
            if (s.SheetCells != null && !s.SheetCells.IsEmpty) parts.Add("sheet style");
            if (s.Borders || (s.DataRows != null && !s.DataRows.IsEmpty)) parts.Add("data row style");
            if (s.DefaultColumnWidth > 0 || (s.DefaultColumn != null && !s.DefaultColumn.IsEmpty)) parts.Add("column defaults");
            if (s.Columns.Count > 0) parts.Add(s.Columns.Count + " column format" + (s.Columns.Count == 1 ? "" : "s"));
            return parts.Count == 0 ? "Nothing to format" : string.Join(", ", parts.ToArray());
        }
    }
}
