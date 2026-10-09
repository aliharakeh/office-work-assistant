using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeWorkAssistant.Excel
{
    public enum CellBorder
    {
        None,
        Thin,
        Thick
    }

    public enum CellAlign
    {
        General,
        Left,
        Center,
        Right
    }

    // How one cell looks. Blank / false / General parts are "not set", so a style laid over
    // another only changes what it sets. Stored in step settings (XML), so plain properties.
    public sealed class CellStyle
    {
        // "#RRGGBB" or blank.
        public string Fill { get; set; }
        public string FontColor { get; set; }
        public bool Bold { get; set; }
        public bool Italic { get; set; }
        public bool Underline { get; set; }
        public bool Strike { get; set; }
        public CellBorder Border { get; set; }
        // An Excel number format such as "#,##0.00" or "dd/MM/yyyy".
        public string NumberFormat { get; set; }
        public CellAlign Align { get; set; }
        public bool Wrap { get; set; }

        public CellStyle()
        {
            Fill = "";
            FontColor = "";
            NumberFormat = "";
        }

        public bool IsEmpty
        {
            get
            {
                return string.IsNullOrWhiteSpace(Fill) && string.IsNullOrWhiteSpace(FontColor) && !Bold && !Italic &&
                    !Underline && !Strike && Border == CellBorder.None && string.IsNullOrWhiteSpace(NumberFormat) &&
                    Align == CellAlign.General && !Wrap;
            }
        }

        public CellStyle Copy()
        {
            return (CellStyle)MemberwiseClone();
        }

        // This style with top laid over it: what top sets wins. Neither is changed.
        public CellStyle Over(CellStyle top)
        {
            if (top == null)
                return this;
            var s = Copy();
            if (!string.IsNullOrWhiteSpace(top.Fill))
                s.Fill = top.Fill;
            if (!string.IsNullOrWhiteSpace(top.FontColor))
                s.FontColor = top.FontColor;
            s.Bold |= top.Bold;
            s.Italic |= top.Italic;
            s.Underline |= top.Underline;
            s.Strike |= top.Strike;
            if (top.Border != CellBorder.None)
                s.Border = top.Border;
            if (!string.IsNullOrWhiteSpace(top.NumberFormat))
                s.NumberFormat = top.NumberFormat;
            if (top.Align != CellAlign.General)
                s.Align = top.Align;
            s.Wrap |= top.Wrap;
            return s;
        }

        // Null when every colour is blank or #RRGGBB; otherwise the reason.
        public string Check()
        {
            if (!NamedColors.IsValid(Fill))
                return "The fill colour '" + Fill + "' is not a colour name or #RRGGBB.";
            if (!NamedColors.IsValid(FontColor))
                return "The text colour '" + FontColor + "' is not a colour name or #RRGGBB.";
            return null;
        }

        // "Light red fill, bold" for lists and step summaries.
        public string Describe()
        {
            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(Fill))
                parts.Add(NamedColors.NameOf(Fill) + " fill");
            if (!string.IsNullOrWhiteSpace(FontColor))
                parts.Add(NamedColors.NameOf(FontColor) + " text");
            if (Bold) parts.Add("bold");
            if (Italic) parts.Add("italic");
            if (Underline) parts.Add("underline");
            if (Strike) parts.Add("strikethrough");
            if (Border != CellBorder.None) parts.Add(Border == CellBorder.Thin ? "thin border" : "thick border");
            if (!string.IsNullOrWhiteSpace(NumberFormat)) parts.Add("format " + NumberFormat);
            if (Align != CellAlign.General) parts.Add(Align.ToString().ToLowerInvariant());
            if (Wrap) parts.Add("wrap");
            return parts.Count == 0 ? "no style" : string.Join(", ", parts.ToArray());
        }
    }

    // Colour names the style pickers offer, and #RRGGBB parsing.
    public static class NamedColors
    {
        public static readonly string[] FillNames =
        {
            "Light red", "Light green", "Light yellow", "Light blue", "Light orange", "Light purple", "Light grey",
            "Grey", "Red", "Green", "Yellow", "Blue", "Orange"
        };

        public static readonly string[] TextNames =
        {
            "Dark red", "Dark green", "Dark yellow", "Red", "Green", "Blue", "Orange", "Grey", "Black", "White"
        };

        static readonly Dictionary<string, string> Hex = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "Light red", "#FFC7CE" },
            { "Light green", "#C6EFCE" },
            { "Light yellow", "#FFEB9C" },
            { "Light blue", "#DDEBF7" },
            { "Light orange", "#FCE4D6" },
            { "Light purple", "#E4DFEC" },
            { "Light grey", "#F2F2F2" },
            { "Grey", "#D9D9D9" },
            { "Red", "#FF0000" },
            { "Green", "#00B050" },
            { "Yellow", "#FFFF00" },
            { "Blue", "#0070C0" },
            { "Orange", "#FFC000" },
            { "Dark red", "#9C0006" },
            { "Dark green", "#006100" },
            { "Dark yellow", "#9C5700" },
            { "Black", "#000000" },
            { "White", "#FFFFFF" }
        };

        // A colour name or #RRGGBB as "#RRGGBB"; blank stays blank; null when it is neither.
        public static string ToHex(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";
            string t = text.Trim();
            string hex;
            if (Hex.TryGetValue(t, out hex))
                return hex;
            if (t[0] != '#')
                t = "#" + t;
            if (t.Length != 7)
                return null;
            int dummy;
            if (!int.TryParse(t.Substring(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out dummy))
                return null;
            return t.ToUpperInvariant();
        }

        public static bool IsValid(string text)
        {
            return ToHex(text) != null;
        }

        // The colour's name when it has one, otherwise its #RRGGBB.
        public static string NameOf(string text)
        {
            string hex = ToHex(text);
            if (string.IsNullOrEmpty(hex))
                return text ?? "";
            foreach (var pair in Hex)
            {
                if (pair.Value == hex)
                    return pair.Key;
            }
            return hex;
        }

        public static bool TryRgb(string text, out byte r, out byte g, out byte b)
        {
            r = g = b = 0;
            string hex = ToHex(text);
            if (string.IsNullOrEmpty(hex))
                return false;
            int v = int.Parse(hex.Substring(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            r = (byte)((v >> 16) & 0xFF);
            g = (byte)((v >> 8) & 0xFF);
            b = (byte)(v & 0xFF);
            return true;
        }
    }

    // How a whole sheet looks when Save writes it. Travels on the DataTable (ExcelFile.GetStyle)
    // from the Highlight and Format sheet steps to Save. Rows and columns are 0-based data
    // indexes (row 0 = the first row under the header). Layers, bottom to top: every cell,
    // columns, banded rows, rows, single cells; the header has its own style.
    // Never change a SheetStyle that is on a table: Copy it first.
    public sealed class SheetStyle
    {
        Dictionary<int, CellStyle> _columns = new Dictionary<int, CellStyle>();
        Dictionary<int, CellStyle> _rows = new Dictionary<int, CellStyle>();
        Dictionary<long, CellStyle> _cells = new Dictionary<long, CellStyle>();
        Dictionary<int, double> _widths = new Dictionary<int, double>();

        public CellStyle Header { get; set; }
        public CellStyle AllCells { get; set; }
        // Every second data row (the 2nd, 4th, ...) gets this.
        public CellStyle Band { get; set; }
        public bool AutoFit { get; set; }
        public int FreezeRows { get; set; }
        public int FreezeColumns { get; set; }
        public bool AutoFilter { get; set; }

        public SheetStyle Copy()
        {
            var s = (SheetStyle)MemberwiseClone();
            s._columns = new Dictionary<int, CellStyle>(_columns);
            s._rows = new Dictionary<int, CellStyle>(_rows);
            s._cells = new Dictionary<long, CellStyle>(_cells);
            s._widths = new Dictionary<int, double>(_widths);
            return s;
        }

        static long Key(int row, int col)
        {
            return ((long)row << 16) | (uint)col;
        }

        static void Lay(Dictionary<int, CellStyle> layer, int key, CellStyle style, bool under)
        {
            CellStyle old;
            if (!layer.TryGetValue(key, out old))
                layer[key] = style;
            else
                layer[key] = under ? style.Over(old) : old.Over(style);
        }

        // under: what is already there wins, so a later step's format stays below earlier highlights.
        public void StyleColumn(int col, CellStyle style, bool under)
        {
            Lay(_columns, col, style, under);
        }

        public void StyleRow(int row, CellStyle style)
        {
            Lay(_rows, row, style, false);
        }

        public void StyleCell(int row, int col, CellStyle style)
        {
            long key = Key(row, col);
            CellStyle old;
            _cells[key] = _cells.TryGetValue(key, out old) ? old.Over(style) : style;
        }

        public void SetWidth(int col, double width)
        {
            if (width > 0)
                _widths[col] = width;
        }

        public IEnumerable<KeyValuePair<int, CellStyle>> Columns { get { return _columns; } }
        public IEnumerable<KeyValuePair<int, CellStyle>> Rows { get { return _rows; } }
        public IEnumerable<KeyValuePair<int, double>> Widths { get { return _widths; } }

        public IEnumerable<KeyValuePair<int, int>> CellKeys
        {
            get
            {
                foreach (var key in _cells.Keys)
                    yield return new KeyValuePair<int, int>((int)(key >> 16), (int)(key & 0xFFFF));
            }
        }

        public CellStyle CellOnly(int row, int col)
        {
            CellStyle s;
            return _cells.TryGetValue(Key(row, col), out s) ? s : null;
        }

        public int StyledCells { get { return _cells.Count; } }
        public int StyledRows { get { return _rows.Count; } }

        // What one data cell ends up looking like; null when nothing styles it.
        public CellStyle Effective(int row, int col)
        {
            CellStyle s = AllCells;
            CellStyle part;
            if (_columns.TryGetValue(col, out part))
                s = s == null ? part : s.Over(part);
            if (Band != null && row % 2 == 1)
                s = s == null ? Band : s.Over(Band);
            if (_rows.TryGetValue(row, out part))
                s = s == null ? part : s.Over(part);
            if (_cells.TryGetValue(Key(row, col), out part))
                s = s == null ? part : s.Over(part);
            return s;
        }
    }
}
