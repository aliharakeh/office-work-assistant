using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using ClosedXML.Excel;

namespace OfficeWorkAssistant.Excel
{
    public sealed class ExcelLoadResult
    {
        public DataTable Table { get; set; }
        public string[] Sheets { get; set; }
        public string Sheet { get; set; }
        public int TotalRows { get; set; }
        public bool Truncated { get; set; }
    }

    // Shared .xlsx read/write. Column headers stay in Caption; ColumnName is a
    // binding-safe id so WPF DataGrid does not choke on '(', '/', and friends.
    public static class ExcelFile
    {
        static ExcelFile()
        {
            var table = new DataTable();
            var a = AddColumn(table, "Amount (USD)");
            var b = AddColumn(table, "Cost/Unit");
            if (a.Caption != "Amount (USD)" || b.Caption != "Cost/Unit")
                throw new InvalidOperationException("ExcelFile caption check failed.");
            if (a.ColumnName.IndexOfAny(new[] { '(', ')', '/' }) >= 0)
                throw new InvalidOperationException("ExcelFile safe-name check failed.");
            if (FindColumn(table, "Amount (USD)") != a || FindColumn(table, "Cost/Unit") != b)
                throw new InvalidOperationException("ExcelFile find-header check failed.");

            var uniq = new DataTable();
            AddColumn(uniq, "Status");
            uniq.Rows.Add("Open");
            uniq.Rows.Add("open");
            uniq.Rows.Add("Closed");
            uniq.Rows.Add("");
            uniq.Rows.Add(DBNull.Value);
            string[] values = UniqueValues(uniq, 0);
            if (values.Length != 2 || values[0] != "Closed" ||
                !string.Equals(values[1], "Open", StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException("ExcelFile unique-values check failed.");
        }

        public static ExcelLoadResult Load(string path, string sheetName)
        {
            return Load(path, sheetName, 0);
        }

        // maxRows <= 0 loads everything; a positive value keeps the UI preview small.
        public static ExcelLoadResult Load(string path, string sheetName, int maxRows)
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

                    return new ExcelLoadResult
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

        // Writes the table to a brand-new workbook with one sheet.
        public static void Save(DataTable table, string path)
        {
            Save(table, path, "Result", false);
        }

        // append = false: create/overwrite the file with one sheet.
        // append = true: add a new sheet to the existing workbook, leaving its other sheets alone.
        public static void Save(DataTable table, string path, string sheetName, bool append)
        {
            string error = CheckSheetName(sheetName);
            if (error != null)
                throw new InvalidOperationException(error);

            if (!append)
            {
                using (var wb = new XLWorkbook())
                {
                    WriteSheet(wb.Worksheets.Add(sheetName.Trim()), table);
                    wb.SaveAs(path);
                }
                return;
            }

            // Copy into memory first: ClosedXML keeps reading from the stream while the workbook is open.
            var copy = new MemoryStream();
            try
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    stream.CopyTo(copy);
                copy.Position = 0;
            }
            catch (Exception ex)
            {
                copy.Dispose();
                throw new IOException("Could not read \"" + Path.GetFileName(path) + "\". " + ex.Message, ex);
            }

            // Save beside the original, then swap, so a failed save never damages the workbook.
            string temp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                using (copy)
                using (var book = new XLWorkbook(copy))
                {
                    foreach (var w in book.Worksheets)
                    {
                        if (string.Equals(w.Name, sheetName.Trim(), StringComparison.OrdinalIgnoreCase))
                            throw new InvalidOperationException("\"" + Path.GetFileName(path) +
                                "\" already has a sheet named \"" + w.Name + "\". Pick another sheet name.");
                    }
                    WriteSheet(book.Worksheets.Add(sheetName.Trim()), table);
                    book.SaveAs(temp);
                }
                File.Replace(temp, path, null);
            }
            catch (IOException ex)
            {
                throw new IOException("Could not update \"" + Path.GetFileName(path) +
                    "\". Close it in Excel and try again. " + ex.Message, ex);
            }
            finally
            {
                if (File.Exists(temp))
                {
                    try { File.Delete(temp); }
                    catch { }
                }
            }
        }

        // Null when the name works as an Excel sheet name; otherwise the reason it does not.
        public static string CheckSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Enter a sheet name.";
            name = name.Trim();
            if (name.Length > 31)
                return "A sheet name can have at most 31 characters.";
            if (name.IndexOfAny(new[] { '[', ']', ':', '*', '?', '/', '\\' }) >= 0)
                return "A sheet name cannot contain [ ] : * ? / \\";
            if (name[0] == '\'' || name[name.Length - 1] == '\'')
                return "A sheet name cannot start or end with an apostrophe.";
            return null;
        }

        static void WriteSheet(IXLWorksheet ws, DataTable table)
        {
            for (var c = 0; c < table.Columns.Count; c++)
                ws.Cell(1, c + 1).Value = Header(table.Columns[c]);

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

            var style = GetStyle(table);
            if (style != null)
                WriteStyle(ws, table, style);
        }

        // ---------- styles from the Highlight and Format sheet steps ----------

        const string StyleKey = "OfficeWorkAssistant.Style";

        // How Save should style this table; null when it is plain.
        public static SheetStyle GetStyle(DataTable table)
        {
            if (table == null || !table.ExtendedProperties.ContainsKey(StyleKey))
                return null;
            return table.ExtendedProperties[StyleKey] as SheetStyle;
        }

        // A copy of the table that carries the style. Tables are shared between steps, so the
        // input itself is never changed.
        public static DataTable WithStyle(DataTable table, SheetStyle style)
        {
            var copy = table.Copy();
            copy.ExtendedProperties[StyleKey] = style;
            return copy;
        }

        // For a step's own new result: rows may have moved, so an old style no longer fits.
        public static void ClearStyle(DataTable table)
        {
            if (table != null && table.ExtendedProperties.ContainsKey(StyleKey))
                table.ExtendedProperties.Remove(StyleKey);
        }

        // Bottom layer first, so each later one only changes what it sets.
        static void WriteStyle(IXLWorksheet ws, DataTable table, SheetStyle style)
        {
            int rows = table.Rows.Count;
            int cols = table.Columns.Count;
            if (cols == 0)
                return;

            if (rows > 0)
            {
                if (style.AllCells != null)
                    Apply(ws.Range(2, 1, rows + 1, cols), style.AllCells);
                foreach (var pair in style.Columns)
                {
                    if (pair.Key < cols)
                        Apply(ws.Range(2, pair.Key + 1, rows + 1, pair.Key + 1), pair.Value);
                }
                if (style.Band != null)
                {
                    for (var r = 1; r < rows; r += 2)
                        Apply(ws.Range(r + 2, 1, r + 2, cols), style.Band);
                }
                foreach (var pair in style.Rows)
                {
                    if (pair.Key < rows)
                        Apply(ws.Range(pair.Key + 2, 1, pair.Key + 2, cols), pair.Value);
                }
                foreach (var cell in style.CellKeys)
                {
                    if (cell.Key < rows && cell.Value < cols)
                        Apply(ws.Range(cell.Key + 2, cell.Value + 1, cell.Key + 2, cell.Value + 1), style.CellOnly(cell.Key, cell.Value));
                }
            }
            if (style.Header != null)
                Apply(ws.Range(1, 1, 1, cols), style.Header);

            if (style.AutoFit)
                ws.Columns(1, cols).AdjustToContents();
            foreach (var pair in style.Widths)
            {
                if (pair.Key < cols)
                    ws.Column(pair.Key + 1).Width = pair.Value;
            }
            if (style.FreezeRows > 0 || style.FreezeColumns > 0)
                ws.SheetView.Freeze(style.FreezeRows, style.FreezeColumns);
            if (style.AutoFilter)
                ws.Range(1, 1, Math.Max(rows, 1) + 1, cols).SetAutoFilter();
        }

        static void Apply(IXLRange range, CellStyle s)
        {
            if (s == null)
                return;
            var st = range.Style;
            string hex = NamedColors.ToHex(s.Fill);
            if (!string.IsNullOrEmpty(hex))
                st.Fill.BackgroundColor = XLColor.FromHtml(hex);
            hex = NamedColors.ToHex(s.FontColor);
            if (!string.IsNullOrEmpty(hex))
                st.Font.FontColor = XLColor.FromHtml(hex);
            if (s.Bold)
                st.Font.Bold = true;
            if (s.Italic)
                st.Font.Italic = true;
            if (s.Underline)
                st.Font.Underline = XLFontUnderlineValues.Single;
            if (s.Strike)
                st.Font.Strikethrough = true;
            if (s.Border != CellBorder.None)
            {
                var line = s.Border == CellBorder.Thick ? XLBorderStyleValues.Medium : XLBorderStyleValues.Thin;
                st.Border.OutsideBorder = line;
                st.Border.InsideBorder = line;
            }
            if (!string.IsNullOrWhiteSpace(s.NumberFormat))
                st.NumberFormat.Format = s.NumberFormat.Trim();
            if (s.Align == CellAlign.Left)
                st.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
            else if (s.Align == CellAlign.Center)
                st.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            else if (s.Align == CellAlign.Right)
                st.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            if (s.Wrap)
                st.Alignment.WrapText = true;
        }

        public static DataColumn AddColumn(DataTable table, string header)
        {
            string caption = string.IsNullOrWhiteSpace(header)
                ? "Column" + (table.Columns.Count + 1)
                : header.Trim();
            var col = table.Columns.Add(UniqueName(table, SafeId(caption)), typeof(object));
            col.Caption = caption;
            return col;
        }

        public static string Header(DataColumn col)
        {
            if (col == null)
                return "";
            if (!string.IsNullOrEmpty(col.Caption))
                return col.Caption;
            return col.ColumnName ?? "";
        }

        // Distinct non-blank display texts, case-insensitive, sorted. Dates match Save.
        public static string[] UniqueValues(DataTable table, int columnIndex)
        {
            if (table == null)
                throw new InvalidOperationException("Choose a data file first.");
            if (columnIndex < 0 || columnIndex >= table.Columns.Count)
                throw new InvalidOperationException("Column is not in the file.");

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var list = new List<string>();
            for (var r = 0; r < table.Rows.Count; r++)
            {
                string text = CellText(table.Rows[r][columnIndex]);
                if (text.Length == 0)
                    continue;
                if (seen.Add(text))
                    list.Add(text);
            }
            list.Sort(StringComparer.OrdinalIgnoreCase);
            return list.ToArray();
        }

        public static DataColumn FindColumn(DataTable table, string name)
        {
            if (table == null || string.IsNullOrEmpty(name))
                return null;
            if (table.Columns.Contains(name))
                return table.Columns[name];
            for (var i = 0; i < table.Columns.Count; i++)
            {
                var col = table.Columns[i];
                if (string.Equals(col.ColumnName, name, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(Header(col), name, StringComparison.OrdinalIgnoreCase))
                    return col;
            }
            return null;
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
                AddColumn(table, name);
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

        static string CellText(object value)
        {
            if (value == null || value == DBNull.Value)
                return "";
            if (value is DateTime)
            {
                DateTime dt = (DateTime)value;
                return dt.TimeOfDay == TimeSpan.Zero
                    ? dt.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)
                    : dt.ToString("dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
            }
            string text = Convert.ToString(value, CultureInfo.InvariantCulture);
            return text != null ? text.Trim() : "";
        }

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

        static string SafeId(string header)
        {
            var sb = new StringBuilder(header.Length);
            for (var i = 0; i < header.Length; i++)
            {
                char c = header[i];
                if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9') || c == '_')
                    sb.Append(c);
                else
                    sb.Append('_');
            }
            if (sb.Length == 0)
                return "Column";
            if (sb[0] >= '0' && sb[0] <= '9')
                sb.Insert(0, 'C');
            return sb.ToString();
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
