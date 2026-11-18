using System;
using System.Collections.Generic;
using System.Data;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.EditHeaders
{
    public enum HeaderEditMode { Keep, Text, Formula, Clear }

    public sealed class HeaderEdit
    {
        // Blank source adds a new, empty data column.
        public string Source { get; set; }
        public bool SourceIsId { get; set; }
        public HeaderEditMode Mode { get; set; }
        public string Value { get; set; }

        public HeaderEdit() { Source = ""; Value = ""; }
    }

    public sealed class EditHeadersSettings
    {
        public string Formula { get; set; }
        public List<HeaderEdit> Columns { get; set; }

        public EditHeadersSettings() { Formula = ""; Columns = new List<HeaderEdit>(); }
    }

    public static class EditHeadersWork
    {
        public static DataTable Run(DataTable input, EditHeadersSettings settings)
        {
            if (input == null) throw new InvalidOperationException("Connect an input first.");
            if (settings == null || settings.Columns == null)
                throw new InvalidOperationException("Set up the headers first.");
            var output = input.Copy();
            if (!string.IsNullOrWhiteSpace(settings.Formula))
            {
                string error = ExpressionEngine.Validate(settings.Formula);
                if (error != null) throw new InvalidOperationException("Header formula: " + error);
            }
            // Formulas always see original headers, independent of edit order.
            if (!string.IsNullOrWhiteSpace(settings.Formula))
                for (var i = 0; i < input.Columns.Count; i++)
                    SetHeader(output.Columns[i], Evaluate(input, i, settings.Formula));
            var changed = new HashSet<int>();
            foreach (var edit in settings.Columns)
            {
                if (edit == null) throw new InvalidOperationException("A header edit is missing.");
                int index;
                if (string.IsNullOrWhiteSpace(edit.Source))
                {
                    if (output.Columns.Count >= 16384)
                        throw new InvalidOperationException("Excel supports at most 16384 columns.");
                    index = output.Columns.Count;
                    ExcelFile.AddColumn(output, "").Caption = "";
                }
                else index = ExcelColumns.Index(input, edit.Source);
                if (!changed.Add(index)) throw new InvalidOperationException("Edit each column only once.");
                switch (edit.Mode)
                {
                    case HeaderEditMode.Keep: break;
                    case HeaderEditMode.Text: SetHeader(output.Columns[index], edit.Value ?? ""); break;
                    case HeaderEditMode.Clear: output.Columns[index].Caption = ""; break;
                    case HeaderEditMode.Formula:
                        SetHeader(output.Columns[index], Evaluate(input, index, edit.Value)); break;
                    default: throw new InvalidOperationException("Choose a valid header action.");
                }
            }
            return output;
        }

        static void SetHeader(DataColumn column, string header)
        {
            // Framework Caption ignores changes that only alter letter case.
            column.Caption = "";
            column.Caption = header;
        }

        static string Evaluate(DataTable input, int index, string formula)
        {
            try
            {
                string error = ExpressionEngine.Validate(formula);
                if (error != null) throw new InvalidOperationException(error);
                return ExpressionEngine.ToText(ExpressionEngine.Evaluate(formula, delegate(string name)
                {
                    if (string.Equals(name, "Value", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(name, "Header", StringComparison.OrdinalIgnoreCase))
                        return index < input.Columns.Count ? ExcelFile.Header(input.Columns[index]) : "";
                    if (string.Equals(name, "Column", StringComparison.OrdinalIgnoreCase)) return ExcelColumns.Letter(index);
                    if (string.Equals(name, "Index", StringComparison.OrdinalIgnoreCase)) return index + 1;
                    int column = ExcelColumns.IndexOf(input, name);
                    if (column >= 0) return ExcelFile.Header(input.Columns[column]);
                    object value;
                    if (ExpressionEngine.TryGetVariable(name, out value)) return value;
                    throw new InvalidOperationException("Unknown header variable $" + name + ".");
                }));
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Header " + ExcelColumns.Letter(index) + ": " + ex.Message, ex);
            }
        }
    }
}
