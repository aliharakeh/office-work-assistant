using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Xml.Serialization;
using ClosedXML.Excel;
using WorkAssistant.Expressions;

namespace WorkAssistant.Features.Templates
{
    public enum TemplateColumnKind
    {
        Fixed,
        Copy,
        Concat,
        Math,
        Conditional
    }

    public sealed class TemplateColumn
    {
        public string Name { get; set; }
        public TemplateColumnKind Kind { get; set; }
        public string FixedValue { get; set; }
        public string SourceColumn { get; set; }
        public string Pattern { get; set; }
        public string Left { get; set; }
        public string Operator { get; set; }
        public string Right { get; set; }
        public string Expression { get; set; }
        public string Condition { get; set; }
        public string TrueExpression { get; set; }
        public string FalseExpression { get; set; }

        public TemplateColumn()
        {
            Name = "";
            Kind = TemplateColumnKind.Fixed;
            FixedValue = "";
            SourceColumn = "";
            Pattern = "";
            Left = "";
            Operator = "+";
            Right = "";
            Expression = "";
            Condition = "";
            TrueExpression = "";
            FalseExpression = "";
        }

        public override string ToString()
        {
            string name = Name != null ? Name : "";
            if (Kind == TemplateColumnKind.Fixed)
                return name + " [Fixed]: " + (FixedValue != null ? FixedValue : "");
            if (Kind == TemplateColumnKind.Copy)
                return name + " [Copy]: " + (SourceColumn != null ? SourceColumn : "");
            if (Kind == TemplateColumnKind.Concat)
                return name + " [Concat]: " + (Pattern != null ? Pattern : "");
            if (Kind == TemplateColumnKind.Conditional)
                return name + " [If]: IF(" + (Condition != null ? Condition : "") + ", " +
                    (TrueExpression != null ? TrueExpression : "") + ", " +
                    (FalseExpression != null ? FalseExpression : "") + ")";
            string expr = Expression != null ? Expression : "";
            if (expr.Length == 0)
                expr = (Left != null ? Left : "") + " " +
                    (Operator != null ? Operator : "") + " " + (Right != null ? Right : "");
            return name + " [Math]: " + expr;
        }
    }

    public sealed class TemplateDefinition
    {
        public string Name { get; set; }
        public bool KeepSourceColumns { get; set; }
        public List<string> SourceColumns { get; set; }
        public List<TemplateColumn> Columns { get; set; }

        public TemplateDefinition()
        {
            Name = "";
            KeepSourceColumns = true;
            SourceColumns = new List<string>();
            Columns = new List<TemplateColumn>();
        }
    }

    public sealed class TemplateSourceResult
    {
        public DataTable Table { get; set; }
        public string[] Sheets { get; set; }
        public string Sheet { get; set; }
    }

    public static class TemplatesWork
    {
        public static bool TryGetVariable(string name, out object value)
        {
            return ExpressionEngine.TryGetVariable(name, out value);
        }

        static string ExpandFixed(string fixedValue)
        {
            return ExpressionEngine.ExpandVariables(fixedValue);
        }

        static string FormatScalar(object value)
        {
            return ExpressionEngine.FormatScalar(value);
        }

        static bool TryToDate(object value, out DateTime result)
        {
            return ExpressionEngine.TryToDate(value, out result);
        }

        static bool IsDateLike(object value)
        {
            return ExpressionEngine.IsDateLike(value);
        }

        public static TemplateSourceResult LoadSource(string path, string sheetName)
        {
            using (var wb = new XLWorkbook(path))
            {
                var sheets = new string[wb.Worksheets.Count];
                var i = 0;
                foreach (var w in wb.Worksheets)
                    sheets[i++] = w.Name;

                var ws = string.IsNullOrEmpty(sheetName)
                    ? wb.Worksheet(1)
                    : wb.Worksheet(sheetName);

                return new TemplateSourceResult
                {
                    Table = ToTable(ws),
                    Sheets = sheets,
                    Sheet = ws.Name
                };
            }
        }

        public static void SaveTable(DataTable table, string path)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Result");
                for (var c = 0; c < table.Columns.Count; c++)
                    ws.Cell(1, c + 1).Value = table.Columns[c].ColumnName;

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

                wb.SaveAs(path);
            }
        }

        public static void SaveTemplate(TemplateDefinition template, string path)
        {
            string error = Validate(template, false);
            if (error != null)
                throw new InvalidOperationException(error);
            var ser = new XmlSerializer(typeof(TemplateDefinition));
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
                ser.Serialize(fs, template);
        }

        public static TemplateDefinition LoadTemplate(string path)
        {
            var ser = new XmlSerializer(typeof(TemplateDefinition));
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                var template = ser.Deserialize(fs) as TemplateDefinition;
                if (template == null)
                    throw new InvalidOperationException("Not a template file.");
                if (template.Columns == null)
                    template.Columns = new List<TemplateColumn>();
                if (template.SourceColumns == null)
                    template.SourceColumns = new List<string>();
                if (template.Name == null)
                    template.Name = "";
                string error = Validate(template, false);
                if (error != null)
                    throw new InvalidOperationException(error);
                return template;
            }
        }

        public static string Validate(TemplateDefinition template, bool requireColumns)
        {
            if (template == null)
                return "Template is empty.";
            if (template.Columns == null || template.Columns.Count == 0)
            {
                if (requireColumns)
                    return "Add at least one extra column.";
                return null;
            }
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var i = 0; i < template.Columns.Count; i++)
            {
                var col = template.Columns[i];
                if (col == null)
                    return "Column #" + (i + 1) + " is empty.";
                string err = ValidateColumn(col);
                if (err != null)
                    return err;
                string name = col.Name != null ? col.Name.Trim() : "";
                if (seen.Contains(name))
                    return "Duplicate column name: " + name + ".";
                seen.Add(name);
            }
            return null;
        }

        public static string ValidateColumn(TemplateColumn col)
        {
            string name = col.Name != null ? col.Name.Trim() : "";
            if (name.Length == 0)
                return "Column name is required.";
            if (col.Kind == TemplateColumnKind.Fixed)
                return null;
            if (col.Kind == TemplateColumnKind.Copy)
            {
                if (string.IsNullOrWhiteSpace(col.SourceColumn))
                    return "Column '" + name + "': choose a source column.";
                return null;
            }
            if (col.Kind == TemplateColumnKind.Concat)
            {
                if (string.IsNullOrWhiteSpace(col.Pattern))
                    return "Column '" + name + "': enter a pattern like {A} {B}.";
                return null;
            }
            if (col.Kind == TemplateColumnKind.Conditional)
            {
                if (string.IsNullOrWhiteSpace(col.Condition))
                    return "Column '" + name + "': enter a condition like {A} > 10.";
                string condErr = ValidateExpression(col.Condition);
                if (condErr != null)
                    return "Column '" + name + "' condition: " + condErr;
                string trueErr = ValidateBranch(col.TrueExpression);
                if (trueErr != null)
                    return "Column '" + name + "' then-value: " + trueErr;
                string falseErr = ValidateBranch(col.FalseExpression);
                if (falseErr != null)
                    return "Column '" + name + "' else-value: " + falseErr;
                return null;
            }
            string expr = col.Expression;
            if (string.IsNullOrWhiteSpace(expr))
            {
                // Back-compat: old templates stored Left/Op/Right only.
                if (string.IsNullOrWhiteSpace(col.Left) || string.IsNullOrWhiteSpace(col.Right))
                    return "Column '" + name + "': enter a formula like {A} * {B}.";
                expr = col.Left + " " + col.Operator + " " + col.Right;
            }
            string err = ValidateExpression(expr);
            if (err != null)
                return "Column '" + name + "': " + err;
            return null;
        }

        public static DataTable ApplyTemplate(DataTable source, TemplateDefinition template)
        {
            if (source == null)
                throw new InvalidOperationException("Choose a data file first.");
            if (template == null)
                throw new InvalidOperationException("Template is empty.");
            if (template.Columns == null)
                template.Columns = new List<TemplateColumn>();
            if (template.SourceColumns == null)
                template.SourceColumns = new List<string>();

            string error = Validate(template, false);
            if (error != null)
                throw new InvalidOperationException(error);

            List<string> kept = ResolveSourceColumns(source, template);

            var output = new DataTable();
            foreach (string name in kept)
                output.Columns.Add(UniqueName(output, name), typeof(object));

            var added = new List<string>();
            for (var i = 0; i < template.Columns.Count; i++)
                added.Add(UniqueName(output, template.Columns[i].Name.Trim()));

            for (var i = 0; i < added.Count; i++)
                output.Columns.Add(added[i], typeof(object));

            foreach (DataRow row in source.Rows)
            {
                var dest = output.NewRow();
                var d = 0;
                foreach (string name in kept)
                {
                    dest[d] = row[name];
                    d++;
                }
                for (var i = 0; i < template.Columns.Count; i++)
                    dest[d + i] = Evaluate(template.Columns[i], source, row);
                output.Rows.Add(dest);
            }

            return output;
        }

        public static List<string> ResolveSourceColumns(DataTable source, TemplateDefinition template)
        {
            var kept = new List<string>();
            if (source == null)
                return kept;
            if (template == null || !template.KeepSourceColumns)
                return kept;
            if (template.SourceColumns == null || template.SourceColumns.Count == 0)
            {
                // Back-compat: old templates kept every source column.
                foreach (DataColumn col in source.Columns)
                    kept.Add(col.ColumnName);
                return kept;
            }
            // New templates store Excel letters ("A", "B", ...). Old ones stored
            // real header names, so accept both. Letters map by position.
            // A pure letter token ("A", "b", "AA") is a position reference
            // only, so a header that happens to be named "A" elsewhere is
            // not also matched by name. Anything else matches by header
            // name (with a letter fallback for old "A - Qty" style labels).
            var wantedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var wantedLetters = new HashSet<int>();
            for (var i = 0; i < template.SourceColumns.Count; i++)
            {
                string entry = template.SourceColumns[i];
                if (string.IsNullOrWhiteSpace(entry))
                    continue;
                string trimmed = entry.Trim();
                int letterIndex;
                bool parses = TryParseColumnLetter(ExtractLetter(trimmed), out letterIndex);
                bool isPureLetter = parses &&
                    string.Equals(trimmed, ExtractLetter(trimmed), StringComparison.OrdinalIgnoreCase);
                if (isPureLetter)
                    wantedLetters.Add(letterIndex);
                else
                {
                    wantedNames.Add(trimmed);
                    if (parses)
                        wantedLetters.Add(letterIndex);
                }
            }
            for (var c = 0; c < source.Columns.Count; c++)
            {
                DataColumn col = source.Columns[c];
                if (wantedLetters.Contains(c) || wantedNames.Contains(col.ColumnName))
                    kept.Add(col.ColumnName);
            }
            return kept;
        }

        public static string[] ColumnNames(DataTable table)
        {
            if (table == null)
                return new string[0];
            var names = new string[table.Columns.Count];
            for (var i = 0; i < table.Columns.Count; i++)
                names[i] = table.Columns[i].ColumnName;
            return names;
        }

        // Excel-style column letters: 0->A, 1->B, ... 25->Z, 26->AA, etc.
        public static string ColumnLetter(int index)
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

        public static bool TryParseColumnLetter(string s, out int index)
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

        // "A - Qty" style labels so users see the letter plus the real header.
        public static string SourceLabel(DataTable table, int index)
        {
            return ColumnLetter(index) + " - " + table.Columns[index].ColumnName;
        }

        public static string[] SourceLabels(DataTable table)
        {
            if (table == null)
                return new string[0];
            var labels = new string[table.Columns.Count];
            for (var i = 0; i < table.Columns.Count; i++)
                labels[i] = SourceLabel(table, i);
            return labels;
        }

        public static string[] ColumnLetters(DataTable table)
        {
            if (table == null)
                return new string[0];
            var letters = new string[table.Columns.Count];
            for (var i = 0; i < table.Columns.Count; i++)
                letters[i] = ColumnLetter(i);
            return letters;
        }

        // Accepts "A", "a", or a UI label like "A - Qty" and returns "A".
        public static string ExtractLetter(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";
            string t = text.Trim();
            // Take the first token before space, '-', or ':' ("A - Qty" -> "A").
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
            if (TryParseColumnLetter(head, out dummy))
                return head.ToUpperInvariant();
            return t;
        }

        static object Evaluate(TemplateColumn col, DataTable source, DataRow row)
        {
            if (col.Kind == TemplateColumnKind.Fixed)
                return ExpandFixed(col.FixedValue);

            if (col.Kind == TemplateColumnKind.Copy)
                return CopyValue(source, row, col.SourceColumn);

            if (col.Kind == TemplateColumnKind.Concat)
                return ConcatValue(source, row, col.Pattern);

            if (col.Kind == TemplateColumnKind.Conditional)
                return ConditionalValue(source, row, col.Condition, col.TrueExpression, col.FalseExpression);

            string expr = col.Expression;
            if (string.IsNullOrWhiteSpace(expr))
            {
                // Back-compat: old templates stored Left/Op/Right only.
                if (string.IsNullOrWhiteSpace(col.Left) || string.IsNullOrWhiteSpace(col.Right))
                    return DBNull.Value;
                expr = col.Left + " " + col.Operator + " " + col.Right;
            }
            return EvaluateExpression(source, row, expr);
        }

        static object CopyValue(DataTable source, DataRow row, string column)
        {
            string name = column != null ? column.Trim() : "";
            if (name.Length == 0 || source == null)
                return DBNull.Value;
            // Letters win: "A" means first column, "B" second, etc.
            int letterIndex;
            if (TryParseColumnLetter(ExtractLetter(name), out letterIndex) &&
                letterIndex >= 0 && letterIndex < source.Columns.Count)
                return row[letterIndex];
            if (source.Columns.Contains(name))
                return row[name];
            // Case-insensitive fallback so old templates survive casing changes.
            foreach (DataColumn c in source.Columns)
            {
                if (string.Equals(c.ColumnName, name, StringComparison.OrdinalIgnoreCase))
                    return row[c.ColumnName];
            }
            return DBNull.Value;
        }

        static object ConcatValue(DataTable source, DataRow row, string pattern)
        {
            return ExpressionEngine.ExpandPlaceholders(pattern, delegate(string key) { return FindValue(source, row, key); });
        }

        static object ConditionalValue(DataTable source, DataRow row, string condition, string ifTrue, string ifFalse)
        {
            bool flag;
            try
            {
                flag = ExpressionEngine.ToBool(
                    ExpressionEngine.Evaluate(condition, delegate(string name) { return FindValue(source, row, name); }));
            }
            catch
            {
                return DBNull.Value;
            }
            string branch = flag ? ifTrue : ifFalse;
            if (string.IsNullOrWhiteSpace(branch))
                return "";
            return EvaluateExpression(source, row, branch);
        }

        static string ValidateBranch(string branch)
        {
            if (string.IsNullOrWhiteSpace(branch))
                return null;
            return ValidateExpression(branch);
        }

        public static string ValidateExpression(string expression)
        {
            return ExpressionEngine.Validate(expression);
        }

        static object EvaluateExpression(DataTable source, DataRow row, string expression)
        {
            try
            {
                return ExpressionEngine.Evaluate(expression, delegate(string name) { return FindValue(source, row, name); });
            }
            catch
            {
                return DBNull.Value;
            }
        }

        static object FindValue(DataTable source, DataRow row, string name)
        {
            if (string.IsNullOrEmpty(name))
                return null;
            if (source != null)
            {
                // Excel letters win: {A} = first column, {B} = second, etc.
                int letterIndex;
                if (TryParseColumnLetter(ExtractLetter(name), out letterIndex) &&
                    letterIndex >= 0 && letterIndex < source.Columns.Count)
                    return row[letterIndex];
                // Back-compat: old templates used real header names.
                if (source.Columns.Contains(name))
                    return row[name];
                foreach (DataColumn c in source.Columns)
                {
                    if (string.Equals(c.ColumnName, name, StringComparison.OrdinalIgnoreCase))
                        return row[c.ColumnName];
                }
            }
            object v;
            if (TryGetVariable(name, out v))
                return v;
            return null;
        }

        static DataTable ToTable(IXLWorksheet ws)
        {
            var table = new DataTable();
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
                table.Columns.Add(UniqueName(table, name), typeof(object));
            }

            foreach (var row in range.RowsUsed())
            {
                if (row.RowNumber() == header.RowNumber())
                    continue;
                var dr = table.NewRow();
                for (var c = 1; c <= colCount; c++)
                {
                    var cell = row.Cell(c);
                    dr[c - 1] = cell.IsEmpty() ? (object)DBNull.Value : cell.Value;
                }
                table.Rows.Add(dr);
            }

            return table;
        }

        static string UniqueName(DataTable table, string name)
        {
            string clean = string.IsNullOrWhiteSpace(name) ? "Column" : name.Trim();
            if (!table.Columns.Contains(clean))
                return clean;
            var n = 2;
            while (table.Columns.Contains(clean + "_" + n))
                n++;
            return clean + "_" + n;
        }
    }
}
