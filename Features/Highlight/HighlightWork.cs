using System;
using System.Collections.Generic;
using System.Data;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.Highlight
{
    public enum HighlightTarget
    {
        // The whole row, when the condition is true for it.
        Row,
        // Each cell of Columns on its own; $Value is that cell.
        Cells,
        // The cells of Columns in a row the condition is true for.
        Columns
    }

    public sealed class HighlightRule
    {
        // A formula, true = style it: $Status == "Late", $Value < 0, $Due < $Today.
        public string Condition { get; set; }
        public HighlightTarget Target { get; set; }
        // Excel letters, for Cells and Columns.
        public List<string> Columns { get; set; }
        public CellStyle Style { get; set; }
        // When it matches, later rules leave that row (or cell) alone.
        public bool StopIfTrue { get; set; }

        public HighlightRule()
        {
            Condition = "";
            Columns = new List<string>();
            Style = new CellStyle();
        }

        public override string ToString()
        {
            string where;
            if (Target == HighlightTarget.Row)
                where = "Row";
            else
                where = (Target == HighlightTarget.Cells ? "Cells " : "Columns ") + string.Join(", ", (Columns ?? new List<string>()).ToArray());
            return where + " when " + Condition + " -> " + (Style ?? new CellStyle()).Describe() + (StopIfTrue ? " (stop)" : "");
        }
    }

    public sealed class HighlightSettings
    {
        public List<HighlightRule> Rules { get; set; }

        public HighlightSettings()
        {
            Rules = new List<HighlightRule>();
        }
    }

    // Pure logic: marks cells and rows for Save to colour. The data is not changed. No WPF.
    public static class HighlightWork
    {
        public static string Validate(HighlightRule rule, int number)
        {
            string at = "Rule " + number + ": ";
            string err = ExpressionEngine.Validate(rule.Condition);
            if (err != null)
                return at + err;
            if (rule.Target != HighlightTarget.Row && (rule.Columns == null || rule.Columns.Count == 0))
                return at + "pick the columns to style.";
            if (rule.Style == null || rule.Style.IsEmpty)
                return at + "choose a style, such as a fill colour.";
            err = rule.Style.Check();
            return err == null ? null : at + err;
        }

        public static DataTable Run(DataTable input, HighlightSettings s)
        {
            if (input == null)
                throw new InvalidOperationException("Connect an input first.");
            if (s == null || s.Rules.Count == 0)
                throw new InvalidOperationException("Add at least one rule.");
            for (var i = 0; i < s.Rules.Count; i++)
            {
                string err = Validate(s.Rules[i], i + 1);
                if (err != null)
                    throw new InvalidOperationException(err);
            }

            // Earlier steps' styles stay; these go on top.
            var old = ExcelFile.GetStyle(input);
            var style = old != null ? old.Copy() : new SheetStyle();
            var lookup = new RowLookup(input);
            Func<string, object> find = lookup.Find;
            var stoppedRows = new HashSet<int>();
            var stoppedCells = new HashSet<long>();

            foreach (var rule in s.Rules)
            {
                var test = ExpressionEngine.Compile(rule.Condition);
                var cols = rule.Target == HighlightTarget.Row ? new List<int>() : ExcelColumns.Indexes(input, rule.Columns);
                var cellStyle = rule.Style.Copy();
                for (var r = 0; r < input.Rows.Count; r++)
                {
                    if (stoppedRows.Contains(r))
                        continue;
                    lookup.Row = input.Rows[r];
                    lookup.HasValue = false;
                    if (rule.Target == HighlightTarget.Cells)
                    {
                        lookup.HasValue = true;
                        foreach (var c in cols)
                        {
                            long key = ((long)r << 16) | (uint)c;
                            if (stoppedCells.Contains(key))
                                continue;
                            lookup.Value = input.Rows[r][c];
                            if (!Passes(test, find))
                                continue;
                            style.StyleCell(r, c, cellStyle);
                            if (rule.StopIfTrue)
                                stoppedCells.Add(key);
                        }
                        continue;
                    }
                    if (!Passes(test, find))
                        continue;
                    if (rule.Target == HighlightTarget.Row)
                        style.StyleRow(r, cellStyle);
                    else
                    {
                        foreach (var c in cols)
                            style.StyleCell(r, c, cellStyle);
                    }
                    if (rule.StopIfTrue)
                        stoppedRows.Add(r);
                }
            }
            return ExcelFile.WithStyle(input, style);
        }

        // A formula that cannot work on this row (text in a sum, say) is simply false, as in Filter.
        static bool Passes(Func<Func<string, object>, object> test, Func<string, object> find)
        {
            try
            {
                return ExpressionEngine.ToBool(test(find));
            }
            catch
            {
                return false;
            }
        }

        // How many rows and cells the result styles, for the editor.
        public static string Describe(DataTable result)
        {
            var style = ExcelFile.GetStyle(result);
            if (style == null)
                return "Nothing styled.";
            return style.StyledRows + " rows and " + style.StyledCells + " single cells styled.";
        }
    }
}
