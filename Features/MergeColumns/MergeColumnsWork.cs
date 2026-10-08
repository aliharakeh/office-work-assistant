using System;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.MergeColumns
{
    // One match rule: the A formula on an A row must equal the B formula on a B row.
    public sealed class MergeColumnsKey
    {
        public string FormulaA { get; set; }
        public string FormulaB { get; set; }
    }

    // One fill rule: when When is true (blank = always), Target gets Value.
    // Rules run top to bottom; the first rule that fires for a column wins that cell.
    public sealed class MergeColumnsRule
    {
        public string Target { get; set; }
        public string When { get; set; }
        public string Value { get; set; }
    }

    public sealed class MergeColumnsResult
    {
        public DataTable Table { get; set; }
        public int Matched { get; set; }
        public int Unmatched { get; set; }
        public int SharedKeys { get; set; }
        public int Filled { get; set; }
    }

    // Names inside fill formulas:
    //   $A_Name / $A_C   file A's matched row (blank when no match)
    //   $B_Name / $B_C   file B's row being filled (sees earlier rules' results)
    //   $Name / $C       B first; a name B lacks falls back to A
    //   $Value           the target cell's current value
    //   $Matched         true when an A row matched
    public static class MergeColumnsWork
    {
        static readonly Regex Names = new Regex(@"\$(\w+)", RegexOptions.Compiled);

        public static ExcelLoadResult Load(string path, string sheetName, int maxRows)
        {
            return ExcelFile.Load(path, sheetName, maxRows);
        }

        // maxRows <= 0 fills every B row (used for saving); a positive value keeps the preview small.
        public static MergeColumnsResult Merge(DataTable a, DataTable b, IList<MergeColumnsKey> keys,
            IList<MergeColumnsRule> rules, int maxRows)
        {
            if (a == null || b == null)
                throw new InvalidOperationException("Open both Excel files first.");
            if (keys == null || keys.Count == 0)
                throw new InvalidOperationException("Add a match rule.");

            // First A row per key wins; later rows with the same key are counted, not used.
            var lookup = new Dictionary<string, DataRow>(StringComparer.Ordinal);
            var shared = new HashSet<string>(StringComparer.Ordinal);
            foreach (DataRow row in a.Rows)
            {
                var k = Key(a, row, keys, true);
                if (k.Length == 0)
                    continue;
                if (lookup.ContainsKey(k))
                    shared.Add(k);
                else
                    lookup[k] = row;
            }

            var table = b.Clone();
            var targets = new DataColumn[rules == null ? 0 : rules.Count];
            for (var i = 0; i < targets.Length; i++)
            {
                var name = rules[i].Target;
                if (string.IsNullOrWhiteSpace(name))
                    throw new InvalidOperationException("Fill rule " + (i + 1) + " needs a target column.");
                if (string.IsNullOrWhiteSpace(rules[i].Value))
                    throw new InvalidOperationException("Fill rule " + (i + 1) + " needs a value formula.");
                targets[i] = ExcelFile.FindColumn(table, name.Trim()) ?? ExcelFile.AddColumn(table, name);
            }

            var result = new MergeColumnsResult { Table = table, SharedKeys = shared.Count };
            var done = new HashSet<DataColumn>();
            for (var r = 0; r < b.Rows.Count; r++)
            {
                if (maxRows > 0 && r >= maxRows)
                    break;
                var dest = table.NewRow();
                for (var c = 0; c < b.Columns.Count; c++)
                    dest[c] = b.Rows[r][c];

                DataRow match = null;
                var k = Key(b, b.Rows[r], keys, false);
                if (k.Length > 0)
                    lookup.TryGetValue(k, out match);
                if (match != null)
                    result.Matched++;
                else
                    result.Unmatched++;

                done.Clear();
                for (var i = 0; i < targets.Length; i++)
                {
                    var col = targets[i];
                    if (done.Contains(col))
                        continue;
                    var rule = rules[i];
                    Func<string, object> names = delegate(string n) { return Find(a, match, table, dest, col, n); };
                    try
                    {
                        if (!string.IsNullOrWhiteSpace(rule.When) &&
                            !ExpressionEngine.ToBool(ExpressionEngine.Evaluate(rule.When, names)))
                            continue;
                        dest[col] = ExpressionEngine.Evaluate(rule.Value, names) ?? DBNull.Value;
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException("B data row " + (r + 1) + ", fill rule " + (i + 1) +
                            " (" + ExcelFile.Header(col) + "): " + ex.Message, ex);
                    }
                    done.Add(col);
                    result.Filled++;
                }
                table.Rows.Add(dest);
            }
            return result;
        }

        // $names that no column or built-in answers. They evaluate as blank, which is usually a typo.
        // b should be Merge's result table so new target columns count as known.
        public static string[] UnknownNames(DataTable a, DataTable b, IList<MergeColumnsKey> keys, IList<MergeColumnsRule> rules)
        {
            var missing = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            object v;
            if (keys != null)
            {
                foreach (var key in keys)
                {
                    foreach (var n in NamesIn(key.FormulaA))
                        if (!TrySide(a, null, n, true, out v) && seen.Add(n)) missing.Add(n);
                    foreach (var n in NamesIn(key.FormulaB))
                        if (!TrySide(b, null, n, false, out v) && seen.Add(n)) missing.Add(n);
                }
            }
            if (rules != null)
            {
                foreach (var rule in rules)
                {
                    var names = new List<string>(NamesIn(rule.When));
                    names.AddRange(NamesIn(rule.Value));
                    foreach (var n in names)
                        if (!TryFind(a, null, b, null, null, n, out v) && seen.Add(n)) missing.Add(n);
                }
            }
            return missing.ToArray();
        }

        static IEnumerable<string> NamesIn(string formula)
        {
            if (string.IsNullOrEmpty(formula))
                yield break;
            foreach (Match m in Names.Matches(formula))
                yield return m.Groups[1].Value;
        }

        static string Key(DataTable table, DataRow row, IList<MergeColumnsKey> keys, bool isA)
        {
            var k = "";
            for (var i = 0; i < keys.Count; i++)
            {
                var formula = isA ? keys[i].FormulaA : keys[i].FormulaB;
                object value;
                try
                {
                    value = ExpressionEngine.Evaluate(formula, delegate(string n) { object v; TrySide(table, row, n, isA, out v); return v; });
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("Match rule " + (i + 1) + " (" + (isA ? "A" : "B") + " side): " + ex.Message, ex);
                }
                var part = ExpressionEngine.ToText(value).Trim().ToLowerInvariant();
                if (part.Length == 0)
                    return "";
                if (i > 0)
                    k += "\x1f";
                k += part;
            }
            return k;
        }

        // Match formulas see only their own file. The file's own prefix ($A_ on A) is accepted too.
        static bool TrySide(DataTable table, DataRow row, string name, bool isA, out object value)
        {
            value = null;
            if (string.IsNullOrEmpty(name))
                return false;
            if (HasPrefix(name, isA ? 'A' : 'B') && TryCell(table, row, name.Substring(2), true, out value))
                return true;
            if (ExpressionEngine.TryGetVariable(name, out value))
                return true;
            return TryCell(table, row, name, true, out value);
        }

        static object Find(DataTable a, DataRow rowA, DataTable b, DataRow rowB, DataColumn target, string name)
        {
            object v;
            TryFind(a, rowA, b, rowB, target, name, out v);
            return v;
        }

        static bool TryFind(DataTable a, DataRow rowA, DataTable b, DataRow rowB, DataColumn target, string name, out object value)
        {
            value = null;
            if (string.IsNullOrEmpty(name))
                return false;
            if (string.Equals(name, "Value", StringComparison.OrdinalIgnoreCase))
            {
                value = target != null && rowB != null ? rowB[target] : null;
                return true;
            }
            if (string.Equals(name, "Matched", StringComparison.OrdinalIgnoreCase))
            {
                value = rowA != null;
                return true;
            }
            if (HasPrefix(name, 'A') && TryCell(a, rowA, name.Substring(2), true, out value))
                return true;
            if (HasPrefix(name, 'B') && TryCell(b, rowB, name.Substring(2), true, out value))
                return true;
            if (ExpressionEngine.TryGetVariable(name, out value))
                return true;
            if (TryCell(b, rowB, name, true, out value))
                return true;
            // Letters stay on B so $E never jumps to file A just because B is narrower.
            return TryCell(a, rowA, name, false, out value);
        }

        static bool HasPrefix(string name, char side)
        {
            return name.Length > 2 && name[1] == '_' && char.ToUpperInvariant(name[0]) == side;
        }

        // Found when the table has the column; the value is blank when there is no row (no match).
        static bool TryCell(DataTable table, DataRow row, string name, bool letters, out object value)
        {
            value = null;
            if (table == null)
                return false;
            DataColumn col = null;
            int index;
            if (letters && TryParseColumnLetter(name, out index) && index < table.Columns.Count)
                col = table.Columns[index];
            if (col == null)
                col = ExcelFile.FindColumn(table, name);
            if (col == null)
                return false;
            value = row != null ? row[col] : null;
            return true;
        }

        static bool TryParseColumnLetter(string text, out int index)
        {
            index = -1;
            if (string.IsNullOrWhiteSpace(text))
                return false;
            var t = text.Trim();
            var n = 0;
            for (var i = 0; i < t.Length; i++)
            {
                var c = char.ToUpperInvariant(t[i]);
                if (c < 'A' || c > 'Z')
                    return false;
                n = n * 26 + (c - 'A' + 1);
                if (n > 16384)
                    return false;
            }
            if (n < 1)
                return false;
            index = n - 1;
            return true;
        }
    }
}
