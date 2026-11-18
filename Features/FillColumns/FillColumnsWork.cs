using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.FillColumns
{
    // One match rule: the A formula on an A row must equal the B formula on a B row.
    public sealed class FillColumnsKey
    {
        // Zero-based lookup index; old keys belong to A.
        public int Source { get; set; }
        public string FormulaA { get; set; }
        public string FormulaB { get; set; }
    }

    // One fill rule: when When is true (blank = always), Target gets Value.
    // Rules run top to bottom; the first rule that fires for a column wins that cell.
    public sealed class FillColumnsRule
    {
        public string Target { get; set; }
        public string When { get; set; }
        public string Value { get; set; }
    }

    // Everything one merge needs, so it can be stored (XML) and run again.
    public sealed class FillColumnsSettings
    {
        public int OutputInput { get; set; }
        public string MatchFormula { get; set; }
        public bool JointLookups { get; set; }
        public List<FillColumnsKey> Keys { get; set; }
        public bool MatchAll { get; set; }
        public List<FillColumnsRule> Rules { get; set; }

        public FillColumnsSettings()
        {
            Keys = new List<FillColumnsKey>();
            MatchAll = true;
            Rules = new List<FillColumnsRule>();
        }
    }

    public sealed class FillColumnsResult
    {
        public DataTable Table { get; set; }
        public int Matched { get; set; }
        public int Unmatched { get; set; }
        public int SharedKeys { get; set; }
        public int Filled { get; set; }
        public int[] FilledByRule { get; set; }
    }

    // Names inside fill formulas:
    //   $A_Name / $A_C   file A's matched row (blank when no match)
    //   $B_Name / $B_C   file B's row being filled (sees earlier rules' results)
    //   $Name / $C       B first; a name B lacks falls back to A
    //   $Value           the target cell's current value
    //   $Matched         true when an A row matched
    public static class FillColumnsWork
    {
        // maxRows <= 0 fills every B row (used for saving); a positive value keeps the preview small.
        public static FillColumnsResult Fill(IList<DataTable> sources, DataTable b, IList<FillColumnsKey> keys,
            IList<FillColumnsRule> rules, int maxRows, bool matchAll = true, int outputInput = 1, string matchFormula = null)
        {
            if (b == null)
                throw new InvalidOperationException("Connect the sheet you want to fill.");
            var inputTables = new List<DataTable>(sources);
            inputTables.Insert(outputInput - 1, b);
            var inputRows = new DataRow[inputTables.Count];
            var searched = new List<int>();
            Func<Func<string, object>, object> match = null;
            if (!string.IsNullOrWhiteSpace(matchFormula))
            {
                match = ExpressionEngine.Compile(matchFormula);
                var references = new SortedSet<int>();
                foreach (Match token in Regex.Matches(matchFormula, "\"[^\"]*\"|'[^']*'|\\$([0-9]+)_"))
                {
                    if (!token.Groups[1].Success) continue;
                    int input;
                    if (!int.TryParse(token.Groups[1].Value, out input) || input < 1 || input > inputTables.Count)
                        throw new InvalidOperationException("Match formula refers to an input that is not connected.");
                    if (input != outputInput) references.Add(input - 1);
                }
                searched.AddRange(references);
            }
            var groups = new List<FillColumnsKey>[sources.Count];
            var indexes = new Dictionary<string, int>[sources.Count];
            int duplicates = 0;
            for (int s = 0; s < sources.Count; s++)
            {
                groups[s] = new List<FillColumnsKey>();
                indexes[s] = new Dictionary<string, int>(StringComparer.Ordinal);
            }
            foreach (var key in keys ?? new List<FillColumnsKey>())
            {
                if (key.Source < 0 || key.Source >= sources.Count)
                    throw new InvalidOperationException("Reconnect input " + (key.Source + 2) + " or remove its match key.");
                groups[key.Source].Add(key);
            }
            for (int s = 0; s < sources.Count; s++)
            {
                if (groups[s].Count == 0)
                    continue;
                int input = s + 2;
                var lookup = sources[s];
                var shared = new HashSet<string>();
                for (int i = 0; i < lookup.Rows.Count; i++)
                {
                    var row = lookup.Rows[i];
                    foreach (var k in MatchKeys(groups[s], true, matchAll, name => LookupKeyValue(lookup, row, input, name)))
                        if (indexes[s].ContainsKey(k)) shared.Add(k);
                        else indexes[s].Add(k, i);
                }
                duplicates += shared.Count;
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

            var result = new FillColumnsResult { Table = table, SharedKeys = duplicates, FilledByRule = new int[targets.Length] };
            var matches = new DataRow[sources.Count];
            var done = new HashSet<DataColumn>();
            for (var r = 0; r < b.Rows.Count; r++)
            {
                if (maxRows > 0 && r >= maxRows)
                    break;
                var dest = table.NewRow();
                for (var c = 0; c < b.Columns.Count; c++)
                    dest[c] = b.Rows[r][c];

                bool any = false;
                for (int s = 0; s < sources.Count; s++)
                {
                    if (groups[s].Count == 0)
                    {
                        matches[s] = r < sources[s].Rows.Count ? sources[s].Rows[r] : null;
                        any |= matches[s] != null;
                        continue;
                    }
                    int best = int.MaxValue;
                    foreach (var k in MatchKeys(groups[s], false, matchAll,
                        name => ComparisonValue(b, b.Rows[r], sources, matches, s, name)))
                    {
                        int index;
                        if (indexes[s].TryGetValue(k, out index)) best = Math.Min(best, index);
                    }
                    matches[s] = best == int.MaxValue ? null : sources[s].Rows[best];
                    any |= matches[s] != null;
                }
                for (int i = 0, source = 0; i < inputRows.Length; i++)
                    inputRows[i] = i == outputInput - 1 ? b.Rows[r] : matches[source++];
                if (match != null)
                {
                    bool found;
                    try
                    {
                        found = FindMatch(inputTables, inputRows, searched, 0, match, b, b.Rows[r]);
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException("Match formula on output data row " + (r + 1) + ": " + ex.Message, ex);
                    }
                    if (!found)
                    {
                        result.Unmatched++;
                        table.Rows.Add(dest);
                        continue;
                    }
                    for (int i = 0, source = 0; i < inputRows.Length; i++)
                        if (i != outputInput - 1) matches[source++] = inputRows[i];
                    any = true;
                }
                inputRows[outputInput - 1] = dest;
                if (any)
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
                    Func<string, object> names = name =>
                    {
                        object value;
                        if (!TryNumberedCell(inputTables, inputRows, outputInput, table, name, out value))
                            Resolve(sources, matches, table, dest, col, name, out value);
                        return value;
                    };
                    try
                    {
                        if (!string.IsNullOrWhiteSpace(rule.When) &&
                            !ExpressionEngine.ToBool(ExpressionEngine.Evaluate(rule.When, names)))
                            continue;
                        dest[col] = ExpressionEngine.Evaluate(rule.Value, names) ?? DBNull.Value;
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException("Input data row " + (r + 1) + ", fill rule " + (i + 1) +
                            " (" + ExcelFile.Header(col) + "): " + ex.Message, ex);
                    }
                    done.Add(col);
                    result.Filled++;
                    result.FilledByRule[i]++;
                }
                table.Rows.Add(dest);
            }
            return result;
        }

        static bool FindMatch(IList<DataTable> inputs, DataRow[] rows, IList<int> searched, int depth,
            Func<Func<string, object>, object> formula, DataTable target, DataRow targetRow)
        {
            if (depth == searched.Count)
            {
                return ExpressionEngine.ToBool(formula(name =>
                {
                    object value;
                    if (!TryNumberedCell(inputs, rows, 0, null, name, out value))
                        TrySide(target, targetRow, name, false, out value);
                    return value;
                }));
            }
            int input = searched[depth];
            if (inputs[input].Rows.Count == 0)
                return FindMatch(inputs, rows, searched, depth + 1, formula, target, targetRow);

            foreach (DataRow row in inputs[input].Rows)
            {
                rows[input] = row;
                if (FindMatch(inputs, rows, searched, depth + 1, formula, target, targetRow))
                    return true;
            }
            return false;
        }

        static bool TryNumberedCell(IList<DataTable> inputs, DataRow[] rows, int outputInput,
            DataTable output, string name, out object value)
        {
            string column;
            int input = InputReference(name, out column);
            value = null;
            if (input == 0) return false;
            if (input > inputs.Count)
                throw new InvalidOperationException("Input " + input + " is not connected.");
            TryCell(input == outputInput ? output : inputs[input - 1], rows[input - 1], column, true, out value);
            return true;
        }

        public static DataTable Run(DataTable a, DataTable b, FillColumnsSettings settings)
        {
            if (settings == null)
                throw new InvalidOperationException("Set up the fill rules first.");
            return Fill(a == null ? new DataTable[0] : new[] { a }, b, settings.Keys, settings.Rules, 0, settings.MatchAll).Table;
        }

        static IEnumerable<string> MatchKeys(IList<FillColumnsKey> keys, bool isA, bool matchAll, Func<string, object> names)
        {
            if (matchAll)
            {
                var key = Key(keys, isA, names);
                if (key.Length > 0) yield return key;
                yield break;
            }
            for (var i = 0; i < keys.Count; i++)
            {
                var key = Key(new[] { keys[i] }, isA, names);
                if (key.Length > 0) yield return i + ":" + key;
            }
        }

        static string Key(IList<FillColumnsKey> keys, bool isA, Func<string, object> names)
        {
            var k = "";
            for (var i = 0; i < keys.Count; i++)
            {
                var formula = isA ? keys[i].FormulaA : keys[i].FormulaB;
                object value;
                try
                {
                    value = ExpressionEngine.Evaluate(formula, names);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException("Match rule " + (i + 1) + " (" + (isA ? "lookup" : "input") + " side): " + ex.Message, ex);
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

        // $1_Name reads the target. $2_Name and higher read lookups in input order.
        static int InputReference(string name, out string column)
        {
            column = null;
            if (string.IsNullOrEmpty(name) || !char.IsDigit(name[0])) return 0;
            int split = name.IndexOf('_');
            int input;
            if (split < 1 || !int.TryParse(name.Substring(0, split), NumberStyles.None,
                CultureInfo.InvariantCulture, out input)) return 0;
            if (input < 1 || split == name.Length - 1)
                throw new InvalidOperationException("Use a numbered input and column, for example $1_Code.");
            column = name.Substring(split + 1);
            return input;
        }

        static object LookupKeyValue(DataTable lookup, DataRow row, int ownInput, string name)
        {
            string column;
            int input = InputReference(name, out column);
            object value;
            if (input == 0) TrySide(lookup, row, name, true, out value);
            else if (input == ownInput) TryCell(lookup, row, column, true, out value);
            else throw new InvalidOperationException("The lookup key for input " + ownInput + " must read that input.");
            return value;
        }

        static object ComparisonValue(DataTable target, DataRow row, IList<DataTable> sources,
            DataRow[] matches, int source, string name)
        {
            string column;
            int input = InputReference(name, out column);
            object value;
            if (input == 0) TrySide(target, row, name, false, out value);
            else if (input == 1) TryCell(target, row, column, true, out value);
            else if (input <= source + 1) TryCell(sources[input - 2], matches[input - 2], column, true, out value);
            else throw new InvalidOperationException("Match input " + (source + 2) + " against input 1 or an earlier matched input.");
            return value;
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

        public static string Alias(int index)
        {
            return index == 0 ? "A" : "A" + (index + 1);
        }

        static bool Resolve(IList<DataTable> sources, DataRow[] rows, DataTable b, DataRow rowB, DataColumn target, string name, out object value)
        {
            for (int i = 1; i < sources.Count; i++)
            {
                var prefix = Alias(i) + "_";
                if (name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    return TryCell(sources[i], rows == null ? null : rows[i], name.Substring(prefix.Length), true, out value);
                if (string.Equals(name, "Matched" + (i + 1), StringComparison.OrdinalIgnoreCase))
                {
                    value = rows != null && rows[i] != null;
                    return true;
                }
            }
            return TryFind(sources.Count == 0 ? null : sources[0], rows == null || rows.Length == 0 ? null : rows[0],
                b, rowB, target, name, out value);
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
