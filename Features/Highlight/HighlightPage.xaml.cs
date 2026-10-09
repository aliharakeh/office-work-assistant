using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.Highlight
{
    public partial class HighlightPage : UserControl
    {
        static readonly string[] Ops = { "==", "!=", ">", ">=", "<", "<=", "contains", "startsWith", "endsWith", "isEmpty", "isNotEmpty" };

        readonly DataTable _input;
        readonly Action<HighlightSettings> _use;
        readonly List<HighlightRule> _rules = new List<HighlightRule>();
        List<Pick> _picks = new List<Pick>();
        int _editing = -1;

        public HighlightPage(DataTable input, string label, HighlightSettings settings, Action<HighlightSettings> use)
        {
            InitializeComponent();
            _input = input;
            _use = use;
            InputInfo.Text = label + " - colours show in the result below and are written by the Save step. " +
                "Put this step right before Save: most other steps drop the colours.";
            FormulaField.Watch(ConditionBox);
            BuildOp.SelectedIndex = 0;
            if (settings != null)
                _rules.AddRange(settings.Rules);
            ShowRules();
            LoadRule(NewRule());
        }

        static HighlightRule NewRule()
        {
            return new HighlightRule { Style = new CellStyle { Fill = NamedColors.ToHex("Light red") } };
        }

        void ShowRules()
        {
            RuleList.ItemsSource = null;
            RuleList.ItemsSource = _rules;
        }

        void LoadRule(HighlightRule rule)
        {
            var on = new HashSet<int>();
            foreach (var c in rule.Columns)
            {
                int i = ExcelColumns.IndexOf(_input, c);
                if (i >= 0)
                    on.Add(i);
            }
            _picks = new List<Pick>();
            for (var i = 0; i < _input.Columns.Count; i++)
                _picks.Add(new Pick { Index = i, Label = ExcelColumns.Label(_input, i), Include = on.Contains(i) });
            ColumnList.ItemsSource = _picks;
            TargetBox.SelectedIndex = (int)rule.Target;
            ConditionBox.Text = rule.Condition ?? "";
            StylePicker.Value = rule.Style;
            StopBox.IsChecked = rule.StopIfTrue;
            UpdateTarget();
        }

        HighlightRule EditorRule()
        {
            var rule = new HighlightRule
            {
                Condition = ConditionBox.Text.Trim(),
                Target = (HighlightTarget)Math.Max(0, TargetBox.SelectedIndex),
                Style = StylePicker.Value,
                StopIfTrue = StopBox.IsChecked == true
            };
            if (rule.Target != HighlightTarget.Row)
            {
                foreach (var p in _picks)
                {
                    if (p.Include)
                        rule.Columns.Add(ExcelColumns.Letter(p.Index));
                }
            }
            return rule;
        }

        void TargetBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateTarget();
        }

        void UpdateTarget()
        {
            if (ColumnsPanel == null || BuildColumn == null)
                return;
            bool cells = TargetBox.SelectedIndex == (int)HighlightTarget.Cells;
            ColumnsPanel.Visibility = TargetBox.SelectedIndex == (int)HighlightTarget.Row ? Visibility.Collapsed : Visibility.Visible;
            ConditionHint.Text = cells
                ? "Checked for each picked cell. $Value is that cell: $Value < 0, $Value > $Today, ISBLANK($Value). Other columns of the row work too."
                : "Checked for each row: $Status == \"Late\", $Qty > 10 && $Price < 5, $Due < $Today, CONTAINS($Name, \"Ltd\").";

            var items = new List<string>();
            if (cells)
                items.Add("This cell ($Value)");
            items.AddRange(ExcelColumns.Labels(_input));
            int keep = BuildColumn.SelectedIndex;
            BuildColumn.ItemsSource = items;
            BuildColumn.SelectedIndex = items.Count == 0 ? -1 : Math.Max(0, Math.Min(keep, items.Count - 1));
        }

        void Build_Click(object sender, RoutedEventArgs e)
        {
            int picked = BuildColumn.SelectedIndex;
            if (picked < 0)
                return;
            bool cells = TargetBox.SelectedIndex == (int)HighlightTarget.Cells;
            string name;
            if (cells && picked == 0)
                name = "$Value";
            else
                name = "$" + ExcelColumns.FormulaName(_input, cells ? picked - 1 : picked);

            string op = Ops[Math.Max(0, BuildOp.SelectedIndex)];
            string value = BuildValue.Text.Trim();
            string text;
            if (op == "isEmpty")
                text = "ISBLANK(" + name + ")";
            else if (op == "isNotEmpty")
                text = "!ISBLANK(" + name + ")";
            else if (op == "contains")
                text = "CONTAINS(" + name + ", " + Quote(value) + ")";
            else if (op == "startsWith")
                text = "STARTSWITH(" + name + ", " + Quote(value) + ")";
            else if (op == "endsWith")
                text = "ENDSWITH(" + name + ", " + Quote(value) + ")";
            else
                text = name + " " + op + " " + Literal(value);
            ConditionBox.Text = text;
        }

        // Numbers and $built-ins as they are; anything else as quoted text.
        static string Literal(string value)
        {
            double d;
            if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out d))
                return value;
            if (value.StartsWith("$", StringComparison.Ordinal) && ExpressionEngine.Validate(value) == null)
                return value;
            return Quote(value);
        }

        static string Quote(string value)
        {
            return value.IndexOf('"') >= 0 ? "'" + value + "'" : "\"" + value + "\"";
        }

        void Help_Click(object sender, RoutedEventArgs e)
        {
            ExpressionHelp.Show(Window.GetWindow(this),
                "A condition is a formula that is true or false. $Qty or $A is a column of the row (A = 1st); $Value is the cell itself when each cell is tested on its own. " +
                "Compare with == != > >= < <=, join with && (and) || (or), and use the functions below.");
        }

        void RuleList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int index = RuleList.SelectedIndex;
            if (index < 0 || index >= _rules.Count)
                return;
            _editing = index;
            LoadRule(_rules[index]);
            AddBtn.Content = "Update rule";
        }

        void Add_Click(object sender, RoutedEventArgs e)
        {
            if (!FormulaField.Check(ConditionBox))
                return;
            var rule = EditorRule();
            string err = HighlightWork.Validate(rule, _editing >= 0 ? _editing + 1 : _rules.Count + 1);
            if (err != null)
            {
                MessageBox.Show(Window.GetWindow(this), err, "Highlight", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (_editing >= 0 && _editing < _rules.Count)
                _rules[_editing] = rule;
            else
                _rules.Add(rule);
            New_Click(null, null);
        }

        void Remove_Click(object sender, RoutedEventArgs e)
        {
            int index = RuleList.SelectedIndex;
            if (index < 0 || index >= _rules.Count)
            {
                Status.Text = "Select a rule first.";
                return;
            }
            _rules.RemoveAt(index);
            New_Click(null, null);
        }

        void Up_Click(object sender, RoutedEventArgs e)
        {
            Move(-1);
        }

        void Down_Click(object sender, RoutedEventArgs e)
        {
            Move(1);
        }

        void Move(int delta)
        {
            int index = RuleList.SelectedIndex;
            int target = index + delta;
            if (index < 0 || target < 0 || target >= _rules.Count)
                return;
            var rule = _rules[index];
            _rules[index] = _rules[target];
            _rules[target] = rule;
            _editing = -1;
            AddBtn.Content = "Add rule";
            ShowRules();
            RuleList.SelectedIndex = target;
        }

        void New_Click(object sender, RoutedEventArgs e)
        {
            _editing = -1;
            AddBtn.Content = "Add rule";
            ShowRules();
            LoadRule(NewRule());
        }

        void Apply_Click(object sender, RoutedEventArgs e)
        {
            var s = new HighlightSettings();
            s.Rules.AddRange(_rules);
            // Nothing added yet but a condition typed: the rule in the editor is meant.
            bool fromEditor = s.Rules.Count == 0 && ConditionBox.Text.Trim().Length > 0;
            if (fromEditor)
            {
                if (!FormulaField.Check(ConditionBox))
                    return;
                s.Rules.Add(EditorRule());
            }
            DataTable result;
            try
            {
                result = HighlightWork.Run(_input, s);
            }
            catch (Exception ex)
            {
                MessageBox.Show(Window.GetWindow(this), ex.Message, "Highlight", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (fromEditor)
            {
                _rules.AddRange(s.Rules);
                New_Click(null, null);
            }
            Status.Text = HighlightWork.Describe(result);
            _use(s);
        }

        public sealed class Pick
        {
            public int Index { get; set; }
            public string Label { get; set; }
            public bool Include { get; set; }
        }
    }
}
