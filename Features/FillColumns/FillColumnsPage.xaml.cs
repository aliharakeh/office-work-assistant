using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.FillColumns
{
    public partial class FillColumnsPage : UserControl
    {
        DataTable _tableA;
        DataTable _tableB;
        bool _cutB;
        int _previewRows = 1000;
        bool _loading;
        readonly List<FillColumnsKey> _keys = new List<FillColumnsKey>();
        readonly List<FillColumnsRule> _rules = new List<FillColumnsRule>();
        bool _restoring;
        // Inputs come from earlier pipeline steps, and "Use in pipeline"
        // hands the settings back.
        DataTable _pipeA;
        DataTable _pipeB;
        string _pipeLabelA;
        string _pipeLabelB;
        Action<FillColumnsSettings> _use;

        public FillColumnsPage(DataTable a, string labelA, DataTable b, string labelB,
            FillColumnsSettings settings, Action<FillColumnsSettings> use)
        {
            InitializeComponent();
            _pipeA = a;
            _pipeB = b;
            _pipeLabelA = labelA;
            _pipeLabelB = labelB;
            _use = use;
            PreviewBox.Items.Add(200);
            PreviewBox.Items.Add(1000);
            PreviewBox.Items.Add(5000);
            PreviewBox.SelectedItem = 1000;
            ExcelGrid.Hook(GridA, GridB, GridOut);
            RebuildKeys();
            RebuildRules();
            InjectInputs();
            if (settings != null)
                ApplySettings(settings);
        }

        void InjectInputs()
        {
            ApplyLoad(true, _pipeA, _pipeLabelA);
            ApplyLoad(false, _pipeB, _pipeLabelB);
        }

        void ApplySettings(FillColumnsSettings s)
        {
            _restoring = true;
            try
            {
                _keys.Clear();
                if (s.Keys != null)
                {
                    foreach (var k in s.Keys)
                        _keys.Add(CopyKey(k));
                }
                _rules.Clear();
                if (s.Rules != null)
                {
                    foreach (var r in s.Rules)
                        _rules.Add(CopyRule(r));
                }
                RebuildKeys();
                RebuildRules();
            }
            finally
            {
                _restoring = false;
            }
            RefreshPreview();
        }

        void UseInPipeline_Click(object sender, RoutedEventArgs e)
        {
            if (!FormulasOk())
            {
                Alert("Fix the formulas marked in red first.", "Formula error", MessageBoxImage.Warning);
                return;
            }
            var s = new FillColumnsSettings { Keys = Keys(), Rules = Rules() };
            try
            {
                // Run it once on every row so a broken rule shows up here, not later in the pipeline.
                FillColumnsWork.Run(FullTable(true), FullTable(false), s);
            }
            catch (Exception ex)
            {
                Alert(ex.Message, "Could not merge", MessageBoxImage.Error);
                return;
            }
            _use(s);
        }

        void PreviewBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (PreviewBox.SelectedItem == null)
                return;
            var rows = (int)PreviewBox.SelectedItem;
            if (rows == _previewRows)
                return;
            _previewRows = rows;
            InjectInputs();
        }

        void AddKey_Click(object sender, RoutedEventArgs e)
        {
            EditKey(-1);
        }

        void AddRule_Click(object sender, RoutedEventArgs e)
        {
            EditRule(-1);
        }

        // The grids show a capped preview of the pipeline input.
        void ApplyLoad(bool isA, DataTable full, string label)
        {
            var preview = full;
            var cut = full.Rows.Count > _previewRows;
            if (cut)
            {
                preview = full.Clone();
                for (var i = 0; i < _previewRows; i++)
                    preview.ImportRow(full.Rows[i]);
            }

            _loading = true;
            try
            {
                if (isA)
                    _tableA = preview;
                else
                {
                    _tableB = preview;
                    _cutB = cut;
                }
                (isA ? PathA : PathB).Text = label;
                (isA ? GridA : GridB).ItemsSource = preview.DefaultView;
                (isA ? InfoA : InfoB).Text = RowInfo(preview.Rows.Count, full.Rows.Count, cut);
            }
            finally
            {
                _loading = false;
            }
            RefreshPreview();
        }

        static string RowInfo(int shown, int total, bool cut)
        {
            var rows = shown.ToString("N0", CultureInfo.InvariantCulture);
            if (!cut)
                return rows + " rows";
            return "Showing " + rows + " of " + total.ToString("N0", CultureInfo.InvariantCulture) +
                   " rows (all rows are used when you save)";
        }

        // Matching needs every A row, not just the capped preview.
        DataTable FullTable(bool isA)
        {
            return isA ? _pipeA : _pipeB;
        }

        void RefreshPreview()
        {
            if (_loading || _restoring)
                return;
            GridOut.ItemsSource = null;
            if (_tableA == null || _tableB == null)
            {
                Status.Text = "Open both files and set a match rule.";
                return;
            }
            if (!FormulasOk())
            {
                Status.Text = "Fix the formulas marked in red. Hover one to see why.";
                return;
            }

            try
            {
                var a = FullTable(true);
                var keys = Keys();
                var rules = Rules();
                var result = FillColumnsWork.Fill(a, _tableB, keys, rules, _previewRows);
                GridOut.ItemsSource = result.Table.DefaultView;

                var text = result.Matched.ToString("N0", CultureInfo.InvariantCulture) + " matched, " +
                           result.Unmatched.ToString("N0", CultureInfo.InvariantCulture) + " not matched, " +
                           result.Filled.ToString("N0", CultureInfo.InvariantCulture) + " cells filled";
                if (_cutB)
                    text += " (preview rows only)";
                text += ".";
                if (result.SharedKeys > 0)
                    text += " " + result.SharedKeys.ToString("N0", CultureInfo.InvariantCulture) +
                            " keys appear on more than one A row; the first A row is used.";
                var unknown = FillColumnsWork.UnknownNames(a, result.Table, keys, rules);
                if (unknown.Length > 0)
                    text += " Unknown names read as blank: $" + string.Join(", $", unknown) + ".";
                Status.Text = text;
            }
            catch (Exception ex)
            {
                Status.Text = ex.Message;
            }
        }

        bool FormulasOk()
        {
            foreach (var k in _keys)
            {
                if (ExpressionEngine.Validate(k.FormulaA) != null || ExpressionEngine.Validate(k.FormulaB) != null)
                    return false;
            }
            foreach (var r in _rules)
            {
                if ((!string.IsNullOrWhiteSpace(r.When) && ExpressionEngine.Validate(r.When) != null) ||
                    ExpressionEngine.Validate(r.Value) != null)
                    return false;
            }
            return true;
        }

        List<FillColumnsKey> Keys()
        {
            return new List<FillColumnsKey>(_keys);
        }

        List<FillColumnsRule> Rules()
        {
            return new List<FillColumnsRule>(_rules);
        }

        void RebuildKeys()
        {
            KeyPanel.Children.Clear();
            if (_keys.Count == 0)
            {
                KeyPanel.Children.Add(new TextBlock
                {
                    Text = "No match keys yet.",
                    Foreground = Brushes.Gray,
                    Margin = new Thickness(0, 0, 0, 4)
                });
                return;
            }
            for (var i = 0; i < _keys.Count; i++)
            {
                var index = i;
                var text = (_keys[i].FormulaA ?? "") + "   ==   " + (_keys[i].FormulaB ?? "");
                var edit = new Button { Content = "Edit", Width = 56, Margin = new Thickness(0, 0, 6, 0) };
                edit.Click += (s, e) => EditKey(index);
                var remove = new Button { Content = "Remove", Width = 64 };
                remove.Click += (s, e) => RemoveKey(index);
                var buttons = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
                buttons.Children.Add(edit);
                buttons.Children.Add(remove);
                var row = new DockPanel { Margin = new Thickness(0, 0, 0, 4) };
                DockPanel.SetDock(buttons, Dock.Right);
                row.Children.Add(buttons);
                row.Children.Add(new TextBlock
                {
                    Text = text,
                    ToolTip = text,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, 0, 8, 0)
                });
                KeyPanel.Children.Add(row);
            }
        }

        static FillColumnsKey CopyKey(FillColumnsKey k)
        {
            return new FillColumnsKey { FormulaA = k.FormulaA, FormulaB = k.FormulaB };
        }

        // index -1 adds a new key.
        void EditKey(int index)
        {
            var seed = index >= 0 ? _keys[index] : null;
            var key = MatchKeyWindow.Edit(Window.GetWindow(this), seed);
            if (key == null)
                return;
            if (index >= 0)
                _keys[index] = key;
            else
                _keys.Add(key);
            KeysChanged();
        }

        void RemoveKey(int index)
        {
            _keys.RemoveAt(index);
            KeysChanged();
        }

        void KeysChanged()
        {
            RebuildKeys();
            RefreshPreview();
            RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
        }

        void RebuildRules()
        {
            RulePanel.Children.Clear();
            if (_rules.Count == 0)
            {
                RulePanel.Children.Add(new TextBlock
                {
                    Text = "No fill rules yet.",
                    Foreground = Brushes.Gray,
                    Margin = new Thickness(0, 0, 0, 4)
                });
                return;
            }
            for (var i = 0; i < _rules.Count; i++)
            {
                var index = i;
                var text = (i + 1).ToString(CultureInfo.InvariantCulture) + ". " + RuleText(_rules[i]);
                var edit = new Button { Content = "Edit", Width = 56, Margin = new Thickness(0, 0, 6, 0) };
                edit.Click += (s, e) => EditRule(index);
                var up = new Button { Content = "Up", Width = 52, Margin = new Thickness(0, 0, 6, 0), ToolTip = "Run this rule earlier." };
                up.Click += (s, e) => MoveRule(index, -1);
                var remove = new Button { Content = "Remove", Width = 64 };
                remove.Click += (s, e) => RemoveRule(index);
                var buttons = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
                buttons.Children.Add(edit);
                buttons.Children.Add(up);
                buttons.Children.Add(remove);
                var row = new DockPanel { Margin = new Thickness(0, 0, 0, 4) };
                DockPanel.SetDock(buttons, Dock.Right);
                row.Children.Add(buttons);
                row.Children.Add(new TextBlock
                {
                    Text = text,
                    ToolTip = text,
                    TextTrimming = TextTrimming.CharacterEllipsis,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, 0, 8, 0)
                });
                RulePanel.Children.Add(row);
            }
        }

        static string RuleText(FillColumnsRule r)
        {
            var text = (r.Target ?? "").Trim();
            if (!string.IsNullOrWhiteSpace(r.When))
                text += " when " + r.When;
            return text + " = " + (r.Value ?? "");
        }

        static FillColumnsRule CopyRule(FillColumnsRule r)
        {
            return new FillColumnsRule { Target = r.Target, When = r.When, Value = r.Value };
        }

        // index -1 adds a new rule.
        void EditRule(int index)
        {
            if (_tableB == null)
            {
                Alert("Open both Excel files first.", "Missing file", MessageBoxImage.Warning);
                return;
            }
            var seed = index >= 0 ? _rules[index] : null;
            var rule = FillRuleWindow.Edit(Window.GetWindow(this), _tableB, seed);
            if (rule == null)
                return;
            if (index >= 0)
                _rules[index] = rule;
            else
                _rules.Add(rule);
            RulesChanged();
        }

        void RemoveRule(int index)
        {
            _rules.RemoveAt(index);
            RulesChanged();
        }

        void MoveRule(int index, int delta)
        {
            var target = index + delta;
            if (target < 0 || target >= _rules.Count)
                return;
            var rule = _rules[index];
            _rules[index] = _rules[target];
            _rules[target] = rule;
            RulesChanged();
        }

        void RulesChanged()
        {
            RebuildRules();
            RefreshPreview();
            RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
        }
        void Alert(string message, string title, MessageBoxImage icon)
        {
            Notices.Show(this, message, title, MessageBoxButton.OK, icon);
        }
    }
}
