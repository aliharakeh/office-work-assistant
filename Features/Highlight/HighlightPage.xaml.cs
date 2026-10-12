using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Media;
using OfficeWorkAssistant.Excel;

namespace OfficeWorkAssistant.Features.Highlight
{
    public partial class HighlightPage : UserControl
    {
        readonly DataTable _input;
        readonly Action<HighlightSettings> _use;
        readonly List<HighlightRule> _rules = new List<HighlightRule>();

        public HighlightPage(DataTable input, string label, HighlightSettings settings, Action<HighlightSettings> use)
        {
            InitializeComponent();
            _input = input;
            _use = use;
            InputInfo.Text = label + " - colours show in the result below and are written by the Save step. " +
                "Put this step right before Save: most other steps drop the colours.";
            if (settings != null)
                _rules.AddRange(settings.Rules);
            RebuildRules();
        }

        void RebuildRules()
        {
            RulePanel.Children.Clear();
            if (_rules.Count == 0)
            {
                RulePanel.Children.Add(new TextBlock { Text = "No rules yet.", Foreground = Brushes.Gray, Margin = new Thickness(0, 0, 0, 4) });
                return;
            }
            for (var i = 0; i < _rules.Count; i++)
            {
                var index = i;
                var text = (i + 1).ToString(CultureInfo.InvariantCulture) + ". " + RuleText(_rules[i]);
                var edit = new Button { Content = "Edit", Width = 56, Margin = new Thickness(0, 0, 6, 0) };
                edit.Click += (s, e) => EditRule(index);
                var up = new Button { Content = "Up", Width = 44, Margin = new Thickness(0, 0, 6, 0), ToolTip = "Run this rule earlier." };
                up.Click += (s, e) => MoveRule(index, -1);
                var down = new Button { Content = "Down", Width = 52, Margin = new Thickness(0, 0, 6, 0), ToolTip = "Run this rule later." };
                down.Click += (s, e) => MoveRule(index, 1);
                var remove = new Button { Content = "Remove", Width = 64 };
                remove.Click += (s, e) => RemoveRule(index);
                var buttons = new StackPanel { Orientation = Orientation.Horizontal, VerticalAlignment = VerticalAlignment.Center };
                buttons.Children.Add(edit);
                buttons.Children.Add(up);
                buttons.Children.Add(down);
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

        // "Columns B - Status when $Qty > 10 -> yellow fill" reads better than the stored letters.
        string RuleText(HighlightRule rule)
        {
            string where;
            if (rule.Target == HighlightTarget.Row)
                where = "Row";
            else
            {
                var names = new List<string>();
                foreach (var c in rule.Columns ?? new List<string>())
                {
                    var at = ExcelColumns.IndexOf(_input, c);
                    names.Add(at >= 0 ? ExcelColumns.Label(_input, at) : c);
                }
                where = (rule.Target == HighlightTarget.Cells ? "Cells " : "Columns ") + string.Join(", ", names.ToArray());
            }
            return where + " when " + rule.Condition + " -> " + (rule.Style ?? new CellStyle()).Describe() + (rule.StopIfTrue ? " (stop)" : "");
        }

        void Add_Click(object sender, RoutedEventArgs e)
        {
            EditRule(-1);
        }

        // index -1 adds a new rule.
        void EditRule(int index)
        {
            var seed = index >= 0 ? _rules[index] : null;
            var rule = HighlightRuleWindow.Edit(Window.GetWindow(this), _input, seed, index >= 0 ? index + 1 : _rules.Count + 1);
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

        // The pipeline applies this editor by itself after a click inside it; the dialog's clicks never reach it.
        void RulesChanged()
        {
            RebuildRules();
            RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
        }

        void Apply_Click(object sender, RoutedEventArgs e)
        {
            var s = new HighlightSettings();
            s.Rules.AddRange(_rules);
            if (s.Rules.Count == 0)
            {
                Status.Text = "Add a rule first.";
                return;
            }
            DataTable result;
            try
            {
                result = HighlightWork.Run(_input, s);
            }
            catch (Exception ex)
            {
                Notices.Show(this, ex.Message, "Highlight", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            Status.Text = HighlightWork.Describe(result);
            _use(s);
        }
    }
}
