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
        readonly IList<DataTable> _inputs;
        readonly List<FillColumnsRule> _rules = new List<FillColumnsRule>();
        FillColumnsSettings _settings;
        readonly Action<FillColumnsSettings> _use;

        public FillColumnsPage(IList<DataTable> inputs, IList<string> labels,
            FillColumnsSettings settings, Action<FillColumnsSettings> use)
        {
            InitializeComponent();
            _inputs = inputs;
            _settings = settings ?? new FillColumnsSettings { OutputInput = 1 };
            _use = use;
            var columns = new List<string>();
            for (int i = 0; i < inputs.Count; i++)
            {
                OutputSheet.Items.Add("Input " + (i + 1) + " — " + labels[i]);
                columns.Add("Input " + (i + 1) + " ($" + (i + 1) + "_A for column A)\n" +
                    string.Join(", ", ExcelColumns.Labels(inputs[i])));
            }
            InputColumns.Text = string.Join("\n\n", columns);
            OutputSheet.SelectedIndex = _settings.OutputInput <= inputs.Count ? Math.Max(0, _settings.OutputInput - 1) : -1;
            MatchFormula.Text = _settings.MatchFormula ?? "";
            FormulaField.WatchOptional(MatchFormula);
            UpdateMatchHint();
            if (_settings.Rules != null) _rules.AddRange(_settings.Rules);
            RebuildRules();
        }

        void UseInPipeline_Click(object sender, RoutedEventArgs e)
        {
            ErrorText.Visibility = Visibility.Collapsed;
            if (OutputSheet.SelectedIndex < 0)
            {
                Alert("Select a connected output sheet.", "Missing output", MessageBoxImage.Warning);
                return;
            }
            if (!FormulaField.Ok(MatchFormula))
            {
                Alert("Enter a valid match formula.", "Formula error", MessageBoxImage.Warning);
                return;
            }
            string formula = MatchFormula.Text.Trim();
            bool keepLegacy = _settings.OutputInput == 0 && OutputSheet.SelectedIndex == 0 && formula.Length == 0;
            var updated = new FillColumnsSettings
            {
                JointLookups = keepLegacy ? _settings.JointLookups : true,
                OutputInput = keepLegacy ? 0 : OutputSheet.SelectedIndex + 1,
                MatchFormula = formula.Length == 0 ? null : formula,
                Keys = keepLegacy ? _settings.Keys : new List<FillColumnsKey>(),
                MatchAll = keepLegacy ? _settings.MatchAll : true,
                Rules = new List<FillColumnsRule>(_rules)
            };
            _settings = updated;
            UpdateMatchHint();
            _use(updated);
        }

        void UpdateMatchHint()
        {
            MatchHint.Text = _settings.OutputInput == 0 && _settings.Keys != null && _settings.Keys.Count > 0
                ? "This step has saved match keys. Enter a formula to replace them, or enter TRUE to use row order."
                : "For example, $1_ID == $2_ID. The first matching rows are used. Leave blank to use row order.";
        }

        void AddRule_Click(object sender, RoutedEventArgs e)
        {
            EditRule(-1);
        }

        void RebuildRules()
        {
            RulePanel.Children.Clear();
            if (_rules.Count == 0)
            {
                RulePanel.Children.Add(new TextBlock
                {
                    Text = "Rows keep their current values until you add a fill rule.",
                    TextWrapping = TextWrapping.Wrap,
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
                up.IsEnabled = i > 0;
                up.Click += (s, e) => MoveRule(index, -1);
                var remove = new Button { Content = "Remove", Width = 64 };
                remove.Click += (s, e) => RemoveRule(index);
                var buttons = new WrapPanel { Margin = new Thickness(0, 4, 0, 4) };
                buttons.Children.Add(edit);
                buttons.Children.Add(up);
                buttons.Children.Add(remove);
                var row = new StackPanel { Margin = new Thickness(0, 0, 0, 8) };
                row.Children.Add(new TextBlock
                {
                    Text = text,
                    ToolTip = text,
                    TextWrapping = TextWrapping.Wrap,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, 0, 8, 0)
                });
                row.Children.Add(buttons);
                RulePanel.Children.Add(row);
            }
        }

        static string RuleText(FillColumnsRule r)
        {
            return (string.IsNullOrWhiteSpace(r.When) ? "Always" : "When " + r.When) +
                ", set " + r.Target + " to " + r.Value;
        }

        // index -1 adds a new rule.
        void EditRule(int index)
        {
            if (OutputSheet.SelectedIndex < 0)
            {
                Alert("Connect the sheet you want to fill.", "Missing target", MessageBoxImage.Warning);
                return;
            }
            var seed = index >= 0 ? _rules[index] : null;
            var rule = FillRuleWindow.Edit(Window.GetWindow(this), _inputs, OutputSheet.SelectedIndex, seed);
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
            RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
        }
        void Alert(string message, string title, MessageBoxImage icon)
        {
            ErrorText.Text = message;
            ErrorText.Visibility = Visibility.Visible;
            Notices.Show(this, message, title, MessageBoxButton.OK, icon);
        }
    }
}
