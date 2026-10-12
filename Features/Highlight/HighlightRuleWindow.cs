using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.Highlight
{
    // Creates or edits one rule of the Highlight step. The step lists its rules; this window holds the form.
    sealed class HighlightRuleWindow : Window
    {
        static readonly string[] Ops = { "==", "!=", ">", ">=", "<", "<=", "contains", "startsWith", "endsWith", "isEmpty", "isNotEmpty" };

        readonly DataTable _input;
        readonly int _number;
        readonly ComboBox _target = new ComboBox();
        readonly StackPanel _columnsPanel = new StackPanel { Margin = new Thickness(90, 0, 0, 4) };
        readonly TextBox _condition = new TextBox { Margin = new Thickness(0, 0, 6, 0) };
        readonly TextBlock _hint = new TextBlock { Foreground = Brushes.Gray, Margin = new Thickness(90, 0, 0, 4), TextWrapping = TextWrapping.Wrap };
        readonly ComboBox _buildColumn = new ComboBox { Margin = new Thickness(0, 0, 0, 4) };
        readonly ComboBox _buildOp = new ComboBox { Margin = new Thickness(0, 0, 6, 0) };
        readonly TextBox _buildValue = new TextBox { Margin = new Thickness(0, 0, 6, 0), ToolTip = "A number, text, or a built-in such as $Today." };
        readonly CellStylePicker _style = new CellStylePicker();
        readonly CheckBox _stop = new CheckBox { Content = "When this rule matches, later rules skip that row (or cell)", Margin = new Thickness(0, 6, 0, 0) };
        readonly TextBlock _error = new TextBlock
        {
            Foreground = Brushes.Firebrick,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 8),
            Visibility = Visibility.Collapsed
        };
        List<Pick> _picks = new List<Pick>();
        HighlightRule _result;

        HighlightRuleWindow(Window owner, DataTable input, HighlightRule seed, int number)
        {
            _input = input;
            _number = number;
            var s = seed ?? NewRule();

            Title = seed == null ? "Add highlight rule" : "Edit highlight rule";
            Owner = owner;
            WindowStartupLocation = owner != null ? WindowStartupLocation.CenterOwner : WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            SizeToContent = SizeToContent.Height;
            Width = 620;

            FormulaField.Watch(_condition);
            _target.Items.Add("The whole row, when the condition is true");
            _target.Items.Add("Each cell in the columns below, tested on its own ($Value)");
            _target.Items.Add("The columns below, in rows where the condition is true");
            _target.SelectedIndex = Math.Max(0, Math.Min((int)s.Target, 2));
            _target.SelectionChanged += (sender, e) => UpdateTarget();

            var columnItems = new ItemsControl { ItemTemplate = PickTemplate() };
            _columnsPanel.Children.Add(new ScrollViewer
            {
                MaxHeight = 90,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled,
                Content = columnItems
            });

            _buildOp.SelectedIndex = 0;
            foreach (var op in Ops)
                _buildOp.Items.Add(op);
            var use = new Button { Content = "Use", Width = 50 };
            use.Click += Build_Click;
            var buildGrid = new Grid();
            buildGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(110) });
            buildGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            buildGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            Grid.SetColumn(_buildOp, 0);
            Grid.SetColumn(_buildValue, 1);
            Grid.SetColumn(use, 2);
            buildGrid.Children.Add(_buildOp);
            buildGrid.Children.Add(_buildValue);
            buildGrid.Children.Add(use);
            var build = new StackPanel { Margin = new Thickness(0, 4, 0, 0) };
            build.Children.Add(_buildColumn);
            build.Children.Add(buildGrid);

            var help = new Button { Content = "Help...", Width = 60 };
            help.Click += (sender, e) => ExpressionHelp.Show(this,
                "A condition is a formula that is true or false. $Qty or $A is a column of the row (A = 1st); $Value is the cell itself when each cell is tested on its own. " +
                "Compare with == != > >= < <=, join with && (and) || (or), and use the functions below.");
            var conditionRow = new Grid { Margin = new Thickness(0, 0, 0, 2) };
            conditionRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(90) });
            conditionRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            conditionRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            var conditionLabel = new TextBlock { Text = "Condition", VerticalAlignment = VerticalAlignment.Center };
            Grid.SetColumn(_condition, 1);
            Grid.SetColumn(help, 2);
            conditionRow.Children.Add(conditionLabel);
            conditionRow.Children.Add(_condition);
            conditionRow.Children.Add(help);

            var targetRow = new Grid { Margin = new Thickness(0, 0, 0, 4) };
            targetRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(90) });
            targetRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            var targetLabel = new TextBlock { Text = "Style", VerticalAlignment = VerticalAlignment.Center };
            Grid.SetColumn(_target, 1);
            targetRow.Children.Add(targetLabel);
            targetRow.Children.Add(_target);

            var ok = new Button { Content = "OK", Width = 80, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
            ok.Click += Ok_Click;
            var cancel = new Button { Content = "Cancel", Width = 80, IsCancel = true };
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 8, 0, 0) };
            buttons.Children.Add(ok);
            buttons.Children.Add(cancel);

            var root = new StackPanel { Margin = new Thickness(12) };
            root.Children.Add(targetRow);
            root.Children.Add(_columnsPanel);
            root.Children.Add(conditionRow);
            root.Children.Add(_hint);
            root.Children.Add(new Expander { Header = "Build a condition", Margin = new Thickness(90, 0, 0, 6), Content = build });
            root.Children.Add(new TextBlock { Text = "Look", FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 4, 0, 4) });
            root.Children.Add(_style);
            root.Children.Add(_stop);
            root.Children.Add(_error);
            root.Children.Add(buttons);
            Content = root;

            LoadRule(s, columnItems);
            UpdateTarget();
        }

        static DataTemplate PickTemplate()
        {
            var template = new DataTemplate();
            var box = new FrameworkElementFactory(typeof(CheckBox));
            box.SetBinding(System.Windows.Controls.Primitives.ToggleButton.IsCheckedProperty, new System.Windows.Data.Binding("Include") { Mode = System.Windows.Data.BindingMode.TwoWay });
            box.SetBinding(ContentControl.ContentProperty, new System.Windows.Data.Binding("Label"));
            box.SetValue(FrameworkElement.MarginProperty, new Thickness(0, 0, 12, 4));
            template.VisualTree = box;
            return template;
        }

        static HighlightRule NewRule()
        {
            return new HighlightRule { Style = new CellStyle { Fill = NamedColors.ToHex("Light red") } };
        }

        void LoadRule(HighlightRule rule, ItemsControl columnItems)
        {
            var on = new HashSet<int>();
            foreach (var c in rule.Columns ?? new List<string>())
            {
                int i = ExcelColumns.IndexOf(_input, c);
                if (i >= 0)
                    on.Add(i);
            }
            _picks = new List<Pick>();
            for (var i = 0; i < _input.Columns.Count; i++)
                _picks.Add(new Pick { Index = i, Label = ExcelColumns.Label(_input, i), Include = on.Contains(i) });
            columnItems.ItemsSource = _picks;
            _condition.Text = rule.Condition ?? "";
            _style.Value = rule.Style;
            _stop.IsChecked = rule.StopIfTrue;
        }

        // Returns the rule the user saved, or null when the window was cancelled.
        public static HighlightRule Edit(Window owner, DataTable input, HighlightRule seed, int number)
        {
            var win = new HighlightRuleWindow(owner, input, seed, number);
            return win.ShowDialog() == true ? win._result : null;
        }

        void UpdateTarget()
        {
            bool cells = _target.SelectedIndex == (int)HighlightTarget.Cells;
            _columnsPanel.Visibility = _target.SelectedIndex == (int)HighlightTarget.Row ? Visibility.Collapsed : Visibility.Visible;
            _hint.Text = cells
                ? "Checked for each picked cell. $Value is that cell: $Value < 0, $Value > $Today, ISBLANK($Value). Other columns of the row work too."
                : "Checked for each row: $Status == \"Late\", $Qty > 10 && $Price < 5, $Due < $Today, CONTAINS($Name, \"Ltd\").";

            var items = new List<string>();
            if (cells)
                items.Add("This cell ($Value)");
            items.AddRange(ExcelColumns.Labels(_input));
            int keep = _buildColumn.SelectedIndex;
            _buildColumn.ItemsSource = items;
            _buildColumn.SelectedIndex = items.Count == 0 ? -1 : Math.Max(0, Math.Min(keep, items.Count - 1));
        }

        void Build_Click(object sender, RoutedEventArgs e)
        {
            int picked = _buildColumn.SelectedIndex;
            if (picked < 0)
                return;
            bool cells = _target.SelectedIndex == (int)HighlightTarget.Cells;
            string name;
            if (cells && picked == 0)
                name = "$Value";
            else
                name = "$" + ExcelColumns.FormulaName(_input, cells ? picked - 1 : picked);

            string op = Ops[Math.Max(0, _buildOp.SelectedIndex)];
            string value = _buildValue.Text.Trim();
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
            _condition.Text = text;
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

        HighlightRule EditorRule()
        {
            var rule = new HighlightRule
            {
                Condition = _condition.Text.Trim(),
                Target = (HighlightTarget)Math.Max(0, _target.SelectedIndex),
                Style = _style.Value,
                StopIfTrue = _stop.IsChecked == true
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

        void Ok_Click(object sender, RoutedEventArgs e)
        {
            var rule = EditorRule();
            var error = HighlightWork.Validate(rule, _number);
            if (error != null)
            {
                _error.Text = error;
                _error.Visibility = Visibility.Visible;
                return;
            }
            _result = rule;
            DialogResult = true;
        }

        public sealed class Pick
        {
            public int Index { get; set; }
            public string Label { get; set; }
            public bool Include { get; set; }
        }
    }
}
