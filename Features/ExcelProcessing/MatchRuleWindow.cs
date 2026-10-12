using System;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.ExcelProcessing
{
    // Creates or edits one match rule of the Compare step. The step lists its rules; this window holds the form.
    sealed class MatchRuleWindow : Window
    {
        readonly SidePart _a = new SidePart();
        readonly SidePart _b = new SidePart();
        readonly ComboBox _op = new ComboBox { Width = 72, Margin = new Thickness(8, 0, 8, 0), VerticalAlignment = VerticalAlignment.Center };
        readonly TextBlock _error = new TextBlock
        {
            Foreground = Brushes.Firebrick,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 8),
            Visibility = Visibility.Collapsed
        };
        ExcelProcessingMatch _result;

        MatchRuleWindow(Window owner, DataTable a, DataTable b, ExcelProcessingMatch seed)
        {
            var s = seed ?? new ExcelProcessingMatch();

            Title = seed == null ? "Add match rule" : "Edit match rule";
            Owner = owner;
            WindowStartupLocation = owner != null ? WindowStartupLocation.CenterOwner : WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            SizeToContent = SizeToContent.Height;
            Width = 640;

            var grid = new Grid { Margin = new Thickness(0, 0, 0, 8) };
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var sideA = SideBox("File A", a, s.ColumnA, s.SplitA, s.ExpressionA ?? "$Value", _a);
            var sideB = SideBox("File B", b, s.ColumnB, s.SplitB, s.ExpressionB ?? "$Value", _b);
            Grid.SetColumn(sideA, 0);
            Grid.SetColumn(_op, 1);
            Grid.SetColumn(sideB, 2);
            grid.Children.Add(sideA);
            grid.Children.Add(_op);
            grid.Children.Add(sideB);

            foreach (var o in new[] { "==", "!=", ">", ">=", "<", "<=" })
                _op.Items.Add(o);
            _op.SelectedItem = string.IsNullOrWhiteSpace(s.Operator) ? "==" : s.Operator;
            if (_op.SelectedIndex < 0)
                _op.SelectedIndex = 0;

            var variables = new Button { Content = "Variables...", Width = 88, HorizontalAlignment = HorizontalAlignment.Left };
            variables.Click += (sender, e) => ExpressionHelp.Show(this,
                "Type $Today in a formula. $Value is the chosen column on that side. After Split, $1 $2 are parts of it. IF(), FIRSTWORD() and LASTWORD() work inside the formula.");
            var ok = new Button { Content = "OK", Width = 80, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
            ok.Click += Ok_Click;
            var cancel = new Button { Content = "Cancel", Width = 80, IsCancel = true };
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            buttons.Children.Add(ok);
            buttons.Children.Add(cancel);
            var footer = new Grid { Margin = new Thickness(0, 4, 0, 0) };
            footer.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            footer.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            footer.Children.Add(variables);
            Grid.SetColumn(buttons, 1);
            footer.Children.Add(buttons);

            var hint = new TextBlock
            {
                Text = "$Value is the chosen column. Split it, then $1 $2 in the formula.",
                FontSize = 11,
                Foreground = Brushes.DimGray,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            };

            var root = new StackPanel { Margin = new Thickness(12) };
            root.Children.Add(grid);
            root.Children.Add(hint);
            root.Children.Add(_error);
            root.Children.Add(footer);
            Content = root;
        }

        // Returns the rule the user saved, or null when the window was cancelled.
        public static ExcelProcessingMatch Edit(Window owner, DataTable a, DataTable b, ExcelProcessingMatch seed)
        {
            var win = new MatchRuleWindow(owner, a, b, seed);
            return win.ShowDialog() == true ? win._result : null;
        }

        static GroupBox SideBox(string header, DataTable table, string column, string split, string formula, SidePart part)
        {
            part.Column = new ComboBox { Margin = new Thickness(0, 0, 0, 8) };
            FillColumns(part.Column, table, column);

            part.Custom = new TextBox { Width = 28, MaxLength = 1, IsEnabled = false, Margin = new Thickness(6, 0, 0, 0) };
            part.Split = NewSplitBox();
            ApplySplit(part.Split, part.Custom, split);
            part.Split.SelectionChanged += (s, e) => part.Custom.IsEnabled = part.Split.SelectedIndex == 4;

            part.Formula = new TextBox { Text = formula, ToolTip = "Formula. $Value is the column." };
            FormulaField.Watch(part.Formula);

            var splitRow = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 8) };
            splitRow.Children.Add(part.Split);
            splitRow.Children.Add(part.Custom);

            var stack = new StackPanel();
            stack.Children.Add(Label("Column"));
            stack.Children.Add(part.Column);
            stack.Children.Add(Label("Split on"));
            stack.Children.Add(splitRow);
            stack.Children.Add(Label("Formula"));
            stack.Children.Add(part.Formula);
            return new GroupBox { Header = header, Padding = new Thickness(8), Content = stack };
        }

        static TextBlock Label(string text)
        {
            return new TextBlock { Text = text, FontSize = 11, Foreground = Brushes.DimGray, Margin = new Thickness(0, 0, 0, 2) };
        }

        static void FillColumns(ComboBox box, DataTable table, string keep)
        {
            var names = table == null ? new string[0] : ExcelProcessingPage.ColumnNames(table);
            box.ItemsSource = names;
            var at = -1;
            for (var i = 0; keep != null && i < names.Length && at < 0; i++)
            {
                if (string.Equals(names[i], keep, StringComparison.OrdinalIgnoreCase))
                    at = i;
            }
            if (at < 0 && names.Length > 0)
                at = 0;
            box.SelectedIndex = at;
        }

        static ComboBox NewSplitBox()
        {
            var box = new ComboBox { Width = 110, ToolTip = "Split the column into $1 $2" };
            box.Items.Add("None");
            box.Items.Add("Space");
            box.Items.Add("Underscore");
            box.Items.Add("Dash");
            box.Items.Add("Custom");
            box.SelectedIndex = 0;
            return box;
        }

        static void ApplySplit(ComboBox sep, TextBox custom, string stored)
        {
            if (stored == " ")
                sep.SelectedIndex = 1;
            else if (stored == "_")
                sep.SelectedIndex = 2;
            else if (stored == "-")
                sep.SelectedIndex = 3;
            else if (!string.IsNullOrEmpty(stored))
            {
                sep.SelectedIndex = 4;
                custom.Text = stored.Substring(0, 1);
                custom.IsEnabled = true;
            }
            else
                sep.SelectedIndex = 0;
        }

        static string ReadSplit(ComboBox sep, TextBox custom)
        {
            switch (sep.SelectedIndex)
            {
                case 1: return " ";
                case 2: return "_";
                case 3: return "-";
                case 4: return custom.Text.Length > 0 ? custom.Text.Substring(0, 1) : "";
                default: return "";
            }
        }

        void Ok_Click(object sender, RoutedEventArgs e)
        {
            var problem = Problem();
            if (problem != null)
            {
                _error.Text = problem;
                _error.Visibility = Visibility.Visible;
                return;
            }
            _result = new ExcelProcessingMatch
            {
                ColumnA = _a.Column.SelectedItem as string,
                ColumnB = _b.Column.SelectedItem as string,
                ExpressionA = _a.Formula.Text,
                ExpressionB = _b.Formula.Text,
                Operator = _op.SelectedItem as string,
                SplitA = ReadSplit(_a.Split, _a.Custom),
                SplitB = ReadSplit(_b.Split, _b.Custom)
            };
            DialogResult = true;
        }

        string Problem()
        {
            if (_a.Column.SelectedItem == null || _b.Column.SelectedItem == null)
                return "Both files need a column.";
            if (!FormulaField.Ok(_a.Formula) || !FormulaField.Ok(_b.Formula))
                return "A formula is not valid. Fix the one shown in red.";
            return null;
        }

        sealed class SidePart
        {
            public ComboBox Column;
            public ComboBox Split;
            public TextBox Custom;
            public TextBox Formula;
        }
    }
}
