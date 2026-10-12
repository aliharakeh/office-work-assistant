using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using OfficeWorkAssistant.Excel;
using OfficeWorkAssistant.Expressions;

namespace OfficeWorkAssistant.Features.FillColumns
{
    // Creates or edits one fill rule of the Fill columns step. The step lists its rules; this window holds the form.
    sealed class FillRuleWindow : Window
    {
        readonly ComboBox _target = new ComboBox { IsEditable = true, ToolTip = "Pick a B column to fill, or type a new column name." };
        readonly TextBox _when = new TextBox { ToolTip = "Optional condition, e.g. $Matched && $Value == \"\". Blank = always." };
        readonly TextBox _value = new TextBox { ToolTip = "Value formula, e.g. $A_Price or IF($A_Qty > 0, $A_Price, $Value)." };
        readonly TextBlock _error = new TextBlock
        {
            Foreground = Brushes.Firebrick,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 8),
            Visibility = Visibility.Collapsed
        };
        FillColumnsRule _result;

        FillRuleWindow(Window owner, DataTable tableB, FillColumnsRule seed)
        {
            var s = seed ?? new FillColumnsRule();

            Title = seed == null ? "Add fill rule" : "Edit fill rule";
            Owner = owner;
            WindowStartupLocation = owner != null ? WindowStartupLocation.CenterOwner : WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            SizeToContent = SizeToContent.Height;
            Width = 560;

            var names = new List<string>();
            if (tableB != null)
            {
                foreach (DataColumn col in tableB.Columns)
                    names.Add(ExcelFile.Header(col));
            }
            _target.ItemsSource = names;
            _target.Text = s.Target ?? "";
            _when.Text = s.When ?? "$Matched";
            _value.Text = s.Value ?? "$Value";
            FormulaField.WatchOptional(_when);
            FormulaField.Watch(_value);

            var variables = new Button { Content = "Variables...", Width = 88, HorizontalAlignment = HorizontalAlignment.Left };
            variables.Click += (sender, e) => ExpressionHelp.Show(this,
                "Example: Target Price, When $Matched && $Value == \"\", Value $A_Price. Add a second Price rule below it with Value $Value * 1.1 to change the rows the first rule skipped.",
                new[]
                {
                    new VariableHelp { Name = "$A_Name", Description = "Column Name (or letter, $A_C) of the matched A row. Blank when no A row matched.", Example = "" },
                    new VariableHelp { Name = "$B_Name", Description = "Column Name (or letter) of the B row being filled, including results of earlier rules.", Example = "" },
                    new VariableHelp { Name = "$Name", Description = "Same as $B_Name. A name B does not have falls back to A.", Example = "" },
                    new VariableHelp { Name = "$Value", Description = "The target cell's current value.", Example = "" },
                    new VariableHelp { Name = "$Matched", Description = "true when an A row matched this B row.", Example = "" }
                });
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

            var root = new StackPanel { Margin = new Thickness(12) };
            root.Children.Add(Label("Target column"));
            root.Children.Add(_target);
            root.Children.Add(Label("When (optional)"));
            root.Children.Add(_when);
            root.Children.Add(Label("Value"));
            root.Children.Add(_value);
            root.Children.Add(new TextBlock
            {
                Text = "$A_Name reads A, $B_Name or $Name reads B, $Value is the target cell, $Matched is true when A matched.",
                FontSize = 11,
                Foreground = Brushes.DimGray,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 8, 0, 8)
            });
            root.Children.Add(_error);
            root.Children.Add(footer);
            Content = root;
        }

        // Returns the rule the user saved, or null when the window was cancelled.
        public static FillColumnsRule Edit(Window owner, DataTable tableB, FillColumnsRule seed)
        {
            var win = new FillRuleWindow(owner, tableB, seed);
            return win.ShowDialog() == true ? win._result : null;
        }

        static TextBlock Label(string text)
        {
            return new TextBlock { Text = text, FontSize = 11, Foreground = Brushes.DimGray, Margin = new Thickness(0, 8, 0, 2) };
        }

        void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_target.Text))
            {
                _error.Text = "Pick a target column, or type a new column name.";
                _error.Visibility = Visibility.Visible;
                return;
            }
            if (!FormulaField.Ok(_when) || !FormulaField.Ok(_value))
            {
                _error.Text = "A formula is not valid. Fix the one shown in red.";
                _error.Visibility = Visibility.Visible;
                return;
            }
            _result = new FillColumnsRule { Target = _target.Text.Trim(), When = _when.Text, Value = _value.Text };
            DialogResult = true;
        }
    }
}
