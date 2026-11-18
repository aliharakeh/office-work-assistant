using System;
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
        readonly ComboBox _target = new ComboBox { IsEditable = true, ToolTip = "Pick a column to fill, or type a new column name." };
        const string EmptyCell = "$Value == \"\"";
        readonly ComboBox _rows = new ComboBox();
        readonly TextBox _when = new TextBox { ToolTip = "Use $Value == \"\" to fill only empty cells." };
        readonly TextBox _value = new TextBox { ToolTip = "Use $ followed by a column letter to copy its value." };
        readonly TextBlock _error = new TextBlock
        {
            Foreground = Brushes.Firebrick,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 8),
            Visibility = Visibility.Collapsed
        };
        FillColumnsRule _result;
        readonly TextBlock _hint = new TextBlock { TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 8, 0, 8) };

        FillRuleWindow(Window owner, IList<DataTable> inputs, int output, FillColumnsRule seed)
        {
            var s = seed ?? new FillColumnsRule();

            Title = seed == null ? "Add fill rule" : "Edit fill rule";
            Owner = owner;
            WindowStartupLocation = owner != null ? WindowStartupLocation.CenterOwner : WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.CanResize;
            Height = 640;
            MinHeight = 300;
            MinWidth = 280;
            Width = 560;

            var names = new List<string>();
            foreach (DataColumn col in inputs[output].Columns)
                names.Add(ExcelFile.Header(col));
            _target.ItemsSource = names;
            _target.Text = s.Target ?? "";
            _when.Text = s.When ?? "";
            _value.Text = s.Value ?? "";
            FormulaField.WatchOptional(_when);
            FormulaField.Watch(_value);

            var variables = new Button { Content = "Variables...", Width = 88, HorizontalAlignment = HorizontalAlignment.Left };
            variables.Click += (sender, e) => ExpressionHelp.Show(this, _hint.Text);
            var ok = new Button { Content = "OK", Width = 80, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
            ok.Click += Ok_Click;
            var cancel = new Button { Content = "Cancel", Width = 80, IsCancel = true };
            var buttons = new WrapPanel { HorizontalAlignment = HorizontalAlignment.Right };
            buttons.Children.Add(ok);
            buttons.Children.Add(cancel);
            var footer = new StackPanel { Margin = new Thickness(0, 4, 0, 0) };

            footer.Children.Add(variables);

            footer.Children.Add(buttons);

            var root = new StackPanel { Margin = new Thickness(12) };
            root.Children.Add(Label("Column to fill — choose an existing header or type a new one"));
            root.Children.Add(_target);
            root.Children.Add(Label("Rows to fill"));
            _rows.Items.Add("Every row");
            _rows.Items.Add("Only empty cells in this column");
            _rows.Items.Add("Rows that meet a condition");
            _rows.SelectedIndex = string.IsNullOrWhiteSpace(_when.Text) ? 0 : _when.Text == EmptyCell ? 1 : 2;
            root.Children.Add(_rows);
            var condition = new StackPanel();
            condition.Children.Add(Label("Condition, for example $Qty > 0"));
            condition.Children.Add(_when);
            condition.Visibility = _rows.SelectedIndex == 2 ? Visibility.Visible : Visibility.Collapsed;
            _rows.SelectionChanged += (sender, e) => condition.Visibility =
                _rows.SelectedIndex == 2 ? Visibility.Visible : Visibility.Collapsed;
            root.Children.Add(condition);
            root.Children.Add(Label("Write this value — quote text, such as \"In stock\", or enter a formula such as $Qty * 2"));
            root.Children.Add(_value);
            _hint.Text = "Use $Name for a column on this row or $Value for the current target cell. Example: when $Qty > 0, set Status to \"In stock\".";
            for (int i = 0; i < inputs.Count; i++)
            {
                int input = i + 1;
                _hint.Text += "\nInput " + input + " ($" + input + "_A for column A): " +
                    string.Join(", ", ExcelColumns.Labels(inputs[i]));
            }
            root.Children.Add(new Expander { Header = "Column and formula reference", Content = _hint, Margin = new Thickness(0, 8, 0, 4) });
            root.Children.Add(_error);
            root.Children.Add(footer);
            Resources.MergedDictionaries.Add(new ResourceDictionary { Source = new Uri("/OfficeWorkAssistant;component/Excel/EditorStyles.xaml", UriKind.Relative) });
            Content = new ScrollViewer { Content = root, VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled };
        }

        // Returns the rule the user saved, or null when the window was cancelled.
        public static FillColumnsRule Edit(Window owner, IList<DataTable> inputs, int output, FillColumnsRule seed)
        {
            var win = new FillRuleWindow(owner, inputs, output, seed);
            return win.ShowDialog() == true ? win._result : null;
        }

        static TextBlock Label(string text)
        {
            return new TextBlock { Text = text, TextWrapping = TextWrapping.Wrap, FontSize = 11, Foreground = Brushes.DimGray, Margin = new Thickness(0, 8, 0, 2) };
        }

        void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_target.Text))
            {
                _error.Text = "Pick a target column, or type a new column name.";
                _error.Visibility = Visibility.Visible;
                return;
            }
            if (string.IsNullOrWhiteSpace(_value.Text))
            {
                _error.Text = "Enter the text or formula to write into the cell.";
                _error.Visibility = Visibility.Visible;
                return;
            }
            if ((_rows.SelectedIndex == 2 && (string.IsNullOrWhiteSpace(_when.Text) || !FormulaField.Ok(_when))) || !FormulaField.Ok(_value))
            {
                _error.Text = "Enter a valid condition and value for this rule.";
                _error.Visibility = Visibility.Visible;
                return;
            }
            _result = new FillColumnsRule { Target = _target.Text.Trim(), When = _rows.SelectedIndex == 2 ? _when.Text : _rows.SelectedIndex == 1 ? EmptyCell : "", Value = _value.Text };
            DialogResult = true;
        }
    }
}
