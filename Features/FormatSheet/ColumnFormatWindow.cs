using System;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using OfficeWorkAssistant.Excel;

namespace OfficeWorkAssistant.Features.FormatSheet
{
    // Creates or edits one column format of the Format sheet step. The step lists its formats; this window holds the form.
    sealed class ColumnFormatWindow : Window
    {
        static readonly string[] Formats =
        {
            "", "0", "0.00", "#,##0", "#,##0.00", "0%", "0.00%", "dd/MM/yyyy", "dd/MM/yyyy HH:mm", "@"
        };

        readonly ComboBox _column = new ComboBox();
        readonly TextBox _width = new TextBox { ToolTip = "About one character per unit. Blank or 0 = leave it." };
        readonly ComboBox _number = new ComboBox { IsEditable = true, ToolTip = "An Excel number format, e.g. #,##0.00 or dd/MM/yyyy." };
        readonly ComboBox _align = new ComboBox();
        readonly CheckBox _wrap = new CheckBox { Content = "Wrap text", VerticalAlignment = VerticalAlignment.Center };
        readonly TextBlock _error = new TextBlock
        {
            Foreground = Brushes.Firebrick,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 8),
            Visibility = Visibility.Collapsed
        };
        ColumnFormat _result;

        ColumnFormatWindow(Window owner, DataTable input, ColumnFormat seed)
        {
            var s = seed ?? new ColumnFormat();

            Title = seed == null ? "Add column format" : "Edit column format";
            Owner = owner;
            WindowStartupLocation = owner != null ? WindowStartupLocation.CenterOwner : WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            SizeToContent = SizeToContent.Height;
            Width = 520;

            _column.ItemsSource = ExcelColumns.Labels(input);
            _column.SelectedIndex = ExcelColumns.IndexOf(input, s.Column);
            if (_column.SelectedIndex < 0 && input != null && input.Columns.Count > 0)
                _column.SelectedIndex = 0;
            _width.Text = s.Width > 0 ? s.Width.ToString(CultureInfo.InvariantCulture) : "";
            foreach (var f in Formats)
                _number.Items.Add(f);
            _number.Text = s.NumberFormat ?? "";
            _align.Items.Add("As Excel does");
            _align.Items.Add("Left");
            _align.Items.Add("Center");
            _align.Items.Add("Right");
            _align.SelectedIndex = Math.Max(0, Math.Min((int)s.Align, _align.Items.Count - 1));
            _wrap.IsChecked = s.Wrap;

            var widthRow = new Grid { Margin = new Thickness(0, 0, 0, 8) };
            widthRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(60) });
            widthRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            widthRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            widthRow.Children.Add(_width);
            var formatLabel = new TextBlock { Text = "Format ", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(8, 0, 6, 0) };
            Grid.SetColumn(formatLabel, 1);
            widthRow.Children.Add(formatLabel);
            Grid.SetColumn(_number, 2);
            widthRow.Children.Add(_number);
            var wrapRow = new Grid { Margin = new Thickness(0, 0, 0, 8) };
            wrapRow.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            wrapRow.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            wrapRow.Children.Add(_align);
            Grid.SetColumn(_wrap, 1);
            wrapRow.Children.Add(_wrap);

            var ok = new Button { Content = "OK", Width = 80, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
            ok.Click += Ok_Click;
            var cancel = new Button { Content = "Cancel", Width = 80, IsCancel = true };
            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            buttons.Children.Add(ok);
            buttons.Children.Add(cancel);

            var root = new StackPanel { Margin = new Thickness(12) };
            root.Children.Add(Label("Column"));
            root.Children.Add(_column);
            root.Children.Add(Label("Width and number format"));
            root.Children.Add(widthRow);
            root.Children.Add(Label("Align"));
            root.Children.Add(wrapRow);
            root.Children.Add(_error);
            root.Children.Add(buttons);
            Content = root;
        }

        // Returns the format the user saved, or null when the window was cancelled.
        public static ColumnFormat Edit(Window owner, DataTable input, ColumnFormat seed)
        {
            var win = new ColumnFormatWindow(owner, input, seed);
            return win.ShowDialog() == true ? win._result : null;
        }

        static TextBlock Label(string text)
        {
            return new TextBlock { Text = text, FontSize = 11, Foreground = Brushes.DimGray, Margin = new Thickness(0, 8, 0, 2) };
        }

        void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (_column.SelectedIndex < 0)
            {
                Fail("Choose a column first.");
                return;
            }
            double width = 0;
            string w = _width.Text.Trim();
            if (w.Length > 0 && (!double.TryParse(w, NumberStyles.Float, CultureInfo.InvariantCulture, out width) || width < 0 || width > 255))
            {
                Fail("The width must be a number from 0 to 255.");
                _width.Focus();
                return;
            }
            _result = new ColumnFormat
            {
                Column = ExcelColumns.Letter(_column.SelectedIndex),
                Width = width,
                NumberFormat = (_number.Text ?? "").Trim(),
                Align = (CellAlign)Math.Max(0, _align.SelectedIndex),
                Wrap = _wrap.IsChecked == true
            };
            DialogResult = true;
        }

        void Fail(string message)
        {
            _error.Text = message;
            _error.Visibility = Visibility.Visible;
        }
    }
}
