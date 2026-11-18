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
        readonly string _column;
        readonly TextBox _width = new TextBox { ToolTip = "About one character per unit. Blank or 0 uses the default width." };
        readonly CellStylePicker _picker = new CellStylePicker { BasicOptions = true };
        readonly TextBlock _error = new TextBlock
        {
            Foreground = Brushes.Firebrick,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 0, 0, 8),
            Visibility = Visibility.Collapsed
        };
        ColumnFormat _result;

        ColumnFormatWindow(Window owner, DataTable input, int column, ColumnFormat seed)
        {
            var s = seed ?? new ColumnFormat();

            Title = "Column style";
            Owner = owner;
            WindowStartupLocation = owner != null ? WindowStartupLocation.CenterOwner : WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.CanResize;
            Height = 640;
            MinHeight = 300;
            MinWidth = 280;
            Width = 520;

            _column = ExcelColumns.Letter(column);
            _width.Text = s.Width > 0 ? s.Width.ToString(CultureInfo.InvariantCulture) : "";
            _width.Margin = new Thickness(0, 0, 0, 8);
            _picker.Value = s.GetStyle();

            var ok = new Button { Content = "OK", Width = 80, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
            ok.Click += Ok_Click;
            var cancel = new Button { Content = "Cancel", Width = 80, IsCancel = true };
            var buttons = new WrapPanel { HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 8, 0, 0) };
            buttons.Children.Add(ok);
            buttons.Children.Add(cancel);

            var root = new StackPanel { Margin = new Thickness(12) };
            root.Children.Add(new TextBlock { Text = ExcelColumns.Label(input, column), FontWeight = FontWeights.Bold, TextWrapping = TextWrapping.Wrap });
            root.Children.Add(Label("Column width (blank = use default)"));
            root.Children.Add(_width);
            root.Children.Add(_picker);
            root.Children.Add(_error);
            root.Children.Add(buttons);
            Resources.MergedDictionaries.Add(new ResourceDictionary { Source = new Uri("/OfficeWorkAssistant;component/Excel/EditorStyles.xaml", UriKind.Relative) });
            Content = new ScrollViewer { Content = root, VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled };
        }

        // Returns the format the user saved, or null when the window was cancelled.
        public static ColumnFormat Edit(Window owner, DataTable input, int column, ColumnFormat seed)
        {
            var win = new ColumnFormatWindow(owner, input, column, seed);
            return win.ShowDialog() == true ? win._result : null;
        }

        static TextBlock Label(string text)
        {
            return new TextBlock { Text = text, FontSize = 11, Foreground = Brushes.DimGray, Margin = new Thickness(0, 8, 0, 2) };
        }

        void Ok_Click(object sender, RoutedEventArgs e)
        {
            double width;
            if (!ColumnFormat.TryWidth(_width.Text, out width))
            {
                Fail("The width must be a number from 0 to 255.");
                _width.Focus();
                return;
            }
            var style = _picker.Value;
            var error = style.Check();
            if (error != null)
            {
                Fail(error);
                return;
            }
            _result = new ColumnFormat
            {
                Column = _column,
                Width = width,
                Style = style
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
