using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using OfficeWorkAssistant.Excel;

namespace OfficeWorkAssistant.Features.FormatSheet
{
    public partial class FormatSheetPage : UserControl
    {
        readonly DataTable _input;
        readonly Action<FormatSheetSettings> _use;
        readonly List<ColumnFormat> _formats = new List<ColumnFormat>();

        public FormatSheetPage(DataTable input, string label, FormatSheetSettings settings, Action<FormatSheetSettings> use)
        {
            InitializeComponent();
            _input = input;
            _use = use;
            InputInfo.Text = label + " · Sheet styles apply to the header and data. Column styles apply to data cells.";
            ColumnBox.ItemsSource = ExcelColumns.Labels(input);
            if (ColumnBox.Items.Count > 0)
                ColumnBox.SelectedIndex = 0;
            else
                EditColumnButton.IsEnabled = false;
            foreach (var name in NamedColors.FillNames)
                BandColor.Items.Add(name);

            var s = settings ?? new FormatSheetSettings();
            HeaderBox.IsChecked = s.StyleHeader;
            HeaderPicker.Value = s.Header;
            FreezeBox.IsChecked = s.FreezeHeader;
            FreezeColsBox.Text = s.FreezeColumns.ToString(CultureInfo.InvariantCulture);
            FilterBox.IsChecked = s.AutoFilter;
            FitBox.IsChecked = s.AutoFit;
            SheetPicker.Value = s.SheetCells;
            BorderValuesBox.IsChecked = s.BorderValuesOnly;
            RowsPicker.Value = s.GetDataRowStyle();
            DefaultColumnPicker.Value = s.DefaultColumn;
            DefaultWidthBox.Text = s.DefaultColumnWidth > 0 ? s.DefaultColumnWidth.ToString(CultureInfo.InvariantCulture) : "";
            BandBox.IsChecked = s.BandRows;
            BandColor.Text = NamedColors.NameOf(s.BandFill);
            _formats.AddRange(s.Columns);
            RebuildFormats();
        }

        void RebuildFormats()
        {
            FormatPanel.Children.Clear();
            for (var i = 0; i < _formats.Count; i++)
            {
                var index = i;
                var format = _formats[i];
                var column = ExcelColumns.IndexOf(_input, format.Column);
                var label = column >= 0 ? ExcelColumns.Label(_input, column) : format.Column;
                var edit = new Button { Content = "Edit", Width = 56, Margin = new Thickness(0, 0, 6, 0) };
                edit.IsEnabled = column >= 0;
                edit.Click += (s, e) => EditFormat(index, column);
                var remove = new Button { Content = "Remove", Width = 64 };
                remove.Click += (s, e) => RemoveFormat(index);
                var buttons = new WrapPanel { Margin = new Thickness(0, 4, 0, 4) };
                buttons.Children.Add(edit);
                buttons.Children.Add(remove);
                var row = new StackPanel { Margin = new Thickness(6) };
                row.Children.Add(new TextBlock
                {
                    Text = format.Describe(),
                    TextWrapping = TextWrapping.Wrap,
                    VerticalAlignment = VerticalAlignment.Center,
                    Margin = new Thickness(0, 0, 8, 0)
                });
                row.Children.Add(buttons);
                FormatPanel.Children.Add(new GroupBox
                {
                    Header = new TextBlock { Text = label, TextWrapping = TextWrapping.Wrap },
                    Content = row,
                    Margin = new Thickness(0, 0, 0, 6)
                });
            }
        }

        void EditSelected_Click(object sender, RoutedEventArgs e)
        {
            int column = ColumnBox.SelectedIndex;
            if (column >= 0)
                EditFormat(_formats.FindIndex(f => ExcelColumns.IndexOf(_input, f.Column) == column), column);
        }

        // index -1 adds a new format.
        void EditFormat(int index, int column)
        {
            var seed = index >= 0 ? _formats[index] : null;
            var f = ColumnFormatWindow.Edit(Window.GetWindow(this), _input, column, seed);
            if (f == null)
                return;
            if (index >= 0)
                _formats[index] = f;
            else
                _formats.Add(f);
            RebuildFormats();
            // The pipeline applies this editor by itself after a click inside it; the dialog's clicks never reach it.
            RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
        }

        void RemoveFormat(int index)
        {
            _formats.RemoveAt(index);
            RebuildFormats();
            RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
        }

        void Apply_Click(object sender, RoutedEventArgs e)
        {
            int freeze;
            if (!int.TryParse(FreezeColsBox.Text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out freeze) || freeze < 0)
            {
                ShowError("Columns to freeze must be 0 or a whole number above 0.");
                return;
            }
            double width;
            if (!ColumnFormat.TryWidth(DefaultWidthBox.Text, out width))
            {
                ShowError("Default column width must be a number from 0 to 255.");
                return;
            }
            var s = new FormatSheetSettings
            {
                StyleHeader = HeaderBox.IsChecked == true,
                Header = HeaderPicker.Value,
                FreezeHeader = FreezeBox.IsChecked == true,
                FreezeColumns = freeze,
                AutoFilter = FilterBox.IsChecked == true,
                AutoFit = FitBox.IsChecked == true,
                SheetCells = SheetPicker.Value,
                BorderValuesOnly = BorderValuesBox.IsChecked == true,
                DataRows = RowsPicker.Value,
                DefaultColumn = DefaultColumnPicker.Value,
                DefaultColumnWidth = width,
                BandRows = BandBox.IsChecked == true,
                BandFill = NamedColors.ToHex(BandColor.Text) ?? BandColor.Text.Trim()
            };
            s.Columns.AddRange(_formats);
            try
            {
                FormatSheetWork.Run(_input, s);
            }
            catch (Exception ex)
            {
                ShowError(ex.Message);
                return;
            }
            Status.Text = FormatSheetWork.Summary(s) + ".";
            _use(s);
        }

        void ShowError(string message)
        {
            Status.Text = message;
            Notices.Show(this, message, "Format sheet", MessageBoxButton.OK, MessageBoxImage.Warning);
        }
    }
}
