using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using OfficeWorkAssistant.Excel;

namespace OfficeWorkAssistant.Features.FormatSheet
{
    public partial class FormatSheetPage : UserControl
    {
        static readonly string[] Formats =
        {
            "", "0", "0.00", "#,##0", "#,##0.00", "0%", "0.00%", "dd/MM/yyyy", "dd/MM/yyyy HH:mm", "@"
        };

        readonly DataTable _input;
        readonly Action<FormatSheetSettings> _use;
        readonly List<ColumnFormat> _formats = new List<ColumnFormat>();
        int _editing = -1;

        public FormatSheetPage(DataTable input, string label, FormatSheetSettings settings, Action<FormatSheetSettings> use)
        {
            InitializeComponent();
            _input = input;
            _use = use;
            InputInfo.Text = label + " - the look Save writes. Highlight colours from an earlier step stay on top.";
            ColumnBox.ItemsSource = ExcelColumns.Labels(input);
            foreach (var f in Formats)
                NumberBox.Items.Add(f);
            foreach (var name in NamedColors.FillNames)
                BandColor.Items.Add(name);

            var s = settings ?? new FormatSheetSettings();
            HeaderBox.IsChecked = s.StyleHeader;
            HeaderPicker.Value = s.Header;
            FreezeBox.IsChecked = s.FreezeHeader;
            FreezeColsBox.Text = s.FreezeColumns.ToString(CultureInfo.InvariantCulture);
            FilterBox.IsChecked = s.AutoFilter;
            FitBox.IsChecked = s.AutoFit;
            BordersBox.IsChecked = s.Borders;
            BandBox.IsChecked = s.BandRows;
            BandColor.Text = NamedColors.NameOf(s.BandFill);
            _formats.AddRange(s.Columns);
            New_Click(null, null);
        }

        void ShowFormats()
        {
            FormatList.ItemsSource = null;
            FormatList.ItemsSource = _formats;
        }

        void FormatList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int index = FormatList.SelectedIndex;
            if (index < 0 || index >= _formats.Count)
                return;
            _editing = index;
            var f = _formats[index];
            ColumnBox.SelectedIndex = ExcelColumns.IndexOf(_input, f.Column);
            WidthBox.Text = f.Width > 0 ? f.Width.ToString(CultureInfo.InvariantCulture) : "";
            NumberBox.Text = f.NumberFormat ?? "";
            AlignBox.SelectedIndex = (int)f.Align;
            WrapBox.IsChecked = f.Wrap;
            AddBtn.Content = "Update format";
        }

        void Add_Click(object sender, RoutedEventArgs e)
        {
            if (ColumnBox.SelectedIndex < 0)
            {
                Status.Text = "Choose a column first.";
                return;
            }
            double width = 0;
            string w = WidthBox.Text.Trim();
            if (w.Length > 0 && (!double.TryParse(w, NumberStyles.Float, CultureInfo.InvariantCulture, out width) || width < 0 || width > 255))
            {
                Status.Text = "The width must be a number from 0 to 255.";
                WidthBox.Focus();
                return;
            }
            var f = new ColumnFormat
            {
                Column = ExcelColumns.Letter(ColumnBox.SelectedIndex),
                Width = width,
                NumberFormat = (NumberBox.Text ?? "").Trim(),
                Align = (CellAlign)Math.Max(0, AlignBox.SelectedIndex),
                Wrap = WrapBox.IsChecked == true
            };
            if (_editing >= 0 && _editing < _formats.Count)
                _formats[_editing] = f;
            else
                _formats.Add(f);
            New_Click(null, null);
        }

        void Remove_Click(object sender, RoutedEventArgs e)
        {
            int index = FormatList.SelectedIndex;
            if (index < 0 || index >= _formats.Count)
            {
                Status.Text = "Select a column format first.";
                return;
            }
            _formats.RemoveAt(index);
            New_Click(null, null);
        }

        void New_Click(object sender, RoutedEventArgs e)
        {
            _editing = -1;
            AddBtn.Content = "Add format";
            ShowFormats();
            ColumnBox.SelectedIndex = _input.Columns.Count > 0 ? 0 : -1;
            WidthBox.Text = "";
            NumberBox.Text = "";
            AlignBox.SelectedIndex = 0;
            WrapBox.IsChecked = false;
        }

        void Apply_Click(object sender, RoutedEventArgs e)
        {
            int freeze;
            if (!int.TryParse(FreezeColsBox.Text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out freeze) || freeze < 0)
            {
                Status.Text = "Columns to freeze must be 0 or a whole number above 0.";
                FreezeColsBox.Focus();
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
                Borders = BordersBox.IsChecked == true,
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
                MessageBox.Show(Window.GetWindow(this), ex.Message, "Format sheet", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            Status.Text = FormatSheetWork.Summary(s) + ".";
            _use(s);
        }
    }
}
