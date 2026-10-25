using System;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using OfficeWorkAssistant.Excel;

namespace OfficeWorkAssistant.Features.SaveText
{
    public partial class SaveTextPage : UserControl
    {
        readonly DataTable _input;
        readonly Action<SaveTextSettings> _use;

        public SaveTextPage(DataTable input, string label, SaveTextSettings settings, Action<SaveTextSettings> use)
        {
            InitializeComponent();
            _input = input;
            _use = use;
            InputInfo.Text = label;
            ColumnBox.ItemsSource = ExcelColumns.Labels(input);
            ValueColumnBox.ItemsSource = ExcelColumns.Labels(input);
            var s = settings ?? new SaveTextSettings();
            int column = ExcelColumns.IndexOf(input, s.Column);
            int value = ExcelColumns.IndexOf(input, s.ValueColumn);
            if (settings == null)
            {
                var preferred = ExcelFile.FindColumn(input, "OCR text") ?? ExcelFile.FindColumn(input, "Key");
                column = preferred != null ? preferred.Ordinal : (input.Columns.Count > 0 ? 0 : -1);
                var mapValue = ExcelFile.FindColumn(input, "Value");
                value = mapValue != null ? mapValue.Ordinal : -1;
                if (preferred != null && ExcelFile.Header(preferred) == "Key" && value >= 0)
                    s.Format = SaveTextFormat.TextMap;
            }
            ColumnBox.SelectedIndex = column;
            ValueColumnBox.SelectedIndex = value;
            PathBox.Text = s.Path;
            FormatBox.SelectedIndex = (int)s.Format;
        }

        public SaveTextFormat SelectedFormat
        {
            get { return (SaveTextFormat)Math.Max(0, FormatBox.SelectedIndex); }
        }

        public void SetOutputPath(string path)
        {
            PathBox.Text = path;
            Apply_Click(null, null);
        }

        void Format_Changed(object sender, SelectionChangedEventArgs e)
        {
            if (ValueColumnBox != null)
                ValueColumnBox.IsEnabled = SaveTextWork.IsMap(SelectedFormat);
        }

        void Apply_Click(object sender, RoutedEventArgs e)
        {
            var settings = new SaveTextSettings
            {
                Path = PathBox.Text.Trim(),
                Format = SelectedFormat,
                Column = ExcelColumns.Letter(ColumnBox.SelectedIndex),
                ValueColumn = ExcelColumns.Letter(ValueColumnBox.SelectedIndex)
            };
            try
            {
                SaveTextWork.Run(_input, settings);
            }
            catch (Exception ex)
            {
                Notices.Show(this, ex.Message, "Save text / JSON", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            Status.Text = "Ready to save " + _input.Rows.Count + " rows. Preview does not write files.";
            _use(settings);
        }
    }
}
