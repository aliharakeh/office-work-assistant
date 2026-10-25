using System;
using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using OfficeWorkAssistant.Excel;

namespace OfficeWorkAssistant.Features.Ocr
{
    public partial class OcrPage : UserControl
    {
        readonly DataTable _input;
        readonly string _dataFolder;
        readonly Action<OcrSettings> _use;

        public OcrPage(DataTable input, string label, OcrSettings settings, Action<OcrSettings> use)
        {
            InitializeComponent();
            _input = input;
            _use = use;
            InputInfo.Text = label;
            PathColumnBox.ItemsSource = ExcelColumns.Labels(input);
            var s = settings ?? new OcrSettings();
            int column = ExcelColumns.IndexOf(input, s.PathColumn);
            if (settings == null)
            {
                var path = ExcelFile.FindColumn(input, "Path");
                column = path != null ? path.Ordinal : (input.Columns.Count > 0 ? 0 : -1);
            }
            PathColumnBox.SelectedIndex = column;
            TextHeaderBox.Text = s.TextHeader;
            _dataFolder = s.DataFolder;
            var options = OcrWork.LanguageOptions(s.DataFolder);
            string current = (s.Languages ?? "").Trim();
            if (current.Length > 0 && !options.Exists(o => o.Value == current))
                options.Add(new KeyValuePair<string, string>(current + " (not installed)", current));
            LanguagesBox.ItemsSource = options;
            LanguagesBox.SelectedValue = current.Length > 0 ? current : "eng";
            if (LanguagesBox.SelectedIndex < 0 && options.Count > 0)
                LanguagesBox.SelectedIndex = 0;
        }

        void Apply_Click(object sender, RoutedEventArgs e)
        {
            var settings = new OcrSettings
            {
                PathColumn = ExcelColumns.Letter(PathColumnBox.SelectedIndex),
                TextHeader = TextHeaderBox.Text.Trim(),
                Languages = Convert.ToString(LanguagesBox.SelectedValue) ?? "",
                DataFolder = _dataFolder ?? ""
            };
            try
            {
                // Auto-apply must not OCR every image just to validate the editor.
                OcrWork.Validate(_input, settings);
            }
            catch (Exception ex)
            {
                Notices.Show(this, ex.Message, "OCR", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            Status.Text = "Recognize " + settings.Languages + " into " + settings.TextHeader + ".";
            _use(settings);
        }
    }
}
