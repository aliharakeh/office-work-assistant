using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using OfficeWorkAssistant.Excel;

namespace OfficeWorkAssistant.Features.SplitCombine
{
    public partial class SplitColumnPage : UserControl
    {
        readonly DataTable _input;
        readonly Action<SplitColumnSettings> _use;
        // What the value box held for each mode, so switching back and forth keeps it.
        readonly string[] _values = new string[3];
        int _mode = -1;

        public SplitColumnPage(DataTable input, string label, SplitColumnSettings settings, Action<SplitColumnSettings> use)
        {
            InitializeComponent();
            _input = input;
            _use = use;
            InputInfo.Text = label;
            ColumnBox.ItemsSource = ExcelColumns.Labels(input);

            var s = settings ?? new SplitColumnSettings();
            int col = ExcelColumns.IndexOf(input, s.Column);
            ColumnBox.SelectedIndex = col >= 0 ? col : (input.Columns.Count > 0 ? 0 : -1);
            _values[0] = s.Delimiter;
            _values[1] = s.Widths;
            _values[2] = s.Pattern;
            ModeBox.SelectedIndex = (int)s.Mode;
            MaxBox.Text = s.MaxParts.ToString(CultureInfo.InvariantCulture);
            HeadersBox.Text = string.Join(", ", s.Headers.ToArray());
            TrimBox.IsChecked = s.TrimParts;
            KeepBox.IsChecked = s.KeepSource;
        }

        void ModeBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (ValueBox == null)
                return;
            if (_mode >= 0)
                _values[_mode] = ValueBox.Text;
            _mode = Math.Max(0, ModeBox.SelectedIndex);
            ValueBox.Text = _values[_mode] ?? "";
            if (_mode == 0)
            {
                ValueLabel.Text = "Delimiter";
                ValueHint.Text = "The text between the parts, e.g. a comma or \" - \". Write \\t for a tab.";
            }
            else if (_mode == 1)
            {
                ValueLabel.Text = "Widths";
                ValueHint.Text = "Characters per part, e.g. 3,5,2. Anything after them is one more part.";
            }
            else
            {
                ValueLabel.Text = "Pattern";
                ValueHint.Text = "A regular expression. Without ( ) the text is cut wherever it matches, e.g. [,;/]. " +
                    "With ( ) groups each group is a part, e.g. (\\w+)@(.+) splits an email.";
            }
        }

        void Apply_Click(object sender, RoutedEventArgs e)
        {
            if (ColumnBox.SelectedIndex < 0)
            {
                Status.Text = "Choose the column to split.";
                return;
            }
            int max;
            if (!int.TryParse(MaxBox.Text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out max) || max < 0)
            {
                Status.Text = "Max parts must be 0 or a whole number above 0.";
                MaxBox.Focus();
                return;
            }
            _values[_mode] = ValueBox.Text;
            var s = new SplitColumnSettings
            {
                Column = ExcelColumns.Letter(ColumnBox.SelectedIndex),
                Mode = (SplitMode)_mode,
                Delimiter = _values[0] ?? "",
                Widths = _values[1] ?? "",
                Pattern = _values[2] ?? "",
                MaxParts = max,
                TrimParts = TrimBox.IsChecked == true,
                KeepSource = KeepBox.IsChecked == true
            };
            foreach (var h in HeadersBox.Text.Split(','))
                s.Headers.Add(h.Trim());
            while (s.Headers.Count > 0 && s.Headers[s.Headers.Count - 1].Length == 0)
                s.Headers.RemoveAt(s.Headers.Count - 1);

            DataTable result;
            try
            {
                result = SplitCombineWork.Split(_input, s);
            }
            catch (Exception ex)
            {
                MessageBox.Show(Window.GetWindow(this), ex.Message, "Split column", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            int parts = result.Columns.Count - _input.Columns.Count + (s.KeepSource ? 0 : 1);
            Status.Text = "Split into " + parts + " part" + (parts == 1 ? "." : "s.");
            _use(s);
        }
    }
}
