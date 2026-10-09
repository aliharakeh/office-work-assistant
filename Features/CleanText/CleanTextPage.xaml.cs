using System;
using System.Collections.Generic;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using OfficeWorkAssistant.Excel;

namespace OfficeWorkAssistant.Features.CleanText
{
    public partial class CleanTextPage : UserControl
    {
        readonly DataTable _input;
        readonly Action<CleanTextSettings> _use;
        readonly List<Pick> _picks = new List<Pick>();

        public CleanTextPage(DataTable input, string label, CleanTextSettings settings, Action<CleanTextSettings> use)
        {
            InitializeComponent();
            _input = input;
            _use = use;
            InputInfo.Text = label + " - only text cells change; numbers and dates stay as they are.";

            var s = settings ?? new CleanTextSettings();
            var on = new HashSet<int>();
            foreach (var c in s.Columns)
            {
                int i = ExcelColumns.IndexOf(input, c);
                if (i >= 0)
                    on.Add(i);
            }
            for (var i = 0; i < input.Columns.Count; i++)
                _picks.Add(new Pick { Letter = ExcelColumns.Letter(i), Label = ExcelColumns.Label(input, i), Include = on.Contains(i) });
            ColumnList.ItemsSource = _picks;
            TrimBox.IsChecked = s.Trim;
            CollapseBox.IsChecked = s.CollapseSpaces;
            BreaksBox.IsChecked = s.RemoveLineBreaks;
            HiddenBox.IsChecked = s.RemoveNonPrinting;
            NumberBox.IsChecked = s.TextToNumber;
            BlankBox.IsChecked = s.EmptyToBlank;
            CaseBox.SelectedIndex = (int)s.Case;
        }

        void Apply_Click(object sender, RoutedEventArgs e)
        {
            var s = new CleanTextSettings
            {
                Trim = TrimBox.IsChecked == true,
                CollapseSpaces = CollapseBox.IsChecked == true,
                RemoveLineBreaks = BreaksBox.IsChecked == true,
                RemoveNonPrinting = HiddenBox.IsChecked == true,
                TextToNumber = NumberBox.IsChecked == true,
                EmptyToBlank = BlankBox.IsChecked == true,
                Case = (TextCase)Math.Max(0, CaseBox.SelectedIndex)
            };
            foreach (var p in _picks)
            {
                if (p.Include)
                    s.Columns.Add(p.Letter);
            }
            try
            {
                CleanTextWork.Run(_input, s);
            }
            catch (Exception ex)
            {
                MessageBox.Show(Window.GetWindow(this), ex.Message, "Clean text", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            Status.Text = s.Describe();
            _use(s);
        }

        public sealed class Pick
        {
            public string Letter { get; set; }
            public string Label { get; set; }
            public bool Include { get; set; }
        }
    }
}
